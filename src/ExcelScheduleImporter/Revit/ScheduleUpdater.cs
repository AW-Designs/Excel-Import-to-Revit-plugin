using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;
using ExcelScheduleImporter.Excel;

namespace ExcelScheduleImporter.Revit
{
    /// <summary>One imported schedule, as displayed in the status panel.</summary>
    public class ManageRow
    {
        public ElementId ViewId;
        public string ViewName;
        public string SheetPlacement;   // "M6.01 - SCHEDULES" or "(not on a sheet)"
        public string FileName;
        public string FilePath;
        public string Worksheet;
        public string Range;
        public DateTime ImportedUtc;
        public bool OutOfDate;
        public bool FileMissing;

        public string Status => FileMissing ? "File not found"
                              : OutOfDate   ? "OUT OF DATE"
                              :               "Up to date";
    }

    /// <summary>
    /// Structured result for an in-dialog update. The external command needs to
    /// know whether model changes were committed so closing the manager cannot
    /// accidentally return Result.Cancelled and make Revit undo them.
    /// </summary>
    public sealed class ScheduleUpdateResult
    {
        public int UpdatedCount { get; set; }
        public bool TransactionCommitted { get; set; }
        public string Summary { get; set; }
    }

    /// <summary>
    /// Queries and updates Excel-linked drafting views. Shared by the import
    /// dialog's status panel (its Update buttons call straight into here).
    /// </summary>
    public static class ScheduleUpdater
    {
        /// <summary>All Excel-linked drafting views with placement + freshness info.</summary>
        public static List<ManageRow> GatherRows(Document doc)
        {
            // Which sheet is each view placed on? (viewports -> sheets)
            var sheetsByView = new Dictionary<ElementId, List<string>>();
            foreach (var vp in new FilteredElementCollector(doc)
                                   .OfClass(typeof(Viewport))
                                   .Cast<Viewport>())
            {
                if (!(doc.GetElement(vp.SheetId) is ViewSheet sheet)) continue;
                if (!sheetsByView.TryGetValue(vp.ViewId, out var list))
                    sheetsByView[vp.ViewId] = list = new List<string>();
                list.Add(sheet.SheetNumber + " - " + sheet.Name);
            }

            var all = AllLinked(doc);
            var copies = FindCopies(all);   // duplicated views: not real imports

            var rows = new List<ManageRow>();
            foreach (var item in all)
            {
                if (copies.ContainsKey(item.View.Id)) continue;
                var v = item.View;
                var link = item.Link;

                rows.Add(new ManageRow
                {
                    ViewId         = v.Id,
                    ViewName       = v.Name,
                    SheetPlacement = sheetsByView.TryGetValue(v.Id, out var s)
                                        ? string.Join(", ", s) : "(not on a sheet)",
                    FileName       = SafeFileName(link.FilePath),
                    FilePath       = link.FilePath,
                    Worksheet      = link.SheetName,
                    Range          = link.AutoRange ? link.RangeAddress + " (auto)" : link.RangeAddress,
                    ImportedUtc    = link.ImportedUtc,
                    FileMissing    = !File.Exists(link.FilePath),
                    OutOfDate      = ScheduleMetadata.IsOutOfDate(link),
                });
            }
            return rows;
        }

        private sealed class Linked
        {
            public View View;
            public ScheduleLink Link;
        }

        /// <summary>Every drafting view carrying an Excel link, in element-id order.</summary>
        private static List<Linked> AllLinked(Document doc)
        {
            var list = new List<Linked>();
            foreach (var v in new FilteredElementCollector(doc)
                                  .OfClass(typeof(ViewDrafting))
                                  .Cast<ViewDrafting>()
                                  .Where(v => !v.IsTemplate)
                                  .OrderBy(v => v.Id.Value))
            {
                var link = ScheduleMetadata.Load(v);
                if (link != null) list.Add(new Linked { View = v, Link = link });
            }
            return list;
        }

        /// <summary>
        /// Views that carry an Excel link only because someone DUPLICATED an
        /// imported schedule (Revit copies the hidden link along with the view).
        /// Maps each copy's id to its original's id (InvalidElementId if the
        /// original no longer exists).
        ///
        /// - Links written by v1.0.8+ record the view's UniqueId; a duplicate has a
        ///   new UniqueId, so a mismatch is conclusive.
        /// - Older links don't have that. For them, views sharing an IDENTICAL stamp
        ///   (same file + sheet + import time to the sub-second) must be copies of
        ///   one another; the lowest element id is the original, since Revit
        ///   assigns ids in creation order.
        /// </summary>
        private static Dictionary<ElementId, ElementId> FindCopies(List<Linked> all)
        {
            var copies = new Dictionary<ElementId, ElementId>();

            var byUniqueId = all.ToDictionary(x => x.View.UniqueId, x => x.View.Id);
            foreach (var x in all.Where(x => ScheduleMetadata.IsCopy(x.View, x.Link)))
                copies[x.View.Id] = byUniqueId.TryGetValue(x.Link.ViewUniqueId, out var orig)
                    ? orig : ElementId.InvalidElementId;

            var legacyGroups = all
                .Where(x => string.IsNullOrEmpty(x.Link.ViewUniqueId))
                .GroupBy(x => x.Link.FilePath + "|" + x.Link.SheetName + "|" + x.Link.ImportedUtc.Ticks)
                .Where(g => g.Count() > 1);
            foreach (var g in legacyGroups)
            {
                var ordered = g.OrderBy(x => x.View.Id.Value).ToList();
                foreach (var copy in ordered.Skip(1))
                    copies[copy.View.Id] = ordered[0].View.Id;
            }
            return copies;
        }

        /// <summary>Can we edit this element right now? (Workshared: not owned by someone else.)</summary>
        private static bool IsEditable(Document doc, ElementId id)
            => !doc.IsWorkshared
               || WorksharingUtils.GetCheckoutStatus(doc, id) != CheckoutStatus.OwnedByOtherUser;

        /// <summary>
        /// Remove the Excel link from one view, leaving its contents untouched.
        /// For views that are not (or no longer) really an import - e.g. a copy
        /// the automatic detection could not recognise.
        /// </summary>
        public static string Unlink(Document doc, ElementId id)
        {
            var view = doc.GetElement(id) as View;
            if (view == null) return "The view no longer exists.";
            if (!IsEditable(doc, id))
                return "\"" + view.Name + "\" is checked out by another user - it cannot be changed right now.";

            using (var tx = new Transaction(doc, "Remove Excel link"))
            {
                tx.Start();
                ScheduleMetadata.Remove(view);
                tx.Commit();
            }
            return null;   // success
        }

        /// <summary>
        /// Rebuild the given views in place from their Excel sources (one
        /// transaction). Views keep their identity, so viewports on sheets
        /// are untouched. Returns a one-line summary; failure details on
        /// subsequent lines.
        /// </summary>
        public static ScheduleUpdateResult UpdateViews(Document doc, ICollection<ElementId> ids,
                                                      Action<int, int, string> progress = null)
        {
            int updated = 0;
            var failures = new List<string>();

            // Several schedules usually come from ONE workbook (one per sheet).
            // Parse each file once for the whole update instead of once per view -
            // on a network drive that is the slowest part of reading.
            var readers = new Dictionary<string, ExcelReader>(StringComparer.OrdinalIgnoreCase);

            using (var tx = new Transaction(doc, "Update Excel Schedules"))
            {
              try
              {
                tx.Start();

                var copies = FindCopies(AllLinked(doc));

                int index = 0;
                foreach (var id in ids)
                {
                    index++;
                    var view = doc.GetElement(id) as View;
                    if (view == null) continue;
                    progress?.Invoke(index, ids.Count, view.Name);

                    try
                    {
                        var link = ScheduleMetadata.Load(view);
                        if (link == null) continue;

                        // Safety: an update wipes every line/text/fill in the view.
                        // On a duplicated view that would destroy someone's own work.
                        if (copies.ContainsKey(view.Id))
                        {
                            failures.Add(view.Name + ":  this is a copy of another imported view "
                                         + "- not updated (use Unlink on it)");
                            continue;
                        }

                        // Re-read the source (auto ranges are re-detected, so
                        // rows/columns added in Excel come along)
                        if (!readers.TryGetValue(link.FilePath, out var reader))
                            readers[link.FilePath] = reader = new ExcelReader(link.FilePath);
                        TableModel table = reader.ReadSheet(link.SheetName,
                            link.AutoRange ? null : link.RangeAddress);

                        // Wipe the view's own contents (lines, text, fills).
                        // The viewport that places it on a sheet is owned by
                        // the SHEET, not this view, so it is never touched.
                        DeleteViewContents(doc, view);

                        // Redraw and re-stamp with the fresh file timestamp
                        var options = link.ToOptions();
                        new DraftingViewBuilder(doc).BuildIntoView(view, table, options);

                        if (link.AutoRange)
                        {
                            string r = table.SourceRange ?? link.RangeAddress;
                            int bang = r.IndexOf('!');
                            options.RangeOverride = bang >= 0 ? r.Substring(bang + 1) : r;
                        }
                        ScheduleMetadata.Save(view, options);

                        // Re-stamping changes this view's import time, which would
                        // break the "identical stamp" detection for any OLD-format
                        // copies of it - they would reappear as imports. Strip the
                        // link from those copies now, while we know they are copies.
                        foreach (var kv in copies.Where(c => c.Value == view.Id).ToList())
                        {
                            if (doc.GetElement(kv.Key) is View copyView && IsEditable(doc, kv.Key))
                                ScheduleMetadata.Remove(copyView);
                        }

                        updated++;
                    }
                    catch (Exception ex)
                    {
                        failures.Add(view.Name + ":  " + ex.Message);
                    }
                }

                tx.Commit();
              }
              catch (Exception ex)
              {
                if (tx.HasStarted() && !tx.HasEnded()) tx.RollBack();
                return new ScheduleUpdateResult
                {
                    UpdatedCount = 0,
                    TransactionCommitted = false,
                    Summary = "Update failed and was rolled back: " + ex.Message,
                };
              }
              finally
              {
                foreach (var r in readers.Values) r.Dispose();
              }
            }

            var sb = new StringBuilder();
            sb.AppendFormat("Updated {0} schedule{1}.", updated, updated == 1 ? "" : "s");
            if (failures.Count > 0)
            {
                sb.AppendFormat("  {0} failed:", failures.Count);
                foreach (var f in failures) sb.Append("\n  -  ").Append(f);
            }
            return new ScheduleUpdateResult
            {
                UpdatedCount = updated,
                TransactionCommitted = true,
                Summary = sb.ToString(),
            };
        }

        /// <summary>
        /// Delete the importer-created contents of a drafting view in as FEW
        /// Revit delete calls as possible.
        ///
        /// Every doc.Delete call triggers a model-wide regeneration/auto-join pass
        /// (~0.2 s each on a real project, per the Revit journal). The previous
        /// one-element-at-a-time loop therefore cost ~0.2 s PER LINE/TEXT, which
        /// made a multi-view update run for many minutes and look frozen.
        ///
        /// Two batches instead:
        ///   1. text notes + filled regions (a filled region takes its own
        ///      boundary sketch lines with it)
        ///   2. whatever detail lines still exist, re-collected AFTER batch 1 so
        ///      no id deleted by the cascade is ever passed in (that double-delete
        ///      is what originally caused InvalidObjectException).
        /// If a batch still fails, fall back to the slow-but-safe per-element loop.
        /// </summary>
        private static void DeleteViewContents(Document doc, View view)
        {
            // Restrict to EXACTLY the element kinds the importer creates. Never a
            // View: deleting one that Revit treats as a view triggers "Deleting all
            // open views in a project is not allowed" when the view is active.
            DeleteBatch(doc, CollectOwned(doc, view, e => e is TextNote || e is FilledRegion));
            DeleteBatch(doc, CollectOwned(doc, view, e => e is CurveElement));
        }

        private static List<ElementId> CollectOwned(Document doc, View view, Func<Element, bool> kind)
            => new FilteredElementCollector(doc, view.Id)
                .WhereElementIsNotElementType()
                .Where(e => e.OwnerViewId == view.Id && !(e is View) && kind(e))
                .Select(e => e.Id)
                .ToList();

        private static void DeleteBatch(Document doc, List<ElementId> ids)
        {
            if (ids.Count == 0) return;
            try
            {
                doc.Delete(ids);
            }
            catch (Autodesk.Revit.Exceptions.ApplicationException)
            {
                // Rare: something in the batch is protected or already gone.
                foreach (var id in ids)
                {
                    if (doc.GetElement(id) == null) continue;
                    try { doc.Delete(id); }
                    catch (Autodesk.Revit.Exceptions.ApplicationException) { /* protected */ }
                }
            }
        }

        private static string SafeFileName(string path)
        {
            try { return Path.GetFileName(path); }
            catch { return path; }
        }
    }
}
