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

            var rows = new List<ManageRow>();
            foreach (var v in new FilteredElementCollector(doc)
                                  .OfClass(typeof(ViewDrafting))
                                  .Cast<ViewDrafting>()
                                  .Where(v => !v.IsTemplate))
            {
                var link = ScheduleMetadata.Load(v);
                if (link == null) continue;

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

        /// <summary>
        /// Rebuild the given views in place from their Excel sources (one
        /// transaction). Views keep their identity, so viewports on sheets
        /// are untouched. Returns a one-line summary; failure details on
        /// subsequent lines.
        /// </summary>
        public static ScheduleUpdateResult UpdateViews(Document doc, ICollection<ElementId> ids)
        {
            int updated = 0;
            var failures = new List<string>();

            using (var tx = new Transaction(doc, "Update Excel Schedules"))
            {
              try
              {
                tx.Start();

                foreach (var id in ids)
                {
                    var view = doc.GetElement(id) as View;
                    if (view == null) continue;

                    try
                    {
                        var link = ScheduleMetadata.Load(view);
                        if (link == null) continue;

                        // Re-read the source (auto ranges are re-detected, so
                        // rows/columns added in Excel come along)
                        TableModel table;
                        using (var reader = new ExcelReader(link.FilePath))
                        {
                            table = reader.ReadSheet(link.SheetName,
                                link.AutoRange ? null : link.RangeAddress);
                        }

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
        /// Delete every element a drafting view owns, safely. Deleting a
        /// FilledRegion cascade-deletes its boundary curves - which are ALSO in
        /// the owned set - so a single batch doc.Delete(list) can hit a
        /// already-deleted id and throw InvalidObjectException at commit time.
        /// Deleting one at a time with an existence check avoids that entirely.
        /// </summary>
        private static void DeleteViewContents(Document doc, View view)
        {
            // Restrict to EXACTLY the element kinds the importer creates:
            // detail lines (CurveElement), text notes and filled regions.
            // Crucially this never includes a View - the collector-by-view can
            // surface view-owned elements, and attempting to delete one that
            // Revit treats as a view triggers "Deleting all open views in a
            // project is not allowed" when the view being updated is active.
            var ids = new FilteredElementCollector(doc, view.Id)
                .WhereElementIsNotElementType()
                .Where(e => e.OwnerViewId == view.Id
                            && !(e is View)
                            && (e is TextNote || e is FilledRegion || e is CurveElement))
                .Select(e => e.Id)
                .ToList();

            foreach (var id in ids)
            {
                if (doc.GetElement(id) == null) continue;   // already gone via a cascade
                try { doc.Delete(id); }
                catch (Autodesk.Revit.Exceptions.ApplicationException) { /* cascade / protected */ }
            }
        }

        private static string SafeFileName(string path)
        {
            try { return Path.GetFileName(path); }
            catch { return path; }
        }
    }
}
