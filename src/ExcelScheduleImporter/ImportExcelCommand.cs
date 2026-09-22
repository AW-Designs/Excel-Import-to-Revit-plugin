using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ExcelScheduleImporter.Excel;
using ExcelScheduleImporter.Revit;
using ExcelScheduleImporter.UI;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;
using View = Autodesk.Revit.DB.View;

namespace ExcelScheduleImporter
{
    /// <summary>
    /// Imports an Excel worksheet into a new drafting view as native
    /// detail lines, text notes and filled regions.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ImportExcelCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            if (uidoc == null)
            {
                message = "Open a project document first.";
                return Result.Failed;
            }
            Document doc = uidoc.Document;
            bool hasCommittedUpdates = false;

            try
            {
                // Existing drafting view names - the dialog warns on collisions.
                var existingNames = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewDrafting))
                    .Cast<ViewDrafting>()
                    .Where(v => !v.IsTemplate)
                    .Select(v => v.Name)
                    .ToList();

                // Status panel data: every Excel-linked view with freshness info.
                // The dialog's Update buttons call straight back into the updater
                // (valid API context - we are inside Execute on the API thread).
                var scheduleRows = ScheduleUpdater.GatherRows(doc);

                // 1. Collect user input and read the workbook (parsed once by the
                //    form when the file was picked; ReadSheet reuses that parse).
                ImportOptions options;
                TableModel table;
                using (var form = new ImportForm(
                    existingNames,
                    scheduleRows,
                    (ids, progress) => ScheduleUpdater.UpdateViews(doc, ids, progress),
                    () => ScheduleUpdater.GatherRows(doc)))
                {
                    DialogResult dialogResult = form.ShowDialog();
                    hasCommittedUpdates = form.HasCommittedUpdates;
                    if (dialogResult != DialogResult.OK)
                        return CancelUnlessChangesCommitted(hasCommittedUpdates);

                    // Multi-sheet import: everything happens here, while the form
                    // (and therefore its already-parsed workbook) is still alive.
                    if (form.BatchOptions != null && form.BatchOptions.Count > 0)
                    {
                        Result batchResult = ImportBatch(uidoc, doc, form, form.BatchOptions);
                        // A failed batch must not reverse schedule updates the user
                        // already committed in the dialog (Revit undoes the whole
                        // command on Failed/Cancelled).
                        return batchResult == Result.Failed && hasCommittedUpdates
                            ? Result.Succeeded : batchResult;
                    }

                    options = form.Options;
                    table = form.Reader.ReadSheet(options.SheetName, options.RangeOverride);
                }

                if (table.Cells.Count == 0)
                {
                    TaskDialog.Show("Excel Schedule Importer",
                        "The selected range contains no visible content.");
                    return CancelUnlessChangesCommitted(hasCommittedUpdates);
                }

                // Sanity warning for very large ranges
                long cellCount = (long)table.Rows * table.Cols;
                if (cellCount > 20000)
                {
                    var td = new TaskDialog("Excel Schedule Importer")
                    {
                        MainInstruction = string.Format(
                            "The selected range spans {0} rows x {1} columns ({2:N0} cells).",
                            table.Rows, table.Cols, cellCount),
                        MainContent = "Importing a very large range can take several minutes " +
                                      "and create many elements. Continue?",
                        CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No,
                        DefaultButton = TaskDialogResult.No,
                    };
                    if (td.Show() != TaskDialogResult.Yes)
                        return CancelUnlessChangesCommitted(hasCommittedUpdates);
                }

                // 3. If replacing, locate the old view now. It cannot simply be
                //    deleted up front: it is very likely the ACTIVE view (typical
                //    re-import flow) and Revit refuses to delete the active view.
                //    Sequence: rename old aside -> build new -> activate new -> delete old.
                ElementId oldViewId = null;
                if (options.ReplaceExisting)
                {
                    var oldView = new FilteredElementCollector(doc)
                        .OfClass(typeof(ViewDrafting))
                        .Cast<ViewDrafting>()
                        .FirstOrDefault(v => !v.IsTemplate && v.Name == options.ViewName);
                    if (oldView != null) oldViewId = oldView.Id;
                }

                // 4. Build the drafting view inside a single transaction
                BuildResult result;
                using (var tx = new Transaction(doc, "Import Excel Schedule"))
                {
                    tx.Start();

                    // Free up the view name for the new view. NOTE: the suffix must
                    // avoid Revit's prohibited name characters ('~' is one of them!).
                    if (oldViewId != null)
                    {
                        var oldView = (View)doc.GetElement(oldViewId);
                        oldView.Name = DraftingViewBuilder.SanitizeViewName(options.ViewName)
                                       + " (replacing " + Guid.NewGuid().ToString("N").Substring(0, 6) + ")";
                    }

                    var builder = new DraftingViewBuilder(doc);
                    result = builder.Build(table, options);

                    // Stamp the view with its Excel source so it can be updated later
                    ScheduleMetadata.Save(result.View, options);

                    tx.Commit();
                }

                // 5. Open the new view - that IS the confirmation, no popup needed.
                uidoc.ActiveView = result.View;

                // 6. Now the old view is no longer active anywhere - delete it.
                if (oldViewId != null)
                {
                    try
                    {
                        using (var tx = new Transaction(doc, "Delete replaced view"))
                        {
                            tx.Start();
                            doc.Delete(oldViewId);
                            tx.Commit();
                        }
                    }
                    catch (Exception ex)
                    {
                        // e.g. workshared model and the view is owned by someone else
                        TaskDialog.Show("Excel Schedule Importer",
                            "The new view was created, but the old view could not be deleted:\n\n"
                            + ex.Message
                            + "\n\nIt remains in the project with a '(replacing ...)' suffix - "
                            + "delete it manually when possible.");
                    }
                }

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return CancelUnlessChangesCommitted(hasCommittedUpdates);
            }
            catch (Exception ex)
            {
                // TEMP diagnostic: full stack trace (incl. inner exceptions) so we
                // can pin down exactly where a failure happens, since Revit's own
                // error dialog only shows ex.Message and hides the call chain.
                App.LogCrash("ImportExcelCommand.Execute", ex);
                message = ex.Message;
                var td = new TaskDialog("Excel Schedule Importer - Error")
                {
                    MainInstruction = ex.Message,
                    ExpandedContent = ex.ToString() + "\n\nAlso logged to:\n" + App.CrashLogPath,
                    CommonButtons = TaskDialogCommonButtons.Close,
                };
                td.Show();
                // A prior in-dialog schedule update is already committed. Returning
                // Failed here would make Revit reverse that valid update along with
                // the unsuccessful import attempt.
                return hasCommittedUpdates ? Result.Succeeded : Result.Failed;
            }
        }

        /// <summary>
        /// Import several worksheets in ONE transaction. Committing once rather
        /// than per sheet is what makes a batch meaningfully faster than repeating
        /// the dialog: Revit's per-transaction overhead is paid a single time.
        /// A failure on one sheet is recorded and the rest still import.
        /// </summary>
        private static Result ImportBatch(UIDocument uidoc, Document doc,
                                          ImportForm form, List<ImportOptions> plan)
        {
            var replacedIds = new List<ElementId>();
            var built = new List<View>();
            var failures = new List<string>();

            using (var tx = new Transaction(doc, "Import Excel Schedules"))
            {
                tx.Start();

                foreach (var options in plan)
                {
                    try
                    {
                        // Free the name first: the old view cannot be deleted yet
                        // (it may be active), so rename it aside and delete later.
                        if (options.ReplaceExisting)
                        {
                            var oldView = new FilteredElementCollector(doc)
                                .OfClass(typeof(ViewDrafting))
                                .Cast<ViewDrafting>()
                                .FirstOrDefault(v => !v.IsTemplate && v.Name == options.ViewName);

                            if (oldView != null)
                            {
                                oldView.Name = DraftingViewBuilder.SanitizeViewName(options.ViewName)
                                    + " (replacing " + Guid.NewGuid().ToString("N").Substring(0, 6) + ")";
                                replacedIds.Add(oldView.Id);
                            }
                        }

                        var table = form.Reader.ReadSheet(options.SheetName, options.RangeOverride);
                        if (table.Cells.Count == 0)
                        {
                            failures.Add(options.SheetName + ":  no visible content");
                            continue;
                        }

                        var result = new DraftingViewBuilder(doc).Build(table, options);
                        ScheduleMetadata.Save(result.View, options);
                        built.Add(result.View);
                    }
                    catch (Exception ex)
                    {
                        failures.Add(options.SheetName + ":  " + ex.Message);
                    }
                }

                if (built.Count == 0)
                {
                    tx.RollBack();
                    TaskDialog.Show("Excel Schedule Importer",
                        "Nothing was imported.\n\n" + string.Join("\n", failures));
                    return Result.Failed;
                }

                tx.Commit();
            }

            // Land on the first new view - immediate visual confirmation.
            uidoc.ActiveView = built[0];

            // The replaced views are no longer active anywhere; remove them now.
            if (replacedIds.Count > 0)
            {
                try
                {
                    using (var tx = new Transaction(doc, "Delete replaced views"))
                    {
                        tx.Start();
                        foreach (var id in replacedIds)
                            if (doc.GetElement(id) != null) doc.Delete(id);
                        tx.Commit();
                    }
                }
                catch (Exception ex)
                {
                    failures.Add("Old views could not all be deleted: " + ex.Message
                                 + "  (they keep a '(replacing ...)' suffix)");
                }
            }

            var sb = new StringBuilder();
            sb.AppendFormat("Imported {0} of {1} worksheet{2}.",
                built.Count, plan.Count, plan.Count == 1 ? "" : "s");
            if (failures.Count > 0)
            {
                sb.AppendLine().AppendLine();
                sb.AppendLine("Not imported:");
                foreach (var f in failures) sb.Append("  -  ").AppendLine(f);
            }

            // Only interrupt with a dialog when something needs attention; a clean
            // batch just leaves the user on the first new view.
            if (failures.Count > 0)
                TaskDialog.Show("Excel Schedule Importer", sb.ToString());

            return Result.Succeeded;
        }

        private static Result CancelUnlessChangesCommitted(bool hasCommittedUpdates)
            => hasCommittedUpdates ? Result.Succeeded : Result.Cancelled;
    }
}
