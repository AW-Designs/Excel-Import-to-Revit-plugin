using System;
using System.Linq;
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
                    ids => ScheduleUpdater.UpdateViews(doc, ids),
                    () => ScheduleUpdater.GatherRows(doc)))
                {
                    if (form.ShowDialog() != DialogResult.OK)
                        return Result.Cancelled;   // updates done in-dialog are already committed
                    options = form.Options;
                    table = form.Reader.ReadSheet(options.SheetName, options.RangeOverride);
                }

                if (table.Cells.Count == 0)
                {
                    TaskDialog.Show("Excel Schedule Importer",
                        "The selected range contains no visible content.");
                    return Result.Cancelled;
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
                        return Result.Cancelled;
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
                return Result.Cancelled;
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
                return Result.Failed;
            }
        }
    }
}
