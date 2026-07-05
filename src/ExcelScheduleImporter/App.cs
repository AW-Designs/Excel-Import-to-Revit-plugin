using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;

namespace ExcelScheduleImporter
{
    /// <summary>Adds the ribbon button: Add-Ins tab -> "Excel Import" panel.</summary>
    public class App : IExternalApplication
    {
        // Revit year this build targets - baked in per build configuration so the
        // updater knows which release asset (R24/R25/R26) to download.
#if R24
        public const int RevitYear = 2024;
#elif R25
        public const int RevitYear = 2025;
#elif R26
        public const int RevitYear = 2026;
#else
        public const int RevitYear = 0;
#endif

        // Log file for exceptions that escape our own try/catch blocks (e.g. thrown
        // from a UI event handler after our command's Execute() has returned, or
        // from anywhere else Revit's own generic error dialog would otherwise hide
        // the real stack trace). TEMP diagnostic - safe to leave in permanently.
        internal static readonly string CrashLogPath =
            Path.Combine(Path.GetTempPath(), "ExcelScheduleImporter_crash.log");

        internal static void LogCrash(string source, Exception ex)
        {
            try
            {
                File.AppendAllText(CrashLogPath,
                    "\n===== " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  [" + source + "] =====\n"
                    + ex + "\n");
            }
            catch { /* logging must never itself throw */ }
        }

        public Result OnStartup(UIControlledApplication application)
        {
            // Safety net: log (don't just swallow) exceptions that occur outside any
            // command's own try/catch - e.g. in a WinForms event handler fired after
            // ShowDialog() has already returned control to Revit.
            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
                LogCrash("AppDomain.UnhandledException", e.ExceptionObject as Exception
                    ?? new Exception("Non-Exception object: " + e.ExceptionObject));
            Application.ThreadException += (s, e) =>
                LogCrash("WinForms.ThreadException", e.Exception);

            try
            {
                RibbonPanel panel = application.GetRibbonPanels()
                    .FirstOrDefault(p => p.Name == "Excel Import")
                    ?? application.CreateRibbonPanel("Excel Import");

                var buttonData = new PushButtonData(
                    "ExcelScheduleImport",
                    "Import Excel\nSchedule",
                    Assembly.GetExecutingAssembly().Location,
                    "ExcelScheduleImporter.ImportExcelCommand")
                {
                    ToolTip = "Recreate an Excel schedule as a drafting view",
                    LongDescription =
                        "Pick an Excel file and worksheet. The used range is auto-detected " +
                        "and reproduced as native detail lines, text notes and filled regions " +
                        "in a new drafting view - fully vector, no raster images.",
                };

                buttonData.LargeImage = LoadEmbeddedPng("icon32.png");   // 32x32 ribbon button
                buttonData.Image      = LoadEmbeddedPng("icon16.png");   // 16x16 stacked/QAT
                buttonData.ToolTipImage = buttonData.LargeImage;

                panel.AddItem(buttonData);

                // Background self-update check against GitHub Releases. Never
                // blocks startup; applies any newer version after Revit closes.
                Updater.CheckInBackground(RevitYear,
                    Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Excel Schedule Importer", "Failed to start: " + ex.Message);
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application) => Result.Succeeded;

        /// <summary>Load a PNG embedded as ExcelScheduleImporter.Resources.&lt;name&gt;.</summary>
        private static BitmapImage LoadEmbeddedPng(string name)
        {
            try
            {
                var asm = Assembly.GetExecutingAssembly();
                Stream stream = asm.GetManifestResourceStream("ExcelScheduleImporter.Resources." + name);
                if (stream == null) return null;

                var image = new BitmapImage();
                image.BeginInit();
                image.CacheOption = BitmapCacheOption.OnLoad;   // read fully, then release the stream
                image.StreamSource = stream;
                image.EndInit();
                image.Freeze();
                stream.Dispose();
                return image;
            }
            catch
            {
                return null;   // no icon is better than no add-in
            }
        }
    }
}
