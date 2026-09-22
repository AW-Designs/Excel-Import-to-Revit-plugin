using System;
using System.Globalization;
using System.IO;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;

namespace ExcelScheduleImporter.Revit
{
    /// <summary>
    /// The persistent link between a drafting view and its Excel source.
    /// Stored on the view itself (extensible storage), so it survives
    /// save/close/synchronize and travels with the model.
    /// </summary>
    public class ScheduleLink
    {
        public string FilePath = "";
        public string SheetName = "";
        public string RangeAddress = "";
        /// <summary>True when the import used the auto-detected range; updates re-detect
        /// so rows/columns added in Excel are picked up automatically.</summary>
        public bool AutoRange = true;
        public DateTime ImportedUtc;
        public DateTime FileWriteUtc;
        public int Scale = 1;
        public bool IncludeFills = true;
        public bool DrawAllGridlines;
        public double TextFactor = 1.0;
        public bool NormalizeBodyText;
        public double BodyTextMm = 2.0;
        public string FontOverride = "";

        /// <summary>
        /// UniqueId of the view this link was written to. Revit copies extensible
        /// storage when a view is duplicated, but gives the duplicate a NEW
        /// UniqueId - so a mismatch means "this is a copy, not the import".
        /// Empty on links written before v1.0.8.
        /// </summary>
        public string ViewUniqueId = "";

        /// <summary>Rebuild the ImportOptions this view was originally created with.</summary>
        public ImportOptions ToOptions() => new ImportOptions
        {
            FilePath          = FilePath,
            SheetName         = SheetName,
            RangeOverride     = AutoRange ? null : RangeAddress,   // null = re-detect
            AutoRange         = AutoRange,
            Scale             = Scale,
            IncludeFills      = IncludeFills,
            DrawAllGridlines  = DrawAllGridlines,
            TextFactor        = TextFactor,
            NormalizeBodyText = NormalizeBodyText,
            BodyTextMm        = BodyTextMm,
            FontOverride      = string.IsNullOrEmpty(FontOverride) ? null : FontOverride,
        };
    }

    /// <summary>Reads/writes the ScheduleLink on drafting views.</summary>
    public static class ScheduleMetadata
    {
        // Never change this GUID - it identifies the storage schema in every model.
        private static readonly Guid SchemaGuid = new Guid("a5c3f7d2-8e14-4b6a-9c05-2f61d43e9b77");

        private static Schema GetSchema()
        {
            var schema = Schema.Lookup(SchemaGuid);
            if (schema != null) return schema;

            var builder = new SchemaBuilder(SchemaGuid);
            builder.SetSchemaName("ExcelScheduleImporterLink");
            builder.SetReadAccessLevel(AccessLevel.Public);
            builder.SetWriteAccessLevel(AccessLevel.Public);
            builder.SetDocumentation("Source link for schedules imported from Excel by Excel Schedule Importer.");
            // One key=value blob keeps the schema forward-compatible:
            // new fields can be added without a schema migration.
            builder.AddSimpleField("Data", typeof(string));
            return builder.Finish();
        }

        /// <summary>Stamp the view with its source link. Must be called inside a transaction.</summary>
        public static void Save(View view, ImportOptions o)
        {
            var ci = CultureInfo.InvariantCulture;

            DateTime writeUtc = DateTime.UtcNow;
            try { if (File.Exists(o.FilePath)) writeUtc = File.GetLastWriteTimeUtc(o.FilePath); }
            catch { /* keep now() */ }

            string data = string.Join("\n", new[]
            {
                "Version=2",
                "ViewUniqueId=" + view.UniqueId,
                "FilePath=" + (o.FilePath ?? ""),
                "SheetName=" + (o.SheetName ?? ""),
                "RangeAddress=" + (o.RangeOverride ?? ""),
                "AutoRange=" + (o.AutoRange ? "1" : "0"),
                "ImportedUtc=" + DateTime.UtcNow.ToString("o", ci),
                "FileWriteUtc=" + writeUtc.ToString("o", ci),
                "Scale=" + o.Scale.ToString(ci),
                "IncludeFills=" + (o.IncludeFills ? "1" : "0"),
                "DrawAllGridlines=" + (o.DrawAllGridlines ? "1" : "0"),
                "TextFactor=" + o.TextFactor.ToString("R", ci),
                "NormalizeBodyText=" + (o.NormalizeBodyText ? "1" : "0"),
                "BodyTextMm=" + o.BodyTextMm.ToString("R", ci),
                "FontOverride=" + (o.FontOverride ?? ""),
            });

            var entity = new Entity(GetSchema());
            entity.Set("Data", data);
            view.SetEntity(entity);
        }

        /// <summary>Read the link from a view, or null when the view was not imported by this add-in.</summary>
        public static ScheduleLink Load(View view)
        {
            try
            {
                var entity = view.GetEntity(GetSchema());
                if (entity == null || !entity.IsValid()) return null;

                string data = entity.Get<string>("Data");
                if (string.IsNullOrEmpty(data)) return null;

                var ci = CultureInfo.InvariantCulture;
                var link = new ScheduleLink();
                foreach (string line in data.Split('\n'))
                {
                    int eq = line.IndexOf('=');
                    if (eq <= 0) continue;
                    string k = line.Substring(0, eq);
                    string v = line.Substring(eq + 1);
                    switch (k)
                    {
                        case "FilePath":          link.FilePath = v; break;
                        case "SheetName":         link.SheetName = v; break;
                        case "RangeAddress":      link.RangeAddress = v; break;
                        case "AutoRange":         link.AutoRange = v == "1"; break;
                        case "ImportedUtc":       DateTime.TryParse(v, ci, DateTimeStyles.RoundtripKind, out link.ImportedUtc); break;
                        case "FileWriteUtc":      DateTime.TryParse(v, ci, DateTimeStyles.RoundtripKind, out link.FileWriteUtc); break;
                        case "Scale":             int.TryParse(v, NumberStyles.Integer, ci, out link.Scale); break;
                        case "IncludeFills":      link.IncludeFills = v == "1"; break;
                        case "DrawAllGridlines":  link.DrawAllGridlines = v == "1"; break;
                        case "TextFactor":        double.TryParse(v, NumberStyles.Float, ci, out link.TextFactor); break;
                        case "NormalizeBodyText": link.NormalizeBodyText = v == "1"; break;
                        case "BodyTextMm":        double.TryParse(v, NumberStyles.Float, ci, out link.BodyTextMm); break;
                        case "FontOverride":      link.FontOverride = v; break;
                        case "ViewUniqueId":      link.ViewUniqueId = v; break;
                    }
                }
                return string.IsNullOrEmpty(link.FilePath) ? null : link;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// True when the link was stamped on a DIFFERENT view, i.e. this view is a
        /// Revit duplicate of an imported schedule. Only decidable for links
        /// written by v1.0.8+; older links return false (see ScheduleUpdater).
        /// </summary>
        public static bool IsCopy(View view, ScheduleLink link)
            => !string.IsNullOrEmpty(link.ViewUniqueId)
               && !string.Equals(link.ViewUniqueId, view.UniqueId, StringComparison.Ordinal);

        /// <summary>Remove the Excel link from a view. Must be called inside a transaction.</summary>
        public static void Remove(View view)
        {
            var schema = Schema.Lookup(SchemaGuid);
            if (schema != null) view.DeleteEntity(schema);
        }

        /// <summary>Is the source Excel file newer than what this view was built from?</summary>
        public static bool IsOutOfDate(ScheduleLink link)
        {
            try
            {
                if (!File.Exists(link.FilePath)) return false;   // "missing" is its own state
                return File.GetLastWriteTimeUtc(link.FilePath) > link.FileWriteUtc.AddSeconds(2);
            }
            catch { return false; }
        }
    }
}
