using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using DrawingColor = System.Drawing.Color;

namespace ExcelScheduleImporter.Excel
{
    /// <summary>
    /// Reads .xlsx / .xlsm files with ClosedXML (no Excel installation required)
    /// and converts a worksheet region into a <see cref="TableModel"/>.
    /// All sizes are converted to printed millimetres.
    ///
    /// The workbook is parsed ONCE in the constructor and kept in memory —
    /// listing sheets, detecting ranges and importing all reuse the same parse
    /// (important for large files on network shares).
    /// </summary>
    public sealed class ExcelReader : IDisposable
    {
        // Excel column width unit = width of one '0' character of the default font
        // (Calibri 11 => 7 px at 96 dpi) plus 5 px of cell padding.
        private const double PixelsPerWidthUnit = 7.0;
        private const double CellPaddingPx = 5.0;
        private const double MmPerPixel = 25.4 / 96.0;   // 96 dpi
        private const double MmPerPoint = 25.4 / 72.0;   // row heights / font sizes are in points

        private readonly XLWorkbook _wb;

        public ExcelReader(string path)
        {
            _wb = OpenWorkbook(path);
        }

        public void Dispose() => _wb.Dispose();

        // Win32 error codes surfaced through IOException.HResult
        private const int SharingViolation = unchecked((int)0x80070020);   // ERROR_SHARING_VIOLATION
        private const int LockViolation    = unchecked((int)0x80070021);   // ERROR_LOCK_VIOLATION

        /// <summary>
        /// Opens the workbook even when the file is open in Excel or sits on a
        /// network share.
        ///
        /// Reading and parsing are deliberately separate steps so a failure says
        /// which one went wrong - previously EVERY I/O error was reported as
        /// "locked", which could be wrong and hid the real cause.
        ///
        /// 1. Copy the file's bytes into memory through a shared read stream
        ///    (works while Excel has the file open). Sharing/lock violations are
        ///    usually momentary - Excel mid-save, or a network lease being
        ///    released - so those are retried with a short backoff.
        /// 2. Parse from memory. We never hold a handle on the source afterwards.
        /// </summary>
        private static XLWorkbook OpenWorkbook(string path)
        {
            if (!File.Exists(path))
                throw new FileNotFoundException("Excel file not found:\n" + path, path);

            var bytes = ReadAllBytesShared(path);

            try
            {
                return new XLWorkbook(new MemoryStream(bytes, writable: false));
            }
            catch (Exception ex)
            {
                // The file was read fine - its CONTENT could not be parsed.
                throw new InvalidDataException(
                    "The Excel file was read but could not be opened as a workbook.\n\n" +
                    "It may be damaged, password-protected, or saved in an unsupported " +
                    "format (only .xlsx / .xlsm are supported - not legacy .xls).\n\n" +
                    "File: " + path + "\nDetails: " + ex.Message, ex);
            }
        }

        /// <summary>
        /// Read the whole file with FileShare.ReadWrite|Delete, retrying briefly
        /// on sharing/lock violations (~5 s total) before giving up.
        /// </summary>
        private static byte[] ReadAllBytesShared(string path)
        {
            int[] backoffMs = { 250, 500, 1000, 1500, 2000 };
            for (int attempt = 0; ; attempt++)
            {
                try
                {
                    using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read,
                                                   FileShare.ReadWrite | FileShare.Delete))
                    using (var ms = new MemoryStream())
                    {
                        fs.CopyTo(ms);
                        return ms.ToArray();
                    }
                }
                catch (IOException ex) when (IsLockError(ex) && attempt < backoffMs.Length)
                {
                    System.Threading.Thread.Sleep(backoffMs[attempt]);
                }
                catch (IOException ex) when (IsLockError(ex))
                {
                    string who = TryGetExcelLockOwner(path);
                    throw new IOException(
                        "Could not read the Excel file - it is locked" +
                        (who != null ? " by " + who : " by another program or user") + ".\n\n" +
                        "Ask whoever has it open to close it (or wait for Excel to finish " +
                        "saving) and try again.\n\n" +
                        "File: " + path + "\nDetails: " + ex.Message, ex);
                }
                catch (IOException ex)
                {
                    // Not a lock - e.g. network drive dropped, path too long.
                    throw new IOException(
                        "Could not read the Excel file.\n\n" +
                        "File: " + path + "\nDetails: " + ex.Message, ex);
                }
            }
        }

        private static bool IsLockError(IOException ex)
            => ex.HResult == SharingViolation || ex.HResult == LockViolation;

        /// <summary>
        /// Excel writes an owner file "~$&lt;name&gt;" beside a workbook it has open;
        /// its first byte is the length of the user name that follows. Best effort -
        /// returns null if there is no owner file or it cannot be read.
        /// </summary>
        private static string TryGetExcelLockOwner(string path)
        {
            try
            {
                string dir = Path.GetDirectoryName(path);
                string name = Path.GetFileName(path);
                string ownerFile = Path.Combine(dir, "~$" + name);
                if (!File.Exists(ownerFile)) return null;

                byte[] data;
                using (var fs = new FileStream(ownerFile, FileMode.Open, FileAccess.Read,
                                               FileShare.ReadWrite | FileShare.Delete))
                {
                    data = new byte[Math.Min(fs.Length, 165)];
                    fs.Read(data, 0, data.Length);
                }
                int len = data.Length > 0 ? data[0] : 0;
                if (len <= 0 || len >= data.Length) return null;
                string user = System.Text.Encoding.Default.GetString(data, 1, len).Trim();
                return user.Length > 0 ? user : null;
            }
            catch { return null; }
        }

        /// <summary>List worksheet names (visible sheets first).</summary>
        public List<string> GetSheetNames()
        {
            return _wb.Worksheets
                      .OrderBy(ws => ws.Visibility != XLWorksheetVisibility.Visible)
                      .ThenBy(ws => ws.Position)
                      .Select(ws => ws.Name)
                      .ToList();
        }

        /// <summary>
        /// Auto-detect the used portion of a sheet (content + formatting).
        /// Returns an A1-style address like "A1:F42", or null when the sheet is empty.
        /// </summary>
        public string DetectUsedRange(string sheetName)
        {
            var ws = _wb.Worksheet(sheetName);
            var used = ws.RangeUsed(XLCellsUsedOptions.AllContents | XLCellsUsedOptions.NormalFormats);
            return used?.RangeAddress.ToString();
        }

        /// <summary>
        /// Read a sheet region into a TableModel.
        /// <paramref name="rangeAddress"/> may be null/empty to auto-detect.
        /// </summary>
        public TableModel ReadSheet(string sheetName, string rangeAddress)
        {
            var ws = _wb.Worksheet(sheetName);

            IXLRange range = string.IsNullOrWhiteSpace(rangeAddress)
                ? ws.RangeUsed(XLCellsUsedOptions.AllContents | XLCellsUsedOptions.NormalFormats)
                : ws.Range(rangeAddress.Trim());

            if (range == null)
                throw new InvalidOperationException(
                    $"Sheet '{sheetName}' appears to be empty - nothing to import.");

            int r0 = range.RangeAddress.FirstAddress.RowNumber;
            int c0 = range.RangeAddress.FirstAddress.ColumnNumber;
            int r1 = range.RangeAddress.LastAddress.RowNumber;
            int c1 = range.RangeAddress.LastAddress.ColumnNumber;
            int rows = r1 - r0 + 1;
            int cols = c1 - c0 + 1;

            var model = new TableModel
            {
                SourceSheet   = sheetName,
                SourceRange   = range.RangeAddress.ToString(),
                RowHeightsMm  = new double[rows],
                ColWidthsMm   = new double[cols],
                HEdges        = new int[rows + 1, cols],
                VEdges        = new int[rows, cols + 1],
            };

            // ── Row heights / column widths (hidden rows & columns collapse to 0) ──
            for (int r = 0; r < rows; r++)
            {
                var xlRow = ws.Row(r0 + r);
                model.RowHeightsMm[r] = xlRow.IsHidden ? 0.0 : xlRow.Height * MmPerPoint;
            }
            for (int c = 0; c < cols; c++)
            {
                var xlCol = ws.Column(c0 + c);
                model.ColWidthsMm[c] = xlCol.IsHidden
                    ? 0.0
                    : (xlCol.Width * PixelsPerWidthUnit + CellPaddingPx) * MmPerPixel;
            }

            // ── Merged-region map: mergeId[r,c] = index of merge, or -1 ──
            var mergeId = new int[rows, cols];
            for (int r = 0; r < rows; r++)
                for (int c = 0; c < cols; c++)
                    mergeId[r, c] = -1;

            var merges = ws.MergedRanges.ToList();
            for (int m = 0; m < merges.Count; m++)
            {
                var a = merges[m].RangeAddress;
                for (int r = Math.Max(a.FirstAddress.RowNumber, r0); r <= Math.Min(a.LastAddress.RowNumber, r1); r++)
                    for (int c = Math.Max(a.FirstAddress.ColumnNumber, c0); c <= Math.Min(a.LastAddress.ColumnNumber, c1); c++)
                        mergeId[r - r0, c - c0] = m;
            }

            // ── Cells ──
            var visited = new bool[rows, cols];
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    if (visited[r, c]) continue;

                    IXLCell cell = ws.Cell(r0 + r, c0 + c);
                    int rowSpan = 1, colSpan = 1;

                    int mid = mergeId[r, c];
                    if (mid >= 0)
                    {
                        var a = merges[mid].RangeAddress;

                        // Only the top-left IN-RANGE cell of a merge produces output.
                        int firstInRangeRow = Math.Max(a.FirstAddress.RowNumber, r0);
                        int firstInRangeCol = Math.Max(a.FirstAddress.ColumnNumber, c0);
                        if (firstInRangeRow != r0 + r || firstInRangeCol != c0 + c)
                        { visited[r, c] = true; continue; }

                        // Content & style always come from the TRUE master cell,
                        // even when it lies outside the selected range.
                        cell = ws.Cell(a.FirstAddress.RowNumber, a.FirstAddress.ColumnNumber);

                        rowSpan = Math.Min(a.LastAddress.RowNumber, r1) - (r0 + r) + 1;
                        colSpan = Math.Min(a.LastAddress.ColumnNumber, c1) - (c0 + c) + 1;
                        for (int rr = r; rr < r + rowSpan; rr++)
                            for (int cc = c; cc < c + colSpan; cc++)
                                visited[rr, cc] = true;
                    }
                    else
                    {
                        visited[r, c] = true;
                    }

                    var data = ReadCell(cell, r, c, rowSpan, colSpan);
                    if (data != null) model.Cells.Add(data);
                }
            }

            // ── Border grid (shared edges take the heavier of the two adjacent borders) ──
            BuildEdges(ws, model, mergeId, r0, c0, rows, cols);

            return model;
        }

        // ─────────────────────────────────────────────────────────────────────

        private CellData ReadCell(IXLCell cell, int r, int c, int rowSpan, int colSpan)
        {
            var style = cell.Style;
            string text;
            try { text = cell.GetFormattedString(); }   // respects number formats: "1,250.00", "50%", dates...
            catch { text = cell.GetString(); }

            DrawingColor? fill = null;
            if (style.Fill.PatternType == XLFillPatternValues.Solid)
            {
                var f = ResolveColor(style.Fill.BackgroundColor);
                // ignore pure white fills - they are indistinguishable from paper
                if (f.HasValue && !(f.Value.R > 250 && f.Value.G > 250 && f.Value.B > 250))
                    fill = f;
            }

            bool hasText = !string.IsNullOrWhiteSpace(text);
            if (!hasText && fill == null && rowSpan == 1 && colSpan == 1)
                return null;   // empty unformatted cell - nothing to draw

            // Horizontal alignment ("General" = numbers right, text left, like Excel)
            int hAlign;
            switch (style.Alignment.Horizontal)
            {
                case XLAlignmentHorizontalValues.Center:
                case XLAlignmentHorizontalValues.CenterContinuous:
                    hAlign = 0; break;
                case XLAlignmentHorizontalValues.Right:
                    hAlign = 1; break;
                case XLAlignmentHorizontalValues.Left:
                    hAlign = -1; break;
                default: // General
                    hAlign = (cell.DataType == XLDataType.Number || cell.DataType == XLDataType.DateTime) ? 1 : -1;
                    break;
            }

            int vAlign;
            switch (style.Alignment.Vertical)
            {
                case XLAlignmentVerticalValues.Top:    vAlign = -1; break;
                case XLAlignmentVerticalValues.Center: vAlign = 0;  break;
                default:                               vAlign = 1;  break; // Excel default = bottom
            }

            return new CellData
            {
                Row        = r,
                Col        = c,
                RowSpan    = rowSpan,
                ColSpan    = colSpan,
                Text       = hasText ? text : null,
                FontName   = style.Font.FontName,
                FontSizePt = style.Font.FontSize,
                Bold       = style.Font.Bold,
                Italic     = style.Font.Italic,
                TextColor  = ResolveColor(style.Font.FontColor) ?? DrawingColor.Black,
                HAlign     = hAlign,
                VAlign     = vAlign,
                Fill       = fill,
                WrapText   = style.Alignment.WrapText,
            };
        }

        private static void BuildEdges(IXLWorksheet ws, TableModel model, int[,] mergeId,
                                       int r0, int c0, int rows, int cols)
        {
            // Horizontal edges: [r, c] = edge above row r  (r == rows -> bottom edge)
            for (int r = 0; r <= rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    // Skip edges inside a merged region
                    if (r > 0 && r < rows && mergeId[r - 1, c] >= 0 && mergeId[r - 1, c] == mergeId[r, c])
                        continue;

                    int w = 0;
                    if (r > 0)    w = Math.Max(w, BorderWeight(ws.Cell(r0 + r - 1, c0 + c).Style.Border.BottomBorder));
                    if (r < rows) w = Math.Max(w, BorderWeight(ws.Cell(r0 + r,     c0 + c).Style.Border.TopBorder));
                    model.HEdges[r, c] = w;
                }
            }

            // Vertical edges: [r, c] = edge left of column c  (c == cols -> right edge)
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c <= cols; c++)
                {
                    if (c > 0 && c < cols && mergeId[r, c - 1] >= 0 && mergeId[r, c - 1] == mergeId[r, c])
                        continue;

                    int w = 0;
                    if (c > 0)    w = Math.Max(w, BorderWeight(ws.Cell(r0 + r, c0 + c - 1).Style.Border.RightBorder));
                    if (c < cols) w = Math.Max(w, BorderWeight(ws.Cell(r0 + r, c0 + c).Style.Border.LeftBorder));
                    model.VEdges[r, c] = w;
                }
            }
        }

        /// <summary>Map Excel border styles to 0 (none) / 1 (thin) / 2 (medium) / 3 (thick).</summary>
        private static int BorderWeight(XLBorderStyleValues style)
        {
            switch (style)
            {
                case XLBorderStyleValues.None:
                    return 0;
                case XLBorderStyleValues.Medium:
                case XLBorderStyleValues.MediumDashed:
                case XLBorderStyleValues.MediumDashDot:
                case XLBorderStyleValues.MediumDashDotDot:
                case XLBorderStyleValues.SlantDashDot:
                    return 2;
                case XLBorderStyleValues.Thick:
                case XLBorderStyleValues.Double:
                    return 3;
                default: // Hair, Thin, Dotted, Dashed, DashDot, DashDotDot
                    return 1;
            }
        }

        /// <summary>Resolve an XLColor (RGB / theme / indexed) to a concrete color.</summary>
        private DrawingColor? ResolveColor(XLColor color)
        {
            try
            {
                if (color == null) return null;

                if (color.ColorType == XLColorType.Theme)
                {
                    var baseColor = _wb.Theme.ResolveThemeColor(color.ThemeColor).Color;
                    return ApplyTint(baseColor, color.ThemeTint);
                }
                return color.Color;   // RGB and indexed colors resolve directly
            }
            catch
            {
                return null;
            }
        }

        /// <summary>Excel theme tint: positive = toward white, negative = toward black.</summary>
        private static DrawingColor ApplyTint(DrawingColor c, double tint)
        {
            if (Math.Abs(tint) < 1e-6) return c;
            Func<int, int> f;
            if (tint > 0) f = v => (int)Math.Round(v + (255 - v) * tint);
            else          f = v => (int)Math.Round(v * (1 + tint));
            return DrawingColor.FromArgb(
                Math.Max(0, Math.Min(255, f(c.R))),
                Math.Max(0, Math.Min(255, f(c.G))),
                Math.Max(0, Math.Min(255, f(c.B))));
        }
    }
}
