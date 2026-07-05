using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using DrawingColor = System.Drawing.Color;
using ExcelScheduleImporter.Excel;

namespace ExcelScheduleImporter.Revit
{
    public class ImportOptions
    {
        public string FilePath { get; set; }
        public string SheetName { get; set; }
        public string RangeOverride { get; set; }
        public string ViewName { get; set; }
        public int Scale { get; set; } = 1;
        public bool IncludeFills { get; set; } = true;
        public bool DrawAllGridlines { get; set; }
        /// <summary>Multiplier on text size, 1.0 = match Excel printed size.</summary>
        public double TextFactor { get; set; } = 1.0;

        /// <summary>
        /// When true, the whole table (geometry + text) is scaled uniformly so the
        /// most common (body) text size lands exactly on <see cref="BodyTextMm"/>.
        /// OFF by default: 1:1 with Excel preserves the original layout best,
        /// because Excel column widths were tuned for Excel's own fonts/sizes.
        /// </summary>
        public bool NormalizeBodyText { get; set; } = false;

        /// <summary>Target paper size of body text in mm (sheet standard, e.g. 2.0).</summary>
        public double BodyTextMm { get; set; } = 2.0;

        /// <summary>Force all text to this font (e.g. "Arial"). Null = keep Excel fonts.</summary>
        public string FontOverride { get; set; }

        /// <summary>User confirmed that an existing view with the same name should be replaced.</summary>
        public bool ReplaceExisting { get; set; }

        /// <summary>True when the range came straight from auto-detection: updates
        /// then re-detect the used range, so growth in Excel is picked up.</summary>
        public bool AutoRange { get; set; }
    }

    public class BuildResult
    {
        public View View { get; set; }
        public int Lines { get; set; }
        public int Texts { get; set; }
        public int Fills { get; set; }
    }

    /// <summary>
    /// Recreates a TableModel inside a new Revit drafting view using
    /// detail lines (borders), text notes (content) and filled regions (cell shading).
    ///
    /// Geometry strategy: the table is drawn in MODEL space of the drafting view at
    /// (printed size x view scale), so when the view is placed on a sheet it appears
    /// exactly at the size Excel would have printed it, regardless of the chosen scale.
    /// </summary>
    public class DraftingViewBuilder
    {
        private const double MmToFt = 1.0 / 304.8;
        private const double PtToMm = 25.4 / 72.0;

        // Excel font size (em) vs Revit text size (roughly cap-height based):
        // an empirical 0.72 factor makes 11pt Calibri in Excel visually match Revit output.
        private const double CapHeightFactor = 0.72;

        // Base text inset from cell borders (printed mm). Excel's own padding is
        // ~0.5mm; staying close keeps text position and usable width faithful.
        private const double CellPadMm = 0.6;

        // Fit-pass tuning:
        // A single-line text that is slightly too wide is shrunk down to at most
        // this factor to stay on one line; anything needing more gets wrapped.
        private const double MinShrinkSingleLine = 0.70;
        // Absolute floor for text size after fitting (paper mm).
        private const double MinTextSizeMm = 1.0;

        private readonly Document _doc;
        private readonly Dictionary<string, TextNoteType> _textTypes = new Dictionary<string, TextNoteType>();
        private readonly Dictionary<int, FilledRegionType> _fillTypes = new Dictionary<int, FilledRegionType>();
        private GraphicsStyle[] _lineStyles;   // index 1..3 = thin / medium / wide

        public DraftingViewBuilder(Document doc) { _doc = doc; }

        public BuildResult Build(TableModel table, ImportOptions options)
        {
            var view = CreateDraftingView(options);
            return BuildIntoView(view, table, options);
        }

        /// <summary>
        /// Draw the table into an EXISTING drafting view. Used by the update
        /// workflow: the view keeps its identity, so any viewport placing it on
        /// a sheet stays exactly where it is.
        /// </summary>
        public BuildResult BuildIntoView(View view, TableModel table, ImportOptions options)
        {
            _fitItems.Clear();   // a builder may be reused; never re-fit deleted notes

            var result = new BuildResult();
            result.View = view;
            view.Scale = Math.Max(1, options.Scale);

            double s = Math.Max(1, options.Scale);

            // Uniform normalization factor: scale the ENTIRE schedule (geometry and
            // text together) so the body text prints at exactly options.BodyTextMm.
            double norm = 1.0;
            if (options.NormalizeBodyText)
            {
                double bodyPt = MostCommonFontSize(table);
                double bodyMm = bodyPt * PtToMm * CapHeightFactor * options.TextFactor;
                if (bodyMm > 1e-6)
                    norm = options.BodyTextMm / bodyMm;
            }

            // Cumulative positions in feet (model space of the drafting view).
            // Origin = top-left of the table; +X right, -Y down.
            var xs = new double[table.Cols + 1];
            for (int c = 0; c < table.Cols; c++)
                xs[c + 1] = xs[c] + table.ColWidthsMm[c] * norm * s * MmToFt;

            var ys = new double[table.Rows + 1];
            for (int r = 0; r < table.Rows; r++)
                ys[r + 1] = ys[r] - table.RowHeightsMm[r] * norm * s * MmToFt;

            ResolveLineStyles();

            // Draw order matters visually: fills first, then lines, then text.
            if (options.IncludeFills)
                result.Fills = DrawFills(view, table, xs, ys);

            result.Lines = DrawBorders(view, table, xs, ys, options.DrawAllGridlines);
            result.Texts = DrawTexts(view, table, xs, ys, s, norm, options);

            // Fit pass: measure every note's real rendered size and fix the ones
            // that don't fit their cell. This is what guarantees clean output -
            // no font-metric guesswork survives contact with reality.
            FitTexts();

            return result;
        }

        /// <summary>
        /// The "body" font size of the table = the most common size among cells
        /// that contain text. Titles/headers are larger but far fewer, so the
        /// mode reliably picks the table body. Ties resolve to the smaller size.
        /// </summary>
        private static double MostCommonFontSize(TableModel table)
        {
            var withText = table.Cells.Where(c => !string.IsNullOrWhiteSpace(c.Text)).ToList();
            if (withText.Count == 0) return 11.0;

            return withText
                .GroupBy(c => Math.Round(c.FontSizePt * 2) / 2.0)   // bucket to 0.5pt
                .OrderByDescending(g => g.Count())
                .ThenBy(g => g.Key)
                .First().Key;
        }

        // ── View ─────────────────────────────────────────────────────────────

        /// <summary>
        /// Strip characters Revit prohibits in element names ( \ : { } [ ] | ; &lt; &gt; ? ~ ` ).
        /// Excel sheet names and user input can contain them; unsanitized names
        /// throw "Name cannot include prohibited characters".
        /// </summary>
        public static string SanitizeViewName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "XLS Import";
            var sb = new System.Text.StringBuilder(name.Length);
            foreach (char c in name)
                sb.Append("\\:{}[]|;<>?~`\r\n\t".IndexOf(c) >= 0 ? '-' : c);
            string clean = sb.ToString().Trim();
            return clean.Length == 0 ? "XLS Import" : clean;
        }

        private View CreateDraftingView(ImportOptions options)
        {
            var vft = new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .FirstOrDefault(t => t.ViewFamily == ViewFamily.Drafting);

            if (vft == null)
                throw new InvalidOperationException(
                    "No Drafting View type found in this project. Load one from your template.");

            var view = ViewDrafting.Create(_doc, vft.Id);

            string baseName = SanitizeViewName(string.IsNullOrWhiteSpace(options.ViewName)
                ? "XLS Import - " + options.SheetName
                : options.ViewName.Trim());

            // Ensure a unique view name
            string name = baseName;
            for (int i = 2; i < 100; i++)
            {
                try { view.Name = name; break; }
                catch { name = baseName + " (" + i + ")"; }
            }
            return view;
        }

        // ── Filled regions (cell shading) ────────────────────────────────────

        private int DrawFills(View view, TableModel t, double[] xs, double[] ys)
        {
            int count = 0;
            foreach (var cell in t.Cells.Where(c => c.Fill.HasValue))
            {
                double x0 = xs[cell.Col];
                double x1 = xs[cell.Col + cell.ColSpan];
                double y0 = ys[cell.Row];
                double y1 = ys[cell.Row + cell.RowSpan];
                if (x1 - x0 < 1e-9 || y0 - y1 < 1e-9) continue;   // collapsed (hidden) cell

                var frType = GetFillType(cell.Fill.Value);
                if (frType == null) continue;

                var loop = new CurveLoop();
                loop.Append(Line.CreateBound(new XYZ(x0, y0, 0), new XYZ(x1, y0, 0)));
                loop.Append(Line.CreateBound(new XYZ(x1, y0, 0), new XYZ(x1, y1, 0)));
                loop.Append(Line.CreateBound(new XYZ(x1, y1, 0), new XYZ(x0, y1, 0)));
                loop.Append(Line.CreateBound(new XYZ(x0, y1, 0), new XYZ(x0, y0, 0)));

                FilledRegion.Create(_doc, frType.Id, view.Id, new List<CurveLoop> { loop });
                count++;
            }
            return count;
        }

        private FilledRegionType GetFillType(DrawingColor color)
        {
            int key = color.ToArgb();
            if (_fillTypes.TryGetValue(key, out var cached)) return cached;

            string name = string.Format("XLS Fill {0:X2}{1:X2}{2:X2}", color.R, color.G, color.B);

            var existing = new FilteredElementCollector(_doc)
                .OfClass(typeof(FilledRegionType))
                .Cast<FilledRegionType>()
                .FirstOrDefault(f => f.Name == name);
            if (existing != null) { _fillTypes[key] = existing; return existing; }

            var solid = new FilteredElementCollector(_doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>()
                .FirstOrDefault(f =>
                {
                    var p = f.GetFillPattern();
                    return p.IsSolidFill && p.Target == FillPatternTarget.Drafting;
                });
            if (solid == null) return null;

            var seed = new FilteredElementCollector(_doc)
                .OfClass(typeof(FilledRegionType))
                .Cast<FilledRegionType>()
                .FirstOrDefault();
            if (seed == null) return null;

            var frType = (FilledRegionType)seed.Duplicate(name);
            frType.ForegroundPatternId = solid.Id;
            frType.ForegroundPatternColor = new Color(color.R, color.G, color.B);
            frType.BackgroundPatternId = ElementId.InvalidElementId;
            frType.IsMasking = false;

            _fillTypes[key] = frType;
            return frType;
        }

        // ── Borders (detail lines) ───────────────────────────────────────────

        private void ResolveLineStyles()
        {
            _lineStyles = new GraphicsStyle[4];
            try
            {
                var linesCat = _doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
                foreach (Category sub in linesCat.SubCategories)
                {
                    // English template names; harmless no-op on localized templates (falls back to default style)
                    if (sub.Name == "Thin Lines")   _lineStyles[1] = sub.GetGraphicsStyle(GraphicsStyleType.Projection);
                    if (sub.Name == "Medium Lines") _lineStyles[2] = sub.GetGraphicsStyle(GraphicsStyleType.Projection);
                    if (sub.Name == "Wide Lines")   _lineStyles[3] = sub.GetGraphicsStyle(GraphicsStyleType.Projection);
                }
            }
            catch { /* keep defaults */ }
        }

        private int DrawBorders(View view, TableModel t, double[] xs, double[] ys, bool allGridlines)
        {
            int count = 0;

            // Horizontal edges - merge consecutive segments of equal weight into single lines
            for (int r = 0; r <= t.Rows; r++)
            {
                int c = 0;
                while (c < t.Cols)
                {
                    int w = EdgeWeight(t.HEdges[r, c], allGridlines);
                    if (w == 0) { c++; continue; }

                    int start = c;
                    while (c < t.Cols && EdgeWeight(t.HEdges[r, c], allGridlines) == w) c++;

                    double y = ys[r];
                    if (xs[c] - xs[start] > 1e-9)
                    {
                        CreateLine(view, new XYZ(xs[start], y, 0), new XYZ(xs[c], y, 0), w);
                        count++;
                    }
                }
            }

            // Vertical edges
            for (int c = 0; c <= t.Cols; c++)
            {
                int r = 0;
                while (r < t.Rows)
                {
                    int w = EdgeWeight(t.VEdges[r, c], allGridlines);
                    if (w == 0) { r++; continue; }

                    int start = r;
                    while (r < t.Rows && EdgeWeight(t.VEdges[r, c], allGridlines) == w) r++;

                    double x = xs[c];
                    if (ys[start] - ys[r] > 1e-9)
                    {
                        CreateLine(view, new XYZ(x, ys[start], 0), new XYZ(x, ys[r], 0), w);
                        count++;
                    }
                }
            }
            return count;
        }

        private static int EdgeWeight(int weight, bool allGridlines)
            => allGridlines ? Math.Max(weight, 1) : weight;

        private void CreateLine(View view, XYZ p0, XYZ p1, int weight)
        {
            var curve = _doc.Create.NewDetailCurve(view, Line.CreateBound(p0, p1));
            var style = _lineStyles[Math.Min(3, Math.Max(1, weight))];
            if (style != null) curve.LineStyle = style;
        }

        // ── Text ─────────────────────────────────────────────────────────────

        // Every created note is remembered here so the fit pass can verify its
        // REAL rendered size against its cell and repair the ones that overflow.
        private class FitItem
        {
            public TextNote Note;
            public string FontName;
            public double SizeMm;
            public double OrigSizeMm;    // size before any fit-shrink (for uniformizing)
            public bool Bold, Italic;
            public DrawingColor Color;
            public double TextAreaWFt;   // model-space usable width (spill box minus padding)
            public double CellHFt;       // model-space cell height
            public bool Wrapped;         // note has a fixed width (wraps at its edge)
        }

        private readonly List<FitItem> _fitItems = new List<FitItem>();
        private double _scale = 1.0;

        private int DrawTexts(View view, TableModel t, double[] xs, double[] ys,
                              double scale, double norm, ImportOptions options)
        {
            int count = 0;
            _scale = scale;
            double textFactor = options.TextFactor * norm;
            double padFt = CellPadMm * scale * MmToFt;      // model-space inset
            double padPaper = CellPadMm * MmToFt;           // paper-space inset

            // Occupancy map: grid squares that contain text block overflow spill,
            // exactly like Excel (text can only spill over EMPTY neighbouring cells).
            var blocked = new bool[t.Rows, t.Cols];
            foreach (var cd in t.Cells)
            {
                if (string.IsNullOrWhiteSpace(cd.Text)) continue;
                for (int rr = cd.Row; rr < cd.Row + cd.RowSpan; rr++)
                    for (int cc = cd.Col; cc < cd.Col + cd.ColSpan; cc++)
                        blocked[rr, cc] = true;
            }

            foreach (var cell in t.Cells.Where(c => !string.IsNullOrWhiteSpace(c.Text)))
            {
                double x0 = xs[cell.Col];
                double x1 = xs[cell.Col + cell.ColSpan];
                double y0 = ys[cell.Row];
                double y1 = ys[cell.Row + cell.RowSpan];
                double cellW = x1 - x0;
                double cellH = y0 - y1;
                if (cellW < 1e-9 || cellH < 1e-9) continue;   // hidden row/col

                string fontName = string.IsNullOrWhiteSpace(options.FontOverride)
                    ? cell.FontName : options.FontOverride;
                double sizeMm = Math.Max(0.5, cell.FontSizePt * PtToMm * CapHeightFactor * textFactor);
                sizeMm = Math.Round(sizeMm * 20) / 20.0;   // 0.05mm buckets share one text type

                var type = GetTextTypeCore(fontName, sizeMm, cell.Bold, cell.Italic, cell.TextColor);
                if (type == null) continue;

                // ── Excel overflow (spill) emulation ─────────────────────────
                // Non-wrapped text may extend over adjacent empty cells, but is
                // hard-clamped at the first occupied cell and at the table edge,
                // so text can never collide with other text or leave the table.
                double availLeft = 0, availRight = 0;
                bool wraps = cell.WrapText || cell.Text.IndexOf('\n') >= 0;
                if (!wraps && cell.RowSpan == 1)
                {
                    int cc = cell.Col + cell.ColSpan;
                    while (cc < t.Cols && !blocked[cell.Row, cc]) cc++;
                    availRight = xs[cc] - x1;

                    cc = cell.Col;
                    while (cc > 0 && !blocked[cell.Row, cc - 1]) cc--;
                    availLeft = x0 - xs[cc];
                }

                double boxW;   // model-space width of the text box
                double ax;     // anchor X per alignment
                if (cell.HAlign < 0)       // left: spill right
                {
                    boxW = cellW + availRight;
                    ax = x0 + padFt;
                }
                else if (cell.HAlign > 0)  // right: spill left
                {
                    boxW = cellW + availLeft;
                    ax = x1 - padFt;
                }
                else                       // center: spill both ways, symmetrically
                {
                    double ext = Math.Min(availLeft, availRight);
                    boxW = cellW + 2 * ext;
                    ax = (x0 + x1) / 2.0;
                }

                double ay = cell.VAlign < 0 ? y0 - padFt
                          : cell.VAlign > 0 ? y1 + padFt
                          : (y0 + y1) / 2.0;
                var anchor = new XYZ(ax, ay, 0);

                var opts = new TextNoteOptions(type.Id)
                {
                    HorizontalAlignment = cell.HAlign < 0 ? HorizontalTextAlignment.Left
                                        : cell.HAlign > 0 ? HorizontalTextAlignment.Right
                                        : HorizontalTextAlignment.Center,
                    VerticalAlignment   = cell.VAlign < 0 ? VerticalTextAlignment.Top
                                        : cell.VAlign > 0 ? VerticalTextAlignment.Bottom
                                        : VerticalTextAlignment.Middle,
                };

                // Creation strategy (matches how Excel actually renders):
                //  - Excel wraps this cell  -> width-constrained note at CELL width
                //  - otherwise              -> free single-line note; the fit pass
                //    verifies its real rendered width and repairs any overflow.
                // TextNote width is PAPER-space feet (not model space).
                TextNote note;
                bool constrained = false;
                if (wraps)
                {
                    double desiredWidth = (boxW / scale) - 2 * padPaper;
                    var lim = GetWidthLimits(type.Id);
                    constrained = desiredWidth > lim.Item1;
                    note = constrained
                        ? TextNote.Create(_doc, view.Id, anchor,
                                          Math.Min(desiredWidth, lim.Item2), cell.Text, opts)
                        : TextNote.Create(_doc, view.Id, anchor, cell.Text, opts);
                }
                else
                {
                    note = TextNote.Create(_doc, view.Id, anchor, cell.Text, opts);
                }

                if (note != null)
                {
                    count++;
                    _fitItems.Add(new FitItem
                    {
                        Note        = note,
                        FontName    = fontName,
                        SizeMm      = sizeMm,
                        OrigSizeMm  = sizeMm,
                        Bold        = cell.Bold,
                        Italic      = cell.Italic,
                        Color       = cell.TextColor,
                        TextAreaWFt = boxW - 2 * padFt,
                        CellHFt     = cellH,
                        Wrapped     = constrained,
                    });
                }
            }
            return count;
        }

        /// <summary>
        /// Fit pass. After all notes exist, regenerate once so Revit computes their
        /// REAL rendered sizes, then repair every note that overflows its cell:
        ///
        ///  1. Single-line text wider than its spill box -> shrink the font a little
        ///     (like Excel's shrink-to-fit), at most down to <see cref="MinShrinkSingleLine"/>.
        ///  2. Still too wide -> give the note a fixed width so it wraps at the box edge.
        ///  3. Anything taller than its cell (wrapped text in a short row) -> shrink
        ///     proportionally, never below <see cref="MinTextSizeMm"/>.
        ///
        /// Runs up to 3 rounds; each round only regenerates when something changed.
        /// Decisions use Revit's own text layout, so no font-metric guesswork.
        /// </summary>
        private void FitTexts()
        {
            if (_fitItems.Count == 0) return;

            for (int round = 0; round < 3; round++)
            {
                _doc.Regenerate();
                bool changed = false;
                double tol = 0.05 * _scale * MmToFt;   // 0.05mm slack

                foreach (var it in _fitItems)
                {
                    // Note dimensions are paper-space; cells are model-space.
                    double wModel = it.Note.Width * _scale;
                    double hModel = it.Note.Height * _scale;

                    if (!it.Wrapped && wModel > it.TextAreaWFt + tol)
                    {
                        double factor = it.TextAreaWFt / wModel;
                        double newSize = Math.Floor(it.SizeMm * factor * 20) / 20.0;

                        if (factor >= MinShrinkSingleLine && newSize >= MinTextSizeMm)
                        {
                            // Gentle shrink keeps it on one line - barely noticeable
                            it.SizeMm = newSize;
                            var nt = GetTextTypeCore(it.FontName, newSize, it.Bold, it.Italic, it.Color);
                            if (nt != null) { it.Note.ChangeTypeId(nt.Id); changed = true; }
                        }
                        else
                        {
                            // Too long for a polite shrink - wrap at the box edge
                            var lim = GetWidthLimits(it.Note.GetTypeId());
                            double widthPaper = it.TextAreaWFt / _scale;
                            if (widthPaper > lim.Item1)
                            {
                                it.Note.Width = Math.Min(widthPaper, lim.Item2);
                                it.Wrapped = true;
                                changed = true;
                            }
                        }
                    }
                    else if (hModel > it.CellHFt + tol && it.SizeMm > MinTextSizeMm)
                    {
                        // Too tall for its row (usually wrapped text) - shrink to fit.
                        // Smaller text also fits more per line, so this converges fast.
                        double factor = it.CellHFt / hModel;
                        double newSize = Math.Max(MinTextSizeMm,
                            Math.Floor(it.SizeMm * factor * 20) / 20.0);

                        if (newSize < it.SizeMm - 1e-9)
                        {
                            it.SizeMm = newSize;
                            var nt = GetTextTypeCore(it.FontName, newSize, it.Bold, it.Italic, it.Color);
                            if (nt != null) { it.Note.ChangeTypeId(nt.Id); changed = true; }
                        }
                    }
                }

                if (!changed) break;
            }

            UniformizeSizes();
        }

        /// <summary>
        /// Make text visually consistent. Excel schedules are usually one font
        /// size for the whole body (e.g. Calibri 11) and a larger size for
        /// titles - but the per-note fit pass can leave many cells at slightly
        /// different shrunk sizes. This collapses each original-size group
        /// (same start size + font + weight) to ONE size: the smallest that any
        /// member needed. So all body text ends up identical, all titles
        /// identical - clean and professional, never a jumble of near-sizes.
        /// Wrapped (multi-line) notes are excluded; their height was tuned
        /// individually and forcing them smaller could clip content.
        /// </summary>
        private void UniformizeSizes()
        {
            var groups = _fitItems
                .Where(it => !it.Wrapped)
                .GroupBy(it => string.Format("{0}|{1:F1}|{2}|{3}",
                    it.FontName, it.OrigSizeMm, it.Bold, it.Italic));

            foreach (var g in groups)
            {
                double target = g.Min(it => it.SizeMm);   // smallest that fits all
                foreach (var it in g)
                {
                    if (Math.Abs(it.SizeMm - target) < 1e-6) continue;
                    var nt = GetTextTypeCore(it.FontName, target, it.Bold, it.Italic, it.Color);
                    if (nt != null) { it.Note.ChangeTypeId(nt.Id); it.SizeMm = target; }
                }
            }
        }

        // Min/max allowed TextNote width per text type - cached, these API calls
        // are too expensive to repeat for every one of hundreds of cells.
        private readonly Dictionary<ElementId, Tuple<double, double>> _widthLimits
            = new Dictionary<ElementId, Tuple<double, double>>();

        private Tuple<double, double> GetWidthLimits(ElementId typeId)
        {
            if (!_widthLimits.TryGetValue(typeId, out var lim))
            {
                lim = Tuple.Create(
                    TextNote.GetMinimumAllowedWidth(_doc, typeId),
                    TextNote.GetMaximumAllowedWidth(_doc, typeId));
                _widthLimits[typeId] = lim;
            }
            return lim;
        }

        /// <summary>
        /// Find or create a TextNoteType for the given style. Called both when
        /// notes are first created and by the fit pass when it shrinks text.
        /// </summary>
        private TextNoteType GetTextTypeCore(string fontName, double sizeMm,
                                             bool bold, bool italic, DrawingColor color)
        {
            string key = string.Format("{0}|{1:F2}|{2}|{3}|{4:X8}",
                fontName, sizeMm, bold, italic, color.ToArgb());
            if (_textTypes.TryGetValue(key, out var cached)) return cached;

            string name = string.Format("XLS {0} {1:0.##}mm{2}{3}{4}",
                fontName, sizeMm,
                bold ? " B" : "",
                italic ? " I" : "",
                color.ToArgb() == DrawingColor.Black.ToArgb()
                    ? "" : string.Format(" #{0:X2}{1:X2}{2:X2}", color.R, color.G, color.B));

            var existing = new FilteredElementCollector(_doc)
                .OfClass(typeof(TextNoteType))
                .Cast<TextNoteType>()
                .FirstOrDefault(x => x.Name == name);
            if (existing != null) { _textTypes[key] = existing; return existing; }

            var seed = new FilteredElementCollector(_doc)
                .OfClass(typeof(TextNoteType))
                .Cast<TextNoteType>()
                .FirstOrDefault();
            if (seed == null) return null;

            var type = (TextNoteType)seed.Duplicate(name);
            SetParam(type, BuiltInParameter.TEXT_FONT, fontName);
            SetParam(type, BuiltInParameter.TEXT_SIZE, sizeMm * MmToFt);
            SetParam(type, BuiltInParameter.TEXT_STYLE_BOLD, bold ? 1 : 0);
            SetParam(type, BuiltInParameter.TEXT_STYLE_ITALIC, italic ? 1 : 0);
            SetParam(type, BuiltInParameter.TEXT_BACKGROUND, 1);                    // transparent
            SetParam(type, BuiltInParameter.TEXT_WIDTH_SCALE, 1.0);
            // Text color: Revit stores color as B*65536 + G*256 + R
            SetParam(type, BuiltInParameter.LINE_COLOR,
                color.R + color.G * 256 + color.B * 65536);

            _textTypes[key] = type;
            return type;
        }

        private static void SetParam(Element e, BuiltInParameter bip, string value)
        { var p = e.get_Parameter(bip); if (p != null && !p.IsReadOnly) p.Set(value); }

        private static void SetParam(Element e, BuiltInParameter bip, int value)
        { var p = e.get_Parameter(bip); if (p != null && !p.IsReadOnly) p.Set(value); }

        private static void SetParam(Element e, BuiltInParameter bip, double value)
        { var p = e.get_Parameter(bip); if (p != null && !p.IsReadOnly) p.Set(value); }
    }
}
