using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using ExcelScheduleImporter.Excel;
using ExcelScheduleImporter.Revit;
using ElementId = Autodesk.Revit.DB.ElementId;

namespace ExcelScheduleImporter.UI
{
    /// <summary>
    /// The one-stop dialog, themed to match Revit (dark/light):
    ///   LEFT  - import a new schedule, grouped in "Import" / "View settings" panels
    ///   RIGHT - card list of every imported schedule (which sheet it's on,
    ///           whether Excel changed) with search and in-place Update buttons.
    /// </summary>
    public class ImportForm : Form
    {
        private const int LeftW = 520;    // width of the import column

        // ── import controls ──────────────────────────────────────────────────
        private TextBox _txtFile;
        private Button _btnBrowse;
        private DarkComboBox _cboSheet;
        private Button _btnPickSheets;
        private TextBox _txtRange;
        private Label _lblDetected;
        private TextBox _txtViewName;
        private DarkNumeric _numScale;
        private DarkNumeric _numTextSize;
        private CheckBox _chkNormalize;
        private DarkNumeric _numBodyMm;
        private CheckBox _chkFont;
        private TextBox _txtFont;
        private CheckBox _chkFills;
        private CheckBox _chkGridlines;
        private Button _btnOk;
        private Button _btnCancel;
        private Label _lblStatus;

        // ── status panel controls ────────────────────────────────────────────
        private TextBox _txtSearch;
        private FlowLayoutPanel _flow;
        private Button _btnSelectAll;
        private Button _btnDeselectAll;
        private Button _btnUpdateChecked;
        private Button _btnUpdateAll;
        private Label _lblUpdateStatus;
        private readonly ToolTip _tip = new ToolTip { AutoPopDelay = 12000 };

        private readonly HashSet<string> _existingViewNames;
        private string _detectedRange;   // last auto-detected range for the selected sheet

        /// <summary>
        /// Worksheets chosen via "Sheets..." for a batch import. Empty = ordinary
        /// single-sheet import driven by the combo box.
        /// </summary>
        private List<string> _batchSheets = new List<string>();

        /// <summary>Guards against the combo's change event clearing a batch selection.</summary>
        private bool _suppressSheetChanged;

        private readonly Func<List<ElementId>, ScheduleUpdateResult> _updateAction;
        private readonly Func<List<ManageRow>> _refreshRows;

        public ImportOptions Options { get; private set; }

        /// <summary>
        /// One entry per worksheet when the user chose several via "Sheets...".
        /// Null for an ordinary single-sheet import (use <see cref="Options"/>).
        /// Ranges are auto-detected and view names auto-generated per sheet.
        /// </summary>
        public List<ImportOptions> BatchOptions { get; private set; }

        /// <summary>
        /// True once at least one linked drafting view has been successfully
        /// rebuilt and committed while this dialog is open. The external command
        /// uses this to preserve those changes when the user closes the manager
        /// without starting a new import.
        /// </summary>
        public bool HasCommittedUpdates { get; private set; }

        /// <summary>
        /// The workbook, parsed once when the user picks a file. The command reuses
        /// this instance for the actual import, so the file is never parsed twice.
        /// </summary>
        public ExcelReader Reader { get; private set; }

        public ImportForm(IEnumerable<string> existingDraftingViewNames,
                          List<ManageRow> scheduleRows,
                          Func<List<ElementId>, ScheduleUpdateResult> updateAction,
                          Func<List<ManageRow>> refreshRows)
        {
            _existingViewNames = new HashSet<string>(
                existingDraftingViewNames ?? Enumerable.Empty<string>(),
                StringComparer.Ordinal);
            _updateAction = updateAction;
            _refreshRows = refreshRows;

            BuildLayout();
            PopulateSchedules(scheduleRows ?? new List<ManageRow>());
            ApplyIcon();
            Theme.Apply(this);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                Reader?.Dispose();
                Reader = null;
                _tip.Dispose();
            }
            base.Dispose(disposing);
        }

        // ── layout ───────────────────────────────────────────────────────────

        private void BuildLayout()
        {
            Text = "Excel Schedule Importer";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(1010, 574);
            Font = new Font("Segoe UI", 9f);

            // ═══ Header band: logo + title + accent hairline ═══════════════
            var header = new Panel { Left = 0, Top = 0, Width = ClientSize.Width, Height = 46, BackColor = Theme.InputBg };
            var hIcon = LoadDimmedIcon(24, 1f);
            if (hIcon != null)
                header.Controls.Add(new PictureBox { Left = 16, Top = 11, Width = 24, Height = 24, Image = hIcon, BackColor = Color.Transparent });
            header.Controls.Add(new Label
            {
                Left = 50, Top = 12, Width = 420, Height = 22,
                Text = "Excel Schedule Importer",
                Font = new Font("Segoe UI Semibold", 11f),
                ForeColor = Theme.Fg, BackColor = Color.Transparent,
            });
            var ver = Assembly.GetExecutingAssembly().GetName().Version;
            header.Controls.Add(new Label
            {
                Left = ClientSize.Width - 116, Top = 16, Width = 100, Height = 16,
                Text = "v" + ver.Major + "." + ver.Minor + "." + ver.Build,
                TextAlign = ContentAlignment.MiddleRight,
                ForeColor = Theme.FgDim, BackColor = Color.Transparent,
            });
            Controls.Add(header);
            Controls.Add(new Panel { Left = 0, Top = 46, Width = ClientSize.Width, Height = 1, BackColor = Theme.Accent });

            // control columns inside the section panels
            const int px = 16;    // label x inside a panel
            const int pc = 116;   // input x inside a panel
            const int pw = 296;   // standard input width

            // ═══ LEFT: "Import" section ════════════════════════════════════
            var pnlImport = new SectionPanel("Import", 14, 62, LeftW - 26, 152);
            Controls.Add(pnlImport);
            int y = 48;

            pnlImport.Controls.Add(MakeLabel("Excel file:", px, y));
            _txtFile = new TextBox { Left = pc, Top = y - 3, Width = pw, ReadOnly = true };
            Theme.SetPlaceholder(_txtFile, "Select the Excel file...");
            _btnBrowse = new Button { Left = pc + pw + 8, Top = y - 5, Width = 56, Height = 26, Text = "..." };
            _btnBrowse.Click += OnBrowse;
            pnlImport.Controls.Add(_txtFile); pnlImport.Controls.Add(_btnBrowse);
            y += 34;

            pnlImport.Controls.Add(MakeLabel("Worksheet:", px, y));
            _cboSheet = new DarkComboBox { Left = pc, Top = y - 3, Width = pw - 66, Height = 24, Enabled = false };
            _cboSheet.SelectedIndexChanged += OnSheetChanged;
            _btnPickSheets = new Button
            {
                Left = pc + pw - 58, Top = y - 5, Width = 58, Height = 26,
                Text = "Sheets", Enabled = false,
            };
            _btnPickSheets.Click += OnPickSheets;
            _tip.SetToolTip(_btnPickSheets,
                "Import several worksheets at once. Each becomes its own drafting view, "
                + "with the range auto-detected and the view named after the sheet.");
            pnlImport.Controls.Add(_cboSheet);
            pnlImport.Controls.Add(_btnPickSheets);
            y += 34;

            pnlImport.Controls.Add(MakeLabel("Cell range:", px, y));
            _txtRange = new TextBox { Left = pc, Top = y - 3, Width = 140, Enabled = false };
            _lblDetected = new Label { Left = pc + 150, Top = y, Width = 210, ForeColor = Color.DimGray, Text = "" };
            pnlImport.Controls.Add(_txtRange); pnlImport.Controls.Add(_lblDetected);

            // ═══ LEFT: "View settings" section ═════════════════════════════
            var pnlView = new SectionPanel("View settings", 14, 226, LeftW - 26, 282);
            Controls.Add(pnlView);
            y = 48;

            pnlView.Controls.Add(MakeLabel("View name:", px, y));
            _txtViewName = new TextBox { Left = pc, Top = y - 3, Width = pw };
            pnlView.Controls.Add(_txtViewName);
            y += 34;

            pnlView.Controls.Add(MakeLabel("View scale  1 :", px, y));
            _numScale = new DarkNumeric { Left = pc, Top = y - 3, Width = 76, Minimum = 1, Maximum = 1000, Value = 1 };
            pnlView.Controls.Add(_numScale);
            pnlView.Controls.Add(new Label { Left = pc + 86, Top = y, Width = 270, ForeColor = Color.DimGray,
                Text = "Printed size is identical at any scale" });
            y += 34;

            pnlView.Controls.Add(MakeLabel("Text size %:", px, y));
            _numTextSize = new DarkNumeric { Left = pc, Top = y - 3, Width = 76, Minimum = 25, Maximum = 400, Value = 100 };
            pnlView.Controls.Add(_numTextSize);
            y += 34;

            // Default OFF: 1:1 with Excel gives the most faithful table layout.
            _chkNormalize = new CheckBox { Left = pc, Top = y - 2, Width = 146, Text = "Match body text to", Checked = false };
            _numBodyMm = new DarkNumeric
            {
                Left = pc + 148, Top = y - 3, Width = 64,
                Minimum = 0.5M, Maximum = 10M, DecimalPlaces = 2, Increment = 0.25M, Value = 2.00M,
                Enabled = false,
            };
            pnlView.Controls.Add(new Label { Left = pc + 216, Top = y, Width = 150, ForeColor = Color.DimGray,
                Text = "mm  (titles scale too)" });
            _chkNormalize.CheckedChanged += (s2, e2) => _numBodyMm.Enabled = _chkNormalize.Checked;
            pnlView.Controls.Add(_chkNormalize); pnlView.Controls.Add(_numBodyMm);
            y += 34;

            _chkFont = new CheckBox { Left = pc, Top = y - 2, Width = 146, Text = "Override font:", Checked = false };
            _txtFont = new TextBox { Left = pc + 148, Top = y - 3, Width = 118, Text = "Arial", Enabled = false };
            _chkFont.CheckedChanged += (s2, e2) => _txtFont.Enabled = _chkFont.Checked;
            pnlView.Controls.Add(_chkFont); pnlView.Controls.Add(_txtFont);
            y += 34;

            _chkFills = new CheckBox { Left = pc, Top = y, Width = pw + 60, Text = "Recreate cell fill colors (filled regions)", Checked = true };
            pnlView.Controls.Add(_chkFills);
            y += 26;
            _chkGridlines = new CheckBox { Left = pc, Top = y, Width = pw + 60, Text = "Draw thin gridlines where Excel has no borders", Checked = false };
            pnlView.Controls.Add(_chkGridlines);

            // Validation: centered red text between the panels and the buttons.
            _lblStatus = new Label
            {
                Left = 14, Top = 514, Width = LeftW - 26, Height = 18,
                ForeColor = Color.Firebrick, Text = "",
                TextAlign = ContentAlignment.MiddleCenter,
            };
            _btnOk = new Button { Left = LeftW - 192, Top = 534, Width = 96, Height = 32, Text = "Import", Enabled = false, Tag = "primary" };
            _btnCancel = new Button { Left = LeftW - 88, Top = 534, Width = 74, Height = 32, Text = "Close", DialogResult = DialogResult.Cancel };
            _btnOk.Click += OnOk;
            Controls.Add(_lblStatus); Controls.Add(_btnOk); Controls.Add(_btnCancel);

            AcceptButton = _btnOk;
            CancelButton = _btnCancel;

            // ═══ separator ═════════════════════════════════════════════════
            Controls.Add(new Panel { Left = LeftW + 6, Top = 62, Width = 1, Height = ClientSize.Height - 76, BackColor = Theme.Border });

            // ═══ RIGHT: imported schedules ═════════════════════════════════
            int rx = LeftW + 22;
            int rw = ClientSize.Width - rx - 16;

            Controls.Add(new Label
            {
                Left = rx, Top = 64, Width = 250, Height = 20,
                Text = "Imported schedules",
                Font = new Font("Segoe UI Semibold", 10f),
            });
            _btnSelectAll = new Button
            {
                Left = rx + rw - 178, Top = 60, Width = 80, Height = 26,
                Text = "Select all",
            };
            _btnDeselectAll = new Button
            {
                Left = rx + rw - 92, Top = 60, Width = 92, Height = 26,
                Text = "Deselect all",
            };
            _btnSelectAll.Click += (s2, e2) => SetVisibleCardsChecked(true);
            _btnDeselectAll.Click += (s2, e2) => SetVisibleCardsChecked(false);
            _tip.SetToolTip(_btnSelectAll, "Select every schedule currently shown by the search filter.");
            _tip.SetToolTip(_btnDeselectAll, "Deselect every schedule currently shown by the search filter.");
            Controls.Add(_btnSelectAll);
            Controls.Add(_btnDeselectAll);
            Controls.Add(new Panel { Left = rx + 1, Top = 85, Width = 32, Height = 3, BackColor = Theme.Accent });

            _txtSearch = new TextBox { Left = rx, Top = 100, Width = rw };
            Theme.SetPlaceholder(_txtSearch, "Search by view, sheet or file...");
            _txtSearch.TextChanged += (s2, e2) => ApplyFilter();
            Controls.Add(_txtSearch);

            _flow = new FlowLayoutPanel
            {
                Left = rx, Top = 132, Width = rw, Height = 334,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true,
                BackColor = Theme.WindowBg,
                Padding = new Padding(0),
            };
            Controls.Add(_flow);

            Controls.Add(new Label
            {
                Left = rx, Top = 474, Width = rw, Height = 16, ForeColor = Color.DimGray,
                Text = "Orange = out of date, pre-checked.  Updated views stay on their sheets.",
                AutoEllipsis = true,
            });

            _btnUpdateChecked = new Button { Left = rx, Top = 500, Width = 134, Height = 32, Text = "Update Checked", Tag = "primary" };
            _btnUpdateAll = new Button { Left = rx + 142, Top = 500, Width = 104, Height = 32, Text = "Update All" };
            _btnUpdateChecked.Click += (s2, e2) => OnUpdate(onlyChecked: true);
            _btnUpdateAll.Click += (s2, e2) => OnUpdate(onlyChecked: false);
            Controls.Add(_btnUpdateChecked); Controls.Add(_btnUpdateAll);

            _lblUpdateStatus = new Label
            {
                Left = rx + 256, Top = 508, Width = rw - 256, Height = 18,
                ForeColor = Color.DarkGreen, Text = "",
                AutoEllipsis = true,
            };
            Controls.Add(_lblUpdateStatus);
        }

        private static Label MakeLabel(string text, int x, int y)
            => new Label { Left = x, Top = y, Width = 100, Text = text };

        /// <summary>Rounded section container with a semibold title and accent underline.</summary>
        private sealed class SectionPanel : Panel
        {
            public SectionPanel(string title, int x, int y, int w, int h)
            {
                Left = x; Top = y; Width = w; Height = h;
                BackColor = Theme.PanelBg;   // child labels blend via ambient BackColor

                Controls.Add(new Label
                {
                    Left = 18, Top = 14, Width = 250, Height = 20, Text = title,
                    Font = new Font("Segoe UI Semibold", 10.5f),
                    ForeColor = Theme.Fg, BackColor = Theme.PanelBg,
                });
                Controls.Add(new Panel { Left = 19, Top = 36, Width = 30, Height = 3, BackColor = Theme.Accent });

                Paint += (s, e) =>
                {
                    e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                    using (var path = Theme.RoundedPath(new Rectangle(0, 0, Width - 1, Height - 1), 10))
                    using (var pen = new Pen(Theme.Border))
                        e.Graphics.DrawPath(pen, path);
                };
                Theme.RoundControl(this, 10);
            }
        }

        // ── schedule cards ───────────────────────────────────────────────────

        private void PopulateSchedules(List<ManageRow> rows)
        {
            _flow.SuspendLayout();
            foreach (Control c in _flow.Controls.OfType<Control>().ToList()) c.Dispose();
            _flow.Controls.Clear();

            int cardW = _flow.ClientSize.Width - SystemInformation.VerticalScrollBarWidth - 4;

            foreach (var row in rows.OrderByDescending(r => r.OutOfDate).ThenBy(r => r.ViewName))
                _flow.Controls.Add(new ScheduleCard(row, cardW, _tip));

            bool any = rows.Count > 0;
            if (!any) _flow.Controls.Add(BuildEmptyState(cardW));

            _btnUpdateChecked.Enabled = any;
            _btnUpdateAll.Enabled = any;
            _btnSelectAll.Enabled = any;
            _btnDeselectAll.Enabled = any;

            _flow.ResumeLayout();
            ApplyFilter();
        }

        /// <summary>Composed empty state: dimmed logo + title + hint, centered.</summary>
        private Control BuildEmptyState(int width)
        {
            var panel = new Panel { Width = width, Height = 300, BackColor = Color.Transparent, Margin = new Padding(0) };

            var img = LoadDimmedIcon(48, 0.30f);
            if (img != null)
            {
                panel.Controls.Add(new PictureBox
                {
                    Left = (width - 48) / 2, Top = 92, Width = 48, Height = 48,
                    Image = img, BackColor = Color.Transparent,
                });
            }
            panel.Controls.Add(new Label
            {
                Left = 0, Top = 152, Width = width, Height = 20,
                Text = "No imported schedules yet",
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI Semibold", 9.5f),
                ForeColor = Theme.Fg, BackColor = Color.Transparent,
            });
            panel.Controls.Add(new Label
            {
                Left = 20, Top = 176, Width = width - 40, Height = 34,
                Text = "Import an Excel schedule on the left.\nIt will appear here with its update status.",
                TextAlign = ContentAlignment.TopCenter,
                ForeColor = Theme.FgDim, BackColor = Color.Transparent,
            });
            return panel;
        }

        private static Bitmap LoadDimmedIcon(int size, float alpha)
        {
            try
            {
                var asm = Assembly.GetExecutingAssembly();
                using (var s = asm.GetManifestResourceStream("ExcelScheduleImporter.Resources.icon32.png"))
                {
                    if (s == null) return null;
                    using (var src = new Bitmap(s))
                    {
                        var bmp = new Bitmap(size, size);
                        using (var g = Graphics.FromImage(bmp))
                        using (var ia = new ImageAttributes())
                        {
                            ia.SetColorMatrix(new ColorMatrix { Matrix33 = alpha });
                            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                            g.DrawImage(src, new Rectangle(0, 0, size, size),
                                        0, 0, src.Width, src.Height, GraphicsUnit.Pixel, ia);
                        }
                        return bmp;
                    }
                }
            }
            catch { return null; }
        }

        private void ApplyFilter()
        {
            string q = (_txtSearch.Text ?? "").Trim();
            bool anyVisible = false;
            foreach (var card in _flow.Controls.OfType<ScheduleCard>())
            {
                bool matches = q.Length == 0 || card.Matches(q);
                card.Visible = matches;
                anyVisible |= matches;
            }

            _btnSelectAll.Enabled = anyVisible;
            _btnDeselectAll.Enabled = anyVisible;
        }

        /// <summary>
        /// Select or deselect the schedules currently shown by the search filter.
        /// Hidden cards keep their selection, making filtered batch updates useful.
        /// </summary>
        private void SetVisibleCardsChecked(bool isChecked)
        {
            foreach (var card in _flow.Controls.OfType<ScheduleCard>().Where(c => c.Visible))
                card.CardChecked = isChecked;
        }

        private void OnUpdate(bool onlyChecked)
        {
            var ids = _flow.Controls.OfType<ScheduleCard>()
                .Where(c => !onlyChecked || c.CardChecked)
                .Select(c => c.Row)
                .Where(r => !r.FileMissing)          // cannot update without the source
                .Select(r => r.ViewId)
                .ToList();

            if (ids.Count == 0)
            {
                _lblUpdateStatus.ForeColor = Theme.Danger;
                _lblUpdateStatus.Text = onlyChecked
                    ? "Nothing checked (or sources missing)."
                    : "Nothing can be updated.";
                return;
            }

            try
            {
                Cursor = Cursors.WaitCursor;
                _btnUpdateChecked.Enabled = _btnUpdateAll.Enabled = false;

                ScheduleUpdateResult updateResult = _updateAction(ids);
                if (updateResult == null)
                    throw new InvalidOperationException("The schedule updater returned no result.");

                if (updateResult.TransactionCommitted && updateResult.UpdatedCount > 0)
                    HasCommittedUpdates = true;

                string summary = updateResult.Summary ?? "Update completed.";

                PopulateSchedules(_refreshRows());

                // Any failure -> readable, resizable-content MessageBox (the inline
                // label is too narrow and cannot be widened). Success -> short label.
                bool failed = summary.IndexOf("fail", StringComparison.OrdinalIgnoreCase) >= 0;
                if (failed)
                {
                    _lblUpdateStatus.ForeColor = Theme.Danger;
                    _lblUpdateStatus.Text = "Update failed - see details.";
                    MessageBox.Show(this, summary, "Update Excel Schedules",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    _lblUpdateStatus.ForeColor = Theme.Success;
                    _lblUpdateStatus.Text = summary;
                }
            }
            catch (Exception ex)
            {
                App.LogCrash("ImportForm.OnUpdate", ex);
                _lblUpdateStatus.ForeColor = Theme.Danger;
                _lblUpdateStatus.Text = "Update failed - see details.";
                MessageBox.Show(this,
                    ex.Message + "\n\nFull details logged to:\n" + App.CrashLogPath,
                    "Update Excel Schedules - Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
                bool any = _flow.Controls.OfType<ScheduleCard>().Any();
                _btnUpdateChecked.Enabled = any;
                _btnUpdateAll.Enabled = any;
            }
        }

        /// <summary>
        /// One imported schedule as a card: colored status edge, bold name,
        /// status text, dim detail lines. Click anywhere toggles the checkbox.
        /// </summary>
        private sealed class ScheduleCard : Panel
        {
            public ManageRow Row { get; }
            private readonly CheckBox _chk;
            private readonly Color _statusColor;
            private Color _fill;

            public bool CardChecked
            {
                get => _chk.Checked;
                set => _chk.Checked = value;
            }

            public ScheduleCard(ManageRow row, int width, ToolTip tip)
            {
                Row = row;
                Width = width;
                Height = 80;
                Margin = new Padding(0, 0, 0, 9);
                BackColor = Theme.WindowBg;   // painted card sits on the window bg
                Cursor = Cursors.Hand;
                _fill = Theme.CardBg;
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer
                         | ControlStyles.ResizeRedraw | ControlStyles.UserPaint, true);

                _statusColor = row.FileMissing ? Theme.Danger
                             : row.OutOfDate   ? Theme.Warning
                             :                   Theme.Success;

                _chk = new CheckBox
                {
                    Left = 16, Top = 12, Width = 18, Height = 18,
                    Checked = row.OutOfDate && !row.FileMissing,
                    BackColor = Theme.CardBg,
                    FlatStyle = FlatStyle.Flat,
                };
                _chk.FlatAppearance.BorderSize = 0;

                var name = new Label
                {
                    Left = 40, Top = 11, Width = width - 178, Height = 18,
                    Text = row.ViewName, AutoEllipsis = true,
                    Font = new Font("Segoe UI Semibold", 9.5f),
                    ForeColor = Theme.Fg, BackColor = Theme.CardBg,
                };

                var sheet = new Label
                {
                    Left = 40, Top = 33, Width = width - 52, Height = 16,
                    Text = "Sheet:    " + row.SheetPlacement,
                    ForeColor = Theme.FgDim, AutoEllipsis = true, BackColor = Theme.CardBg,
                };

                var src = new Label
                {
                    Left = 40, Top = 53, Width = width - 52, Height = 16,
                    Text = "Source:  " + row.FileName + "   [" + row.Worksheet + "  " + row.Range + "]",
                    ForeColor = Theme.FgDim, AutoEllipsis = true, BackColor = Theme.CardBg,
                };

                Controls.AddRange(new Control[] { _chk, name, sheet, src });

                string tipText = row.FilePath
                    + "\nWorksheet: " + row.Worksheet + "   Range: " + row.Range
                    + "\nImported: " + (row.ImportedUtc == DateTime.MinValue
                        ? "-" : row.ImportedUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm"));
                tip.SetToolTip(this, tipText);
                tip.SetToolTip(name, tipText);
                tip.SetToolTip(sheet, tipText);
                tip.SetToolTip(src, tipText);

                // Rounded card: fill + subtle border, rounded colored status edge,
                // and a tinted rounded status pill.
                var pillFont = new Font("Segoe UI Semibold", 8.25f);
                Paint += (s, e) =>
                {
                    var g = e.Graphics;
                    g.Clear(Theme.WindowBg);   // corners blend into the flow background
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    var r = new Rectangle(0, 0, Width - 1, Height - 1);

                    using (var path = RoundedRect(r, 9))
                    {
                        using (var fill = new SolidBrush(_fill)) g.FillPath(fill, path);
                        using (var pen = new Pen(Theme.Border)) g.DrawPath(pen, path);

                        // colored status edge, clipped to the rounded card shape
                        var clip = g.Clip;
                        g.SetClip(path);
                        using (var edge = new SolidBrush(_statusColor))
                            g.FillRectangle(edge, 0, 0, 4, Height);
                        g.Clip = clip;
                    }

                    var textSize = TextRenderer.MeasureText(g, Row.Status, pillFont);
                    var pill = new Rectangle(Width - textSize.Width - 32, 10, textSize.Width + 16, 20);
                    using (var path = RoundedRect(pill, 10))
                    {
                        using (var bg = new SolidBrush(Color.FromArgb(40, _statusColor)))
                            g.FillPath(bg, path);
                        using (var pen = new Pen(Color.FromArgb(120, _statusColor)))
                            g.DrawPath(pen, path);
                    }
                    TextRenderer.DrawText(g, Row.Status, pillFont, pill, _statusColor,
                        TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                };
                Theme.RoundControl(this, 9);
                WireHover(this);
                foreach (Control c in Controls)
                {
                    WireHover(c);
                    c.Cursor = Cursors.Hand;
                    if (!(c is CheckBox))
                        c.Click += (s, e) => _chk.Checked = !_chk.Checked;
                }
                Click += (s, e) => _chk.Checked = !_chk.Checked;
            }

            private static GraphicsPath RoundedRect(Rectangle r, int rad)
            {
                var p = new GraphicsPath();
                int d = rad * 2;
                p.AddArc(r.X, r.Y, d, d, 180, 90);
                p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
                p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
                p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
                p.CloseFigure();
                return p;
            }

            private void SetFill(Color c)
            {
                _fill = c;
                foreach (Control child in Controls) child.BackColor = c;
                Invalidate();
            }

            private void WireHover(Control c)
            {
                c.MouseEnter += (s, e) => SetFill(Theme.CardHover);
                c.MouseLeave += (s, e) =>
                {
                    if (!ClientRectangle.Contains(PointToClient(Cursor.Position)))
                        SetFill(Theme.CardBg);
                };
            }

            public bool Matches(string query)
                => (Row.ViewName ?? "").IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0
                || (Row.FileName ?? "").IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0
                || (Row.Worksheet ?? "").IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0
                || (Row.SheetPlacement ?? "").IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        // ── import events ────────────────────────────────────────────────────

        private void OnBrowse(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog
            {
                Title = "Select Excel file",
                Filter = "Excel files (*.xlsx;*.xlsm)|*.xlsx;*.xlsm",
                CheckFileExists = true,
            })
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;

                try
                {
                    Cursor = Cursors.WaitCursor;
                    _lblStatus.ForeColor = Color.Firebrick;
                    _lblStatus.Text = "Reading workbook...";
                    _lblStatus.Refresh();
                    // Drop the previous workbook BEFORE opening the new one, so a
                    // failed open never leaves a disposed reader behind.
                    Reader?.Dispose();
                    Reader = null;
                    Reader = new ExcelReader(dlg.FileName);   // parse once, reuse everywhere
                    var sheets = Reader.GetSheetNames();
                    _txtFile.Text = dlg.FileName;
                    ClearBatch();                     // a new workbook invalidates any batch
                    _cboSheet.Items.Clear();
                    foreach (var s in sheets) _cboSheet.Items.Add(s);
                    _cboSheet.Enabled = sheets.Count > 0;
                    _btnPickSheets.Enabled = sheets.Count > 1;
                    if (sheets.Count > 0) _cboSheet.SelectedIndex = 0;
                    _lblStatus.Text = "";
                }
                catch (Exception ex)
                {
                    // Reset the import side so nothing refers to the old workbook.
                    _txtFile.Text = "";
                    _cboSheet.Items.Clear();
                    _cboSheet.Enabled = false;
                    _btnPickSheets.Enabled = false;
                    _btnOk.Enabled = false;
                    ClearBatch();

                    App.LogCrash("ImportForm.OnBrowse", ex);
                    _lblStatus.ForeColor = Color.Firebrick;
                    _lblStatus.Text = "Cannot read file - see details.";
                    // The one-line label truncated the actual Windows error; show all of it.
                    MessageBox.Show(this,
                        ex.Message + "\n\nFull details logged to:\n" + App.CrashLogPath,
                        "Cannot read Excel file",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                finally { Cursor = Cursors.Default; }
            }
        }

        /// <summary>
        /// Choose several worksheets to import in one go. Picking a single sheet
        /// simply selects it in the combo (ordinary single import); picking two or
        /// more switches the dialog into batch mode.
        /// </summary>
        private void OnPickSheets(object sender, EventArgs e)
        {
            if (Reader == null) return;

            var all = _cboSheet.Items.Cast<object>().Select(o => o.ToString()).ToList();
            var preselected = _batchSheets.Count > 0
                ? _batchSheets
                : (_cboSheet.SelectedItem != null
                    ? new List<string> { _cboSheet.SelectedItem.ToString() }
                    : new List<string>());

            using (var picker = new SheetPickerForm(all, preselected))
            {
                if (picker.ShowDialog(this) != DialogResult.OK) return;
                var chosen = picker.SelectedSheets;
                if (chosen.Count == 0) return;

                if (chosen.Count == 1)
                {
                    // Not really a batch - fall back to the normal single-sheet flow
                    // so the user keeps range and view-name editing.
                    ClearBatch();
                    bool alreadySelected = _cboSheet.SelectedItem != null
                        && string.Equals(_cboSheet.SelectedItem.ToString(), chosen[0], StringComparison.Ordinal);

                    if (alreadySelected)
                        OnSheetChanged(null, EventArgs.Empty);   // no event fires; refresh by hand
                    else
                        _cboSheet.SelectedItem = chosen[0];
                    return;
                }
                EnterBatchMode(chosen);
            }
        }

        /// <summary>
        /// Batch mode: per-sheet range and view name are derived automatically, so
        /// those two inputs are disabled and clearly labelled as such. Every other
        /// setting (scale, text size, fills, gridlines, font) still applies to all.
        /// </summary>
        private void EnterBatchMode(List<string> sheets)
        {
            _batchSheets = sheets;

            _suppressSheetChanged = true;
            try { _cboSheet.SelectedItem = sheets[0]; }
            finally { _suppressSheetChanged = false; }

            _txtRange.Enabled = false;
            _txtRange.Text = "";
            _lblDetected.Text = "auto-detected per sheet";
            _txtViewName.Enabled = false;
            _txtViewName.Text = "(named after each worksheet)";

            _btnOk.Enabled = true;
            _btnOk.Text = "Import " + sheets.Count;
            _lblStatus.ForeColor = Theme.Accent;
            _lblStatus.Text = sheets.Count + " worksheets selected - each becomes its own view.";
        }

        /// <summary>Leave batch mode and restore ordinary single-sheet editing.</summary>
        private void ClearBatch()
        {
            if (_batchSheets.Count == 0) return;
            _batchSheets = new List<string>();
            _txtViewName.Enabled = true;
            _btnOk.Text = "Import";
            _lblStatus.ForeColor = Color.Firebrick;
            _lblStatus.Text = "";
        }

        private void OnSheetChanged(object sender, EventArgs e)
        {
            if (_suppressSheetChanged) return;
            if (_cboSheet.SelectedItem == null) return;
            ClearBatch();   // an explicit combo pick means "just this one sheet"
            string sheet = _cboSheet.SelectedItem.ToString();

            try
            {
                Cursor = Cursors.WaitCursor;
                string range = Reader.DetectUsedRange(sheet);
                if (range == null)
                {
                    _txtRange.Text = "";
                    _lblDetected.Text = "Sheet is empty";
                    _btnOk.Enabled = false;
                }
                else
                {
                    // Strip sheet prefix if present ('Sheet1'!A1:F42 -> A1:F42)
                    int bang = range.IndexOf('!');
                    if (bang >= 0) range = range.Substring(bang + 1);
                    _detectedRange = range;
                    _txtRange.Text = range;
                    _lblDetected.Text = "auto-detected (editable)";
                    _btnOk.Enabled = true;
                }
                _txtRange.Enabled = range != null;
                _txtViewName.Text = DraftingViewBuilder.SanitizeViewName("XLS Import - " + sheet);
                _lblStatus.Text = "";
            }
            catch (Exception ex)
            {
                _lblStatus.Text = "Cannot read sheet: " + ex.Message;
                _btnOk.Enabled = false;
            }
            finally { Cursor = Cursors.Default; }
        }

        private void OnOk(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_txtFile.Text) || !File.Exists(_txtFile.Text))
            { _lblStatus.ForeColor = Color.Firebrick; _lblStatus.Text = "Select an Excel file first."; return; }
            if (_cboSheet.SelectedItem == null)
            { _lblStatus.ForeColor = Color.Firebrick; _lblStatus.Text = "Select a worksheet."; return; }

            if (_batchSheets.Count > 0) { OnOkBatch(); return; }

            if (string.IsNullOrWhiteSpace(_txtRange.Text))
            { _lblStatus.ForeColor = Color.Firebrick; _lblStatus.Text = "Enter a cell range (e.g. A1:F42)."; return; }

            string sheetName = _cboSheet.SelectedItem.ToString();
            string viewName = DraftingViewBuilder.SanitizeViewName(
                string.IsNullOrWhiteSpace(_txtViewName.Text)
                    ? "XLS Import - " + sheetName
                    : _txtViewName.Text.Trim());

            // Name collision: let the user replace the existing view or go back
            // and type a different name.
            bool replaceExisting = false;
            if (_existingViewNames.Contains(viewName))
            {
                var choice = MessageBox.Show(this,
                    "A drafting view named \"" + viewName + "\" already exists in this project.\n\n" +
                    "OK      - replace it (the existing view is deleted; if it is placed " +
                    "on a sheet it will be removed from that sheet)\n" +
                    "Cancel - keep it and enter a different view name",
                    "View already exists",
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button2);

                if (choice != DialogResult.OK)
                {
                    _txtViewName.Focus();
                    _txtViewName.SelectAll();
                    return;   // stay in the dialog
                }
                replaceExisting = true;
            }

            Options = new ImportOptions
            {
                FilePath = _txtFile.Text,
                SheetName = sheetName,
                RangeOverride = _txtRange.Text.Trim(),
                AutoRange = string.Equals(_txtRange.Text.Trim(), _detectedRange,
                                          StringComparison.OrdinalIgnoreCase),
                ViewName = viewName,
                ReplaceExisting = replaceExisting,
                Scale = (int)_numScale.Value,
                TextFactor = (double)_numTextSize.Value / 100.0,
                NormalizeBodyText = _chkNormalize.Checked,
                BodyTextMm = (double)_numBodyMm.Value,
                FontOverride = _chkFont.Checked && !string.IsNullOrWhiteSpace(_txtFont.Text)
                    ? _txtFont.Text.Trim() : null,
                IncludeFills = _chkFills.Checked,
                DrawAllGridlines = _chkGridlines.Checked,
            };
            DialogResult = DialogResult.OK;
            Close();
        }

        /// <summary>
        /// Build one ImportOptions per selected worksheet. Ranges are auto-detected
        /// here (cheap - the workbook is already parsed) so empty sheets can be
        /// dropped before the user waits on a long import. Name collisions are
        /// resolved once for the whole batch rather than sheet by sheet.
        /// </summary>
        private void OnOkBatch()
        {
            var plan = new List<ImportOptions>();
            var skippedEmpty = new List<string>();
            var collisions = new List<string>();
            // Two worksheets can sanitise to the same view name (they may differ
            // only by characters Revit forbids). Revit would reject the second,
            // so make names unique within the batch up front.
            var namesInBatch = new HashSet<string>(StringComparer.Ordinal);

            try
            {
                Cursor = Cursors.WaitCursor;

                foreach (var sheet in _batchSheets)
                {
                    string range;
                    try { range = Reader.DetectUsedRange(sheet); }
                    catch { range = null; }

                    if (string.IsNullOrWhiteSpace(range)) { skippedEmpty.Add(sheet); continue; }

                    int bang = range.IndexOf('!');
                    if (bang >= 0) range = range.Substring(bang + 1);

                    string viewName = DraftingViewBuilder.SanitizeViewName("XLS Import - " + sheet);
                    if (!namesInBatch.Add(viewName))
                    {
                        string unique;
                        int n = 2;
                        do { unique = viewName + " (" + n++ + ")"; }
                        while (!namesInBatch.Add(unique));
                        viewName = unique;
                    }
                    if (_existingViewNames.Contains(viewName)) collisions.Add(viewName);

                    plan.Add(new ImportOptions
                    {
                        FilePath = _txtFile.Text,
                        SheetName = sheet,
                        RangeOverride = range,
                        AutoRange = true,          // batch always uses detection
                        ViewName = viewName,
                        Scale = (int)_numScale.Value,
                        TextFactor = (double)_numTextSize.Value / 100.0,
                        NormalizeBodyText = _chkNormalize.Checked,
                        BodyTextMm = (double)_numBodyMm.Value,
                        FontOverride = _chkFont.Checked && !string.IsNullOrWhiteSpace(_txtFont.Text)
                            ? _txtFont.Text.Trim() : null,
                        IncludeFills = _chkFills.Checked,
                        DrawAllGridlines = _chkGridlines.Checked,
                    });
                }
            }
            finally { Cursor = Cursors.Default; }

            if (plan.Count == 0)
            {
                _lblStatus.ForeColor = Color.Firebrick;
                _lblStatus.Text = "All selected worksheets are empty - nothing to import.";
                return;
            }

            // One decision for every colliding name in the batch.
            bool replace = false;
            if (collisions.Count > 0)
            {
                string list = string.Join("\n   - ", collisions.Take(10));
                if (collisions.Count > 10) list += "\n   - ... and " + (collisions.Count - 10) + " more";

                var choice = MessageBox.Show(this,
                    collisions.Count + " of the " + plan.Count + " views already exist:\n\n   - " + list +
                    "\n\nOK      - replace them (existing views are deleted; any that are " +
                    "placed on a sheet will be removed from that sheet)\n" +
                    "Cancel - skip those and import only the new ones",
                    "Some views already exist",
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button1);

                replace = choice == DialogResult.OK;
                if (replace)
                {
                    foreach (var o in plan)
                        if (_existingViewNames.Contains(o.ViewName)) o.ReplaceExisting = true;
                }
                else
                {
                    plan = plan.Where(o => !_existingViewNames.Contains(o.ViewName)).ToList();
                    if (plan.Count == 0)
                    {
                        _lblStatus.ForeColor = Color.Firebrick;
                        _lblStatus.Text = "Every selected worksheet already has a view - nothing to import.";
                        return;
                    }
                }
            }

            if (skippedEmpty.Count > 0)
            {
                MessageBox.Show(this,
                    "These worksheets are empty and will be skipped:\n\n   - "
                    + string.Join("\n   - ", skippedEmpty)
                    + "\n\n" + plan.Count + " view(s) will be imported.",
                    "Empty worksheets skipped",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            BatchOptions = plan;
            DialogResult = DialogResult.OK;
            Close();
        }

        /// <summary>
        /// Checklist of the workbook's worksheets, with select-all/none. Kept
        /// deliberately plain: it inherits the dialog theme via Theme.Apply.
        /// </summary>
        private sealed class SheetPickerForm : Form
        {
            private readonly CheckedListBox _list;

            public List<string> SelectedSheets =>
                _list.CheckedItems.Cast<object>().Select(o => o.ToString()).ToList();

            public SheetPickerForm(List<string> sheets, List<string> preselected)
            {
                Text = "Select worksheets";
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MinimizeBox = MaximizeBox = false;
                StartPosition = FormStartPosition.CenterParent;
                ClientSize = new Size(360, 420);
                Font = new Font("Segoe UI", 9f);

                Controls.Add(new Label
                {
                    Left = 14, Top = 12, Width = 332, Height = 32,
                    Text = "Each checked worksheet becomes its own drafting view.",
                });

                _list = new CheckedListBox
                {
                    Left = 14, Top = 48, Width = 332, Height = 296,
                    CheckOnClick = true,
                    IntegralHeight = false,
                    BorderStyle = BorderStyle.FixedSingle,
                };
                foreach (var s in sheets)
                    _list.Items.Add(s, preselected.Contains(s));
                Controls.Add(_list);

                var btnAll = new Button { Left = 14, Top = 352, Width = 84, Height = 28, Text = "All" };
                var btnNone = new Button { Left = 104, Top = 352, Width = 84, Height = 28, Text = "None" };
                btnAll.Click += (s, e) => SetAll(true);
                btnNone.Click += (s, e) => SetAll(false);

                var ok = new Button
                {
                    Left = 176, Top = 384, Width = 84, Height = 28,
                    Text = "OK", DialogResult = DialogResult.OK, Tag = "primary",
                };
                var cancel = new Button
                {
                    Left = 266, Top = 384, Width = 80, Height = 28,
                    Text = "Cancel", DialogResult = DialogResult.Cancel,
                };
                Controls.Add(btnAll); Controls.Add(btnNone);
                Controls.Add(ok); Controls.Add(cancel);
                AcceptButton = ok;
                CancelButton = cancel;

                Theme.Apply(this);
                // CheckedListBox ignores the ambient theme colors, so set them directly.
                _list.BackColor = Theme.InputBg;
                _list.ForeColor = Theme.Fg;
            }

            private void SetAll(bool value)
            {
                for (int i = 0; i < _list.Items.Count; i++)
                    _list.SetItemChecked(i, value);
            }
        }

        /// <summary>Title-bar icon from the embedded ribbon PNG.</summary>
        private void ApplyIcon()
        {
            try
            {
                var asm = Assembly.GetExecutingAssembly();
                using (var s = asm.GetManifestResourceStream("ExcelScheduleImporter.Resources.icon32.png"))
                {
                    if (s == null) return;
                    using (var bmp = new Bitmap(s))
                        Icon = Icon.FromHandle(bmp.GetHicon());
                }
            }
            catch { /* cosmetic only */ }
        }
    }
}
