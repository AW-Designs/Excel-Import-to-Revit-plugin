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
        private Button _btnUpdateChecked;
        private Button _btnUpdateAll;
        private Label _lblUpdateStatus;
        private readonly ToolTip _tip = new ToolTip { AutoPopDelay = 12000 };

        private readonly HashSet<string> _existingViewNames;
        private string _detectedRange;   // last auto-detected range for the selected sheet

        private readonly Func<List<ElementId>, string> _updateAction;
        private readonly Func<List<ManageRow>> _refreshRows;

        public ImportOptions Options { get; private set; }

        /// <summary>
        /// The workbook, parsed once when the user picks a file. The command reuses
        /// this instance for the actual import, so the file is never parsed twice.
        /// </summary>
        public ExcelReader Reader { get; private set; }

        public ImportForm(IEnumerable<string> existingDraftingViewNames,
                          List<ManageRow> scheduleRows,
                          Func<List<ElementId>, string> updateAction,
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
            _cboSheet = new DarkComboBox { Left = pc, Top = y - 3, Width = pw, Height = 24, Enabled = false };
            _cboSheet.SelectedIndexChanged += OnSheetChanged;
            pnlImport.Controls.Add(_cboSheet);
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
            foreach (var card in _flow.Controls.OfType<ScheduleCard>())
                card.Visible = q.Length == 0 || card.Matches(q);
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

                string summary = _updateAction(ids);

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

            public bool CardChecked => _chk.Checked;

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
                    Reader?.Dispose();
                    Reader = new ExcelReader(dlg.FileName);   // parse once, reuse everywhere
                    var sheets = Reader.GetSheetNames();
                    _txtFile.Text = dlg.FileName;
                    _cboSheet.Items.Clear();
                    foreach (var s in sheets) _cboSheet.Items.Add(s);
                    _cboSheet.Enabled = sheets.Count > 0;
                    if (sheets.Count > 0) _cboSheet.SelectedIndex = 0;
                    _lblStatus.Text = "";
                }
                catch (Exception ex)
                {
                    _lblStatus.Text = "Cannot read file: " + ex.Message;
                }
                finally { Cursor = Cursors.Default; }
            }
        }

        private void OnSheetChanged(object sender, EventArgs e)
        {
            if (_cboSheet.SelectedItem == null) return;
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
            { _lblStatus.Text = "Select an Excel file first."; return; }
            if (_cboSheet.SelectedItem == null)
            { _lblStatus.Text = "Select a worksheet."; return; }
            if (string.IsNullOrWhiteSpace(_txtRange.Text))
            { _lblStatus.Text = "Enter a cell range (e.g. A1:F42)."; return; }

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
