using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Windows.Forms;

namespace ExcelScheduleImporter.UI
{
    /// <summary>
    /// Fully owner-drawn dark combo box (DropDownList). Paints its own border,
    /// arrow and item list so no system-gray chrome shows in a dark theme.
    /// </summary>
    internal class DarkComboBox : ComboBox
    {
        private const int WM_PAINT = 0x000F;
        private const int ArrowW = 20;

        public DarkComboBox()
        {
            DrawMode = DrawMode.OwnerDrawFixed;
            FlatStyle = FlatStyle.Flat;
            DropDownStyle = ComboBoxStyle.DropDownList;
            BackColor = Theme.InputBg;
            ForeColor = Theme.Fg;
            ItemHeight = 20;
        }

        protected override void OnDrawItem(DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            bool sel = (e.State & DrawItemState.Selected) != 0;
            using (var bg = new SolidBrush(sel ? Theme.Accent : Theme.InputBg))
                e.Graphics.FillRectangle(bg, e.Bounds);
            var r = e.Bounds; r.X += 4;
            TextRenderer.DrawText(e.Graphics, Items[e.Index].ToString(), Font, r,
                sel ? Color.White : Theme.Fg,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            if (m.Msg != WM_PAINT) return;

            using (var g = Graphics.FromHwnd(Handle))
            {
                var r = ClientRectangle;
                // repaint the arrow well and the border over the system chrome
                using (var bg = new SolidBrush(Theme.InputBg))
                    g.FillRectangle(bg, r.Right - ArrowW, r.Top, ArrowW, r.Height);
                using (var pen = new Pen(Theme.Border))
                    g.DrawRectangle(pen, 0, 0, r.Width - 1, r.Height - 1);

                g.SmoothingMode = SmoothingMode.AntiAlias;
                int cx = r.Right - ArrowW / 2 - 1;
                int cy = r.Height / 2 - 1;
                using (var b = new SolidBrush(Enabled ? Theme.Fg : Theme.FgDim))
                    g.FillPolygon(b, new[]
                    {
                        new Point(cx - 4, cy - 2), new Point(cx + 4, cy - 2), new Point(cx, cy + 3)
                    });
            }
        }
    }

    /// <summary>
    /// Custom dark numeric spinner. Composite (borderless TextBox + painted
    /// up/down arrows) exposing the subset of the NumericUpDown API the dialog
    /// uses, so it can be a drop-in replacement while looking themed.
    /// </summary>
    internal class DarkNumeric : Control
    {
        private readonly TextBox _box;
        private decimal _value, _min = 0m, _max = 100m, _inc = 1m;
        private int _decimals;
        private Rectangle _upRect, _downRect;
        private const int ArrowW = 17;

        public DarkNumeric()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer
                     | ControlStyles.ResizeRedraw | ControlStyles.UserPaint, true);

            // IMPORTANT: _box must exist before ANY property set that can trigger
            // layout (Height/Width/Bounds). Setting Height below fires OnLayout
            // synchronously (SetBoundsCore -> UpdateBounds -> OnResize -> PerformLayout),
            // and OnLayout touches _box - so _box has to be built first, or that
            // first layout pass hits a null reference.
            _box = new TextBox
            {
                BorderStyle = BorderStyle.None,
                BackColor = Theme.InputBg,
                ForeColor = Theme.Fg,
            };
            _box.Leave += (s, e) => Commit();
            _box.KeyPress += (s, e) =>
            {
                if (e.KeyChar == (char)Keys.Return) { Commit(); e.Handled = true; }
            };
            Controls.Add(_box);

            BackColor = Theme.InputBg;
            ForeColor = Theme.Fg;
            Height = 24;   // safe now: OnLayout can find _box

            SyncText();
        }

        // ── NumericUpDown-compatible surface ─────────────────────────────────
        public decimal Minimum { get => _min; set { _min = value; Value = _value; } }
        public decimal Maximum { get => _max; set { _max = value; Value = _value; } }
        public decimal Increment { get => _inc; set => _inc = value; }
        public int DecimalPlaces { get => _decimals; set { _decimals = value; SyncText(); } }

        public decimal Value
        {
            get => _value;
            set
            {
                decimal v = Math.Max(_min, Math.Min(_max, value));
                if (v != _value) { _value = v; SyncText(); }
                else { _value = v; }
            }
        }

        public new bool Enabled
        {
            get => base.Enabled;
            set
            {
                base.Enabled = value;
                if (_box == null) return;   // guard: base class can set this before our ctor finishes
                _box.Enabled = value;
                _box.ForeColor = value ? Theme.Fg : Theme.FgDim;
                Invalidate();
            }
        }

        private void SyncText()
        {
            if (_box == null) return;
            string t = _value.ToString("F" + _decimals, CultureInfo.CurrentCulture);
            if (_box.Text != t) _box.Text = t;
        }

        private void Commit()
        {
            if (decimal.TryParse(_box.Text, NumberStyles.Any, CultureInfo.CurrentCulture, out var v))
                Value = v;
            SyncText();
        }

        protected override void OnLayout(LayoutEventArgs e)
        {
            base.OnLayout(e);
            if (_box == null) return;   // guard: layout can fire before the ctor finishes building _box
            _box.SetBounds(5, (Height - _box.PreferredHeight) / 2 + 1,
                           Width - ArrowW - 8, _box.PreferredHeight);
            _upRect = new Rectangle(Width - ArrowW, 1, ArrowW - 1, Height / 2 - 1);
            _downRect = new Rectangle(Width - ArrowW, Height / 2, ArrowW - 1, Height / 2 - 1);
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (!base.Enabled) return;
            if (_upRect.Contains(e.Location)) Value = _value + _inc;
            else if (_downRect.Contains(e.Location)) Value = _value - _inc;
        }

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            base.OnMouseWheel(e);
            if (base.Enabled) Value = _value + (e.Delta > 0 ? _inc : -_inc);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            using (var bg = new SolidBrush(Theme.InputBg))
                g.FillRectangle(bg, ClientRectangle);

            // separator before the arrow column + outer border
            using (var pen = new Pen(Theme.Border))
            {
                g.DrawLine(pen, Width - ArrowW - 1, 1, Width - ArrowW - 1, Height - 2);
                g.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
            }

            g.SmoothingMode = SmoothingMode.AntiAlias;
            Color arrow = base.Enabled ? Theme.Fg : Theme.FgDim;
            int cx = Width - ArrowW / 2 - 1;
            using (var b = new SolidBrush(arrow))
            {
                int uy = _upRect.Top + _upRect.Height / 2;
                g.FillPolygon(b, new[] { new Point(cx - 3, uy + 2), new Point(cx + 3, uy + 2), new Point(cx, uy - 2) });
                int dy = _downRect.Top + _downRect.Height / 2;
                g.FillPolygon(b, new[] { new Point(cx - 3, dy - 2), new Point(cx + 3, dy - 2), new Point(cx, dy + 2) });
            }
        }
    }
}
