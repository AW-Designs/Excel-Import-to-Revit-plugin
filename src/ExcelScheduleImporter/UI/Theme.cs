using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelScheduleImporter.UI
{
    /// <summary>
    /// One palette for the whole add-in. Follows Revit's current UI theme
    /// (dark/light, Revit 2024+) so dialogs blend in instead of flashing a
    /// bright gray WinForms window inside a dark Revit.
    /// </summary>
    internal static class Theme
    {
        public static bool Dark { get; private set; }

        public static Color WindowBg, PanelBg, InputBg, CardBg, CardHover, ButtonBg,
                            Fg, FgDim, Border, Accent, Success, Warning, Danger;

        static Theme()
        {
            Dark = DetectRevitDarkTheme();
            if (Dark) SetDark(); else SetLight();
        }

        private static bool DetectRevitDarkTheme()
        {
            try
            {
                return Autodesk.Revit.UI.UIThemeManager.CurrentTheme
                       == Autodesk.Revit.UI.UITheme.Dark;
            }
            catch { return true; }   // no theme API -> assume dark (matches user's setup)
        }

        private static void SetDark()
        {
            // Slightly blue-tinted charcoal family (modern, not flat neutral gray)
            WindowBg  = FromHex("#1F2124");   // window base
            PanelBg   = FromHex("#282B30");   // section containers (raised)
            InputBg   = FromHex("#17181B");   // fields (recessed)
            CardBg    = FromHex("#2A2D33");
            CardHover = FromHex("#343841");
            ButtonBg  = FromHex("#31353C");
            Fg        = FromHex("#EDEEF0");
            FgDim     = FromHex("#8B9099");
            Border    = FromHex("#3A3E45");
            Accent    = FromHex("#3B93E6");   // clean modern blue
            Success   = FromHex("#4FB36A");
            Warning   = FromHex("#E7A13B");
            Danger    = FromHex("#EA5F5A");
        }

        private static void SetLight()
        {
            WindowBg  = FromHex("#F2F2F2");
            PanelBg   = FromHex("#FAFAFA");
            InputBg   = Color.White;
            CardBg    = Color.White;
            CardHover = FromHex("#EAF4F9");
            ButtonBg  = FromHex("#E4E4E6");
            Fg        = FromHex("#1E1E1E");
            FgDim     = FromHex("#6A6A70");
            Border    = FromHex("#C4C4C8");
            Accent    = FromHex("#1F87AD");
            Success   = FromHex("#3B8C3B");
            Warning   = FromHex("#B87413");
            Danger    = FromHex("#C0392B");
        }

        private static Color FromHex(string hex)
            => ColorTranslator.FromHtml(hex);

        // ── application ─────────────────────────────────────────────────────

        /// <summary>Style a form and everything currently on it.</summary>
        public static void Apply(Form form)
        {
            form.BackColor = WindowBg;
            form.ForeColor = Fg;
            ApplyRecursive(form);

            if (form.IsHandleCreated) EnableImmersiveTitleBar(form.Handle);
            else form.HandleCreated += (s, e) => EnableImmersiveTitleBar(form.Handle);
        }

        private static void ApplyRecursive(Control root)
        {
            if (root is DarkNumeric) return;   // self-styled; leave its inner TextBox alone

            foreach (Control c in root.Controls)
            {
                switch (c)
                {
                    case TextBox t:
                        t.BackColor = InputBg;
                        t.ForeColor = Fg;
                        t.BorderStyle = BorderStyle.FixedSingle;
                        break;

                    case ComboBox cb:
                        cb.BackColor = InputBg;
                        cb.ForeColor = Fg;
                        cb.FlatStyle = FlatStyle.Flat;
                        break;

                    case NumericUpDown n:
                        n.BackColor = InputBg;
                        n.ForeColor = Fg;
                        n.BorderStyle = BorderStyle.FixedSingle;
                        break;

                    case Button b:
                        StyleButton(b, Equals(b.Tag, "primary"));
                        break;

                    case CheckBox chk:
                        chk.ForeColor = Fg;
                        chk.FlatStyle = FlatStyle.Flat;   // dark glyph instead of system white
                        chk.FlatAppearance.BorderSize = 0;
                        break;

                    case Label lbl:
                        // remap the stock construction colors onto the palette
                        if (lbl.ForeColor == Color.DimGray)        lbl.ForeColor = FgDim;
                        else if (lbl.ForeColor == Color.Firebrick) lbl.ForeColor = Danger;
                        else if (lbl.ForeColor == Color.DarkGreen) lbl.ForeColor = Success;
                        else if (lbl.ForeColor == SystemColors.ControlText) lbl.ForeColor = Fg;
                        break;
                }
                if (c.HasChildren) ApplyRecursive(c);
            }
        }

        public static void StyleButton(Button b, bool primary)
        {
            b.FlatStyle = FlatStyle.Flat;
            b.UseVisualStyleBackColor = false;
            b.Cursor = Cursors.Hand;
            b.FlatAppearance.BorderSize = 0;
            if (primary)
            {
                b.BackColor = Accent;
                b.ForeColor = Color.White;
                b.Font = new Font("Segoe UI Semibold", b.Font.Size);
                b.FlatAppearance.MouseOverBackColor = ControlPaint.Light(Accent, 0.18f);
                b.FlatAppearance.MouseDownBackColor = ControlPaint.Dark(Accent, 0.05f);
            }
            else
            {
                b.BackColor = ButtonBg;
                b.ForeColor = Fg;
                b.FlatAppearance.MouseOverBackColor = CardHover;
                b.FlatAppearance.MouseDownBackColor = WindowBg;
            }
            RoundControl(b, 6);
        }

        // ── rounded corners ─────────────────────────────────────────────────

        public static GraphicsPath RoundedPath(Rectangle r, int radius)
        {
            int d = radius * 2;
            var p = new GraphicsPath();
            if (radius <= 0) { p.AddRectangle(r); return p; }
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }

        /// <summary>Clip a control to a rounded rectangle; re-clips on resize.</summary>
        public static void RoundControl(Control c, int radius)
        {
            void Apply(object s, EventArgs e)
            {
                if (c.Width <= 0 || c.Height <= 0) return;
                using (var path = RoundedPath(new Rectangle(0, 0, c.Width, c.Height), radius))
                    c.Region = new Region(path);
            }
            Apply(c, EventArgs.Empty);
            c.Resize += Apply;
        }

        // ── native touches ──────────────────────────────────────────────────

        [DllImport("dwmapi.dll", PreserveSig = true)]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);

        /// <summary>Dark window title bar (Windows 10 1809+ / Windows 11).</summary>
        public static void EnableImmersiveTitleBar(IntPtr handle)
        {
            try
            {
                int dark = Dark ? 1 : 0;
                // 20 = DWMWA_USE_IMMERSIVE_DARK_MODE (19 on very old Win10 builds)
                if (DwmSetWindowAttribute(handle, 20, ref dark, 4) != 0)
                    DwmSetWindowAttribute(handle, 19, ref dark, 4);
            }
            catch { /* cosmetic only */ }
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, string lParam);

        /// <summary>Gray placeholder text inside an empty TextBox (EM_SETCUEBANNER).</summary>
        public static void SetPlaceholder(TextBox box, string text)
        {
            void Set() { try { SendMessage(box.Handle, 0x1501, (IntPtr)1, text); } catch { } }
            if (box.IsHandleCreated) Set();
            else box.HandleCreated += (s, e) => Set();
        }
    }
}
