// Screen grid (2026-11-11): dot-matrix addressing for the desktop.
// usage:
//   grid info                       -> monitors + virtual screen + cell size
//   grid cell <A1>                  -> rect of that cell
//   grid at <x> <y>                 -> cell id containing point
//   grid overlay on|off             -> click-through grid overlay
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;

static class Grid
{
    public const int CELL = 64;

    static string ColName(int i)
    {
        string s = "";
        i++;
        while (i > 0) { int r = (i - 1) % 26; s = (char)('A' + r) + s; i = (i - 1) / 26; }
        return s;
    }

    static bool ParseCell(string id, out int col, out int row)
    {
        col = row = -1;
        if (string.IsNullOrEmpty(id)) return false;
        id = id.Trim().ToUpperInvariant();
        int i = 0;
        while (i < id.Length && id[i] >= 'A' && id[i] <= 'Z') i++;
        if (i == 0 || i == id.Length) return false;
        string letters = id.Substring(0, i);
        string digits = id.Substring(i);
        col = 0;
        foreach (char c in letters) col = col * 26 + (c - 'A' + 1);
        col -= 1;
        return int.TryParse(digits, out row) && row > 0;
    }

    [STAThread]
    static int Main(string[] args)
    {
        if (args.Length == 0) { Console.WriteLine("usage: grid info|cell <A1>|at <x> <y>|overlay on|off"); return 2; }
        Rectangle vs = SystemInformation.VirtualScreen;
        switch (args[0])
        {
            case "info":
                {
                    Console.WriteLine("virtual=" + vs.X + "," + vs.Y + "," + vs.Width + "x" + vs.Height + " cell=" + CELL);
                    int n = 0;
                    foreach (Screen s in Screen.AllScreens)
                    {
                        Rectangle b = s.Bounds;
                        int cols = (int)Math.Ceiling((double)b.Width / CELL);
                        int rows = (int)Math.Ceiling((double)b.Height / CELL);
                        Console.WriteLine("monitor[" + n + "] " + (s.Primary ? "primary " : "") + b.X + "," + b.Y + "," + b.Width + "x" + b.Height + " cols=" + cols + " rows=" + rows);
                        n++;
                    }
                    return 0;
                }
            case "cell":
                {
                    if (args.Length < 2) return 2;
                    int col, row;
                    if (!ParseCell(args[1], out col, out row)) { Console.WriteLine("bad cell id"); return 2; }
                    int x = vs.X + col * CELL;
                    int y = vs.Y + (row - 1) * CELL;
                    Console.WriteLine("CELL " + args[1].ToUpperInvariant() + " x=" + x + " y=" + y + " w=" + CELL + " h=" + CELL + " cx=" + (x + CELL / 2) + " cy=" + (y + CELL / 2));
                    return 0;
                }
            case "at":
                {
                    if (args.Length < 3) return 2;
                    int x, y;
                    if (!int.TryParse(args[1], out x) || !int.TryParse(args[2], out y)) return 2;
                    int col = (x - vs.X) / CELL;
                    int row = (y - vs.Y) / CELL + 1;
                    int cx = vs.X + col * CELL, cy = vs.Y + (row - 1) * CELL;
                    Console.WriteLine("AT " + x + "," + y + " -> " + ColName(col) + row + " rect=" + cx + "," + cy + "," + CELL + "x" + CELL);
                    return 0;
                }
            case "overlay":
                {
                    bool on = args.Length > 1 && args[1].ToLowerInvariant() == "on";
                    if (!on) { Console.WriteLine("OFF (overlay closes itself when the process exits)"); return 0; }
                    var f = new OverlayForm(vs);
                    Application.Run(f);
                    return 0;
                }
            default:
                Console.WriteLine("unknown command");
                return 2;
        }
    }

    class OverlayForm : Form
    {
        Rectangle vs;
        public OverlayForm(Rectangle v)
        {
            vs = v;
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            Bounds = v;
            TopMost = true;
            ShowInTaskbar = false;
            BackColor = Color.Black;
            Opacity = 0.25;
            SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);
        }
        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |= 0x20; // WS_EX_TRANSPARENT -> click-through
                return cp;
            }
        }
        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            using (var penMinor = new Pen(Color.FromArgb(70, 0, 255, 200), 1))
            using (var penMajor = new Pen(Color.FromArgb(140, 0, 255, 200), 2))
            using (var font = new Font("Consolas", 8f))
            using (var brush = new SolidBrush(Color.FromArgb(170, 160, 255, 240)))
            {
                int cols = (int)Math.Ceiling((double)vs.Width / CELL);
                int rows = (int)Math.Ceiling((double)vs.Height / CELL);
                for (int c = 0; c <= cols; c++)
                {
                    int x = c * CELL;
                    g.DrawLine(c % 4 == 0 ? penMajor : penMinor, x, 0, x, vs.Height);
                }
                for (int r = 0; r <= rows; r++)
                {
                    int y = r * CELL;
                    g.DrawLine(r % 4 == 0 ? penMajor : penMinor, 0, y, vs.Width, y);
                }
                for (int c = 0; c < cols; c++)
                    for (int r = 0; r < rows; r++)
                    {
                        if ((c + r) % 3 != 0) continue;
                        g.DrawString(ColName(c) + (r + 1), font, brush, c * CELL + 2, r * CELL + 1);
                    }
            }
        }
    }
}
