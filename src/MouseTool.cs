using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

class MouseTool
{
    [DllImport("user32.dll")] static extern bool SetCursorPos(int X, int Y);
    [DllImport("user32.dll")] static extern void mouse_event(uint f, uint dx, uint dy, uint d, UIntPtr e);
    const uint LEFTDOWN = 0x0002, LEFTUP = 0x0004;

    static void Shot(string path)
    {
        var b = SystemInformation.VirtualScreen;
        using (var bmp = new Bitmap(b.Width, b.Height))
        using (var g = Graphics.FromImage(bmp))
        {
            g.CopyFromScreen(b.Left, b.Top, 0, 0, bmp.Size);
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            bmp.Save(path, ImageFormat.Png);
        }
    }

    static void To(int x, int y) { SetCursorPos(x, y); Thread.Sleep(4); }

    static int Main(string[] args)
    {
        if (args.Length == 0) { Console.WriteLine("usage: mouse click x y | drag x1 y1 x2 y2 [--steps N] [--delay ms] [--shot dir] [--mid] | move x y"); return 2; }
        string action = args[0];
        int steps = 20, delay = 15;
        string shotDir = null;
        bool mid = false;
        var nums = new System.Collections.Generic.List<int>();
        for (int i = 1; i < args.Length; i++)
        {
            if (args[i] == "--steps") { steps = int.Parse(args[++i]); }
            else if (args[i] == "--delay") { delay = int.Parse(args[++i]); }
            else if (args[i] == "--shot") { shotDir = args[++i]; }
            else if (args[i] == "--mid") { mid = true; }
            else { int v; if (int.TryParse(args[i], out v)) nums.Add(v); }
        }
        string before = null, after = null;
        if (shotDir != null)
        {
            Directory.CreateDirectory(shotDir);
            before = Path.Combine(shotDir, "before-" + DateTime.Now.ToString("HHmmssfff") + ".png");
            Shot(before);
        }
        switch (action)
        {
            case "move":
                To(nums[0], nums[1]);
                break;
            case "click":
                To(nums[0], nums[1]); Thread.Sleep(80);
                mouse_event(LEFTDOWN, 0, 0, 0, UIntPtr.Zero); Thread.Sleep(60);
                mouse_event(LEFTUP, 0, 0, 0, UIntPtr.Zero);
                break;
            case "drag":
                {
                    int x1 = nums[0], y1 = nums[1], x2 = nums[2], y2 = nums[3];
                    To(x1, y1); Thread.Sleep(120);
                    mouse_event(LEFTDOWN, 0, 0, 0, UIntPtr.Zero); Thread.Sleep(120);
                    for (int i = 1; i <= steps; i++)
                    {
                        int x = x1 + (x2 - x1) * i / steps;
                        int y = y1 + (y2 - y1) * i / steps;
                        To(x, y); Thread.Sleep(delay);
                        if (mid && shotDir != null && i == (int)(steps * 0.8))
                        {
                            var mp = Path.Combine(shotDir, "mid-" + DateTime.Now.ToString("HHmmssfff") + ".png");
                            Shot(mp);
                            Console.WriteLine("MID " + Path.GetFileName(mp));
                        }
                    }
                    Thread.Sleep(150);
                    mouse_event(LEFTUP, 0, 0, 0, UIntPtr.Zero);
                    break;
                }
            default:
                Console.WriteLine("unknown action " + action); return 2;
        }
        if (shotDir != null)
        {
            Thread.Sleep(250);
            after = Path.Combine(shotDir, "after-" + DateTime.Now.ToString("HHmmssfff") + ".png");
            Shot(after);
            bool changed = true;
            using (var a = new FileStream(before, FileMode.Open, FileAccess.Read))
            using (var b = new FileStream(after, FileMode.Open, FileAccess.Read))
            {
                changed = a.Length != b.Length;
                if (!changed)
                {
                    var ha = System.Security.Cryptography.MD5.Create().ComputeHash(a);
                    var hb = System.Security.Cryptography.MD5.Create().ComputeHash(b);
                    changed = !Convert.ToBase64String(ha).Equals(Convert.ToBase64String(hb));
                }
            }
            Console.WriteLine("OK " + action + " before=" + Path.GetFileName(before) + " after=" + Path.GetFileName(after) + " changed=" + changed);
        }
        else Console.WriteLine("OK " + action);
        return 0;
    }
}
