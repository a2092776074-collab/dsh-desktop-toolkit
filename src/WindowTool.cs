// window.exe — 窗口枚举与取矩形（2026-11-11）
// usage: window list | window rect <title-substring>
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

class WindowTool
{
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr p);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] static extern int GetWindowTextLength(IntPtr h);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);

    delegate bool EnumProc(IntPtr h, IntPtr l);
    struct RECT { public int L, T, R, B; }

    class WinInfo { public IntPtr H; public string Title; public uint Pid; public RECT R; }
    static List<WinInfo> All()
    {
        var list = new List<WinInfo>();
        EnumWindows((h, l) =>
        {
            if (!IsWindowVisible(h)) return true;
            int len = GetWindowTextLength(h);
            if (len == 0) return true;
            var sb = new StringBuilder(len + 1);
            GetWindowText(h, sb, sb.Capacity);
            RECT r; GetWindowRect(h, out r);
            if (r.R - r.L <= 0 || r.B - r.T <= 0) return true;
            uint pid; GetWindowThreadProcessId(h, out pid);
            list.Add(new WinInfo { H = h, Title = sb.ToString(), Pid = pid, R = r });
            return true;
        }, IntPtr.Zero);
        return list;
    }

    static string Esc(string s) { return s == null ? "" : s.Replace("\\", "\\\\").Replace("\"", "\\\""); }

    static int Main(string[] args)
    {
        var wins = All();
        if (args.Length == 0 || args[0] == "list")
        {
            Console.WriteLine("{\"count\":" + wins.Count + ",\"windows\":[");
            for (int i = 0; i < wins.Count; i++)
            {
                var w = wins[i];
                Console.WriteLine("{\"title\":\"" + Esc(w.Title) + "\",\"pid\":" + w.Pid + ",\"x\":" + w.R.L + ",\"y\":" + w.R.T + ",\"w\":" + (w.R.R - w.R.L) + ",\"h\":" + (w.R.B - w.R.T) + "}" + (i < wins.Count - 1 ? "," : ""));
            }
            Console.WriteLine("]}");
            return 0;
        }
        if (args[0] == "rect")
        {
            if (args.Length < 2) { Console.WriteLine("{\"ok\":false,\"error\":\"need title\"}"); return 2; }
            string kw = args[1].ToLowerInvariant();
            foreach (var w in wins)
            {
                if (w.Title.ToLowerInvariant().Contains(kw))
                {
                    Console.WriteLine("{\"ok\":true,\"title\":\"" + Esc(w.Title) + "\",\"pid\":" + w.Pid + ",\"x\":" + w.R.L + ",\"y\":" + w.R.T + ",\"w\":" + (w.R.R - w.R.L) + ",\"h\":" + (w.R.B - w.R.T) + "}");
                    return 0;
                }
            }
            Console.WriteLine("{\"ok\":false,\"error\":\"no window matched\"}");
            return 1;
        }
        if (args[0] == "rectpid")
        {
            uint want = uint.Parse(args[1]);
            foreach (var w in wins)
            {
                if (w.Pid == want)
                {
                    Console.WriteLine("{\"ok\":true,\"title\":\"" + Esc(w.Title) + "\",\"pid\":" + w.Pid + ",\"x\":" + w.R.L + ",\"y\":" + w.R.T + ",\"w\":" + (w.R.R - w.R.L) + ",\"h\":" + (w.R.B - w.R.T) + "}");
                    return 0;
                }
            }
            Console.WriteLine("{\"ok\":false,\"error\":\"pid not found\"}");
            return 1;
        }
        Console.WriteLine("usage: window list | window rect <title-substring> | window rectpid <pid>");
        return 2;
    }
}
