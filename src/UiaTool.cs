using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Automation;
using System.Runtime.InteropServices;

class UiaTool
{
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);

    static int ForegroundPid()
    {
        uint pid;
        GetWindowThreadProcessId(GetForegroundWindow(), out pid);
        return (int)pid;
    }

    static string Esc(string s)
    {
        if (s == null) return "";
        return s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r", " ").Replace("\n", " ");
    }

    static List<AutomationElement> Collect(AutomationElement root, int maxDepth, int maxNodes)
    {
        var list = new List<AutomationElement>();
        var q = new Queue<Tuple<AutomationElement, int>>();
        q.Enqueue(Tuple.Create(root, 0));
        while (q.Count > 0 && list.Count < maxNodes)
        {
            var cur = q.Dequeue();
            var el = cur.Item1;
            int d = cur.Item2;
            list.Add(el);
            if (d >= maxDepth) continue;
            AutomationElementCollection kids = null;
            try { kids = el.FindAll(TreeScope.Children, Condition.TrueCondition); } catch { continue; }
            foreach (AutomationElement k in kids)
            {
                if (k == null) continue;
                q.Enqueue(Tuple.Create(k, d + 1));
                if (list.Count + q.Count > maxNodes * 2) break;
            }
        }
        return list;
    }

    static void Main(string[] args)
    {
        try
        {
            if (args.Length < 2) { Console.WriteLine("{\"ok\":false,\"error\":\"usage: tree|find <pid> [name] [maxDepth] [maxNodes]\"}"); return; }
            string cmd = args[0];
            int pid = args[1] == "fg" ? ForegroundPid() : int.Parse(args[1]);
            string want = args.Length > 2 ? args[2] : "";
            int maxDepth = args.Length > 3 ? int.Parse(args[3]) : 8;
            int maxNodes = args.Length > 4 ? int.Parse(args[4]) : 400;

            var cond = new PropertyCondition(AutomationElement.ProcessIdProperty, pid);
            AutomationElement win = AutomationElement.RootElement.FindFirst(TreeScope.Children, cond);
            if (win == null) { Console.WriteLine("{\"ok\":false,\"error\":\"window not found\"}"); return; }

            var nodes = Collect(win, maxDepth, maxNodes);
            var sb = new StringBuilder();
            int hit = 0;
            bool truncated = nodes.Count >= maxNodes;
            sb.Append("{\"ok\":true,\"pid\":" + pid + ",\"scanned\":" + nodes.Count + ",\"truncated\":" + (truncated ? "true" : "false") + ",\"items\":[");
            string w = want.Replace(" ", "").ToLower();
            foreach (AutomationElement el in nodes)
            {
                string name = ""; string type = ""; System.Windows.Rect r = System.Windows.Rect.Empty;
                try
                {
                    name = el.Current.Name ?? "";
                    type = el.Current.ControlType.ProgrammaticName.Replace("ControlType.", "");
                    if (!el.Current.IsOffscreen) r = el.Current.BoundingRectangle;
                }
                catch { continue; }
                if (name.Length == 0 || r.IsEmpty || r.Width < 2 || r.Height < 2) continue;
                if (w.Length > 0 && !name.Replace(" ", "").ToLower().Contains(w)) continue;
                if (hit > 0) sb.Append(",");
                sb.Append("{\"name\":\"" + Esc(name) + "\",\"type\":\"" + Esc(type) + "\",");
                sb.Append("\"x\":" + (int)r.X + ",\"y\":" + (int)r.Y + ",\"w\":" + (int)r.Width + ",\"h\":" + (int)r.Height + "}");
                hit++;
                if (hit >= 80) break;
            }
            sb.Append("],\"hits\":" + hit + "}");
            Console.WriteLine(sb.ToString());
        }
        catch (Exception ex)
        {
            Console.WriteLine("{\"ok\":false,\"error\":\"" + Esc(ex.Message) + "\"}");
        }
    }
}
