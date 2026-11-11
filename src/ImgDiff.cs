// 图片区块差异（2026-11-11）：比较两张同尺寸截图的指定区域，输出差异百分比
// 用法：imgdiff.exe a.png b.png [x y w h]
using System;
using System.Drawing;
using System.IO;

class ImgDiff
{
    static int Main(string[] args)
    {
        if (args.Length < 2) { Console.WriteLine("usage: imgdiff a.png b.png [x y w h]"); return 2; }
        using (var a = new Bitmap(args[0]))
        using (var b = new Bitmap(args[1]))
        {
            int x = 0, y = 0, w = Math.Min(a.Width, b.Width), h = Math.Min(a.Height, b.Height);
            if (args.Length >= 6) { x = int.Parse(args[2]); y = int.Parse(args[3]); w = int.Parse(args[4]); h = int.Parse(args[5]); }
            if (x + w > Math.Min(a.Width, b.Width)) w = Math.Min(a.Width, b.Width) - x;
            if (y + h > Math.Min(a.Height, b.Height)) h = Math.Min(a.Height, b.Height) - y;
            long diff = 0, total = 0;
            for (int j = y; j < y + h; j += 2)
            {
                for (int i = x; i < x + w; i += 2)
                {
                    var pa = a.GetPixel(i, j); var pb = b.GetPixel(i, j);
                    total++;
                    if (Math.Abs(pa.R - pb.R) + Math.Abs(pa.G - pb.G) + Math.Abs(pa.B - pb.B) > 30) diff++;
                }
            }
            double pct = total == 0 ? 0 : diff * 100.0 / total;
            Console.WriteLine("DIFF " + pct.ToString("F2") + "% region=" + x + "," + y + "," + w + "x" + h + " changed=" + (pct >= 1.0));
            return 0;
        }
    }
}
