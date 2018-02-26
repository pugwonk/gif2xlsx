using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gif2xlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            BookWriter bw = new BookWriter("out.xlsx");
            Image img = Image.FromFile(@"..\..\giphy.gif");
            img.SelectActiveFrame(FrameDimension.Time, 0);
            Bitmap single = new Bitmap(img);
            bw.AddSheet("Frame1", single);
            bw.Save();
        }
    }
}
