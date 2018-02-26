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

            FrameDimension dimension = new FrameDimension(img.FrameDimensionsList[0]);
            for (int i = 0; i < img.GetFrameCount(dimension); i++)
            {
                img.SelectActiveFrame(FrameDimension.Time, i);
                Bitmap single = new Bitmap(img);
                bw.AddSheet("Frame" + i.ToString(), single);
            }

            bw.Save();
        }
    }
}
