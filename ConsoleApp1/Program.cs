using System;
using System.Collections.Generic;
using System.Diagnostics;
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
            if (args.Count() == 0)
            {
                // Help
                Console.WriteLine("Usage:");
                Console.WriteLine();
                Console.WriteLine(" gif2xlsx myfile.gif");
                Console.WriteLine();
                Console.WriteLine("Output will be saved as out.xlsx in current folder. If there is any output. If not you can have a full refund.");
            }
            else
            {
                BookWriter bw = new BookWriter("out.xlsx");
                Image img = Image.FromFile(args[0]);

                FrameDimension dimension = new FrameDimension(img.FrameDimensionsList[0]);
                int doFrames = img.GetFrameCount(dimension);
                //doFrames = 1;
                for (int i = 0; i < doFrames; i++)
                {
                    Console.WriteLine("Processing frame " + (i + 1).ToString() + " of " + doFrames.ToString() + "...");
                    img.SelectActiveFrame(FrameDimension.Time, i);
                    Bitmap single = new Bitmap(img);
                    bw.AddSheet("Frame" + i.ToString(), single);
                }

                Console.WriteLine("Saving spreadsheet...");
                bw.Save();
                Console.WriteLine("Done. Opening spreadsheet...");
                Process.Start("out.xlsx");
            }
        }
    }
}
