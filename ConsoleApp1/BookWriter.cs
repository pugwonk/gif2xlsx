using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.Drawing;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Gif2xlsx
{
    internal class BookWriter
    {
        SpreadsheetDocument sd;
        Workbook workbook1;
        WorkbookPart workbookPart1;
        Sheets sheets1;
        UInt32Value curRid = 1;

        public BookWriter(string filePath)
        {
            sd = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
            workbookPart1 = sd.AddWorkbookPart();
            workbook1 = new Workbook();
            workbookPart1.Workbook = workbook1;
            workbook1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            sheets1 = new Sheets();
            workbook1.Append(sheets1);

            var stylesPart = sd.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new Stylesheet();
            DifferentialFormats dxfs = new DifferentialFormats();
            DifferentialFormat dxf = new DifferentialFormat();
            Fill f = new Fill();
            PatternFill p = new PatternFill();
            BackgroundColor b = new BackgroundColor();
            b.Rgb = "FFFFFF00";
            p.Append(b);
            f.Append(p);
            dxf.Append(f);
            dxfs.Append(dxf);

            stylesPart.Stylesheet.DifferentialFormats = dxfs;
        }

        internal void AddSheet(string name, Bitmap img)
        {
            Sheet sheet1 = new Sheet() { Name = name, SheetId = curRid, Id = "rId" + curRid };
            sheets1.Append(sheet1);

            WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId" + curRid);

            Worksheet worksheet1 = new Worksheet();
            SheetData sheetData1 = new SheetData();

            int nextFreePal = 0;
            Dictionary<System.Drawing.Color, int> palette = new Dictionary<System.Drawing.Color, int>();
            for (int y = 0; y < img.Height/10; y++)
            {
                Row row1 = new Row();
                List<float> brights = new List<float>();
                for (int x = 0; x < img.Width/10; x++)
                {
                    System.Drawing.Color pix = img.GetPixel(x, y);
                    if (!palette.ContainsKey(pix))
                        palette[pix] = (nextFreePal++);
                    brights.Add(palette[pix]);
                    Cell cell1 = new Cell();
                    cell1.CellReference = xyToRef(x + 1, y + 1);
                    cell1.DataType = CellValues.Number;
                    cell1.CellValue = new CellValue(XmlConvert.ToString(palette[pix]));
                    row1.Append(cell1);
                }
                sheetData1.Append(row1);
            }

            worksheet1.Append(sheetData1);
            // Add the CF stuff
            ConditionalFormatting cf = new ConditionalFormatting();
            cf.SetAttribute(new OpenXmlAttribute("sqref", "", "A1:C4"));
            ConditionalFormattingRule cfr = new ConditionalFormattingRule();
            cfr.Type = ConditionalFormatValues.Expression;
            cfr.SetAttribute(new OpenXmlAttribute("dxfId", "", "1"));
            cfr.Priority = 1;
            Formula f = new Formula("A1=3");
            cfr.Append(f);
            cf.Append(cfr);
            worksheet1.Append(cf);
            worksheetPart1.Worksheet = worksheet1;

            curRid++;
        }

        private string GetColumnName(int start)
        {
            string result = "";
            int remainder;

            while (start > 0)
            {
                start = System.Math.DivRem(start, 26, out remainder);
                if (remainder == 0)
                {
                    result = "Z" + result;
                    start -= 1;
                }
                else
                {
                    result = (char)(64 + remainder) + result;
                }
            }

            return result;
        }

        private StringValue xyToRef(int x, int y)
        {
            return GetColumnName(x) + y;
        }

        internal void Save()
        {
            sd.Save();
            sd.Close();
        }
    }
}