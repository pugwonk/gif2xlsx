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
        int nextFreePal;
        Dictionary<System.Drawing.Color, int> palette;

        public BookWriter(string filePath)
        {
            sd = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
            workbookPart1 = sd.AddWorkbookPart();
            workbook1 = new Workbook();
            workbookPart1.Workbook = workbook1;
            workbook1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            sheets1 = new Sheets();
            workbook1.Append(sheets1);

            nextFreePal = 0;
            palette = new Dictionary<System.Drawing.Color, int>();
        }

        internal void AddSheet(string name, Bitmap img)
        {
            Sheet sheet1 = new Sheet() { Name = name, SheetId = curRid, Id = "rId" + curRid };
            sheets1.Append(sheet1);

            WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId" + curRid);

            Worksheet worksheet1 = new Worksheet();
            // Set zoom
            SheetViews sheetViews = new SheetViews();
            SheetView sheetView = new SheetView();
            sheetView.ZoomScale = 13;
            sheetView.WorkbookViewId = 0;
            sheetViews.Append(sheetView);
            worksheet1.Append(sheetViews);

            SheetData sheetData1 = new SheetData();

            List<int> paletteUsedInSheet = new List<int>();
            int runWidth = 150;
            int runHeight = 200;
            //runWidth = 10;
            //runHeight = 15;
            for (int y = 0; y < runHeight; y++)
            {
                Row row1 = new Row();
                for (int x = 0; x < runWidth; x++)
                {
                    System.Drawing.Color pix = img.GetPixel((int)((double)x/runWidth * img.Width), (int)((double)y /runHeight * img.Height));
                    if (!palette.ContainsKey(pix))
                        palette[pix] = (nextFreePal++);
                    if (!paletteUsedInSheet.Contains(palette[pix]))
                        paletteUsedInSheet.Add(palette[pix]);
                    Cell cell1 = new Cell();
                    cell1.CellReference = xyToRef(x + 1, y + 1);
                    cell1.DataType = CellValues.InlineString;
                    InlineString ils = new InlineString();
                    Text cval = new Text { Text = ColorTranslator.ToHtml(pix).Replace("#", "") };
                    //Text cval = new Text { Text = palette[pix].ToString() };
                    ils.AppendChild(cval);
                    //cell1.AppendChild(ils);
                    cell1.StyleIndex = (UInt32)(2+palette[pix]);
                    row1.Append(cell1);
                }
                sheetData1.Append(row1);
            }

            worksheet1.Append(sheetData1);
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
            var stylesPart = sd.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new Stylesheet();
            CellFormats cellFormats = new CellFormats();

            Fills fills = new Fills();
            // There are two fills that seemingly need to be always in there
            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill();
            patternFill1.PatternType = PatternValues.None;
            fill1.Append(patternFill1);
            fills.Append(fill1);
            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill();
            patternFill2.PatternType = PatternValues.Gray125;
            fill2.Append(patternFill2);
            fills.Append(fill2);
            // And two XFs
            CellFormat cellFormat1 = new CellFormat();
            cellFormat1.FillId = (UInt32)0;
            cellFormats.Append(cellFormat1);
            CellFormat cellFormat2 = new CellFormat();
            cellFormat2.FillId = (UInt32)1;
            cellFormats.Append(cellFormat2);

            foreach (var pal in palette)
            {
                CellFormat cellFormat = new CellFormat();
                cellFormat.FillId = (UInt32)(pal.Value + 2);
                cellFormats.Append(cellFormat);

                Fill fill = new Fill();
                PatternFill patternFill = new PatternFill();
                patternFill.PatternType = PatternValues.Solid;
                ForegroundColor fgc = new ForegroundColor();
                fgc.Rgb = "FF" + ColorTranslator.ToHtml(pal.Key).Replace("#", "");
                patternFill.Append(fgc);
                fill.Append(patternFill);
                fills.Append(fill);
            }

            // Add dummy borders and fonts, which we need for the cellXfs to work
            Borders borders = new Borders();
            Border border = new Border();
            borders.Append(border);
            Fonts fonts = new Fonts();
            DocumentFormat.OpenXml.Spreadsheet.Font font = new DocumentFormat.OpenXml.Spreadsheet.Font();
            fonts.Append(font);

            // Add the other things we made to the stylesheet
            stylesPart.Stylesheet.Append(fonts);
            stylesPart.Stylesheet.Append(fills);
            stylesPart.Stylesheet.Append(borders);
            stylesPart.Stylesheet.Append(cellFormats);
            sd.Save();
            sd.Close();
        }
    }
}