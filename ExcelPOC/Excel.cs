using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelPOC
{
    public class Excel
    {
        private readonly List<ISheet> excelSheets = new List<ISheet>(); 
        
        public Sheet<TData> AddSheet<TData>(string sheetname)
        {
            var sheet = new Sheet<TData>(sheetname);
            excelSheets.Add(sheet);
            return sheet;
        }

        public void SaveTo(string stream)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = document.AddWorkbookPart();
                document.WorkbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = new CustomStylesheet();
                workbookPart.Workbook = new Workbook();

                var sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());

                //var sharedStringTablePart = document.WorkbookPart.AddNewPart<SharedStringTablePart>();
                //sharedStringTablePart.SharedStringTable = new SharedStringTable();

                uint sheetIndex = 1;
                foreach (var excelSheet in excelSheets)
                {
                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet();

                    worksheetPart.Worksheet.Append(excelSheet.GetColumns());

                    var relId = document.WorkbookPart.GetIdOfPart(worksheetPart);

                    var sheet = new Sheet()
                    {
                        Name = excelSheet.SheetName,
                        Id = relId,
                        SheetId = sheetIndex++
                    };

                    sheets.Append(sheet);
                    worksheetPart.Worksheet.Append(excelSheet.GetSheetData());
                }

                document.WorkbookPart.Workbook.Save();
                document.Close();
            }
        }
    }




    



    public class CustomStylesheet : Stylesheet
    {
        
        public CustomStylesheet()
        {
            // blank font list
            var fonts = new Fonts();
            fonts.AppendChild(new Font());
            fonts.Count = 1;
            Append(fonts);

            // create fills
            var fills = new Fills();

            // create a solid blue fill
            var solidBlue = new PatternFill() { PatternType = PatternValues.Solid };
            solidBlue.ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("397FDB") }; // blue fill
            solidBlue.BackgroundColor = new BackgroundColor { Indexed = 64 };

            fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } }); // required, reserved by Excel
            fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } }); // required, reserved by Excel
            fills.AppendChild(new Fill { PatternFill = solidBlue });
            fills.Count = 3;
            Append(fills);

            // blank border list
            var borders = new Borders();
            borders.AppendChild(new Border());
            borders.AppendChild(new Border()
            {
                TopBorder = new TopBorder() { Style = BorderStyleValues.Thin },
                RightBorder = new RightBorder() { Style = BorderStyleValues.Thin },
                BottomBorder = new BottomBorder() { Style = BorderStyleValues.Thin },
                LeftBorder = new LeftBorder() { Style = BorderStyleValues.Thin }
            });
            borders.Count = 2;
            Append(borders);

            // blank cell format list
            var cellStyleFormats = new CellStyleFormats();
            cellStyleFormats.AppendChild(new CellFormat());
            cellStyleFormats.Count = 1;
            Append(cellStyleFormats);

            // cell format list
            var cellFormats = new CellFormats();
            // empty one for index 0, seems to be required
            cellFormats.AppendChild(new CellFormat());
            // cell format default with border
            cellFormats.AppendChild(new CellFormat() { FormatId = 0, FontId = 0, BorderId = 1, FillId = 0 }).AppendChild(new Alignment() { WrapText = true });
            // cell format for header (blue with border)
            cellFormats.AppendChild(new CellFormat { FormatId = 0, FontId = 0, BorderId = 1, FillId = 2, ApplyFill = true }).AppendChild(new Alignment { Horizontal = HorizontalAlignmentValues.Center });
            cellFormats.Count = 2;
            Append(cellFormats);
        }
    }
}