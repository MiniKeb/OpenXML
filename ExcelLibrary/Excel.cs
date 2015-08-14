using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelLibrary
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

        public void SaveTo(Stream stream)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                ProcessDocument(document);
                document.Close();
            }
        }

        public void SaveTo(string filename)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook))
            {
                ProcessDocument(document);
                document.Close();
            }
        }

        private void ProcessDocument(SpreadsheetDocument document)
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
        }
    }
}