using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelPOC.CodeProjectSample
{
    public class ExcelHelper
    {
        /// <summary>
        /// Write excel file of a list of object as T
        /// Assume that maximum of 24 columns 
        /// </summary>
        /// <typeparam name="T">Object type to pass in</typeparam>
        /// <param name="fileName">Full path of the file name of excel spreadsheet</param>
        /// <param name="objects">list of the object type</param>
        /// <param name="sheetName">Sheet names of Excel File</param>
        /// <param name="headerNames">Header names of the object</param>
        public void Create<T>(
            string fileName,
            List<T> objects,
            string sheetName,
            List<string> headerNames)
        {
            //Open the copied template workbook. 
            using (SpreadsheetDocument document =
                SpreadsheetDocument.Create(fileName,
                    SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                // Create Styles and Insert into Workbook
                var stylesPart =
                    document.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                Stylesheet styles = new CustomStylesheet();
                styles.Save(stylesPart);
                string relId = workbookPart.GetIdOfPart(worksheetPart);
                var worksheet = new Worksheet();
                int numCols = headerNames.Count;
                var columns = new Columns();
                for (int col = 0; col < numCols; col++)
                {
                    int width = headerNames[col].Length + 5;
                    Column c = new CustomColumn((UInt32)col + 1,
                        (UInt32)numCols + 1, width);
                    columns.Append(c);
                }
                worksheet.Append(columns);
                var sheets = new Sheets();
                var sheet = new Sheet { Name = sheetName, SheetId = 1, Id = relId };
                sheets.Append(sheet);
                var workbook = new Workbook();
                workbook.Append(new FileVersion
                {
                    ApplicationName = "Microsoft Office Excel"
                });
                workbook.Append(sheets);
                SheetData sheetData = CreateSheetData(objects, headerNames);
                worksheet.Append(sheetData);
                worksheetPart.Worksheet = worksheet;
                worksheetPart.Worksheet.Save();
                document.WorkbookPart.Workbook = workbook;
                document.WorkbookPart.Workbook.Save();
                document.Close();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T">Object type to pass in</typeparam>
        /// <param name="objects">list of the object type</param>
        /// <param name="headerNames">Header names of the object</param>
        /// <returns></returns>
        private static SheetData CreateSheetData<T>(List<T> objects,
            List<string> headerNames)
        {
            var sheetData = new SheetData();
            if (objects != null)
            {
                //Get fields names of object
                List<string> fields = GetPropertyInfo<T>();
                //Get a list of A to Z
                var az = new List<Char>(Enumerable.Range('A', 'Z' -
                                                              'A' + 1).Select(i => (Char)i).ToArray());
                //A to E number of columns 
                List<Char> headers = az.GetRange(0, fields.Count);
                int numRows = objects.Count;
                int numCols = fields.Count;
                var header = new Row();
                int index = 1;
                header.RowIndex = (uint)index;
                for (int col = 0; col < numCols; col++)
                {
                    var c = new HeaderCell(headers[col].ToString(),
                        headerNames[col], index);
                    header.Append(c);
                }
                sheetData.Append(header);
                for (int i = 0; i < numRows; i++)
                {
                    index++;
                    var obj1 = objects[i];
                    var r = new Row { RowIndex = (uint)index };
                    for (int col = 0; col < numCols; col++)
                    {
                        string fieldName = fields[col];
                        PropertyInfo myf = obj1.GetType().GetProperty(fieldName);
                        if (myf != null)
                        {
                            object obj = myf.GetValue(obj1, null);
                            if (obj != null)
                            {
                                if (obj.GetType() == typeof(string))
                                {
                                    var c = new TextCell(headers[col].ToString(),
                                        obj.ToString(), index);
                                    r.Append(c);
                                }
                                else if (obj.GetType() == typeof(bool))
                                {
                                    string value =
                                        (bool)obj ? "Yes" : "No";
                                    var c = new TextCell(headers[col].ToString(),
                                        value, index);
                                    r.Append(c);
                                }
                                else if (obj.GetType() == typeof(DateTime))
                                {
                                    var c = new DateCell(headers[col].ToString(),
                                        (DateTime)obj, index);
                                    r.Append(c);
                                }
                                else if (obj.GetType() == typeof(decimal) ||
                                         obj.GetType() == typeof(double))
                                {
                                    var c = new FormatedNumberCell(
                                        headers[col].ToString(),
                                        obj.ToString(), index);
                                    r.Append(c);
                                }
                                else
                                {
                                    long value;
                                    if (long.TryParse(obj.ToString(), out value))
                                    {
                                        var c = new NumberCell(headers[col].ToString(),
                                            obj.ToString(), index);
                                        r.Append(c);
                                    }
                                    else
                                    {
                                        var c = new TextCell(headers[col].ToString(),
                                            obj.ToString(), index);
                                        r.Append(c);
                                    }
                                }
                            }
                        }
                    }
                    sheetData.Append(r);
                }
                index++;
                Row total = new Row();
                total.RowIndex = (uint)index;
                for (int col = 0; col < numCols; col++)
                {
                    var obj1 = objects[0];
                    string fieldName = fields[col];
                    PropertyInfo myf = obj1.GetType().GetProperty(fieldName);
                    if (myf != null)
                    {
                        object obj = myf.GetValue(obj1, null);
                        if (obj != null)
                        {
                            if (col == 0)
                            {
                                var c = new TextCell(headers[col].ToString(),
                                    "Total", index);
                                c.StyleIndex = 10;
                                total.Append(c);
                            }
                            else if (obj.GetType() == typeof(decimal) ||
                                     obj.GetType() == typeof(double))
                            {
                                string headerCol = headers[col].ToString();
                                string firstRow = headerCol + "2";
                                string lastRow = headerCol + (numRows + 1);
                                string formula = "=SUM(" + firstRow + " : " + lastRow + ")";
                                //Console.WriteLine(formula);
                                var c = new FomulaCell(headers[col].ToString(),
                                    formula, index);
                                c.StyleIndex = 9;
                                total.Append(c);
                            }
                            else
                            {
                                var c = new TextCell(headers[col].ToString(),
                                    string.Empty, index);
                                c.StyleIndex = 10;
                                total.Append(c);
                            }
                        }
                    }
                }
                sheetData.Append(total);
            }
            return sheetData;
        }
        private static List<string> GetPropertyInfo<T>()
        {
            PropertyInfo[] propertyInfos = typeof(T).GetProperties();
            // write property names
            return propertyInfos.Select(propertyInfo => propertyInfo.Name).ToList();
        }
    }
}