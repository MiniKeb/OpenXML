using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelPOC
{
    internal interface ISheet
    {
        string SheetName { get; }

        SheetData GetSheetData();

        Columns GetColumns();
    }
}