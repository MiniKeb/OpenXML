using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelPOC.CodeProjectSample
{
    public class DateCell : Cell
    {
        public DateCell(string header, DateTime dateTime, int index)
        {
            this.DataType = CellValues.Date;
            this.CellReference = header + index;
            this.StyleIndex = 1;
            this.CellValue = new CellValue { Text = dateTime.ToOADate().ToString() }; ;
        }
    }
}