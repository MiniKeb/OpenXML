using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelPOC.CodeProjectSample
{
    public class CustomColumn : Column
    {

        public CustomColumn(UInt32 startColumnIndex,
            UInt32 endColumnIndex, double columnWidth)
        {
            this.Min = startColumnIndex;
            this.Max = endColumnIndex;
            this.Width = columnWidth;
            this.CustomWidth = true;
        }
    }
}