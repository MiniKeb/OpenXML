using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelPOC.CodeProjectSample
{
    public class NumberCell : Cell
    {
        public NumberCell(string header, string text, int index)
        {
            this.DataType = CellValues.Number;
            this.CellReference = header + index;
            this.CellValue = new CellValue(text);
        }
    }
}