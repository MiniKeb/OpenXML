using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelPOC.CodeProjectSample
{
    public class TextCell : Cell
    {
        public TextCell(string header, string text, int index)
        {
            this.DataType = CellValues.InlineString;
            this.CellReference = header + index;
            //Add text to the text cell.
            this.InlineString = new InlineString { Text = new Text { Text = text } };
        }
    }
}