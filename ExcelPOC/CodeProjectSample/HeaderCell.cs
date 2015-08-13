namespace ExcelPOC.CodeProjectSample
{
    public class HeaderCell : TextCell
    {
        public HeaderCell(string header, string text, int index) :
               base(header, text, index)
        {
            this.StyleIndex = 11;
        }
    }

}