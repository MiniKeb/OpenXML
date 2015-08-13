namespace ExcelPOC.CodeProjectSample
{
    public class FormatedNumberCell : NumberCell
    {
        public FormatedNumberCell(string header, string text, int index)
            : base(header, text, index)
        {
            this.StyleIndex = 2;
        }
    }
}