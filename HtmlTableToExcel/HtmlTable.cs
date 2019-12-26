using AngleSharp.Html.Dom;
using HtmlTableToExcel.Parsers;

namespace HtmlTableToExcel
{
    public class HtmlTable : HtmlElement
    {
        public HtmlTable(IHtmlTableParser parser, HtmlElement parent, IHtmlElement htmlElement) 
            : base(parser, parent, htmlElement)
        {
            Parser.ScanTableNode(this, this.El);
        }
        protected override int CalculateRowCount()
        {
            int rowCount = 0;
            foreach (var child in this.Children)
            {
                rowCount += child.RowCount;
            }
            return rowCount;
        }
        protected override int CalculateColCount()
        {
            int colCount = 0;
            foreach (var child in this.Children)
            {
                int tempColCount = child.ColCount;
                if (tempColCount > colCount) // if td has table and table col count bigger than another table col count
                {
                    colCount = tempColCount;
                }
            }
            return colCount;
        }
    }
}
