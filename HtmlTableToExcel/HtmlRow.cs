using AngleSharp.Html.Dom;
using HtmlTableToExcel.Parsers;

namespace HtmlTableToExcel
{
    public class HtmlRow : HtmlElement
    {
        public HtmlRow(IHtmlTableParser parser, HtmlElement parent, IHtmlElement htmlElement) 
            : base(parser, parent, htmlElement)
        {
            Parser.ScanRowNodes(this, this.El);
        }
        protected override int CalculateRowCount()
        {
            int rowCount = 1;
            foreach (var child in this.Children)
            {
                int tempRowCount = child.RowCount;
                if (tempRowCount > rowCount) // if td has table and table row count bigger than another table row count
                {
                    rowCount = tempRowCount;
                }

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
                else
                {
                    ++colCount;
                }
            }
            return colCount;
        }
    }
}
