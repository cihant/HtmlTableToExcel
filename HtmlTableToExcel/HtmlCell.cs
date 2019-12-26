using AngleSharp.Html.Dom;
using HtmlTableToExcel.Parsers;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace HtmlTableToExcel
{
    public class HtmlCell : HtmlElement
    {
        public string Text { get { return GetText(); } }
        public Dictionary<string, string> StyleAttributes { get { return GetStyleText(); } }
        private Dictionary<string, string> GetStyleText()
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            const string stylePattern = "style=\"([^\\\"]*)\"";
            const string attributePattern = @"([\w\-]+):\s*([\w\#\s]+\s*)";
            var match = Regex.Match(this.El.OuterHtml, stylePattern);

            if (match.Success)
            {
                var attributeMatches = Regex.Matches(match.Value, attributePattern);
                foreach (Match m in attributeMatches)
                {
                    string key = m.Groups[1].Value.Trim();
                    if (!result.ContainsKey(key))
                    {
                        result.Add(m.Groups[1].Value.Trim(), m.Groups[2].Value.Trim());
                    }
                }
            }
            return result;
        }

        public HtmlCell(IHtmlTableParser parser, HtmlElement parent, IHtmlElement htmlElement)
            : base(parser, parent, htmlElement)
        {
            Parser.ScanCellNodes(this);
        }

        protected override int CalculateRowCount()
        {
            int rowCount = 0;
            foreach (var child in this.Children)
            {
                rowCount += child.RowCount;
            }
            int rowSpan = GetSpan(this.El.OuterHtml, 0);
            if (rowSpan > rowCount)
            {
                rowCount = rowSpan;
            }
            return rowCount;
        }
        protected override int CalculateColCount()
        {
            int colCount = 1; // because td is a column
            foreach (var child in this.Children)
            {
                int tempColCount = child.ColCount;
                if (tempColCount > colCount) // if td has table and table col count bigger than another table col count
                {
                    colCount = tempColCount;
                }
            }
            int colSpan = GetSpan(this.El.OuterHtml, 1);
            if (colSpan > colCount)
            {
                colCount = colSpan;
            }
            return colCount;
        }

        private string GetText()
        {
            if (this.Children.Count == 0)
            {
                return this.El.InnerHtml.Trim();
            }

            return "";
        }
        private int GetSpan(string html, int spanType = 0)
        {
            string spanTypeText;
            int span;

            if (spanType == 0)
            {
                spanTypeText = "row";
            }
            else
            {
                spanTypeText = "col";
            }
            string equation = Regex.Match(html.ToLower(), spanTypeText + @"span=.*?\d{1,}").ToString();
            if (!int.TryParse(Regex.Match(equation, @"\d{1,}").ToString(), out span))
            {
                span = 1;
            }

            return span;
        }
    }
}
