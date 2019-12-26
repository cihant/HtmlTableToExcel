using AngleSharp.Html.Dom;
using HtmlTableToExcel.Parsers;
using System.Collections.Generic;

namespace HtmlTableToExcel
{
    public abstract class HtmlElement
    {
        public int RowCount { get { return CalculateRowCount(); } }
        public int ColCount { get { return CalculateColCount(); } }

        public IHtmlTableParser Parser { get; }
        public HtmlElement Parent { get; set; }
        public List<HtmlElement> Children { get; set; } = new List<HtmlElement>();
        public IHtmlElement El { get; protected set; }

        protected HtmlElement(IHtmlTableParser parser, HtmlElement parent, IHtmlElement htmlElement)
        {
            this.Parser = parser;
            this.Parent = parent;
            this.El = htmlElement;
        }
        protected abstract int CalculateRowCount();
        protected abstract int CalculateColCount();
    }
}
