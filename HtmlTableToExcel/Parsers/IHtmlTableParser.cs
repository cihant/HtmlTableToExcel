using AngleSharp.Html.Dom;
using System.Collections.Generic;

namespace HtmlTableToExcel.Parsers
{
    public interface IHtmlTableParser
    {
        List<HtmlTable> GetTablesFromRootNode(IHtmlDocument doc);
        void ScanCellNodes(HtmlElement parent);
        void ScanRowNodes(HtmlElement parent, IHtmlElement el);
        void ScanTableNode(HtmlElement parent, IHtmlElement el);
    }
}