using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using System.Collections.Generic;

namespace HtmlTableToExcel.Parsers
{
    public class TableTagParser: IHtmlTableParser
    {
        public TableTagParser()
        {
        }
        public List<HtmlTable> GetTablesFromRootNode(IHtmlDocument doc)
        {
            List<HtmlTable> tables = new List<HtmlTable>();
            foreach (IElement table in doc.GetElementsByTagName("table"))
            {
                if (!(table.ParentElement is IHtmlTableCellElement))
                {
                    HtmlTable htmlTable = new HtmlTable(this, null, table as IHtmlTableElement);
                    tables.Add(htmlTable);
                }
            }
            return tables;
        }
        public void ScanTableNode(HtmlElement parent, IHtmlElement el)
        {
            if (!(el is IHtmlElement))
            {
                return;
            }
            foreach (var cn in el.ChildNodes)
            {
                if (!(cn is IHtmlTableElement))
                {
                    if (cn is IHtmlTableRowElement)
                    {
                        parent.Children.Add(new HtmlRow(this, parent, cn as IHtmlTableRowElement));
                    }
                    else
                    {
                        ScanTableNode(parent, cn as IHtmlElement);
                    }
                }
            }
        }

        public void ScanRowNodes(HtmlElement parent, IHtmlElement el)
        {
            foreach (var cn in parent.El.ChildNodes)
            {
                if (!(cn is IHtmlElement))
                {
                    continue;
                }
                if (cn is IHtmlTableCellElement)
                {
                    parent.Children.Add(new HtmlCell(this, parent, cn as IHtmlTableCellElement));
                }
            }
        }

        public void ScanCellNodes(HtmlElement parent)
        {
            foreach (var cn in parent.El.ChildNodes)
            {
                if (!(cn is IHtmlElement))
                {
                    continue;
                }
                if (cn is IHtmlTableElement)
                {
                    parent.Children.Add(new HtmlTable(this, parent, cn as IHtmlTableElement));
                }
            }
        }

    }
}
