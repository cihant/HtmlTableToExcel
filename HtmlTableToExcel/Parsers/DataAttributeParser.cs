using AngleSharp.Html.Dom;
using System.Collections.Generic;

namespace HtmlTableToExcel.Parsers
{
    public class DataAttributeParser : IHtmlTableParser
    {
        const string ATTRIBUTE_NAME = "data-export";

        const string ATTRIBUTE_VALUE_TABLE = "table";
        const string ATTRIBUTE_VALUE_ROW = "row";
        const string ATTRIBUTE_VALUE_CELL = "cell";

        public DataAttributeParser()
        {
        }

        public List<HtmlTable> GetTablesFromRootNode(IHtmlDocument doc)
        {
            List<HtmlTable> tables = new List<HtmlTable>();
            foreach (var cn in doc.Body.ChildNodes)
            {
                IHtmlElement tempElement = cn as IHtmlElement;
                if (tempElement == null)
                {
                    continue;
                }

                if (tempElement.HasAttribute(ATTRIBUTE_NAME) && tempElement.GetAttribute(ATTRIBUTE_NAME) == ATTRIBUTE_VALUE_TABLE)
                {
                    tables.Add(new HtmlTable(this, null, tempElement));
                }
            }

            return tables;
        }
        public void ScanTableNode(HtmlElement parent, IHtmlElement el)
        {
            foreach (var cn in el.ChildNodes)
            {
                if (!(cn is IHtmlElement))
                {
                    continue;
                }
                IHtmlElement tempElement = (IHtmlElement)cn;
                if (tempElement.HasAttribute(ATTRIBUTE_NAME) && tempElement.GetAttribute(ATTRIBUTE_NAME) == ATTRIBUTE_VALUE_ROW)
                {
                    parent.Children.Add(new HtmlRow(this, parent, tempElement));
                }
                else
                {
                    ScanTableNode(parent, tempElement);
                }
            }
        }
        public void ScanRowNodes(HtmlElement parent, IHtmlElement el)
        {
            if (!(el is IHtmlElement))
            {
                return;
            }
            foreach (var cn in el.ChildNodes)
            {
                if (!(cn is IHtmlElement))
                {
                    continue;
                }
                IHtmlElement tempElement = cn as IHtmlElement;
                if (tempElement == null)
                {
                    continue;
                }
                if (tempElement.HasAttribute(ATTRIBUTE_NAME) && tempElement.GetAttribute(ATTRIBUTE_NAME) == ATTRIBUTE_VALUE_CELL)
                {
                    parent.Children.Add(new HtmlCell(this, parent, tempElement));
                }
            }
        }
        public void ScanCellNodes(HtmlElement parent)
        {
            foreach (var cn in parent.El.ChildNodes)
            {
                IHtmlElement tempElement = cn as IHtmlElement;
                if (tempElement == null)
                {
                    continue;
                }

                if (tempElement.HasAttribute(ATTRIBUTE_NAME) && tempElement.GetAttribute(ATTRIBUTE_NAME) == ATTRIBUTE_VALUE_TABLE)
                {
                    parent.Children.Add(new HtmlTable(this, parent, cn as IHtmlTableElement));
                }
            }
        }
    }
}
