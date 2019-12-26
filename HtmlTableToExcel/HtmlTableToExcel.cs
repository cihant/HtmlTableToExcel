using AngleSharp;
using AngleSharp.Css.Dom;
using AngleSharp.Html.Parser;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Linq;
using System.Collections.Generic;
using HtmlTableToExcel.Parsers;

namespace HtmlTableToExcel
{
    public class HtmlTableToExcel
    {
        const string tagWhiteSpace = @"(>|$)(\W|\n|\r)+<";//matches one or more (white space or line breaks) between '>' and '<'
        const string stripFormatting = @"<[^>]*(>|$)";//match any character between '<' and '>', even when end tag is missing
        const string lineBreak = @"<(br|BR)\s{0,1}\/{0,1}>";//matches: <br>,<br/>,<br />,<BR>,<BR/>,<BR />

        readonly Regex lineBreakRegex = new Regex(lineBreak, RegexOptions.Multiline);
        readonly Regex stripFormattingRegex = new Regex(stripFormatting, RegexOptions.Multiline);
        readonly Regex tagWhiteSpaceRegex = new Regex(tagWhiteSpace, RegexOptions.Multiline);

        readonly ExcelPackage excel = new ExcelPackage();
        readonly ExcelWorksheet sheet;

        readonly bool useCssStyles = false;
        readonly IHtmlTableParser htmlTableParser;

        public HtmlTableToExcel(IHtmlTableParser parser=null, string sheetName = "sheet1", eOrientation orientation = eOrientation.Portrait)
        {
            sheet = excel.Workbook.Worksheets.Add(sheetName);
            sheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            sheet.PrinterSettings.Orientation = orientation;

            if (parser == null)
            {
                htmlTableParser = new TableTagParser();
            }
            else
            {
                htmlTableParser = parser;
            }
        }

        public byte[] Process(string html)
        {
            MemoryStream stream = null;
            try
            {
                Process(html, out stream);
                return stream.ToArray();
            }
            finally
            {
                if (stream != null)
                {
                    try
                    {
                        stream.Close();
                    }
                    catch (IOException e)
                    {
                        throw;
                    }
                }
            }
        }
        public void Process(string html, out MemoryStream output)
        {
            IBrowsingContext context;
            if (useCssStyles)
            {
                context = BrowsingContext.New(Configuration.Default.WithCss());
            }
            else
            {
                context = BrowsingContext.New(Configuration.Default);
            }

            var parser = context.GetService<IHtmlParser>();
            var doc = parser.ParseDocument(html);

            ExcelRange startTableRange = sheet.Cells[1, 1];
            List<HtmlTable> tables = htmlTableParser.GetTablesFromRootNode(doc);
            foreach (var htmlTable in tables)
            {
                startTableRange = ProcessTable(htmlTable, startTableRange);
                startTableRange = sheet.Cells[startTableRange.Start.Row + 1, 1];
            }

            // sheet.Cells[sheet.Dimension.Address].AutoFitColumns(10, 50)
            // sheet.Cells[sheet.Dimension.Start.Row, sheet.Dimension.End.Column].AutoFitColumns(10, 50)

            try
            {
                output = new MemoryStream();
                excel.SaveAs(output);
            }
            catch (Exception e)
            {
                throw;
            }
        }

        private ExcelRange ProcessTable(HtmlTable htmlTable, ExcelRange containerRange)
        {
            foreach (HtmlRow row in htmlTable.Children.OfType<HtmlRow>())
            {
                ExcelRange range = ProcessCells(row, containerRange);
                containerRange = sheet.Cells[containerRange.Start.Row + range.Rows, containerRange.Start.Column];
                Log($"Row Address:{containerRange.Address} - {range.Address}");
            }

            return containerRange;
        }
        private ExcelRange ProcessCells(HtmlRow row, ExcelRange containerRange)
        {
            int relativeRowIndex = containerRange.Start.Row;
            int tempColIndex = containerRange.Start.Column;
            int rowSpan = row.RowCount, colSpan;
            foreach (HtmlCell cell in row.Children.OfType<HtmlCell>())
            {
                ExcelRange range = sheet.Cells[relativeRowIndex, tempColIndex];
                colSpan = cell.ColCount;
                if (rowSpan > 1 || colSpan > 1)
                {
                    range = sheet.Cells[relativeRowIndex, tempColIndex, relativeRowIndex + rowSpan - 1, tempColIndex + colSpan - 1];
                    tempColIndex += colSpan - 1;
                    range.Merge = true;
                    Log($"------------Cell Merged:{range.Address}");
                }

                containerRange = range;

                if (cell.Children.Count == 0)
                {
                    range.Value = HtmlToPlainText(cell.Text);
                    SetRangeStyle(range, cell);
                }
                else
                {
                    containerRange.Merge = false;
                    foreach (HtmlTable table in cell.Children.OfType<HtmlTable>())
                    {
                        Log($"Processing subtable: {range.Address}");
                        ProcessTable(table, range);
                    }
                }
                ++tempColIndex;

                Log($"Cell Address:{containerRange.Address}");
            }
            return containerRange;
        }
        private void SetRangeStyle(ExcelRange range, HtmlCell cell)
        {
            if (useCssStyles)
            {
                var css = cell.El.GetStyle();
                foreach (var sa in css)
                {
                    PropertyToStyle(range, sa.Name, sa.Value);
                }
            }
            else
            {
                var cellStyleAttributes = cell.StyleAttributes;

                foreach (var sa in cellStyleAttributes)
                {
                    PropertyToStyle(range, sa.Key, sa.Value);
                }
            }
        }
        private void PropertyToStyle(ExcelRange range, string cssProperty, string cssPropertyValue)
        {
            if (string.IsNullOrEmpty(cssProperty) || string.IsNullOrEmpty(cssPropertyValue))
                return;

            switch (cssProperty)
            {

                case "text-align":
                    switch (cssPropertyValue)
                    {
                        case "left": range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left; break;
                        case "right": range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right; break;
                        case "center": range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; break;
                        case "justify": range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify; break;
                    }
                    break;
                case "font-weight":
                    switch (cssPropertyValue)
                    {
                        case "normal": range.Style.Font.Bold = false; break;
                        case "bold":
                        case "bolder": range.Style.Font.Bold = true; break;
                    }
                    break;

                case "border":
                    switch (cssPropertyValue)
                    {
                        case "solid 1px": range.Style.Border.BorderAround(ExcelBorderStyle.Thin); break;
                        case "solid 2px": range.Style.Border.BorderAround(ExcelBorderStyle.Medium); break;
                        case "solid 3px": range.Style.Border.BorderAround(ExcelBorderStyle.Thick); break;
                        case "dashed 1px": range.Style.Border.BorderAround(ExcelBorderStyle.Dashed); break;
                        case "dotted 1px": range.Style.Border.BorderAround(ExcelBorderStyle.Dotted); break;
                    }
                    break;
                case "size":
                case "width":
                    double width;
                    if (double.TryParse(cssPropertyValue.Replace("px", ""), out width))
                    {
                        sheet.Column(range.Columns).Width = width;
                    }
                    break;
                case "background-color":
                    // range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(Int32.Parse(cssPropertyValue.Replace("#", ""), System.Globalization.NumberStyles.HexNumber)));
                    break;
                case "border-color":
                    // range.Style.Border.BorderAround(range.Style.Border.Left.Style, System.Drawing.Color.FromArgb(Int32.Parse(cssPropertyValue.Replace("#", ""), System.Globalization.NumberStyles.HexNumber)));
                    break;
                default:
                    break;
            }
        }
        private string HtmlToPlainText(string html)
        {
            //https://stackoverflow.com/a/16407272

            //Decode html specific characters
            html = System.Net.WebUtility.HtmlDecode(html);
            //Remove tag whitespace/line breaks
            html = tagWhiteSpaceRegex.Replace(html, "><");
            //Replace <br /> with line breaks
            html = lineBreakRegex.Replace(html, Environment.NewLine);
            //Strip formatting
            html = stripFormattingRegex.Replace(html, string.Empty);

            return html;
        }
        private void Log(string v)
        {
            Console.WriteLine(v);
        }
    }

}
