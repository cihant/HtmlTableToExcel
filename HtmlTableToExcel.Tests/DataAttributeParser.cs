using AngleSharp;
using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;
using OfficeOpenXml;
using System.IO;
using Xunit;

namespace HtmlTableToExcel.Tests
{
    public class DataAttributeParserTests
    {
        [Fact]
        public void DataAttributeParser_ShouldTableNotNull()
        {
            var html = ReadFromFile("htmltable.html");
            Parsers.IHtmlTableParser htmlTableParser = new Parsers.DataAttributeParser();
            IHtmlDocument doc = GetHtmlDoc(html);
            var table = doc.GetElementsByTagName("table")[0];

            HtmlTable htmlTable = new HtmlTable(htmlTableParser, null, table as IHtmlTableElement);
            Assert.NotNull(htmlTable);
        }

        [Fact]
        public void DataAttributeParser_ShouldHaveThreeRows()
        {
            string html = ReadFromFile("htmltable.html");

            HtmlTableToExcel tableToExcel = new HtmlTableToExcel();
            byte[] bytes = tableToExcel.Process(html);

            ExcelPackage excel = new ExcelPackage();
            using (MemoryStream ms = new MemoryStream(bytes))
            {
                excel.Load(ms);
            }

            Assert.Equal(3, excel.Workbook.Worksheets[0].Dimension.Rows);
        }

        [Theory]
        [InlineData(1, 1, "Name")]
        [InlineData(1, 2, "Favorite Color")]
        [InlineData(2, 1, "Bob")]
        [InlineData(2, 2, "Yellow")]
        [InlineData(3, 1, "Michelle")]
        [InlineData(3, 2, "Purple")]
        public void DataAttributeParser_ShouldCellTextHaveCorrectValues(int row, int col, string cellValue)
        {
            string html = ReadFromFile("htmldiv.html");
            var parser = new Parsers.DataAttributeParser();
            HtmlTableToExcel tableToExcel = new HtmlTableToExcel(parser);
            byte[] bytes = tableToExcel.Process(html);

            ExcelPackage excel = new ExcelPackage();
            using (MemoryStream ms = new MemoryStream(bytes))
            {
                excel.Load(ms);
            }

            Assert.Equal(cellValue, excel.Workbook.Worksheets[0].Cells[row, col].Text);
        }


        //// Commented.Tests should not generate file.
        //[Fact]
        //public void DataAttributeParser_ShouldGenerateExcelFile()
        //{
        //    string html = ReadFromFile("htmltable.html");
        //    string targetFileName = "htmldiv.xlsx";

        //    HtmlTableToExcel tableToExcel = new HtmlTableToExcel();
        //    byte[] bytes = tableToExcel.Process(html);

        //    File.WriteAllBytes(targetFileName, bytes);

        //    Assert.True(File.Exists(targetFileName));
        //}

        static string ReadFromFile(string file)
        {
            using (var stream = new StreamReader(file))
            {
                return stream.ReadToEnd();
            }
        }
        private static IHtmlDocument GetHtmlDoc(string html)
        {
            var context = BrowsingContext.New(Configuration.Default);
            var parser = context.GetService<IHtmlParser>();
            var doc = parser.ParseDocument(html);
            return doc;
        }
    }
}
