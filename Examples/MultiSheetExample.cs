using DocumentFormat.OpenXml.Spreadsheet;
using Examples.CollectionExample;
using Examples.MonthlyReportExample;
using InStack.Excel.Builder;
using InStack.Excel.OpenXmlStyles;
using System.Globalization;
using TableStyles = Examples.MonthlyReportExample.TableStyles;

namespace Examples;

public static class MultiSheetExample
{
    public static void CreateMultiSheetExcel()
    {
        var monthlyReportHeaderStyles = new TableHeaderStyles("0F6CBD", "FFFFFF", "FFFFFF", BorderStyleValues.Thick);
        var monthlyReportTableStyles = new TableStyles("DCEBF7", "F7FBFF");
        var collectionExampleStyles = new CollectionExampleStyles();

        var stylesXml = Style.Compile(
            monthlyReportHeaderStyles, 
            monthlyReportTableStyles);

        var fileName = $"D:\\xlsx\\{nameof(CreateMultiSheetExcel)}{DateTime.UtcNow.ToString("T", CultureInfo.InvariantCulture).Replace(":", "-")}.xlsx";

        using (var builder = new XlsxDocument(fileName))
        {
            using (var styleXmlWriter = builder.CreateStyleEntry())
            {
                using var xmlWriter = System.Xml.XmlWriter.Create(styleXmlWriter, new System.Xml.XmlWriterSettings { CloseOutput = true });
                stylesXml.WriteTo(xmlWriter);
            }

            builder.CollectionExampleSheet(Collection.GetItems(10), collectionExampleStyles);
            builder.MonthlyReportExampleSheet(monthlyReportHeaderStyles, monthlyReportTableStyles);
        }

        System.IO.Compression.ZipFile.ExtractToDirectory(fileName, fileName.Replace(".xlsx", ""));
    }
}
