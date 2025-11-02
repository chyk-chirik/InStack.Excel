using InStack.Excel.Builder;
using Shouldly;
using System.Xml;
using Tests.Helpers;

namespace Tests;

[TestClass]
public class CompoundWriteCellTests
{
    [TestMethod]
    public void WriteSeveralCells_ValidXmlProduced()
    {
        var testStreamManager = new TestZipStreamManager();

        using (var xlsx = new XlsxDocument(testStreamManager))
        using (var sheet = xlsx.AddSheet("sheetName"))
        {
            sheet.StartRow(row: 1, column: 1, height: 12);

            sheet.Write("String");
            sheet.Write(true, style: 1);
            sheet.Write<int>(3, style: 2);
            sheet.Write<decimal>(3.12M, style: 3);
            sheet.Write(DateTime.UtcNow, style: 4);

            sheet.MergePrevCellToBottom();

            sheet.EndRowAndStartNew();

            sheet.Write(TestHelper.StringLengthOf_4096, 1000);
            sheet.Write(true);
            sheet.Write<int>(3);
            sheet.Write<decimal>(3.12M);
            sheet.Write(DateTime.UtcNow);

            sheet.MergePrevCellToRight();

            sheet.EndRow();
        }

        Should.NotThrow(() =>
        {
            var stream = testStreamManager.Entries["xl/worksheets/sheet1.xml"];
            var sheetData = System.Text.Encoding.UTF8.GetString(stream.ToArray());
            using var sr = new StringReader(sheetData);
            using var xr = XmlReader.Create(sr);
            
            while (xr.Read()) { }
        });
    }

    [TestMethod]
    public void NotClosedRowTag_InvalidXmlProduced()
    {
        var testStreamManager = new TestZipStreamManager();

        using (var xlsx = new XlsxDocument(testStreamManager))
        using (var sheet = xlsx.AddSheet("sheetName"))
        {
            sheet.StartRow(row: 1, column: 1, height: 12);
        }

        Should.Throw<XmlException>(() => 
        {
            var stream = testStreamManager.Entries["xl/worksheets/sheet1.xml"];
            var sheetData = System.Text.Encoding.UTF8.GetString(stream.ToArray());
            using var sr = new StringReader(sheetData);
            using var xr = XmlReader.Create(sr);

            while (xr.Read()) { }
        });
    }
}
