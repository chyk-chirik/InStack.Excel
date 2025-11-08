using InStack.Excel.Builder;
using Shouldly;
using System.Xml;
using Tests.Helpers;
using InStack.Excel.Builder.Extensions.Cell;

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
            uint column = 1;

            sheet.StartRow(row: 1, height: 12);

            sheet.Write("String", column++);
            sheet.WriteBool(true, column++, style: 1);
            sheet.Write<int>(3, column++, style: 2);
            sheet.Write<decimal>(3.12M, column++, style: 3);
            sheet.Write(DateTime.UtcNow, column++, style: 4);

            sheet.MergeCellToBottom(column);

            sheet.StartRow();

            sheet.Write(TestHelper.StringLengthOf_4096, 1000);
            sheet.WriteBool(true, column++);
            sheet.Write<int>(3, column++);
            sheet.Write<decimal>(3.12M, column++);
            sheet.Write(DateTime.UtcNow, column++);

            sheet.MergeCellToRight(column);
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
}
