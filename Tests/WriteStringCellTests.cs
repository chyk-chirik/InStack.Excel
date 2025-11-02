using InStack.Excel.Builder;
using Shouldly;
using Tests.Helpers;

namespace Tests;

[TestClass]
public class WriteStringCellTests
{
    public const int SheetBufferSize = 512;

    [TestMethod]
    [DataRow("withstyle", (uint)1, (uint)322, (uint)10, @"<c t=""inlineStr"" r=""LJ1"" s=""10""><is><t>withstyle</t></is></c>")]
    [DataRow("nostyle", (uint)1, (uint)1, null, @"<c t=""inlineStr"" r=""A1""><is><t>nostyle</t></is></c>")]
    [DataRow(null, (uint)1, (uint)1, null, @"<c t=""inlineStr"" r=""A1""/>")]
    [DataRow("", (uint)1, (uint)1, null, @"<c t=""inlineStr"" r=""A1""/>")]
    public void WriteShortInlineString_CorrectCellGenerated(string? val, uint? row, uint? column, uint? style, string expectedOutput)
    {
        TestHelper.TestCellWrite(row, column, expectedOutput, (sheet) =>
        {
            sheet.Write(val, style: style, escape: false);
        });
    }

    [TestMethod]
    [DataRow("<Big> \"Boy\" & 'co'", (uint)1, (uint)1, null, @"<c t=""inlineStr"" r=""A1""><is><t>&lt;Big&gt; &quot;Boy&quot; &amp; &apos;co&apos;</t></is></c>")]
    public void WriteShortInlineStringEscaped_CorrectCellGenerated(string? val, uint? row, uint? column, uint? style, string expectedOutput)
    {
        TestHelper.TestCellWrite(row, column, expectedOutput, (sheet) =>
        {
            sheet.Write(val, style: style, escape: true);
        });
    }

    [TestMethod]
    [DataRow()]
    public void WriteInlineStringLongerThenBuffer_Success()
    {
        var testStreamManager = new TestZipStreamManager();

        MemoryStream sheetStream;

        using (var xlsx = new XlsxDocument(testStreamManager))
        using (var sheet = xlsx.AddSheet("sheetName", new SheetConfig {  BufferSizeInBytes = 2048 }))
        {
            sheetStream = testStreamManager.Entries["xl/worksheets/sheet1.xml"];

            sheet.StartRow(row: 1, column: 1, height: 12);

            sheet.Write(TestHelper.StringLengthOf_4096);
           
            var bufferHelper = new BufferHelper(sheet);
            bufferHelper.FlushBuffer();

            var streamBytes = sheetStream.GetBuffer();
            var sheetData = System.Text.Encoding.UTF8.GetString(streamBytes.AsSpan());

            sheetData.ShouldContain(TestHelper.StringLengthOf_4096);
        }
    }
}
