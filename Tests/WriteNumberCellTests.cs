using Tests.Helpers;
using InStack.Excel.Builder.Extensions.Cell;

namespace Tests;

[TestClass]
public class WriteNumberCellTests
{
    [TestMethod]
    [DataRow(1234567, (uint)1, (uint)322, (uint)10, @"<c t=""n"" r=""LJ1"" s=""10""><v>1234567</v></c>")]
    [DataRow(123, (uint)1, (uint)1, null, @"<c t=""n"" r=""A1""><v>123</v></c>")]
    [DataRow(null, (uint)1, (uint)1, null, @"<c t=""n"" r=""A1""/>")]
    public void WriteIntNumber_CorrectCellGenerated(int? val, uint? row, uint column, uint? style, string expectedOutput)
    {
        TestHelper.TestCellWrite(row, expectedOutput, (sheet) =>
        {
            sheet.Write(val, column, style: style);
        });
    }

    [TestMethod]
    [DataRow(123456.7, (uint)1, (uint)322, (uint)10, @"<c t=""n"" r=""LJ1"" s=""10""><v>123456.7</v></c>")]
    [DataRow(12.3, (uint)1, (uint)1, null, @"<c t=""n"" r=""A1""><v>12.3</v></c>")]
    [DataRow(null, (uint)1, (uint)1, null, @"<c t=""n"" r=""A1""/>")]
    public void WriteDoubleNumber_CorrectCellGenerated(double? val, uint? row, uint column, uint? style, string expectedOutput)
    {
        TestHelper.TestCellWrite(row, expectedOutput, (sheet) =>
        {
            sheet.Write(val, column, style: style);
        });
    }

    [TestMethod]
    [DataRow(1234, (uint)1, (uint)322, (uint)10, @"<c t=""n"" r=""LJ1"" s=""10""><v>1234</v></c>")]
    [DataRow(123, (uint)1, (uint)1, null, @"<c t=""n"" r=""A1""><v>123</v></c>")]
    [DataRow(null, (uint)1, (uint)1, null, @"<c t=""n"" r=""A1""/>")]
    public void WriteDate_CorrectCellGenerated(int? oaDate, uint? row, uint column, uint? style, string expectedOutput)
    {
        TestHelper.TestCellWrite(row, expectedOutput, (sheet) =>
        {
            var oaDateTime = oaDate.HasValue
                ? (DateTime?)DateTime.FromOADate(oaDate.Value)
                : null;
            sheet.Write(oaDateTime, column, style: style);
        });
    }
}
