using Tests.Helpers;
using InStack.Excel.Builder.Extensions.Cell;

namespace Tests;

[TestClass]
public class WriteBoolCellTests
{
    [TestMethod]
    [DataRow(true, (uint)1, (uint)1, (uint)10, @"<c t=""b"" r=""A1"" s=""10""><v>1</v></c>")]
    [DataRow(false, (uint)1, (uint)1, (uint)10, @"<c t=""b"" r=""A1"" s=""10""><v>0</v></c>")]
    [DataRow(false, (uint)1, (uint)322, (uint)10, @"<c t=""b"" r=""LJ1"" s=""10""><v>0</v></c>")]
    [DataRow(false, (uint)1, (uint)1, null, @"<c t=""b"" r=""A1""><v>0</v></c>")]
    [DataRow(null, (uint)1, (uint)1, null, @"<c t=""b"" r=""A1""/>")]
    public void WriteBool_CorrectCellGenerated(bool? val, uint? row, uint? column, uint? style, string expectedOutput)
    {
        TestHelper.TestCellWrite(row, column, expectedOutput, (sheet) =>
        {
            sheet.WriteBool(val, style: style);
        });
    }
}
