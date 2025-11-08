using Tests.Helpers;
using InStack.Excel.Builder.Extensions.Cell;

namespace Tests;

[TestClass]
public class WriteSumFormulaCellTests
{
    [TestMethod]
    [DataRow((uint)1, (uint)1, (uint)1, (uint)10, (uint)3, (uint)78, @"<c r=""A1"" s=""78""><f>SUM(C1:C10)</f></c>")]
    public void WriteSumFormula_CorrectCellGenerated(uint cellRow, uint cellColumn, uint rowStart, uint rowEnd, uint formulaColumn, uint? style, string expectedOutput)
    {
        TestHelper.TestCellWrite(cellRow, expectedOutput, (sheet) =>
        {
            sheet.WriteSum(cellColumn, formulaColumn, rowStart, rowEnd, style: style);
        });
    }
}