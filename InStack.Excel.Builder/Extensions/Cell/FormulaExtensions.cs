using ExcelUtils.Builder.RowExtensions;

namespace InStack.Excel.Builder.Extensions.Cell;

public static class FormulaExtensions
{
    public static void WriteNumberFormula(this Sheet sheet, ReadOnlySpan<byte> formula, uint shift = 0, uint? style = null)
    {
        sheet._writer.Write("<c r=\""u8);

        sheet._writer.FormatCellRefAndStyle(sheet.Row, sheet.Column, style);

        sheet._writer.Write("\"><f>"u8);

        sheet._writer.Write(formula);

        sheet._writer.Write("</f></c>"u8);

        sheet.Column++;
    }
}
