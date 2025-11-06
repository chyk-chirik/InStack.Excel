using ExcelUtils.Builder.RowExtensions;

namespace InStack.Excel.Builder.Extensions.Cell;

public static class FormulaExtensions
{
    public static void WriteNumberFormula(this Sheet sheet, ReadOnlySpan<byte> formula, uint shift = 0, uint? style = null)
    {
        sheet.Writer.Write("<c r=\""u8);

        sheet.Writer.FormatCellRefAndStyle(sheet.Row, sheet.Column, style);

        sheet.Writer.Write("\"><f>"u8);

        sheet.Writer.Write(formula);

        sheet.Writer.Write("</f></c>"u8);

        sheet.Column++;
    }
}
