using ExcelUtils.Builder.RowExtensions;
using System.Text;

namespace InStack.Excel.Builder.Extensions.Cell;

public static class FormulaExtensions
{
    public static void WriteNumberFormula(this Sheet sheet, ReadOnlySpan<byte> formula, uint column, uint? style = null)
    {
        sheet.Writer.Write("<c r=\""u8);

        sheet.Writer.FormatCellRefAndStyle(sheet.Row, column, style);

        sheet.Writer.Write("\"><f>"u8);

        sheet.Writer.Write(formula);

        sheet.Writer.Write("</f></c>"u8);
    }

    public static void WriteSum(this Sheet sheet, uint column, uint formulaColumn, uint rowStart, uint rowEnd, uint? style = null)
    {
        Span<byte> formula = stackalloc byte[16];

        "SUM("u8.CopyTo(formula);

        var bytesWritten = CellReferenceFormatter.Format(formula[4..], rowStart, formulaColumn) + 4;
        formula[bytesWritten++] = (byte)':';
        bytesWritten += CellReferenceFormatter.Format(formula[bytesWritten..], rowEnd, formulaColumn);
        formula[bytesWritten++] = (byte)')';
      
        sheet.WriteNumberFormula(formula[..bytesWritten], column, style: style);
    }
}
