using InStack.Excel.Builder.Extensions;

namespace ExcelUtils.Builder.RowExtensions;

public sealed partial class Sheet
{
    public void WriteNumberFormula(ReadOnlySpan<byte> formula, uint shift = 0, uint? style = null)
    {
        _writer.Write("<c r=\""u8);

        _writer.FormatCellRefAndStyle(Row, Column, style);

        _writer.Write("\"><f>"u8);

        _writer.Write(formula);

        _writer.Write("</f></c>"u8);

        Column++;
    }
}
