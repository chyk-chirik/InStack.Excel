using InStack.Excel.Builder.Extensions;

namespace ExcelUtils.Builder.RowExtensions;

public sealed partial class Sheet
{
    public void Write(bool? value, uint shift = 0, uint? style = null)
    {
        Column += shift;

        _writer.Write("<c t=\"b\" r=\""u8);
        _writer.FormatCellRefAndStyle(Row, Column, style);

        if (value is null)
        {
            _writer.Write("\"/>"u8);
        }
        else
        {
            switch (value)
            {
                case true:
                    _writer.Write("\"><v>1</v></c>"u8);
                    break;
                case false:
                    _writer.Write("\"><v>0</v></c>"u8);
                    break;
            }
        }

        Column++;
    }
}
