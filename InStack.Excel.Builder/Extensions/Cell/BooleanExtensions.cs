using ExcelUtils.Builder.RowExtensions;

namespace InStack.Excel.Builder.Extensions.Cell;

public static class BooleanExtensions
{
    public static void WriteBool(this Sheet sheet, bool? value, uint shift = 0, uint? style = null)
    {
        sheet.Column += shift;

        sheet._writer.WriteUnsafe("<c t=\"b\" r=\""u8);
        sheet._writer.FormatCellRefAndStyle(sheet.Row, sheet.Column, style);

        if (value is null)
        {
            sheet._writer.Write("\"/>"u8);
        }
        else
        {
            switch (value)
            {
                case true:
                    sheet._writer.WriteUnsafe("\"><v>1</v></c>"u8);
                    break;
                case false:
                    sheet._writer.WriteUnsafe("\"><v>0</v></c>"u8);
                    break;
            }
        }

        sheet.Column++;
    }
}
