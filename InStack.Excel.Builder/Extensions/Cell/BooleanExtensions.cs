using ExcelUtils.Builder.RowExtensions;

namespace InStack.Excel.Builder.Extensions.Cell;

public static class BooleanExtensions
{
    public static void WriteBool(this Sheet sheet, bool? value, uint column, uint? style = null)
    {
        sheet.Writer.WriteUnsafe("<c t=\"b\" r=\""u8);
        sheet.Writer.FormatCellRefAndStyle(sheet.Row, column, style);

        if (value is null)
        {
            sheet.Writer.Write("\"/>"u8);
        }
        else
        {
            switch (value)
            {
                case true:
                    sheet.Writer.WriteUnsafe("\"><v>1</v></c>"u8);
                    break;
                case false:
                    sheet.Writer.WriteUnsafe("\"><v>0</v></c>"u8);
                    break;
            }
        }
    }
}
