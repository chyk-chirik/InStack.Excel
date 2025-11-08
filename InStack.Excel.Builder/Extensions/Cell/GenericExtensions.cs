using ExcelUtils.Builder.RowExtensions;
namespace InStack.Excel.Builder.Extensions.Cell;

public static class GenericExtensions
{
    public static void WriteEmpty(this Sheet sheet, uint column, uint count = 1, uint? style = null)
    {

        for (int i = 0; i < count; i++)
        {
            sheet.Writer.Write("<c r=\""u8);
            sheet.Writer.FormatCellRefAndStyle(sheet.Row, column++, style);
            sheet.Writer.Write("\"/>"u8);
        }
    }
}
