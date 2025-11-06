using ExcelUtils.Builder.RowExtensions;
namespace InStack.Excel.Builder.Extensions.Cell;

public static class GenericExtensions
{
    public static void WriteEmpty(this Sheet sheet, uint shift = 0, uint count = 1, uint? style = null)
    {
        sheet.Column += shift;

        for (int i = 0; i < count; i++)
        {
            sheet.Writer.Write("<c r=\""u8);
            sheet.Writer.FormatCellRefAndStyle(sheet.Row, sheet.Column, style);
            sheet.Writer.Write("\"/>"u8);

            sheet.Column++;
        }
    }
}
