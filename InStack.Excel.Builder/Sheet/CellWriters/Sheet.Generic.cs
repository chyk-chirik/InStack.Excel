using InStack.Excel.Builder.Extensions;

namespace ExcelUtils.Builder.RowExtensions;

public sealed partial class Sheet
{
    /// <summary>
    /// Creates empty cells un-typed cells
    /// </summary>
    /// <param name="shift">Current cell column increment</param>
    /// <param name="count"></param>
    /// <param name="style"></param>
    public void WriteEmpty(uint shift = 0, uint count = 1, uint? style = null)
    {
        Column += shift;

        for (int i = 0; i < count; i++)
        {
            _writer.Write("<c r=\""u8);
            _writer.FormatCellRefAndStyle(Row, Column, style);
            _writer.Write("\"/>"u8);

            Column++;
        }
    }
}
