using ExcelUtils.Builder.RowExtensions;
using System.Globalization;
using System.Numerics;
using static InStack.Excel.Builder.Extensions.DateExtensions;

namespace InStack.Excel.Builder.Extensions.Cell;

public static class NumberExtensions
{
    public static void Write<T>(this Sheet sheet, T? number, uint column, uint? style = null)
      where T : struct, INumber<T>
    {
        sheet.Writer.FlushBufferIfNoSpace(64);

        sheet.Writer.WriteUnsafe("<c t=\"n\" r=\""u8);
        sheet.Writer.FormatCellRefAndStyle(sheet.Row, column, style);

        if (number is null)
        {
            sheet.Writer.WriteUnsafe("\"/>"u8);
        }
        else
        {
            sheet.Writer.WriteUnsafe("\"><v>"u8);

            number.Value.TryFormat(sheet.Writer.AsSpanUnsafe(), out var bytesWritten, ['G'], CultureInfo.InvariantCulture);
            sheet.Writer.SpanUsed(bytesWritten);

            sheet.Writer.WriteUnsafe("</v></c>"u8);
        }
    }

    public static void Write(this Sheet sheet, DateTime? date, uint shift = 0, uint? style = null)
    {
        if (date.HasValue)
        {
            sheet.Write(
                (long?)TicksToOaDate(date.Value.Ticks),
                shift,
                style: style);
        }
        else sheet.Write((long?)null, shift, style);
    }
}
