using ExcelUtils.Builder.RowExtensions;
using InStack.Excel.Builder.Extensions;
using System;
using System.Globalization;
using System.Numerics;
using static InStack.Excel.Builder.Extensions.DateExtensions;

namespace InStack.Excel.Builder.Extensions.Cell;

public static class NumberExtensions
{
    public static void Write<T>(this Sheet sheet, T? number, uint shift = 0, uint? style = null)
      where T : struct, INumber<T>
    {
        sheet.Column += shift;

        sheet._writer.FlushBufferIfNoSpace(64);

        sheet._writer.WriteUnsafe("<c t=\"n\" r=\""u8);
        sheet._writer.FormatCellRefAndStyle(sheet.Row, sheet.Column, style);

        if (number is null)
        {
            sheet._writer.WriteUnsafe("\"/>"u8);
        }
        else
        {
            sheet._writer.WriteUnsafe("\"><v>"u8);

            number.Value.TryFormat(sheet._writer.AsSpanUnsafe(), out var bytesWritten, ['G'], CultureInfo.InvariantCulture);
            sheet._writer.SpanUsed(bytesWritten);

            sheet._writer.WriteUnsafe("</v></c>"u8);
        }
        sheet.Column++;
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
