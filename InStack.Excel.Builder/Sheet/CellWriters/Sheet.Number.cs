using InStack.Excel.Builder.Extensions;
using System;
using System.Globalization;
using System.Numerics;
using static InStack.Excel.Builder.Extensions.DateExtensions;

namespace ExcelUtils.Builder.RowExtensions;

public sealed partial class Sheet
{
    public void Write<T>(T? number, uint shift = 0, uint? style = null)
      where T : struct, INumber<T>
    {
        Column += shift;

        _writer.FlushBufferIfNoSpace(64);

        _writer.WriteUnsafe("<c t=\"n\" r=\""u8);
        _writer.FormatCellRefAndStyle(Row, Column, style);

        if (number is null)
        {
            _writer.WriteUnsafe("\"/>"u8);
        }
        else
        {
            _writer.WriteUnsafe("\"><v>"u8);

            number.Value.TryFormat(_writer.AsSpanUnsafe(), out var bytesWritten, ['G'], CultureInfo.InvariantCulture);
            _writer.SpanUsed(bytesWritten);

            _writer.WriteUnsafe("</v></c>"u8);
        }
        Column++;
    }

    public void Write(DateTime? date, uint shift = 0, uint? style = null)
    {
        if (date.HasValue)
        {
            Write(
                (long?)TicksToOaDate(date.Value.Ticks),
                shift,
                style: style);
        }
        else Write((long?)null, shift, style);
    }
}
