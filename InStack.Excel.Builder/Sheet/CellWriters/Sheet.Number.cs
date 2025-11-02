using InStack.Excel.Builder.Extensions;
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

        _writer.Write("<c t=\"n\" r=\""u8);
        _writer.FormatCellRefAndStyle(Row, Column, style);

        if (number is null)
        {
            _writer.Write("\"/>"u8);
        }
        else
        {
            _writer.Write("\"><v>"u8);

            _writer.Format((buffer) => {
                number.Value.TryFormat(buffer, out var bytesWritten, ['G'], CultureInfo.InvariantCulture);
                return bytesWritten;
            }, 64);

            _writer.Write("</v></c>"u8);
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
