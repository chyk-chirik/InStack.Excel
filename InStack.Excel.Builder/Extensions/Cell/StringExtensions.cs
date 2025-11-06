using ExcelUtils.Builder.RowExtensions;
using System.Runtime.CompilerServices;

namespace InStack.Excel.Builder.Extensions.Cell;

public static class StringExtensions
{
    public static void Write(this Sheet sheet, string? value, uint shift = 0, uint? style = null, bool escape = false)
    {
        sheet.Column += shift;

        sheet._writer.FlushBufferIfNoSpace(64);

        sheet._writer.WriteUnsafe("<c t=\"inlineStr\" r=\""u8);
        sheet._writer.FormatCellRefAndStyle(sheet.Row, sheet.Column, style);

        if (string.IsNullOrEmpty(value))
        {
            sheet._writer.WriteUnsafe("\"/>"u8);
        }
        else
        {
            sheet._writer.WriteUnsafe("\"><is><t>"u8);

            if (escape)
            {
                Escape(sheet, value.AsSpan());
            }
            else
            {
                sheet._writer.Write(value.AsSpan());
            }

            sheet._writer.WriteUnsafe("</t></is></c>"u8);
        }
        sheet.Column++;
    }

    private static void Escape(Sheet sheet, ReadOnlySpan<char> valueSpan)
    {
        var rangeIndexStart = 0;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        void Copy(ReadOnlySpan<char> replacement, int currentPosition, ref ReadOnlySpan<char> valueSpan)
        {
            if (rangeIndexStart != currentPosition)
            {
                sheet._writer.Write(valueSpan.Slice(rangeIndexStart, currentPosition - rangeIndexStart));
            }

            sheet._writer.Write(replacement);
            rangeIndexStart = currentPosition + 1;
        }

        for (var i = 0; i < valueSpan.Length; i++) 
        {
            if (valueSpan[i] == '<')
            {
                Copy(['&', 'l', 't', ';'], i, ref valueSpan);
            }
            else if (valueSpan[i] == '>')
            {
                Copy(['&', 'g', 't', ';'], i, ref valueSpan);
            }
            else if (valueSpan[i] == '&')
            {
                Copy(['&', 'a', 'm', 'p', ';'], i, ref valueSpan);
            }
            else if (valueSpan[i] == '"')
            {
                Copy(['&', 'q', 'u', 'o', 't', ';'], i, ref valueSpan);
            }
            else if (valueSpan[i] == '\'')
            {
                Copy(['&', 'a', 'p', 'o', 's', ';'], i, ref valueSpan);
            }
        }

        if(rangeIndexStart != valueSpan.Length)
        {
            sheet._writer.Write(valueSpan.Slice(rangeIndexStart, valueSpan.Length - rangeIndexStart));
        }
    }
}
