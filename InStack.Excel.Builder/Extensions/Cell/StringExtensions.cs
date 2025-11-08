using ExcelUtils.Builder.RowExtensions;
using System.Runtime.CompilerServices;

namespace InStack.Excel.Builder.Extensions.Cell;

public static class StringExtensions
{
    public static void Write(this Sheet sheet, string? value, uint column, uint? style = null, bool escape = false)
    {
        sheet.Writer.FlushBufferIfNoSpace(64);

        sheet.Writer.WriteUnsafe("<c t=\"inlineStr\" r=\""u8);
        sheet.Writer.FormatCellRefAndStyle(sheet.Row, column, style);

        if (string.IsNullOrEmpty(value))
        {
            sheet.Writer.WriteUnsafe("\"/>"u8);
        }
        else
        {
            sheet.Writer.WriteUnsafe("\"><is><t>"u8);

            if (escape)
            {
                Escape(sheet, value.AsSpan());
            }
            else
            {
                sheet.Writer.Write(value.AsSpan());
            }

            sheet.Writer.Write("</t></is></c>"u8);
        }
    }

    private static void Escape(Sheet sheet, ReadOnlySpan<char> valueSpan)
    {
        var rangeIndexStart = 0;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        void Copy(ReadOnlySpan<char> replacement, int currentPosition, ref ReadOnlySpan<char> valueSpan)
        {
            if (rangeIndexStart != currentPosition)
            {
                sheet.Writer.Write(valueSpan.Slice(rangeIndexStart, currentPosition - rangeIndexStart));
            }

            sheet.Writer.Write(replacement);
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
            sheet.Writer.Write(valueSpan.Slice(rangeIndexStart, valueSpan.Length - rangeIndexStart));
        }
    }
}
