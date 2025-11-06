using InStack.Excel.Builder.Extensions;
using System.Runtime.CompilerServices;

namespace ExcelUtils.Builder.RowExtensions;

public sealed partial class Sheet
{
    public void Write(string? value, uint shift = 0, uint? style = null, bool escape = false)
    {
        Column += shift;

        _writer.FlushBufferIfNoSpace(64);

        _writer.WriteUnsafe("<c t=\"inlineStr\" r=\""u8);
        _writer.FormatCellRefAndStyle(Row, Column, style);

        if (string.IsNullOrEmpty(value))
        {
            _writer.WriteUnsafe("\"/>"u8);
        }
        else
        {
            _writer.WriteUnsafe("\"><is><t>"u8);

            if (escape)
            {
                Escape(value.AsSpan());
            }
            else
            {
                _writer.Write(value.AsSpan());
            }

            _writer.WriteUnsafe("</t></is></c>"u8);
        }
        Column++;
    }

    private void Escape(ReadOnlySpan<char> valueSpan)
    {
        var rangeIndexStart = 0;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        void Copy(ReadOnlySpan<char> replacement, int currentPosition, ref ReadOnlySpan<char> valueSpan)
        {
            if (rangeIndexStart != currentPosition)
            {
                _writer.Write(valueSpan.Slice(rangeIndexStart, currentPosition - rangeIndexStart));
            }

            _writer.Write(replacement);
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
            _writer.Write(valueSpan.Slice(rangeIndexStart, valueSpan.Length - rangeIndexStart));
        }
    }
}
