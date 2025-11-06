using System;
using System.Runtime.CompilerServices;

namespace InStack.Excel.Builder.Extensions
{
    internal static class CellBuilderExtensions
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static void FormatCellRefAndStyle(this StreamBuffer wrapper, uint row, uint column, uint? style)
        {
            var span = wrapper.AsSpan(64);

            var bytesWritten = CellReferenceFormatter.Format(span, row, column);

            if (style.HasValue)
            {
                "\" s=\""u8.CopyTo(span.Slice(bytesWritten));
                bytesWritten += 5;

                bytesWritten += UintByteFormatter.Format(span.Slice(bytesWritten), style.Value);
            }

            wrapper.SpanUsed(bytesWritten);
        }
    }
}
