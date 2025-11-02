using System.Runtime.CompilerServices;

namespace InStack.Excel.Builder.Extensions
{
    internal static class CellBuilderExtensions
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static void FormatCellRefAndStyle(this StreamBufferedWrapper wrapper, uint row, uint column, uint? style)
        {
            wrapper.Format((buffer) =>
            {
                var bytesWritten = CellReferenceFormatter.Format(buffer, row, column);

                if (style.HasValue)
                {
                    "\" s=\""u8.CopyTo(buffer.Slice(bytesWritten));
                    bytesWritten += 5;

                    bytesWritten += UintByteFormatter.Format(buffer.Slice(bytesWritten), style.Value);
                }

                return bytesWritten;

            }, 32);
        }

    }
}
