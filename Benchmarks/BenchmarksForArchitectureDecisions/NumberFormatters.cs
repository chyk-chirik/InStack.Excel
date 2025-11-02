using BenchmarkDotNet.Attributes;
using InStack.Excel.Builder.Extensions;
using System.Globalization;
using System.Numerics;
using System.Runtime.CompilerServices;

namespace Benchmarks.BenchmarksForArchitectureDecisions
{
    public class NumberFormatters
    {
        public const int Iterations = 1000000;

        [Benchmark]
        public void TryFormat_UintAsInumber()
        {
            Span<byte> span = stackalloc byte[36];

            for (uint i = 0; i < Iterations; i++)
            {
                FormatINumber(span, i);
            }
        }

        [Benchmark]
        public void TryFormat_Uint()
        {
            Span<byte> span = stackalloc byte[36];

            for (uint i = 0; i < Iterations; i++)
            {
                FormatINumber(span, i);
            }
        }
       
        [Benchmark]
        public void CustomFormat_Uint()
        {
            Span<byte> span = stackalloc byte[36];

            for (uint i = 0; i < Iterations; i++)
            {
                UintByteFormatter.Format(span, i);
            }
        }

        [Benchmark]
        public void TryFormat_DecimalAsInumber()
        {
            Span<byte> span = stackalloc byte[36];

            for (decimal i = 0.22M; i < Iterations; i++)
            {
                FormatINumber(span, i);
            }
        }

        [Benchmark]
        public void TryFormat_Decimal()
        {
            Span<byte> span = stackalloc byte[36];

            for (decimal i = 0.22M; i < Iterations; i++)
            {
                FormatDecimal(span, i);
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void FormatINumber<T>(Span<byte> span, INumber<T> number)
            where T: struct, INumber<T>
        {
            number.TryFormat(span, out var bytesWritten, ['G'], CultureInfo.InvariantCulture);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void FormatInt(Span<byte> span, int number)
        {
            number.TryFormat(span, out var bytesWritten, ['G'], CultureInfo.InvariantCulture);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void FormatDecimal(Span<byte> span, decimal number)
        {
            number.TryFormat(span, out var bytesWritten, ['G'], CultureInfo.InvariantCulture);
        }
    }
}
