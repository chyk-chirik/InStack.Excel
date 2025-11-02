using BenchmarkDotNet.Attributes;

namespace Benchmarks.BenchmarksForArchitectureDecisions;


public class SliceVsOffsetTest
{
    public const int Iterations = 1000000;

    [Benchmark]
    public void FormatterWithOffset()
    {
        Span<byte> buffer = stackalloc byte[16];
        uint val = 123456;

        for (int i = 0; i < Iterations; i++)
        {
            SomeFormatterWithOffset(buffer, 6, val, out var bytesWritten);
        }
    }


    [Benchmark]
    public void FormatterNoOffset()
    {
        Span<byte> buffer = stackalloc byte[16];
        uint val = 123456;

        for (int i = 0; i < Iterations; i++)
        {
            SomeFormatterNoOffset(buffer.Slice(6), val, out var bytesWritten);
        }
    }

    private void SomeFormatterNoOffset(Span<byte> slicedBuffer, uint val, out int bytesWritten)
    {
        slicedBuffer[5] = (byte)(val % 10 + 48);
        slicedBuffer[4] = (byte)((val /= 10) % 10 + 48);
        slicedBuffer[3] = (byte)((val /= 10) % 10 + 48);
        slicedBuffer[2] = (byte)((val /= 10) % 10 + 48);
        slicedBuffer[1] = (byte)((val /= 10) % 10 + 48);
        slicedBuffer[0] = (byte)((val /= 10) % 10 + 48);

        bytesWritten = 6;
    }

    private void SomeFormatterWithOffset(Span<byte> buffer, int offset, uint val, out int bytesWritten)
    {
        buffer[5 + offset] = (byte)(val % 10 + 48);
        buffer[4 + offset] = (byte)((val /= 10) % 10 + 48);
        buffer[3 + offset] = (byte)((val /= 10) % 10 + 48);
        buffer[2 + offset] = (byte)((val /= 10) % 10 + 48);
        buffer[1 + offset] = (byte)((val /= 10) % 10 + 48);
        buffer[offset] = (byte)((val /= 10) % 10 + 48);

        bytesWritten = 6;
    }
}
