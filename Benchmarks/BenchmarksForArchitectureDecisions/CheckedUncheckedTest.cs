using BenchmarkDotNet.Attributes;

namespace Benchmarks.BenchmarksForArchitectureDecisions;


public class CheckedUncheckedTest
{

    [Benchmark]
    public void Format6DigitIntegerChecked()
    {
        Span<byte> buffer = stackalloc byte[6];
        uint val = 123456;

        for (int i = 0; i < 100000000; i++)
        {
            buffer[5] = (byte)(val % 10 + 48);
            buffer[4] = (byte)((val /= 10) % 10 + 48);
            buffer[3] = (byte)((val /= 10) % 10 + 48);
            buffer[2] = (byte)((val /= 10) % 10 + 48);
            buffer[1] = (byte)((val /= 10) % 10 + 48);
            buffer[0] = (byte)((val /= 10) % 10 + 48);
        }
    }


    [Benchmark]
    public void Format6DigitIntegerUnchecked()
    {
        Span<byte> buffer = stackalloc byte[6];
        uint val = 123456;

        for (int i = 0; i < 100000000; i++)
        {
            unchecked
            {
                buffer[5] = (byte)(val % 10 + 48);
                buffer[4] = (byte)((val /= 10) % 10 + 48);
                buffer[3] = (byte)((val /= 10) % 10 + 48);
                buffer[2] = (byte)((val /= 10) % 10 + 48);
                buffer[1] = (byte)((val /= 10) % 10 + 48);
                buffer[0] = (byte)((val /= 10) % 10 + 48);
            }
        }
    }
}
