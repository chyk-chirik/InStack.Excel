using BenchmarkDotNet.Attributes;

namespace Benchmarks.BenchmarksForArchitectureDecisions;


public class SliceVsAsSpanTest
{
    private byte[] heapArray = new byte[1000];

    public const int Iterations = 1000000;

    [Benchmark]
    public void Slice()
    {
        Span<byte> buffer = heapArray.AsSpan(100);

        for (int i = 0; i < Iterations; i++)
        {
            buffer.Slice(10);
        }
    }


    [Benchmark]
    public void AsSpan()
    {
        for (int i = 0; i < Iterations; i++)
        {
            heapArray.AsSpan(100);
        }
    }
}
