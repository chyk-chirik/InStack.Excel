using BenchmarkDotNet.Attributes;

namespace Benchmarks.BenchmarksForArchitectureDecisions;

public class CreateSpanUsingAsSpanVsPointer
{
    public const int Iterations = 1000000;

    private byte[] heapArray = new byte[1000];


    [Benchmark]
    public void AsSpan()
    {
        for (int i = 0; i < Iterations; i++)
        {
            Span<byte> span = heapArray.AsSpan(100, 6);
        }
    }


    [Benchmark]
    public void Pointer()
    {
        for (int i = 0; i < Iterations; i++)
        {
            unsafe
            {
                fixed (byte* ptr = &heapArray[100])
                {
                    var span = new Span<byte>(ptr, 6);
                }
            }
        }
    }
}
