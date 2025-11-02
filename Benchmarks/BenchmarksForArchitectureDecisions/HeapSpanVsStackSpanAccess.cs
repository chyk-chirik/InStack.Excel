using BenchmarkDotNet.Attributes;
using System.Runtime.CompilerServices;

namespace Benchmarks.BenchmarksForArchitectureDecisions;


public class HeapSpanVsStackSpanAccess
{
    public const int Iterations = 1000000;

    private byte[] heapArray = new byte[1000];


    [Benchmark]
    public void StackIndexerAccess()
    {
        Span<byte> span = stackalloc byte[6];

        for (int i = 0; i < Iterations; i++)
        {
            span[0] = (byte)'1';
            span[1] = (byte)'2';
            span[2] = (byte)'3';
            span[3] = (byte)'4';
            span[4] = (byte)'5';
            span[5] = (byte)'6';
        }
    }


    [Benchmark]
    public void HeapIndexerAccess()
    {
        Span<byte> span = heapArray.AsSpan(100, 6);

        for (int i = 0; i < Iterations; i++)
        {
            span[0] = (byte)'1';
            span[1] = (byte)'2';
            span[2] = (byte)'3';
            span[3] = (byte)'4';
            span[4] = (byte)'5';
            span[5] = (byte)'6';
        }
    }

    [Benchmark]
    public void CopyTo_FromStackToStack()
    {
        Span<byte> span = stackalloc byte[6];

        for (int i = 0; i < Iterations; i++)
        {
            "123456"u8.CopyTo(span);
        }
    }


    [Benchmark]
    public void CopyTo_FromStackToHeap()
    {
        Span<byte> span = heapArray.AsSpan(100, 6);

        for (int i = 0; i < Iterations; i++)
        {
            "123456"u8.CopyTo(span);
        }
    }

    [Benchmark]
    public void CopyTo_FromHeapToHeap()
    {
        Span<byte> spanSource = heapArray.AsSpan(100, 6);
        Span<byte> spanDest = heapArray.AsSpan(106, 6);

        for (int i = 0; i < Iterations; i++)
        {
            spanSource.CopyTo(spanDest);
        }
    }

    [Benchmark]
    public void CopyTo_FromHeapToStack()
    {
        Span<byte> spanSource = heapArray.AsSpan(100, 6);
        Span<byte> spanDest = stackalloc byte[6];

        for (int i = 0; i < Iterations; i++)
        {
            spanSource.CopyTo(spanDest);
        }
    }

    [Benchmark]
    public void MemCopy_FromHeapToHeap()
    {
        unsafe
        {
            fixed (byte* source = &heapArray[100])
            fixed (byte* dest = &heapArray[106])
            {
                for (int i = 0; i < Iterations; i++)
                {
                    Buffer.MemoryCopy(source, dest, 6, 6);
                }
            }
        }
    }

    [Benchmark]
    public void CopyBlock_FromHeapToHeap()
    {
        unsafe
        {
            fixed (byte* source = &heapArray[100])
            fixed (byte* dest = &heapArray[106])
            {
                for (int i = 0; i < Iterations; i++)
                {
                    Unsafe.CopyBlock(dest, source, 6);
                }
            }
        }
    }

    [Benchmark]
    public void CycleCopy_FromHeapToHeap()
    {
        for (int i = 0; i < Iterations; i++)
        {
            for (int j = 0; j < 6; j++)
            {
                heapArray[100 + j] = heapArray[106 + j];
            }
        }
    }

    [Benchmark]
    public void CycleCopy_FromStackToHeap()
    {
        Span<byte> span = heapArray.AsSpan(100, 6);

        for (int i = 0; i < Iterations; i++)
        {
            for (int j = 0; j < 6; j++)
            {
                heapArray[100 + j] = span[j];
            }
        }
    }
}
