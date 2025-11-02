using System.Buffers;
using System.Runtime.CompilerServices;
using System.Text.Unicode;

namespace InStack.Excel.Builder;

public sealed class StreamBufferedWrapper : IDisposable
{
    private readonly ArrayPool<byte> _pool;
    internal readonly Stream _stream;
    private byte[] _buffer;
    private int _position;

    public StreamBufferedWrapper(Stream zipStream, int bufferSize)
    {
        
        _pool = ArrayPool<byte>.Shared;
        _buffer = _pool.Rent(bufferSize);
        _stream = zipStream;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal void Write(ReadOnlySpan<char> valueSpan)
    {
        var reuseBuffer = _buffer.AsSpan(_position);

        while (true)
        {
            Utf8.FromUtf16(valueSpan, reuseBuffer, out var charsRead, out var bytesWritten);

            _position += bytesWritten;

            if (charsRead < valueSpan.Length)
            {
                FlushBuffer();
                reuseBuffer = _buffer.AsSpan();
                valueSpan = valueSpan.Slice(charsRead);
            }
            else
            {
                return;
            }
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal void Write(ReadOnlySpan<byte> bytes)
    {
        var maxBytesToWrite = _buffer.Length - _position;

        if (bytes.Length < maxBytesToWrite)
        {
            bytes.CopyTo(_buffer.AsSpan(_position));
            _position += bytes.Length;
        }
        else if (bytes.Length > _buffer.Length)
        {
            FlushBuffer();
            _stream.Write(bytes);
        }
        else
        {
            bytes
                .Slice(0, maxBytesToWrite)
                .CopyTo(_buffer.AsSpan(_position));
            _position += maxBytesToWrite;

            FlushBuffer();

            bytes
                .Slice(maxBytesToWrite)
                .CopyTo(_buffer);

            _position += bytes.Length - maxBytesToWrite;
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal void Format(Func<Span<byte>, int> formatter, int minLength)
    {
        if (_buffer.Length - _position <= minLength)
        {
            FlushBuffer();
        }

        _position += formatter(_buffer.AsSpan(_position));
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal void FlushBuffer()
    {
        _stream.Write(_buffer, 0, _position);
        _position = 0;
    }

    public void Dispose()
    {
        FlushBuffer();
        _pool.Return(_buffer);
        _buffer = null!;
        _stream.Dispose();
    }
}
