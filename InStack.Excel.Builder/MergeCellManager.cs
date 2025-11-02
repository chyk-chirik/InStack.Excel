using InStack.Excel.Builder.Extensions;
using System.Buffers;

namespace InStack.Excel.Builder;


public sealed class MergeCellManager: IDisposable
{
    private readonly ArrayPool<uint> _pool = ArrayPool<uint>.Shared;
    private readonly List<RentedArray> _chunks = new List<RentedArray>(32);

    public MergeCellManager(int initialCapacityInBytes = 1024)
    {
        _chunks.Add(new RentedArray(_pool.Rent(initialCapacityInBytes)));
    }

    public void Add(uint rowStart, uint columnStart, uint rowEnd, uint columnEnd)
    {
        var currentChunk = _chunks[_chunks.Count - 1];

        if (currentChunk.IsFull())
        {
             currentChunk = new RentedArray(_pool.Rent(currentChunk.Length * 2));
            _chunks.Add(currentChunk);
        }

        currentChunk.Add(rowStart);
        currentChunk.Add(columnStart);
        currentChunk.Add(rowEnd);
        currentChunk.Add(columnEnd);
    }

    public void Dispose()
    {
        for (var chunkIndex = 0; chunkIndex < _chunks.Count; chunkIndex++)
        {
            _pool.Return(_chunks[chunkIndex].Array);
        }
    }

    internal void Write(Stream _doc)
    {
        var lastChunk = _chunks[_chunks.Count - 1];

        if (_chunks.Count == 1 && lastChunk.IsEmpty())
            return;

        _doc.Write("<mergeCells>"u8);

        Span<byte> mergeCellBytes = stackalloc byte[64];
        "<mergeCell ref=\""u8.CopyTo(mergeCellBytes);

        for (var chunkIndex = 0; chunkIndex < _chunks.Count; chunkIndex++)
        {
            var chunk = _chunks[chunkIndex];
            var lastIndexInChunk = _chunks.Count - 1 == chunkIndex
                ? chunk.Position
                : chunk.Length;

            var nextIndex = 16;

            for (var lastIndexInBlock = 3; lastIndexInBlock <= lastIndexInChunk; lastIndexInBlock += 4)
            {
                
                nextIndex += CellReferenceFormatter.Format(mergeCellBytes[nextIndex..], chunk[lastIndexInBlock - 3], chunk[lastIndexInBlock - 2]);

                mergeCellBytes[nextIndex++] = (byte)':';
                
                nextIndex += CellReferenceFormatter.Format(mergeCellBytes[nextIndex..], chunk[lastIndexInBlock - 1], chunk[lastIndexInBlock]);

                "\"/>"u8.CopyTo(mergeCellBytes[nextIndex..]);
                nextIndex += 3;

                _doc.Write(mergeCellBytes[..nextIndex]);
                nextIndex = 16;
            }
        }

        _doc.Write("</mergeCells>"u8);
    }

    internal class RentedArray(uint[] array)
    {
        public uint[] Array => array;
        public int Length { get; } = array.Length - array.Length % 4;
        public int Position { get; set; } = -1;

        public void Add(uint value)
        {
            array[++Position] = value;
        }

        public bool IsFull() => array.Length == Position + 1;
        public bool IsEmpty() => Position == -1;

        public uint this[int index]
        {
            get => array[index];
        }
    }
}
