using ExcelUtils.Builder.RowExtensions;
using InStack.Excel.Builder;
using System.Reflection;
using System.Text;

namespace Tests.Helpers;

internal class BufferHelper
{
    StreamBuffer wrapper;
    public BufferHelper(Sheet sheet)
    {
        wrapper = TestHelper.GetField<StreamBuffer, Sheet>(sheet, "_writer");
        _buffer = TestHelper.GetField<byte[], StreamBuffer>(wrapper, "_buffer");
    }
    private int _startPosition;
    private byte[] _buffer;

    internal void StartTrackChanges()
    {
        _startPosition =  TestHelper.GetField<int, StreamBuffer>(wrapper, "_position");
    }

    internal string GetStringResult()
    {
        var endPosition = TestHelper.GetField<int, StreamBuffer>(wrapper, "_position");

        return Encoding.UTF8.GetString(_buffer.AsSpan(_startPosition, endPosition - _startPosition));
    }

    internal void FlushBuffer()
    {
        MethodInfo? method = typeof(StreamBuffer).GetMethod(
           "FlushBuffer",
           BindingFlags.Instance | BindingFlags.NonPublic);

        method!.Invoke(wrapper, []);
    }
}
