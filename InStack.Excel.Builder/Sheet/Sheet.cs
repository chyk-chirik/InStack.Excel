using InStack.Excel.Builder;
using InStack.Excel.Builder.Extensions;
using System.Buffers;
using System.Buffers.Text;
using System.Runtime.CompilerServices;
using System.Text;

namespace ExcelUtils.Builder.RowExtensions;

public sealed partial class Sheet : IDisposable
{
    private readonly SheetConfig _config;
    private readonly StandardFormat _floatRowHeightFormat = new('F', 2);
    private readonly StreamBufferedWrapper _writer;

    public uint Row { get; private set; }
    public uint Column { get; private set; }

    public Sheet(Stream stream, SheetConfig config)
    {
        _writer = new StreamBufferedWrapper(stream, config.BufferSizeInBytes);
        _config = config;

        _writer.Write(@"<?xml version=""1.0"" encoding=""utf-8""?>
<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">"u8);

        if (config.Columns.Count > 0)
        {
            _writer.Write("<cols>"u8);

            foreach (var column in config.Columns)
            {
                _writer.Write(Encoding.UTF8.GetBytes(
                    $"<col min=\"{column.Min}\" max=\"{column.Max}\" width=\"{column.Width}\" customWidth=\"1\" />"));
            }

            _writer.Write("</cols>"u8);
        }
        
        _writer.Write("<sheetData>"u8);
    }

    public void StartRow(uint? row = null, uint? column = null, double? height = null)
    {
        Row = row ?? Row + 1;
        Column = column ?? 1;

        _writer.Write("<row r=\""u8);

        _writer.Format((buffer) => UintByteFormatter.Format(buffer, Row), 8);

        _writer.Write("\""u8);

        if (height.HasValue)
        {
            _writer.Write(" customHeight=\"1\" ht=\""u8);

            _writer.Format((buffer) => {
                Utf8Formatter.TryFormat(height.Value, buffer, out var bytesWritten, _floatRowHeightFormat);
                return bytesWritten;
            }, 64);

            _writer.Write("\""u8);
        }

        _writer.Write(">"u8);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void EndRow()
    {
        _writer.Write("</row>"u8);
    }

    public void EndRowAndStartNew(uint? row = null, uint? column = null, double? height = null)
    {
        EndRow();
        StartRow(row, column, height);
    }

    public void Dispose()
    {
        //	<autoFilter ref="A1:E5" />

        _writer.Write("</sheetData>"u8);

        _writer.FlushBuffer();

        _mergeCellManager.Write(_writer._stream);

        _writer.Write("</worksheet>"u8);
        _writer.Dispose();
        _mergeCellManager.Dispose();
    }
}