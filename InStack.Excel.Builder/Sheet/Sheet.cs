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
    public readonly StreamBuffer Writer;

    public uint Row { get; private set; }
    public uint Column;

    public Sheet(Stream stream, SheetConfig config)
    {
        Writer = new StreamBuffer(stream, config.BufferSizeInBytes);
        _config = config;

        Writer.Write(@"<?xml version=""1.0"" encoding=""utf-8""?>
<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">"u8);

        if (config.Columns.Count > 0)
        {
            Writer.Write("<cols>"u8);

            foreach (var column in config.Columns)
            {
                Writer.Write(Encoding.UTF8.GetBytes(
                    $"<col min=\"{column.Min}\" max=\"{column.Max}\" width=\"{column.Width}\" customWidth=\"1\" />"));
            }

            Writer.Write("</cols>"u8);
        }
        
        Writer.Write("<sheetData>"u8);
    }

    public void StartRow(uint? row = null, uint? column = null, double? height = null)
    {
        Row = row ?? Row + 1;
        Column = column ?? 1;

        Writer.Write("<row r=\""u8);

        Writer.SpanUsed(UintByteFormatter.Format(Writer.AsSpan(8), Row));

        Writer.Write("\""u8);

        if (height.HasValue)
        {
            Writer.Write(" customHeight=\"1\" ht=\""u8);

            Utf8Formatter.TryFormat(height.Value, Writer.AsSpan(16), out var bytesWritten, _floatRowHeightFormat);
            Writer.SpanUsed(bytesWritten);

            Writer.Write("\""u8);
        }

        Writer.Write(">"u8);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void EndRow()
    {
        Writer.Write("</row>"u8);
    }

    public void EndRowAndStartNew(uint? row = null, uint? column = null, double? height = null)
    {
        EndRow();
        StartRow(row, column, height);
    }

    public void Dispose()
    {
        //	<autoFilter ref="A1:E5" />

        Writer.Write("</sheetData>"u8);

        Writer.FlushBuffer();

        _mergeCellManager.Write(Writer._stream);

        Writer.Write("</worksheet>"u8);
        Writer.Dispose();
        _mergeCellManager.Dispose();
    }
}