namespace InStack.Excel.Builder;

public record ColumnDescriptor(int Min, int Max, int Width);

public class SheetConfig
{
    public List<ColumnDescriptor> Columns = [];
    public int BufferSizeInBytes { get; set; } = 4096;
}