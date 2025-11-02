using InStack.Excel.Builder;

namespace ExcelUtils.Builder.RowExtensions;

/// <summary>
/// Since top left cell wins, only functionality to merge right and bottom exposed. 
/// </summary>
public sealed partial class Sheet
{
    private readonly MergeCellManager _mergeCellManager = new MergeCellManager();

    /// <summary>
    /// Merges previous cell with cells standing to the right. To avoid styling collisions
    /// empty cells created here, they should have same style as the main cell.
    /// </summary>
    /// <param name="count">Amount of cells to merge with main</param>
    /// <param name="style">Style of the main cell</param>
    public void MergePrevCellToRight(uint count = 1, uint? style = null)
    {
        _mergeCellManager.Add(
            rowStart: Row,
            columnStart: Column - 1,
            rowEnd: Row,
            columnEnd: Column + count - 1);

        WriteEmpty(count:  count, style: style);
    }

    /// <summary>
    /// Merges previous cell with cells standing to the bottom. To avoid styling collisions 
    /// they should be created with the same style as the main cell.
    /// </summary>
    /// <param name="count">Amount of cells to merge with main</param>
    /// <param name="style">Style of the main cell</param>
    public void MergePrevCellToBottom(uint count = 1)
    {
        _mergeCellManager.Add(
            rowStart: Row,
            columnStart: Column - 1,
            rowEnd: Row + count,
            columnEnd: Column - 1);
    }

    /// <summary>
    /// Merges a range of cells, where previous (main) cell stands in the left top corner.
    /// To avoid styling collision all cells in range should be created with the same style as the main cell.
    /// In this method only cells withing same row are created.
    /// </summary>
    /// <param name="rightCount"></param>
    /// <param name="bottomCount"></param>
    /// <param name="style">Style of the main cell</param>
    public void MergePrevCellToRightAndBottom(uint rightCount, uint bottomCount, uint? style = null)
    {
        _mergeCellManager.Add(
            rowStart: Row,
            columnStart: Column - 1,
            rowEnd: Row + bottomCount,
            columnEnd: Column + rightCount - 1);

        WriteEmpty(count: rightCount, style: style);
    }
}
