using DocumentFormat.OpenXml.Spreadsheet;

namespace InStack.Excel.OpenXmlStyles;

public class SystemStyles : Style
{
    public static readonly uint DefaultStyleIndex = 0;

    protected override List<Font> GetFonts() => [new ()];

    protected override List<Fill> GetFills() =>
    [
        new(new PatternFill { PatternType = PatternValues.None }),
        new(new PatternFill { PatternType = PatternValues.Gray125 })
    ];

    protected override List<Border> GetBorders() => [new ()];

    protected override List<CellFormat> GetCellFormats() => [new ()];
    protected override void BaseIndexUpdated() {}
}