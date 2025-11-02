using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using InStack.Excel.OpenXmlStyles;

namespace Examples.MonthlyReportExample;

public class TableStyles(string evenRowBgColor, string oddRowBgColor) : Style
{
    public uint Even { get; private set; }
    public uint EvenCategory { get; private set; }
    public uint Odd { get; private set; }
    public uint OddCategory { get; private set; }
    public uint Total { get; private set; }
    public uint Divider { get; private set; }

    protected override List<Font> GetFonts() =>
    [
        new(new Bold())
    ];
    protected override List<Fill> GetFills() =>
    [
        new(new PatternFill
            {
                PatternType = PatternValues.Solid,
                ForegroundColor = new() { Rgb = HexBinaryValue.FromString(evenRowBgColor) }
            }
        ),
         new(
            new PatternFill
            {
                PatternType = PatternValues.Solid,
                ForegroundColor = new() { Rgb = HexBinaryValue.FromString(oddRowBgColor) }
            }
        )
    ];

    protected override List<Border> GetBorders() =>
    [
        new Border(
        new TopBorder{
            Style = BorderStyleValues.Thin,
        },
        new BottomBorder{
            Style = BorderStyleValues.Thin,
        })
    ];

    protected override List<CellFormat> GetCellFormats() =>
    [
        new()
        {
            FillId = 1,
            ApplyFill = true
        },
        new()
        {
            FillId = 2,
            ApplyFill = true
        },
        new()
        {
            BorderId = 1,
            ApplyBorder = true
        },
        new()
        {
            FontId = 1,
            ApplyFont = true,
            FillId = 1,
            ApplyFill = true
        },
        new()
        {
            FontId = 1,
            ApplyFont = true,
            FillId = 2,
            ApplyFill = true
        },
        new()
        {
            FontId = 1,
            ApplyFont = true,
        }
    ];

    protected override void BaseIndexUpdated()
    {
        Even = BaseIndex + 1;
        Odd = BaseIndex + 2;
        Divider = BaseIndex + 3;
        EvenCategory = BaseIndex + 4;
        OddCategory = BaseIndex + 5;
        Total = BaseIndex + 6;
    }
}
