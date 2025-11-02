using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using InStack.Excel.OpenXmlStyles;

namespace Examples.MonthlyReportExample;

public class TableHeaderStyles(
    string bgColor, 
    string borderColor,
    string fontColor,
    BorderStyleValues borderStyle) : Style
{
    public uint NoBorder { get; private set; }
    public uint LeftAndBottom { get; private set; }
    public uint Left { get; private set; }

    protected override List<Font> GetFonts() =>
    [
        new(
            new Bold(),
            new Color() {
                Rgb = new HexBinaryValue() { Value = fontColor }
            }
        )
    ];

    protected override List<Fill> GetFills() =>
    [
        new(new PatternFill
        {
            PatternType = PatternValues.Solid,
            ForegroundColor = new() { Rgb = HexBinaryValue.FromString(bgColor) }
        })
    ];

    protected override List<Border> GetBorders() =>
    [
         new Border(
            new LeftBorder{
                Style = borderStyle,
                Color = new Color{
                    Rgb = new HexBinaryValue() { Value = borderColor }
                }
            },
            new BottomBorder{
                Style = borderStyle,
                Color = new Color{
                    Rgb = new HexBinaryValue() { Value = borderColor }
                }
            }),
         new Border(
            new LeftBorder{
                Style = borderStyle,
                Color = new Color{
                    Rgb = new HexBinaryValue() { Value = borderColor }
                }
            })
    ];

    protected override List<CellFormat> GetCellFormats() =>
    [
        new()
        {
            FontId = 1,
            ApplyFont = true,
            FillId = 1,
            ApplyFill = true,
            Alignment = new Alignment{
                Horizontal = HorizontalAlignmentValues.Center,
                Vertical = VerticalAlignmentValues.Center
            },
            ApplyAlignment = true,
            BorderId = 1,
            ApplyBorder = true,
        },
        new()
        {
            FontId = 1,
            ApplyFont = true,
            FillId = 1,
            ApplyFill = true,
            Alignment = new Alignment{
                Horizontal = HorizontalAlignmentValues.Center,
                Vertical = VerticalAlignmentValues.Center
            },
            ApplyAlignment = true,
            BorderId = 2,
            ApplyBorder = true,
        },
        new()
        {
            FontId = 1,
            ApplyFont = true,
            FillId = 1,
            ApplyFill = true,
            Alignment = new Alignment{
                Horizontal = HorizontalAlignmentValues.Center,
                Vertical = VerticalAlignmentValues.Center
            },
            ApplyAlignment = true,
        }
    ];

    protected override void BaseIndexUpdated()
    {
        LeftAndBottom = BaseIndex + 1;
        Left = BaseIndex + 2;
        NoBorder = BaseIndex + 3;
    }
}
