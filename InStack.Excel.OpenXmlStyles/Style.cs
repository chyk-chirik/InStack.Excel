using DocumentFormat.OpenXml.Spreadsheet;

namespace InStack.Excel.OpenXmlStyles;

public abstract class Style
{
    protected uint BaseIndex;
    protected virtual List<NumberingFormat> GetNumberingFormats() { return []; }
    protected virtual List<Font> GetFonts() { return []; }
    protected virtual List<Fill> GetFills() { return []; }
    protected virtual List<Border> GetBorders() { return []; }
    protected virtual List<CellFormat> GetCellFormats() { return []; }

    protected abstract void BaseIndexUpdated();

    public static Stylesheet Compile(params Style[] styles)
    {
        var defaultStyle = new SystemStyles();

        var numberingFormats = defaultStyle.GetNumberingFormats(); // start custom IDs at 164+
        var fills = defaultStyle.GetFills();  // start custom IDs at 2+
        var borders = defaultStyle.GetBorders(); // start custom IDs at 1+
        var fonts = defaultStyle.GetFonts(); // start custom IDs at 1+
        var cellFormats = defaultStyle.GetCellFormats();

        foreach (var style in styles)
        {
            var currentStyleNumberingFormats = style.GetNumberingFormats();
            SetNumberFormatIds(currentStyleNumberingFormats, numberingFormats.Count);

            var currentStyleFills = style.GetFills();
            var currentStyleBorders = style.GetBorders();
            var currentStyleFonts = style.GetFonts();

            var currentStyleCellFormats = style.GetCellFormats();

            foreach (var currentStyleCellFormat in currentStyleCellFormats)
            {
                if ((currentStyleCellFormat.ApplyNumberFormat ?? false) && currentStyleCellFormat.NumberFormatId is not null)
                {
                    currentStyleCellFormat.NumberFormatId += (uint)(numberingFormats.Count + 163);
                }

                if ((currentStyleCellFormat.ApplyFill ?? false) && currentStyleCellFormat.FillId is not null)
                {
                    currentStyleCellFormat.FillId += (uint)(fills.Count - 1);
                }

                if ((currentStyleCellFormat.ApplyFont ?? false) && currentStyleCellFormat.FontId is not null)
                {
                    currentStyleCellFormat.FontId += (uint)(fonts.Count - 1);
                }

                if ((currentStyleCellFormat.ApplyBorder ?? false) && currentStyleCellFormat.BorderId is not null)
                {
                    currentStyleCellFormat.BorderId += (uint)(borders.Count - 1);
                }
            }

            style.BaseIndex = (uint)cellFormats.Count - 1;
            style.BaseIndexUpdated();

            numberingFormats.AddRange(currentStyleNumberingFormats);
            borders.AddRange(currentStyleBorders);
            fonts.AddRange(currentStyleFonts);
            fills.AddRange(currentStyleFills);
            cellFormats.AddRange(currentStyleCellFormats);
        }

        var stylesheet = new Stylesheet();
        stylesheet.Append(new NumberingFormats(numberingFormats));
        stylesheet.Append(new Fonts(fonts));
        stylesheet.Append(new Fills(fills));
        stylesheet.Append(new Borders(borders));
        stylesheet.Append(new CellFormats(cellFormats));

        return stylesheet;
    }

    private static void SetNumberFormatIds(IReadOnlyList<NumberingFormat> formats, int processedFormats)
    {
        for (var formatPosition = 1; formatPosition <= formats.Count; formatPosition++)
        {
            formats[formatPosition - 1].NumberFormatId = (uint)(processedFormats + formatPosition + 163);
        }
    }
}