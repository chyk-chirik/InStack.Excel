using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using InStack.Excel.OpenXmlStyles;

namespace Examples.CollectionExample;

public class CollectionExampleStyles : Style
{
    public uint DateFormatStyleId { get; set; }
    protected override List<NumberingFormat> GetNumberingFormats() =>
    [
        new()
        {
            FormatCode = StringValue.FromString("yyyy-mm-dd")
        }
    ];

    protected override List<CellFormat> GetCellFormats() =>
    [
        new()
        {
            NumberFormatId = 1,
            ApplyNumberFormat = true
        }
    ];

    protected override void BaseIndexUpdated()
    {
        DateFormatStyleId = BaseIndex + 1;
    }
}