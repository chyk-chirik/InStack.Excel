# InStack.Excel
Generate xlsx almost without memory allocations. Simplified work with styles

## Styles

Work with styles is always a pain, it's easy to get corrupted file and spend a lot of time in troubleshoot. This was solved in next way:

- Split styles across several classes for usability and reusability, like MyTableHeaderStyles, MyTableSpecificRowStyles ..
- When you reference specific style in CellFormat object you reference it's local (relative to class) position. 

  Look at existing example from repo, when we need to populate CellFormat *XId*, we specify it's index from corresponding array. 

```
public class TableHeaderStyles(...) : Style
{
    ...
    public uint LeftAndBottom { get; private set; }
    ...

    protected override List<Font> GetFonts() =>
    [
        new(
            new Bold(),
            new Color() {
                Rgb = new HexBinaryValue() { Value = fontColor }
            }
        )
    ];
    ...

    protected override List<CellFormat> GetCellFormats() =>
    [
        new()
        {
            FontId = 1, // GetFonts array, index 1
            ApplyFont = true,
            ...
        },
        ...
    ];

    protected override void BaseIndexUpdated()
    {
        LeftAndBottom = BaseIndex + 1; // Cellformat which we reference here has index 1, so must be BaseIndex + 1
        ..
    }
}
```
