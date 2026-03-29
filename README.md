# InStack.Excel
Generate xlsx almost without memory allocations. Simplified work with styles.

## Minimal example
```
var styles = new CollectionExampleStyles();
var stylesXml = Style.Compile(styles);

using (var builder = new XlsxDocument(_stream))
{
    using (var styleXmlWriter = builder.CreateStyleEntry())
    {
        using var xmlWriter = System.Xml.XmlWriter.Create(styleXmlWriter, new System.Xml.XmlWriterSettings { CloseOutput = true });
        stylesXml.WriteTo(xmlWriter);
    }

    builder.CollectionExampleSheet(_source, styles);
}

 public static void CollectionExampleSheet(this XlsxDocument builder, IEnumerable<CollectionItem> source, CollectionExampleStyles styles)
 {
     using var sheet = builder.AddSheet("Collection Example");

     sheet.StartRow();
     uint column = 1;
     
     sheet.Write("Id", column++);
     sheet.Write("Name", column++);
     sheet.Write("Nickname", column++);
     sheet.Write("Salary", column++);
     sheet.Write("Birth Date", column++);
     sheet.Write("Has Kids", column++);

     foreach (var item in source)
     {
         sheet.StartRow();
         column = 1;

         sheet.Write(item.Id, column++);
         sheet.Write(item.Name, column++);
         sheet.Write(item.NickName, column++, escape: true);
         sheet.Write<decimal>(item.Salary, column++);
         sheet.Write(item.BirthDate, column++, styles.DateFormatStyleId);
         sheet.WriteBool(item.HasKids, column++);
     }
 }
```

## Styles

Work with styles is always a pain, it's easy to get corrupted file and spend a lot of time in troubleshoot. This was solved in next way:

- Split styles across several classes for usability and reusability, like MyTableHeaderStyles, MyTableSpecificRowStyles ..
- When you reference specific style in CellFormat object you reference it's local (relative to class) position. 

  Look at existing example from repo, when we need to populate CellFormat *XId*, we specify it's index from corresponding array. 

```
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
            NumberFormatId = 1, // we specify here index in GetNumberingFormats() array, 1-based
            ApplyNumberFormat = true
        }
    ];

    protected override void BaseIndexUpdated()
    {
        DateFormatStyleId = BaseIndex + 1; // position in GetCellFormats() array
    }
}
```
## More advanced example
Code contains example **MonthlyReportExample**, outcome is 
![Code contains example **MonthlyReportExample**, outcome is](/readme_images/example.png)


