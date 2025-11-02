using InStack.Excel.Builder;

namespace Examples.CollectionExample;

public static class CollectionExample
{
    public static void CollectionExampleSheet(this XlsxDocument builder, IEnumerable<CollectionItem> source, CollectionExampleStyles styles)
    {
        using var sheet = builder.AddSheet("Collection Example");

        sheet.StartRow();

        sheet.Write("Id");
        sheet.Write("Name");
        sheet.Write("Nickname");
        sheet.Write("Salary");
        sheet.Write("Birth Date");
        sheet.Write("Has Kids");

        sheet.EndRow();

        foreach (var item in source)
        {
            sheet.StartRow();

            sheet.Write(item.Id);
            sheet.Write(item.Name);
            sheet.Write(item.NickName, escape: true);
            sheet.Write<decimal>(item.Salary);
            sheet.Write(item.BirthDate, styles.DateFormatStyleId);
            sheet.Write(item.HasKids);

            sheet.EndRow();
        }
    }
}

