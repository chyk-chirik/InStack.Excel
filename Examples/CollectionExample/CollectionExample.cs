using InStack.Excel.Builder;
using InStack.Excel.Builder.Extensions.Cell;

namespace Examples.CollectionExample;

public static class CollectionExample
{
    public static void CollectionExampleSheet(this XlsxDocument builder, IEnumerable<CollectionItem> source, CollectionExampleStyles styles)
    {
        using var sheet = builder.AddSheet("Collection Example");

        sheet.StartRow();
        uint column = 1;
        
        sheet.Write("Id", column++);
        sheet.Write("Name", column++);
       // sheet.Write("Nickname", column++);
        sheet.Write("Salary", column++);
        sheet.Write("Birth Date", column++);
        sheet.Write("Has Kids", column++);

        foreach (var item in source)
        {
            sheet.StartRow();
            column = 1;

            sheet.Write(item.Id, column++);
            sheet.Write(item.Name, column++);
           // sheet.Write(item.NickName, column++, escape: true);
            sheet.Write<decimal>(item.Salary, column++);
            sheet.Write(item.BirthDate, column++, styles.DateFormatStyleId);
            sheet.WriteBool(item.HasKids, column++);
        }
    }
}

