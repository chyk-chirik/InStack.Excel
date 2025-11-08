using InStack.Excel.Builder;
using InStack.Excel.Builder.Extensions;
using System.Text;
using InStack.Excel.Builder.Extensions.Cell;

namespace Examples.MonthlyReportExample;

public record Category(string Name, int Jan, int Feb, int Mar, int Apr, int May, int Jun, int Jul, int Aug,
    int Sep, int Oct, int Nov, int Dec);
public static class MonthlyReportExample
{
    public static IEnumerable<Category> GetCategories()
    {
        var rand = new Random();
        var categories = new List<string> {
            "Tomatoes",
            "Pork",
            "Chicken",
            "Cheese",
            "Milk", 
            "Water",
            "Wine", 
            "Apples" 
        };

        foreach (var item in categories)
        {
            yield return new Category(item, rand.Next(0, 10000), rand.Next(0, 10000), rand.Next(0, 10000), rand.Next(0, 10000), rand.Next(0, 10000),
                rand.Next(0, 10000), rand.Next(0, 10000), rand.Next(0, 10000), rand.Next(0, 10000), rand.Next(0, 10000), rand.Next(0, 10000), rand.Next(0, 10000));
        }
    }
    public static void MonthlyReportExampleSheet(this XlsxDocument builder, TableHeaderStyles headerStyles, TableStyles tableStyles)
    {
        using var sheet = builder.AddSheet("Monthly Report", new SheetConfig { Columns = [new(3, 3, 30)] });

        uint column = 3;
        sheet.StartRow(row:3);

        sheet.MergeCellToBottom(column);
        sheet.Write("Category", column++, style: headerStyles.NoBorder);

        sheet.Write("Q1", column, style: headerStyles.LeftAndBottom);
        column = sheet.MergeCellToRight(column , count: 2, style: headerStyles.LeftAndBottom);

        sheet.Write("Q2", column, style: headerStyles.LeftAndBottom);
        column = sheet.MergeCellToRight(column, count: 2, style: headerStyles.LeftAndBottom);

        sheet.Write("Q3", column, style: headerStyles.LeftAndBottom);
        column = sheet.MergeCellToRight(column, count: 2, style: headerStyles.LeftAndBottom);

        sheet.Write("Q4", column, style: headerStyles.LeftAndBottom);
        column = sheet.MergeCellToRight(column, count: 2, style: headerStyles.LeftAndBottom);

        sheet.StartRow();
        column = 3;

        sheet.WriteEmpty(column++, style: headerStyles.NoBorder);
        sheet.Write("Jan", column++, style: headerStyles.Left);
        sheet.Write("Feb", column++, style: headerStyles.Left);
        sheet.Write("Mar", column++, style: headerStyles.Left);
        sheet.Write("Apr", column++, style: headerStyles.Left);
        sheet.Write("May", column++, style: headerStyles.Left);
        sheet.Write("Jun", column++, style: headerStyles.Left);
        sheet.Write("Jul", column++, style: headerStyles.Left);
        sheet.Write("Aug", column++, style: headerStyles.Left);
        sheet.Write("Sep", column++, style: headerStyles.Left);
        sheet.Write("Oct", column++, style: headerStyles.Left);
        sheet.Write("Nov", column++, style: headerStyles.Left);
        sheet.Write("Dec", column++, style: headerStyles.Left);

        var categories = GetCategories();
        uint firstCategoryRow = sheet.Row + 1;
        uint lastCategoryRow = sheet.Row + (uint)categories.Count();

        foreach (var category in categories)
        {
            sheet.StartRow();
            column = 3;

            var rowStyle = sheet.Row % 2 == 0
                ? tableStyles.Even
                : tableStyles.Odd;

            var categoryStyle = sheet.Row % 2 == 0
                ? tableStyles.EvenCategory
                : tableStyles.OddCategory;

            sheet.Write(category.Name, column++, style: categoryStyle);
            sheet.Write<int>(category.Jan, column++, style: rowStyle);
            sheet.Write<int>(category.Feb, column++, style: rowStyle);
            sheet.Write<int>(category.Mar, column++, style: rowStyle);
            sheet.Write<int>(category.Apr, column++, style: rowStyle);
            sheet.Write<int>(category.May, column++, style: rowStyle);
            sheet.Write<int>(category.Jun, column++, style: rowStyle);
            sheet.Write<int>(category.Jul, column++, style: rowStyle);
            sheet.Write<int>(category.Aug, column++, style: rowStyle);
            sheet.Write<int>(category.Sep, column++, style: rowStyle);
            sheet.Write<int>(category.Oct, column++, style: rowStyle);
            sheet.Write<int>(category.Nov, column++, style: rowStyle);
            sheet.Write<int>(category.Dec, column++, style: rowStyle);
        }

        sheet.StartRow(height: 5);

        sheet.WriteEmpty(3, count: 13, style: tableStyles.Divider);

        sheet.StartRow();
        column = 3;

        sheet.Write("Number formula:", column++, style: tableStyles.Total);

        Span<byte> formula = stackalloc byte[16];

        for (var formulaColumn = column; formulaColumn <= column + 12; formulaColumn++)
        {
            sheet.WriteSum(formulaColumn, formulaColumn, firstCategoryRow, lastCategoryRow, tableStyles.Total);
        }
    }
}
