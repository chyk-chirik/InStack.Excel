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

        uint columnBegin = 3;
        sheet.StartRow(row:3, column:columnBegin);

        sheet.Write("Category", style: headerStyles.NoBorder);
        sheet.MergePrevCellToBottom();

        sheet.Write("Q1", style: headerStyles.LeftAndBottom);
        sheet.MergePrevCellToRight(count: 2, style: headerStyles.LeftAndBottom);

        sheet.Write("Q2", style: headerStyles.LeftAndBottom);
        sheet.MergePrevCellToRight(count: 2, style: headerStyles.LeftAndBottom);

        sheet.Write("Q3", style: headerStyles.LeftAndBottom);
        sheet.MergePrevCellToRight(count: 2, style: headerStyles.LeftAndBottom);

        sheet.Write("Q4", style: headerStyles.LeftAndBottom);
        sheet.MergePrevCellToRight(count: 2, style: headerStyles.LeftAndBottom);

        sheet.StartRow(column: columnBegin);

        sheet.WriteEmpty(style: headerStyles.NoBorder);
        sheet.Write("Jan", style: headerStyles.Left);
        sheet.Write("Feb", style: headerStyles.Left);
        sheet.Write("Mar", style: headerStyles.Left);
        sheet.Write("Apr", style: headerStyles.Left);
        sheet.Write("May", style: headerStyles.Left);
        sheet.Write("Jun", style: headerStyles.Left);
        sheet.Write("Jul", style: headerStyles.Left);
        sheet.Write("Aug", style: headerStyles.Left);
        sheet.Write("Sep", style: headerStyles.Left);
        sheet.Write("Oct", style: headerStyles.Left);
        sheet.Write("Nov", style: headerStyles.Left);
        sheet.Write("Dec", style: headerStyles.Left);

        var categories = GetCategories();
        uint firstCategoryRow = sheet.Row + 1;
        uint lastCategoryRow = sheet.Row + (uint)categories.Count();

        foreach (var category in categories)
        {
            sheet.StartRow(column: columnBegin);

            var rowStyle = sheet.Row % 2 == 0
                ? tableStyles.Even
                : tableStyles.Odd;

            var categoryStyle = sheet.Row % 2 == 0
                ? tableStyles.EvenCategory
                : tableStyles.OddCategory;

            sheet.Write(category.Name, style: categoryStyle);
            sheet.Write<int>(category.Jan, style: rowStyle);
            sheet.Write<int>(category.Feb, style: rowStyle);
            sheet.Write<int>(category.Mar, style: rowStyle);
            sheet.Write<int>(category.Apr, style: rowStyle);
            sheet.Write<int>(category.May, style: rowStyle);
            sheet.Write<int>(category.Jun, style: rowStyle);
            sheet.Write<int>(category.Jul, style: rowStyle);
            sheet.Write<int>(category.Aug, style: rowStyle);
            sheet.Write<int>(category.Sep, style: rowStyle);
            sheet.Write<int>(category.Oct, style: rowStyle);
            sheet.Write<int>(category.Nov, style: rowStyle);
            sheet.Write<int>(category.Dec, style: rowStyle);
        }

        sheet.StartRow(column: columnBegin, height: 5);

        sheet.WriteEmpty(count: 13, style: tableStyles.Divider);

        sheet.StartRow(column: columnBegin);

        sheet.Write("Number formula:", style: tableStyles.Total);

        Span<byte> formula = stackalloc byte[16];

        for (var column = sheet.Column; column <= columnBegin + 12; column++)
        {
            var offset = Encoding.UTF8.GetBytes("SUM(", formula);
            offset += CellReferenceFormatter.Format(formula[offset..], firstCategoryRow, column);
            formula[offset++] = (byte)':';
            offset += CellReferenceFormatter.Format(formula[offset..], lastCategoryRow, column);
            formula[offset++] = (byte)')';
            sheet.WriteNumberFormula(formula[..offset], style: tableStyles.Total);
        }

        sheet.StartRow(column: columnBegin);
    }
}
