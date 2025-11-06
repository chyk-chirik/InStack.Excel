using ExcelUtils.Builder.RowExtensions;
using InStack.Excel.Builder;
using Shouldly;
using System.Reflection;

namespace Tests.Helpers;

internal static class TestHelper
{
    internal static string StringLengthOf_4096 = string.Create(4096, "some random value", (span, st) => { span.Fill('0'); });

    internal static TField GetField<TField, TObject>(TObject obj, string fieldName)
    {
        var field = typeof(TObject).GetField(fieldName, BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic)!;
        return (TField)field.GetValue(obj)!;
    }

    internal static void TestCellWrite(uint? row, uint? column, string expectedOutput, Action<Sheet> writeCell, SheetConfig? sheetConfig = null)
    {
        using var stream = new MemoryStream();
        using var xlsx = new XlsxDocument(stream);
        using var sheet = xlsx.AddSheet("sheetName", sheetConfig);

        var bufferHelper = new BufferHelper(sheet);

        sheet.StartRow(row: row, column: column);

        bufferHelper.StartTrackChanges();
        writeCell(sheet);
        bufferHelper.GetStringResult().ShouldBe(expectedOutput);
    }
}
