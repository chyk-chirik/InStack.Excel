using InStack.Excel.Builder;
using Shouldly;
using Tests.Helpers;

namespace Tests
{
    [TestClass]
    public class WriteRowTests
    {
        [TestMethod]
        [DataRow(null, (uint)1, @"<row r=""1"">")]
        [DataRow(23.2, (uint)1, @"<row r=""1"" customHeight=""1"" ht=""23.20"">")]
        public void StartRow_RowOpenTagCreated(double? rowHeight, uint? row, string expectedOutput)
        {
            using var stream = new MemoryStream();
            using var xlsx = new XlsxDocument(stream);
            using var sheet = xlsx.AddSheet("sheetName");

            var bufferHelper = new BufferHelper(sheet);

            bufferHelper.StartTrackChanges();
            sheet.StartRow(row: row, column: 123456, height: rowHeight); // random column
            bufferHelper.GetStringResult().ShouldBe(expectedOutput);
        }


        [TestMethod]
        public void StartRow_CloseRowTagAndNewRowOpenTagCreated()
        {
            using var stream = new MemoryStream();
            using var xlsx = new XlsxDocument(stream);
            using var sheet = xlsx.AddSheet("sheetName");

            var bufferHelper = new BufferHelper(sheet);

            uint row = 2;
            sheet.StartRow(row: row);

            bufferHelper.StartTrackChanges();
            sheet.StartRow();
            bufferHelper.GetStringResult().ShouldBe(@$"</row><row r=""{row + 1}"">");
        }
    }
}
