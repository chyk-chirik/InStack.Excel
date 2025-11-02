using InStack.Excel.Builder;
using Shouldly;
using Tests.Helpers;

namespace Tests;

[TestClass]
public sealed class MergeCellManagerTests
{
    [TestMethod]
    [DataRow(4, 16, 1)]
    [DataRow(5, 16, 2)]
    [DataRow(8, 16, 2)]
    [DataRow(9, 16, 3)]
    [DataRow(9, 24, 2)]
    public void InitialCapacitySmall_MoreSpaceRented(int numberOfAddOperations, int initialCapacity, int maxAmountChunks)
    {
        // 16 is minimum length of the array pool
        // https://github.com/dotnet/runtime/blob/346706614bc9e3345906c4696106c554602e9bf6/src/libraries/System.Private.CoreLib/src/System/Buffers/ConfigurableArrayPool.cs#L27
        var cellManager = new MergeCellManager(initialCapacity);

        for (var i = 0; i < numberOfAddOperations; i++)
        {
            cellManager.Add(1, 1, 1, 1);
        }

        var chunks = TestHelper.GetField<List<MergeCellManager.RentedArray>, MergeCellManager>(cellManager, "_chunks");

        chunks.ShouldNotBeNull();
        chunks.Count.ShouldBeLessThanOrEqualTo(maxAmountChunks);
    }

    [TestMethod]
    public void AddMergeCells_CorrectStreamOutput()
    {
        var cellManager = new MergeCellManager(16);

        cellManager.Add(1, 1, 2, 1);

        using var stream = new MemoryStream();
        cellManager.Write(stream);

        var output = System.Text.Encoding.UTF8.GetString(stream.ToArray());
        var expectedOutput = @"<mergeCells><mergeCell ref=""A1:A2""/></mergeCells>";

        output.ShouldBe(expectedOutput);
    }
}
