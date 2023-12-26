using ExcelInfo;

namespace ExcelInfoTests;

public class WorksheetInfoTests
{
    private const string DifferentTypesFile = "testFiles/differentTypes.xlsx";
    [Fact]
    public void GetInfoOnWorksheets()
    {
        var result = WorksheetInfo.GetInfoOnWorksheets(DifferentTypesFile);
        Assert.Equal(2, result.Count);
        Assert.NotEqual(result.First(), result.Last());
        Assert.NotEqual(result.First().Columns, result.Last().Columns);
    }
}