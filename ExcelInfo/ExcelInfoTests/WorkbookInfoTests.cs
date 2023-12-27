using ExcelInfo;

namespace ExcelInfoTests;

public class WorkbookInfoTests
{
    private const string DifferentTypesFile = "testFiles/differentTypes.xlsx";
    private const string DifferentTypesFileStartFrom = "testFiles/differentTypesStartFrom.xlsx";
    
    [Fact]
    public void GetInfoOnWorksheets()
    {
        var result = WorkbookInfo.GetInfoOnWorksheets(DifferentTypesFile).ToList();
        Assert.Equal(2, result.Count());
        Assert.NotEqual(result.First(), result.Last());
        Assert.NotEqual(result.First().Columns, result.Last().Columns);
    }
    
    [Fact]
    public void GetInfoOnWorksheetsStartFrom()
    {
        var result = WorkbookInfo.GetInfoOnWorksheets(DifferentTypesFileStartFrom, 2).ToList();
        Assert.Equal(2, result.Count());
        Assert.NotEqual(result.First(), result.Last());
        Assert.NotEqual(result.First().Columns, result.Last().Columns);
    }
}