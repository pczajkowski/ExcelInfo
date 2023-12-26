using ClosedXML.Excel;

namespace ExcelInfo
{
    public record ColumnInfo(string Letter, string? Name, int NumberOfCells, XLDataType Type, int Position);
}