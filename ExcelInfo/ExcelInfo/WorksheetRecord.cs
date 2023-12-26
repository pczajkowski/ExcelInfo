namespace ExcelInfo
{
    public record WorksheetRecord(int Position, string Name, IEnumerable<ColumnInfo> Columns);
}
