using ClosedXML.Excel;

namespace ExcelInfo
{
    public static class WorksheetInfo
    {
        private static XLDataType EstablishType(IXLCells? cells)
        {
            if (cells == null || !cells.Any())
                return XLDataType.Error;

            var numberOfCells = cells.Count();
            return numberOfCells > 2 ? cells.Last().DataType : cells.First().DataType;
        }

        public static IEnumerable<WorksheetRecord> GetInfoOnWorksheets(string workbookPath)
        {
            var workbook = new XLWorkbook(workbookPath);
            return GetInfoOnWorksheets(workbook);
        }

        private static IEnumerable<WorksheetRecord> GetInfoOnWorksheets(IXLWorkbook workbook)
        {
            foreach (var sheet in workbook.Worksheets)
            {
                var columns = new List<ColumnInfo>();
                var firstRow = sheet.FirstRowUsed();
                foreach (var cell in firstRow.CellsUsed())
                {
                    var columnCells = cell.WorksheetColumn().CellsUsed();
                    var cellsType = EstablishType(columnCells);
                    columns.Add(new ColumnInfo(cell.Address.ColumnLetter, cell.Value.ToString(), columnCells.Count() - 1, cellsType, cell.Address.ColumnNumber));
                }

                yield return new WorksheetRecord(sheet.Position, sheet.Name, columns);
            }
        }
    }
}