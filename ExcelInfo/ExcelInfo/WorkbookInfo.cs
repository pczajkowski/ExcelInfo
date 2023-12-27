using ClosedXML.Excel;

namespace ExcelInfo
{
    public static class WorkbookInfo
    {
        private static XLDataType EstablishType(IEnumerable<IXLCell>? cells)
        {
            if (cells == null || !cells.Any())
                return XLDataType.Error;

            var numberOfCells = cells.Count();
            return numberOfCells > 2 ? cells.Last().DataType : cells.First().DataType;
        }

        public static IEnumerable<WorksheetRecord> GetInfoOnWorksheets(string workbookPath, uint startFrom = 1)
        {
            var workbook = new XLWorkbook(workbookPath);
            return GetInfoOnWorksheets(workbook, startFrom);
        }

        private static IEnumerable<WorksheetRecord> GetInfoOnWorksheets(IXLWorkbook workbook, uint startFrom)
        {
            foreach (var worksheet in workbook.Worksheets)
            {
                var columns = new List<ColumnInfo>();
                var firstRow = worksheet.Row((int)startFrom);
                if (firstRow.IsEmpty())
                    firstRow = worksheet.RowsUsed().First(x => x.RowNumber() > startFrom && !x.IsEmpty());
                
                foreach (var cell in firstRow.CellsUsed())
                {
                    var columnCells = cell.WorksheetColumn().Cells()
                        .Where(x => x.Address.RowNumber > firstRow.RowNumber());
                    var cellsType = EstablishType(columnCells);
                    columns.Add(new ColumnInfo(cell.Address.ColumnLetter, cell.Value.ToString(), columnCells.Count() - 1, cellsType, cell.Address.ColumnNumber));
                }

                yield return new WorksheetRecord(worksheet.Position, worksheet.Name, columns);
            }
        }
    }
}