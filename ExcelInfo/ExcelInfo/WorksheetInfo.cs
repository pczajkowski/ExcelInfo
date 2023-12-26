﻿using ClosedXML.Excel;

namespace ExcelInfo
{
    public class WorksheetInfo
    {
        private static XLDataType EstablishType(IXLCells? cells)
        {
            if (cells == null || !cells.Any())
                return XLDataType.Error;

            var numberOfCells = cells.Count();
            return numberOfCells > 2 ? cells.Last().DataType : cells.First().DataType;
        }

        public static List<WorksheetRecord> GetInfoOnWorksheets(string workbookPath)
        {
            var workbook = new XLWorkbook(workbookPath);
            return GetInfoOnWorksheets(workbook);
        }

        private static List<WorksheetRecord> GetInfoOnWorksheets(IXLWorkbook workbook)
        {
            var worksheets = new List<WorksheetRecord>();
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

                worksheets.Add(new WorksheetRecord(sheet.Position, sheet.Name, columns));
            }

            return worksheets;
        }
    }
}