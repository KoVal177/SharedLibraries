using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExtensions
{
    class XlXml
    {
        public class XlXmlReader
        {
            string SourceFileNameFull { get; }

            public XlXmlReader(string sourceFileNameFull) =>
                this.SourceFileNameFull = sourceFileNameFull;

            public List<string> GetValuesColumnAsString(
                string sheet,
                string range)
            {
                var columnValues = new Dictionary<int, string>();
                string startCell = range.Split(':')[0].ToUpper();
                string endCell = range.Split(':')[1].ToUpper();

                string startCellLetter = Xl.GetCellColumnLetter(startCell);
                int startRow = startCell.Any(char.IsDigit) ? Xl.GetCellRowIndex(startCell) : 1;
                int endRow = endCell.Any(char.IsDigit) ? Xl.GetCellRowIndex(endCell) : -1;
                int maxRowWithValue = startRow;

                using (var sourceFile = SpreadsheetDocument.Open(this.SourceFileNameFull, false))
                {
                    var wb = sourceFile.WorkbookPart.Workbook;
                    var wsPart = (WorksheetPart)sourceFile.WorkbookPart
                        .GetPartById(wb.Descendants<Sheet>()
                        .First(s => s.Name == sheet).Id);
                    var lastRowToCheck = (endRow != -1)
                        ? endRow
                        : Xl.GetCellRowIndex(wsPart.Worksheet.SheetDimension.Reference);

                    foreach (var cell in wsPart.Worksheet.Descendants<Cell>())
                    {
                        int cellRowIndex = Xl.GetCellRowIndex(cell.CellReference.Value);
                        if (startCellLetter != Xl.GetCellColumnLetter(cell.CellReference.Value)
                            || cellRowIndex < startRow
                            || cellRowIndex > lastRowToCheck)
                            continue;
                        string cellValue = GetCellValueAsString(cell, wb.WorkbookPart);
                        if (string.IsNullOrEmpty(cellValue)) continue;

                        if (cellRowIndex > maxRowWithValue) maxRowWithValue = cellRowIndex;
                        columnValues.Add(cellRowIndex, cellValue);
                    }
                }
                return Enumerable.Range(startRow, maxRowWithValue - startRow + 1)
                    .Select(i => columnValues.ContainsKey(i) ? columnValues[i] : "")
                    .ToList();
            }
            public List<string> GetValuesRowAsString(
                string sheet,
                string range)
            {
                var rowValues = new Dictionary<int, string>();
                string startCell = range.Split(':')[0].ToUpper();
                string endCell = range.Split(':')[1].ToUpper();

                int startRow = Xl.GetCellRowIndex(startCell);
                int startColIndex = (startCell.Any(char.IsLetter)) ? Xl.GetCellColumnIndex(startCell) : 1;
                int endColIndex = (endCell.Any(char.IsLetter)) ? Xl.GetCellColumnIndex(endCell) : -1;
                int maxColWithValueIndex = startColIndex;

                using (var sourceFile = SpreadsheetDocument.Open(this.SourceFileNameFull, false))
                {
                    var wb = sourceFile.WorkbookPart.Workbook;
                    var wsPart = (WorksheetPart)sourceFile.WorkbookPart
                        .GetPartById(wb.Descendants<Sheet>()
                        .First(s => s.Name == sheet).Id);
                    var lastColToCheckIndex = (endColIndex != -1)
                        ? endColIndex
                        : Xl.GetCellColumnIndex(
                            wsPart.Worksheet.SheetDimension.Reference.ToString().Split(':')[1]);

                    foreach (var cell in wsPart.Worksheet.Descendants<Cell>())
                    {
                        int cellColIndex = Xl.GetCellColumnIndex(cell.CellReference.Value);
                        if (startRow != Xl.GetCellRowIndex(cell.CellReference.Value)
                            || cellColIndex < startColIndex
                            || cellColIndex > lastColToCheckIndex)
                            continue;
                        string cellValue = GetCellValueAsString(cell, wb.WorkbookPart);
                        if (string.IsNullOrEmpty(cellValue)) continue;

                        if (cellColIndex > maxColWithValueIndex) maxColWithValueIndex = cellColIndex;
                        rowValues.Add(cellColIndex, cellValue);
                    }
                }
                return Enumerable.Range(startColIndex, maxColWithValueIndex - startColIndex + 1)
                    .Select(i => rowValues.ContainsKey(i) ? rowValues[i] : "")
                    .ToList();
            }

            public static string GetCellValueAsString(Cell cell, WorkbookPart? workbookPart)
            {
                string cellValue = string.Empty;
                if (cell.DataType is null)
                    if (cell.CellValue != null) cellValue = cell.CellValue.Text;
                    else { }
                else if (cell.DataType == CellValues.SharedString)
                {
                    int id = -1;
                    if (Int32.TryParse(cell.InnerText, out id))
                    {
                        var item = workbookPart.SharedStringTablePart.SharedStringTable
                            .Elements<SharedStringItem>().ElementAt(id);
                        if (item.Text != null)
                            cellValue = item.Text.Text;
                        else if (item.InnerText != null)
                            cellValue = item.InnerText;
                        else if (item.InnerXml != null)
                            cellValue = item.InnerXml;
                    }
                }
                else cellValue = cell.CellValue.Text;

                return cellValue;
            }
        }
    }
}
