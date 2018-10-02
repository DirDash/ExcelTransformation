using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelTransformation.TableClasses
{
    public class OpenXMLTable : ITable
    {
        private const string _sheetName = "Sheet1";

        private SpreadsheetDocument _document;
        private SheetData _sheetData;
        private SharedStringTable _sharedStringTable;

        private bool _autosave;
        private int _rowCount;

        public OpenXMLTable(bool autosave = false)
        {
            _autosave = autosave;
        }

        public void Open(string fileUrl, bool editable)
        {
            _document = SpreadsheetDocument.Open(Path.GetFullPath(fileUrl), editable, new OpenSettings() { AutoSave = _autosave });

            var workbookPart = _document.WorkbookPart;

            _sheetData = GetSheetData(workbookPart);

            _rowCount = _sheetData.ChildElements.Count;

            _sharedStringTable = GetSharedStringTable(workbookPart);
        }

        public void Create(string fileUrl)
        {
            _document = SpreadsheetDocument.Create(Path.GetFullPath(fileUrl), SpreadsheetDocumentType.Workbook, _autosave);

            var workbookPart = _document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            _sheetData = new SheetData();
            worksheetPart.Worksheet = new Worksheet(_sheetData);

            AddSheet(workbookPart, worksheetPart);

            _sharedStringTable = AddSharedStringTable(workbookPart);
        }

        public void SaveAndClose()
        {
            _document.Save();
            _document.Close();
        }

        public IEnumerable<TableCell> GetRow(int rowIndex)
        {
            if (rowIndex >= _rowCount) return null;

            var row = _sheetData.ElementAt(rowIndex);

            return row.OfType<Cell>().Select(ConvertToTableCell);
        }

        public void AddRow(IEnumerable<TableCell> cells)
        {
            _rowCount++;
            var row = new Row { RowIndex = (uint)_rowCount };

            foreach (var cell in cells)
            {
                var cellReference = ConvertToColumnName(cell.ColumnIndex) + _rowCount;

                var newCell = new Cell();
                newCell.CellReference = cellReference;

                newCell.CellValue = new CellValue(cell.Value);
                newCell.DataType = new EnumValue<CellValues>(CellValues.String);

                row.Append(newCell);
            }

            _sheetData.Append(row);
        }

        private SheetData GetSheetData(WorkbookPart workbookPart)
        {
            var firstSheetId = workbookPart.Workbook.Descendants<Sheet>().First().Id;
            var firstWorksheet = ((WorksheetPart)workbookPart.GetPartById(firstSheetId)).Worksheet;

            return firstWorksheet.GetFirstChild<SheetData>();
        }

        private SharedStringTable GetSharedStringTable(WorkbookPart workbookPart)
        {
            var shareStringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

            if (shareStringTablePart == null)
            {
                shareStringTablePart = workbookPart.AddNewPart<SharedStringTablePart>();
            }

            if (shareStringTablePart.SharedStringTable == null)
            {
                shareStringTablePart.SharedStringTable = new SharedStringTable();
            }

            return shareStringTablePart.SharedStringTable;
        }

        private void AddSheet(WorkbookPart workbookPart, WorksheetPart worksheetPart)
        {
            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            sheets.Append(new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = _sheetName
            });
        }

        private SharedStringTable AddSharedStringTable(WorkbookPart workbookPart)
        {
            var shareStringTablePart = workbookPart.AddNewPart<SharedStringTablePart>();
            shareStringTablePart.SharedStringTable = new SharedStringTable();

            return shareStringTablePart.SharedStringTable;
        }

        private TableCell ConvertToTableCell(Cell cell)
        {
            var columnPart = string.Empty;
            foreach (char c in cell.CellReference.Value)
            {
                if (!char.IsDigit(c))
                    columnPart += c;
            }
            
            var columnIndex = ConvertToColumnIndex(columnPart);
            var cellValue = GetCellValue(cell);

            return new TableCell(columnIndex, cellValue);
        }

        private string GetCellValue(Cell cell)
        {
            if (cell == null) return null;

            if (cell.DataType == CellValues.SharedString)
            {
                var sharedStringIndex = int.Parse(cell.InnerText);
                var sharedStringItem = _sharedStringTable.ElementAt(sharedStringIndex);

                return sharedStringItem.InnerText;
            }

            return cell.InnerText;
        }

        private int ConvertToColumnIndex(string columnName)
        {
            var columnIndex = 0;

            var i = 0;
            while (i < columnName.Length)
            {
                int remainder = columnName[i] - 65;
                columnIndex += remainder;
                if (i > 0)
                {
                    columnIndex += 26 - remainder;
                }
                i++;
            }

            return columnIndex;
        }

        private string ConvertToColumnName(int columnIndex)
        {
            columnIndex++;
            var columnName = string.Empty;

            while (columnIndex > 0)
            {
                var remainder = (columnIndex - 1) % 26;
                columnName = Convert.ToChar(65 + remainder).ToString() + columnName;
                columnIndex = (columnIndex - remainder) / 26;
            }

            return columnName;
        }
    }
}