using System;
using System.Collections.Generic;
using System.Linq;

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
        private int _rowsCount;

        public OpenXMLTable(bool autosave = false)
        {
            _autosave = autosave;
        }

        public void Open(string fileUrl, bool editable)
        {
            _document = SpreadsheetDocument.Open(fileUrl, editable, new OpenSettings() { AutoSave = _autosave });

            var workbookPart = _document.WorkbookPart;

            _sheetData = GetSheetData(workbookPart);

            _rowsCount = _sheetData.ChildElements.Count;

            _sharedStringTable = GetSharedStringTable(workbookPart);
        }

        public void Create(string fileUrl)
        {
            _document = SpreadsheetDocument.Create(fileUrl, SpreadsheetDocumentType.Workbook, _autosave);

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
            if (rowIndex >= _rowsCount) return null;

            var row = _sheetData.ElementAt(rowIndex);

            return row.OfType<Cell>().Select(ConvertToTableCell);
        }

        public void AddRow(IEnumerable<TableCell> cells)
        {
            _rowsCount++;
            var row = new Row { RowIndex = (uint)_rowsCount };

            foreach (var cell in cells)
            {
                var cellReference = ConvertToColumnName(cell.ColumnIndex) + _rowsCount;

                var newCell = new Cell();
                newCell.CellReference = cellReference;

                //var sharedStringIndex = InsertSharedStringItem(cell.Value);

                //newCell.CellValue = new CellValue(sharedStringIndex.ToString());
                //newCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

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
            string rowPart = string.Empty;
            string columnPart = string.Empty;
            foreach (char c in cell.CellReference.Value)
            {
                if (char.IsDigit(c))
                {
                    rowPart += c;
                }
                else
                {
                    columnPart += c;
                }
            }

            int rowIndex = int.Parse(rowPart) - 1;
            int columnIndex = ConvertToColumnIndex(columnPart);
            string cellValue = GetCellValue(cell);

            return new TableCell(rowIndex, columnIndex, cellValue);
        }

        private string GetCellValue(Cell cell)
        {
            if (cell == null) return null;

            if (cell.DataType == CellValues.SharedString)
            {
                var sharedStringIndex = int.Parse(cell.InnerText);
                //var sharedStringItem = sharedStringsTable.ChildElements.GetElementSafe(sharedStringIndex);
                var sharedStringItem = _sharedStringTable.ElementAt(sharedStringIndex);

                return sharedStringItem.InnerText;
            }

            return cell.InnerText;
        }

        private int InsertSharedStringItem(string text)
        {
            int itemIndex = 0;
            var shr = _sharedStringTable;
            foreach (var item in _sharedStringTable.Elements())
            {
                if (item.InnerText == text)
                {
                    return itemIndex;
                }
                itemIndex++;
            }

            _sharedStringTable.AppendChild(new SharedStringItem(new Text(text)));

            return itemIndex;
        }

        private int ConvertToColumnIndex(string columnName)
        {
            int columnIndex = 0;

            int i = 0;
            while (i < columnName.Length)
            {
                if (i > 0)
                {
                    columnIndex += 26;
                }

                columnIndex += columnName[i] - 65;
                i++;
            }

            return columnIndex;
        }

        private string ConvertToColumnName(int columnIndex)
        {
            columnIndex++;
            string columnName = string.Empty;

            while (columnIndex > 0)
            {
                int remainder = (columnIndex - 1) % 26;
                columnName = Convert.ToChar(65 + remainder).ToString() + columnName;
                columnIndex = (columnIndex - remainder) / 26;
            }

            return columnName;
        }
    }
}