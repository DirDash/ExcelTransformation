using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelTransformation.TableClasses
{
    public class OpenXMLTable : ITable
    {
        private const string _sheetName = "Sheet1";

        private bool _autosave = false;

        private SpreadsheetDocument _document;
        private WorkbookPart _workbookPart;
        private WorksheetPart _worksheetPart;
        private SheetData _sheetData;
        private SharedStringTablePart _shareStringTablePart;

        private int _rowsCount;

        public OpenXMLTable(string fileUrl, bool editable)
        {
            if (File.Exists(fileUrl))
            {
                OpenExistingFile(fileUrl, editable);
            }
            else
            {
                CreateNewFile(fileUrl);
            }

            _rowsCount = _sheetData.ChildElements.Count;
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

            return row.OfType<Cell>().Select(ConvertToTableCell); ;
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

        public void AddRow(IEnumerable<TableCell> cells)
        {
            _rowsCount++;
            var row = new Row { RowIndex = (uint)(_rowsCount) };

            foreach (var cell in cells)
            {
                var cellReference = ConvertToColumnName(cell.ColumnIndex) + _rowsCount;

                var newCell = new Cell();
                newCell.CellReference = cellReference;

                var sharedStringIndex = InsertSharedStringItem(cell.Value);

                newCell.CellValue = new CellValue(sharedStringIndex.ToString());
                newCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                row.Append(newCell);
            }

            _sheetData.Append(row);
        }
        
        private string GetCellValue(Cell cell)
        {
            if (cell == null) return null;

            if (cell.DataType == CellValues.SharedString)
            {
                var sharedStringIndex = int.Parse(cell.InnerText);
                var sharedStringsTable = _workbookPart.SharedStringTablePart.SharedStringTable;
                var sharedStringItem = sharedStringsTable.ChildElements.GetElementSafe(sharedStringIndex);
                //var sharedStringItem = sharedStringsTable.ElementAt(sharedStringIndex);

                return sharedStringItem.InnerText;
            }

            return cell.InnerText;
        }

        private void OpenExistingFile(string fileUrl, bool editable)
        {
            _document = SpreadsheetDocument.Open(fileUrl, editable, new OpenSettings() { AutoSave = _autosave });

            _workbookPart = _document.WorkbookPart;

            string firstSheetId = _workbookPart.Workbook.Descendants<Sheet>().First().Id;
            _worksheetPart = (WorksheetPart)_workbookPart.GetPartById(firstSheetId);
            _sheetData = _worksheetPart.Worksheet.GetFirstChild<SheetData>();

            _shareStringTablePart = _workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            if (_shareStringTablePart == null)
            {
                _shareStringTablePart = _workbookPart.AddNewPart<SharedStringTablePart>();
            }
            if (_shareStringTablePart.SharedStringTable == null)
            {
                _shareStringTablePart.SharedStringTable = new SharedStringTable();
            }
        }

        private void CreateNewFile(string fileUrl)
        {
            _document = SpreadsheetDocument.Create(fileUrl, SpreadsheetDocumentType.Workbook, _autosave);

            _workbookPart = _document.AddWorkbookPart();
            _workbookPart.Workbook = new Workbook();

            _worksheetPart = _workbookPart.AddNewPart<WorksheetPart>();
            _sheetData = new SheetData();
            _worksheetPart.Worksheet = new Worksheet(_sheetData);
            _shareStringTablePart = _workbookPart.AddNewPart<SharedStringTablePart>();
            _shareStringTablePart.SharedStringTable = new SharedStringTable();

            var sheets = _document.WorkbookPart.Workbook.AppendChild(new Sheets());

            var sheet = new Sheet
            {
                Id = _document.WorkbookPart.GetIdOfPart(_worksheetPart),
                SheetId = 1,
                Name = _sheetName
            };

            sheets.Append(sheet);
        }

        private int InsertSharedStringItem(string text)
        {
            int itemIndex = 0;
            foreach (var item in _shareStringTablePart.SharedStringTable.Elements())
            {
                if (item.InnerText == text)
                {
                    return itemIndex;
                }
                itemIndex++;
            }

            _shareStringTablePart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text)));

            return itemIndex;
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

        // Only for first 26 columns (A..Z)
        // TODO: extend for all columns
        private int ConvertToColumnIndex(string columnName) 
        {
            return columnName[0] - 65;
        }
    }
}