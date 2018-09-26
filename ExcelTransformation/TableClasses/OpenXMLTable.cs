using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelTransformation.Utils;

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

            RowsCount = _sheetData.ChildElements.Count;
        }

        public int RowsCount { get; }

        public string GetCellValue(int rowIndex, int columnIndex)
        {
            if (rowIndex >= RowsCount) return null;

            var row = _sheetData.ChildElements.GetElementSafe(rowIndex);
            var cell = row?.ChildElements.GetElementSafe<Cell>(columnIndex);

            return GetCellValue(cell);
        }

        public string GetCellValue(Cell cell)
        {
            if (cell == null) return null;

            if (cell.DataType == CellValues.SharedString)
            {
                var sharedStringIndex = int.Parse(cell.InnerText);
                var sharedStringsTable = _workbookPart.SharedStringTablePart.SharedStringTable;
                var sharedStringItem = sharedStringsTable.ChildElements.GetElementSafe(sharedStringIndex);

                return sharedStringItem.InnerText;
            }

            return cell.InnerText;
        }

        public IEnumerable<string> GetCellValues(int rowIndex)
        {
            if (rowIndex >= RowsCount) return null;

            var row = _sheetData.ElementAt(rowIndex);

            return row.OfType<Cell>().Select(GetCellValue);
        }

        public void SetCellValue(int rowIndex, int columnIndex, string value)
        {
            var sharedStringIndex = InsertSharedStringItem(value);

            InsertCell(rowIndex, columnIndex, sharedStringIndex);
        }

        public void SaveAndClose()
        {
            _document.Save();
            _document.Close();
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
            foreach (var item in _shareStringTablePart.SharedStringTable.Elements<SharedStringItem>())
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

        private void InsertCell(int rowIndex, int columnIndex, int sharedStringIndex)
        {
            var cell = GetCell(rowIndex, columnIndex);

            cell.CellValue = new CellValue(sharedStringIndex.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
        }

        private Cell GetCell(int rowIndex, int columnIndex)
        {
            var row = GetRowOrCreateNew(rowIndex);
            var cell = (Cell)row.ElementAt(columnIndex);

            if (cell == null)
            {
                var cellReference = ConvertToColumnName(columnIndex);

                cell = new Cell { CellReference = cellReference };
                row.Append(cell);
            }

            return cell;
        }

        private Row GetRowOrCreateNew(int rowIndex)
        {
            var row = (Row)_sheetData.ElementAt(rowIndex);

            if (row == null)
            {
                row = new Row { RowIndex = (uint)(rowIndex + 1) };
                _sheetData.Append(row);
            }

            return row;
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