using System;
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
        }

        public string GetCellValue(int rowIndex, int columnIndex)
        {
            var row = _sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex + 1);

            if (row == null) return null;

            string cellReference = ConvertToColumnName(columnIndex) + (rowIndex + 1);

            var cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference == cellReference);

            if (cell == null) return null;   

            if (cell.DataType == CellValues.SharedString)
            {
                var sharedStringItem = _workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(cell.InnerText));
                return sharedStringItem.InnerText;
            }
            else
            {
                return cell.InnerText;
            }
        }

        public void SetCellValue(int rowIndex, int columnIndex, string value)
        {
            int sharedStringIndex = InsertSharedStringItem(value);            
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

            Sheet sheet = new Sheet()
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
            var row = GetRow(rowIndex);

            string cellReference = ConvertToColumnName(columnIndex) + (rowIndex + 1);
            var cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference == cellReference);

            if (cell == null)
            {
                cell = new Cell() { CellReference = cellReference };
                row.Append(cell);
            }

            return cell;
        }

        private Row GetRow(int rowIndex)
        {
            var row = _sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex + 1);

            if (row == null)
            {
                row = new Row() { RowIndex = (uint)(rowIndex + 1) };
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