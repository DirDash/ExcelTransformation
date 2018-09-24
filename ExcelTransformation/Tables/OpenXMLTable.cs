using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelTransformation.Tables
{
    // TODO
    public class OpenXMLTable : ITable
    {
        private SpreadsheetDocument _table;

        public OpenXMLTable(string fileUrl)
        {
            if (File.Exists(fileUrl))
            {
                _table = SpreadsheetDocument.Open(fileUrl, true);
            }
            else
            {
                _table = SpreadsheetDocument.Create(fileUrl, SpreadsheetDocumentType.Workbook);
            }
            
        }

        public void LoadFromFile(string fileUrl)
        {
            _table = SpreadsheetDocument.Open(fileUrl, true);
        }

        public void SaveToFile(string fileUrl)
        {
            _table.SaveAs(fileUrl);
        }

        public string GetValue(int rowIndex, int columnIndex)
        {
            WorkbookPart wbPart = _table.WorkbookPart;
            Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault();
            WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
            Cell theCell = wsPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == ConvertToCellReference(rowIndex, columnIndex)).FirstOrDefault();
            return theCell.InnerText;
        }

        public void SetValue(int rowIndex, int columnIndex, string value)
        {
        }

        private string ConvertToCellReference(int rowIndex, int columnIndex)
        {
            string columnName = string.Empty;
            int remainder;

            while (columnIndex > 0)
            {
                remainder = (columnIndex - 1) % 26;
                columnName = Convert.ToChar(65 + remainder).ToString() + columnName;
                columnIndex = (columnIndex - remainder) / 26;
            }

            return columnName + rowIndex.ToString();
        }
    }
}