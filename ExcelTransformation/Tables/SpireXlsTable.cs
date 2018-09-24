using Spire.Xls;

namespace ExcelTransformation.Tables
{
    public class SpireXlsTable : ITable
    {
        private Workbook _book;
        private int _workSheetIndex;
        private Worksheet _workSheet;

        public SpireXlsTable(int workSheetIndex = 0)
        {
            _book = new Workbook();
            _workSheetIndex = workSheetIndex;
            _workSheet = _book.Worksheets[_workSheetIndex];
        }

        public void LoadFromFile(string fileUrl)
        {
            _book.LoadFromFile(fileUrl);
            _workSheet = _book.Worksheets[_workSheetIndex];
        }

        public void SaveToFile(string fileUrl)
        {
            _book.SaveToFile(fileUrl);
        }

        public string GetValue(int rowIndex, int columnIndex)
        { 
            if (rowIndex < _workSheet.Rows.Length && columnIndex < _workSheet.Rows[rowIndex].Cells.Length)
            {
                return _workSheet.Rows[rowIndex].Cells[columnIndex].Value;
            }
            else
            {
                return null;
            }
        }

        public void SetValue(int rowIndex, int columnIndex, string value)
        {
            _workSheet.Rows[rowIndex].Cells[columnIndex].Value = value; 
        }
    }
}