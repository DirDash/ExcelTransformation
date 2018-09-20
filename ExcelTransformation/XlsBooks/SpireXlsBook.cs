using Spire.Xls;

namespace ExcelTransformation.XlsBooks
{
    public class SpireXlsBook : IXlsBook
    {
        private Workbook _book;
        private int _workSheetIndex = 0;

        public SpireXlsBook(int workSheetIndex = 0)
        {
            _book = new Workbook();
            _workSheetIndex = workSheetIndex;
        }

        public void LoadFromFile(string fileUrl)
        {
            _book.LoadFromFile(fileUrl);
        }

        public void SaveToFile(string fileUrl)
        {
            _book.SaveToFile(fileUrl);
        }

        public string GetValue(int rowIndex, int cellIndex)
        {
            var sheet = _book.Worksheets[_workSheetIndex];

            if (rowIndex < sheet.Rows.Length && cellIndex < sheet.Rows[rowIndex].Cells.Length)
            {
                return sheet.Rows[rowIndex].Cells[cellIndex].Value;
            }
            else
            {
                return null;
            }
        }

        public void SetValue(int rowIndex, int cellIndex, string value)
        {
            var sheet = _book.Worksheets[_workSheetIndex];
            sheet.Rows[rowIndex].Cells[cellIndex].Value = value; 
        }
    }
}