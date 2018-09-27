namespace ExcelTransformation
{
    public class TableCell
    {
        public int RowIndex { get; private set; }
        public int ColumnIndex { get; private set; }
        public string Value;

        public TableCell (int rowIndex, int columnIndex, string value = "")
        {
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
            Value = value;
        }
    }
}