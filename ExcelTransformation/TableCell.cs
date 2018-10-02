namespace ExcelTransformation
{
    public class TableCell
    {
        public int ColumnIndex { get; private set; }
        public string Value;

        public TableCell (int columnIndex, string value = "")
        {
            ColumnIndex = columnIndex;
            Value = value;
        }
    }
}