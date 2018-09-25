namespace ExcelTransformation
{
    public interface ITable
    {
        string GetCellValue(int rowIndex, int columnIndex);
        void SetCellValue(int rowIndex, int columnIndex, string value);
        void SaveAndClose();
    }
}