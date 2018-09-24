namespace ExcelTransformation
{
    public interface ITable
    {
        void LoadFromFile(string fileUrl);
        void SaveToFile(string fileUrl);
        string GetValue(int rowIndex, int columnIndex);
        void SetValue(int rowIndex, int columnIndex, string value);
    }
}