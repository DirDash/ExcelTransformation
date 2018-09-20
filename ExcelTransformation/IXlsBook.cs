namespace ExcelTransformation
{
    public interface IXlsBook
    {
        void LoadFromFile(string fileUrl);
        void SaveToFile(string fileUrl);
        string GetValue(int rowIndex, int cellIndex);
        void SetValue(int rowIndex, int cellIndex, string value);
    }
}