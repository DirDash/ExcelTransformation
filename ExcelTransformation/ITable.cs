using System.Collections.Generic;

namespace ExcelTransformation
{
    public interface ITable
    {
        int RowsCount { get; }

        IEnumerable<string> GetCellValues(int rowIndex);

        string GetCellValue(int rowIndex, int columnIndex);

        void SetCellValue(int rowIndex, int columnIndex, string value);

        void SaveAndClose();
    }
}