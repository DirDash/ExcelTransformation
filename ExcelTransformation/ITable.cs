using System.Collections.Generic;

namespace ExcelTransformation
{
    public interface ITable
    {
        IEnumerable<TableCell> GetRow(int rowIndex);

        void AddRow(IEnumerable<TableCell> cells);

        void SaveAndClose();


        //TODO: remove
        int RowsCount { get; }

        //TODO: remove
        IEnumerable<string> GetCellValues(int rowIndex);

        //TODO: remove
        string GetCellValue(int rowIndex, int columnIndex);

        //TODO: remove
        void SetCellValue(int rowIndex, int columnIndex, string value);
    }
}