using System.Collections.Generic;

namespace ExcelTransformation
{
    public interface ITable
    {
        IEnumerable<TableCell> GetRow(int rowIndex);

        void AddRow(IEnumerable<TableCell> cells);

        void SaveAndClose();
    }
}