using System;
using System.Collections.Generic;

namespace ExcelTransformation
{
    public class AccountManagerNormalizer
    {
        private const string _initialTableManagerColumnTitle = "__l9district_mgrs";
        private const string _initialTableRegionColumnTitle = "__l9region_vps";
        private const string _initialTableAreaColumnTitle = "__l9area_vps";
        private const string _initialTableDivisionColumnTitle = "__l9division";

        private const string _managerTableManagerColumnTitle = "Manager";

        private const string _relationTableAccountColumnTitle = "id";
        private const string _relationTableManagerColumnTitle = "Manager";
        private const string _relationTableTypeColumnTitle = "Type";

        private char[] _cellContentDividers = new char[] { '|' };

        private ITable _initialTable;
        private ITable _accountTable;
        private ITable _managerTable;
        private ITable _relationTable;
        
        private Dictionary<string, int> _accountTableColumnTitles;
        private HashSet<string> _managerSet;

        private int _initialTableColumnAmount;
        private int _managerTableRowIndex;
        private int _relationTableRowIndex;

        public void Normalize(ITable initialTable, ITable accountTable, ITable managerTable, ITable relationTable)
        {
            _initialTable = initialTable;
            _accountTable = accountTable;
            _managerTable = managerTable;
            _relationTable = relationTable;

            var initialTableColumnTitles = GetInitialTableColumnTitles();
            _initialTableColumnAmount = initialTableColumnTitles.Count;

            FormatAccountTable(initialTableColumnTitles);
            FormatManagerTable();
            FormatRelationTable();

            ProccessInitialTable();
        }

        private List<string> GetInitialTableColumnTitles()
        {
            var columnTitles = new List<string>();

            int columnIndex = 0;
            string columnTitle;
            while (!string.IsNullOrEmpty(columnTitle = _initialTable.GetCellValue(0, columnIndex)))
            {
                columnTitles.Add(columnTitle);
                columnIndex++;
            }

            return columnTitles;
        }

        private void FormatAccountTable(List<string> initialTableColumnTitles)
        {
            _accountTableColumnTitles = new Dictionary<string, int>();

            int columnIndex = 0;
            foreach (var columnTitle in initialTableColumnTitles)
            {
                if (columnTitle != _initialTableManagerColumnTitle
                && columnTitle != _initialTableRegionColumnTitle
                && columnTitle != _initialTableAreaColumnTitle
                && columnTitle != _initialTableDivisionColumnTitle)
                {
                    _accountTableColumnTitles.Add(columnTitle, columnIndex);
                    _accountTable.SetCellValue(0, columnIndex, columnTitle);
                    columnIndex++;
                }
            }
        }

        private void FormatManagerTable()
        {
            _managerSet = new HashSet<string>();

            _managerTable.SetCellValue(0, 0, _managerTableManagerColumnTitle);

            _managerTableRowIndex = 1;
        }

        private void FormatRelationTable()
        {
            _relationTable.SetCellValue(0, 0, _relationTableAccountColumnTitle);
            _relationTable.SetCellValue(0, 1, _relationTableManagerColumnTitle);
            _relationTable.SetCellValue(0, 2, _relationTableTypeColumnTitle);

            _relationTableRowIndex = 1;
        }

        private void ProccessInitialTable()
        {
            int rowIndex = 1;
            string accountId;

            while (!string.IsNullOrEmpty(accountId = _initialTable.GetCellValue(rowIndex, 0)))
            {
                int columnIndex = 0;
                while (columnIndex < _initialTableColumnAmount)
                {
                    string cellContent = _initialTable.GetCellValue(rowIndex, columnIndex);
                    if (cellContent != null)
                    {
                        ProccesInitialTableCell(rowIndex, columnIndex, accountId, cellContent);
                    }
                    columnIndex++;
                }
                rowIndex++;
            }
        }

        private void ProccesInitialTableCell(int rowIndex, int columnIndex, string accountId, string cellContent)
        {
            string columnTitle = _initialTable.GetCellValue(0, columnIndex);

            cellContent = FormatCellContent(cellContent, columnIndex);

            if (columnTitle == _initialTableManagerColumnTitle
                || columnTitle == _initialTableRegionColumnTitle
                || columnTitle == _initialTableAreaColumnTitle
                || columnTitle == _initialTableDivisionColumnTitle)
            {
                ProcessAsManagerCell(accountId, columnTitle, cellContent);
            }
            else
            {
                ProcessAsAccountCell(rowIndex, columnTitle, cellContent);
            }
        }

        private string FormatCellContent(string content, int columnIndex)
        {
            if (columnIndex < 9 && columnIndex != 1)
            {
                return content.ToUpper().Trim();
            }
            return content;
        }

        private void ProcessAsManagerCell(string rowId, string columnTitle, string cellContent)
        {
            var managers = cellContent.Split(_cellContentDividers, StringSplitOptions.RemoveEmptyEntries);

            foreach (var manager in managers)
            {
                if (!_managerSet.Contains(manager))
                {
                    InsertInRowManagerTable(manager);
                    _managerSet.Add(manager);
                }
                InsertRowInRelationTable(rowId, manager, GetRelationType(columnTitle));
            }
        }

        private void ProcessAsAccountCell(int rowIndex, string columnTitle, string cellContent)
        {
            int columnIndex = _accountTableColumnTitles[columnTitle];
            _accountTable.SetCellValue(rowIndex, columnIndex, cellContent);
        }

        private string GetRelationType(string columnTitle)
        {
            switch (columnTitle)
            {
                case _initialTableManagerColumnTitle:
                    return "District";
                case _initialTableRegionColumnTitle:
                    return "Region";
                case _initialTableAreaColumnTitle:
                    return "Area";
                case _initialTableDivisionColumnTitle:
                    return "Division";
                default:
                    return "Undefined";
            }
        }

        private void InsertInRowManagerTable(string manager)
        {
            _managerTable.SetCellValue(_managerTableRowIndex, 0, manager);
            _managerTableRowIndex++;
        }

        private void InsertRowInRelationTable(string accountId, string manager, string type)
        {
            _relationTable.SetCellValue(_relationTableRowIndex, 0, accountId);
            _relationTable.SetCellValue(_relationTableRowIndex, 1, manager);
            _relationTable.SetCellValue(_relationTableRowIndex, 2, type);

            _relationTableRowIndex++;
        }
    }
}