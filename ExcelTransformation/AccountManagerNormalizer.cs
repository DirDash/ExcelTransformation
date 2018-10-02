using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelTransformation
{
    public class AccountManagerNormalizer
    {
        private const string _initialTableManagerHeader = "__l9district_mgrs";
        private const string _initialTableRegionHeader = "__l9region_vps";
        private const string _initialTableAreaHeader = "__l9area_vps";
        private const string _initialTableDivisionHeader = "__l9division";

        private const string _managerTableManagerHeader = "Manager";

        private const string _relationTableAccountHeader = "id";
        private const string _relationTableManagerHeader = "Manager";
        private const string _relationTableTypeHeader = "Type";

        private char[] _cellContentDividers = new char[] { '|' };
        private List<string> _managerHeaders = new List<string>()
        {
            _initialTableManagerHeader,
            _initialTableRegionHeader,
            _initialTableAreaHeader,
            _initialTableDivisionHeader
        };
        private Dictionary<string, string> _managerHeaderRelationTypeDictionary = new Dictionary<string, string>()
        {
            { _initialTableManagerHeader, "District" },
            { _initialTableRegionHeader, "Region" },
            { _initialTableAreaHeader, "Area" },
            { _initialTableDivisionHeader, "Division" },
        };

        private ITable _initialTable;
        private ITable _accountTable;
        private ITable _managerTable;
        private ITable _relationTable;

        private List<string> _initialTableHeaders;
        private Dictionary<string, int> _accountTableHeaders;
        private HashSet<string> _managerSet;

        public void Normalize(ITable initialTable, ITable accountTable, ITable managerTable, ITable relationTable)
        {
            _initialTable = initialTable;
            _accountTable = accountTable;
            _managerTable = managerTable;
            _relationTable = relationTable;

            _initialTableHeaders = GetInitialTableHeaders();

            FormatAccountTable();
            FormatManagerTable();
            FormatRelationTable();
            
            ProccessInitialTable();
        }

        private List<string> GetInitialTableHeaders()
        {
            var headerCells = _initialTable.GetRow(0);

            return headerCells.Select(h => h.Value).ToList();
        }

        private void FormatAccountTable()
        {
            _accountTableHeaders = new Dictionary<string, int>();

            var accountTableHeaderCells = new List<TableCell>();

            var columnIndex = 0;

            foreach (var header in _initialTableHeaders)
            {
                if (!_managerHeaders.Contains(header))
                {
                    _accountTableHeaders.Add(header, columnIndex);
                    accountTableHeaderCells.Add(new TableCell(columnIndex, header));
                    columnIndex++;
                }
            }

            _accountTable.AddRow(accountTableHeaderCells);
        }

        private void FormatManagerTable()
        {
            _managerSet = new HashSet<string>();

            var managerTableHeaders = new List<TableCell>();

            managerTableHeaders.Add(new TableCell(0, _managerTableManagerHeader));

            _managerTable.AddRow(managerTableHeaders);
        }

        private void FormatRelationTable()
        {
            var relationTableHeaders = new List<TableCell>();

            relationTableHeaders.Add(new TableCell(0, _relationTableAccountHeader));
            relationTableHeaders.Add(new TableCell(1, _relationTableManagerHeader));
            relationTableHeaders.Add(new TableCell(2, _relationTableTypeHeader));

            _relationTable.AddRow(relationTableHeaders);
        }

        private void ProccessInitialTable()
        {
            var rowIndex = 1;

            IEnumerable<TableCell> rowCells;
            while ((rowCells = _initialTable.GetRow(rowIndex)) != null)
            {
                var accountId = rowCells.First().Value;

                var accountCells = new List<TableCell>();
                foreach (var cell in rowCells)
                {
                    var columnHeader = _initialTableHeaders[cell.ColumnIndex];

                    cell.Value = FormatCellValue(cell);

                    if (_managerHeaders.Contains(columnHeader))
                    {
                        ProcessAsManagerCell(cell, columnHeader, accountId);
                    }
                    else
                    {
                        ProcessAsAccountCell(cell, columnHeader, rowIndex, accountCells);
                    }
                }

                _accountTable.AddRow(accountCells);

                rowIndex++;
            }
        }

        private string FormatCellValue(TableCell cell)
        {
            if (cell.ColumnIndex < 9 && cell.ColumnIndex != 1)
            {
                return cell.Value.ToUpper().Trim();
            }
            return cell.Value;
        }

        private void ProcessAsManagerCell(TableCell cell, string columnHeader, string accountId)
        {
            var managers = cell.Value.Split(_cellContentDividers, StringSplitOptions.RemoveEmptyEntries);

            foreach (var manager in managers)
            {
                if (!_managerSet.Contains(manager))
                {
                    InsertRowInManagerTable(manager);
                    _managerSet.Add(manager);
                }

                InsertRowInRelationTable(accountId, manager, _managerHeaderRelationTypeDictionary[columnHeader]);
            }
        }

        private void ProcessAsAccountCell(TableCell cell, string columnHeader, int rowIndex, List<TableCell> accountCells)
        {
            var accountTableColumnIndex = _accountTableHeaders[columnHeader];
            var accountCell = new TableCell(accountTableColumnIndex, cell.Value);
            accountCells.Add(accountCell);
        }

        private void InsertRowInManagerTable(string manager)
        {
            var rowCells = new List<TableCell>();
            rowCells.Add(new TableCell(0, manager));
            _managerTable.AddRow(rowCells);
        }

        private void InsertRowInRelationTable(string accountId, string manager, string type)
        {
            var rowCells = new List<TableCell>();
            rowCells.Add(new TableCell(0, accountId));
            rowCells.Add(new TableCell(1, manager));
            rowCells.Add(new TableCell(2, type));

            _relationTable.AddRow(rowCells);           
        }
    }
}