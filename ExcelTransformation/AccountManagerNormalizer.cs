using System;
using System.Collections.Generic;
using System.Linq;
using ExcelTransformation.Utils;

namespace ExcelTransformation
{
    public class AccountManagerNormalizer
    {
        private const string _initialTableManagerColumnTitle = "__l9district_mgrs";
        private const string _initialTableRegionColumnTitle = "__l9region_vps";
        private const string _initialTableAreaColumnTitle = "__l9area_vps";
        private const string _initialTableDivisionColumnTitle = "__l9division";

        private const string _managerTableManagerHeader = "Manager";

        private const string _relationTableAccountHeader = "id";
        private const string _relationTableManagerHeader = "Manager";
        private const string _relationTableTypeHeader = "Type";

        private char[] _cellContentDividers = new char[] { '|' };

        private ITable _initialTable;
        private ITable _accountTable;
        private ITable _managerTable;
        private ITable _relationTable;

        private List<string> _initialTableHeaders;
        private Dictionary<string, int> _accountTableHeaders;
        private HashSet<string> _managerSet;
        
        //private int _managerTableRowIndex;
        //private int _relationTableRowIndex;

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

            using (ExecutionTimer.StartNew("ProccessInitialTable"))
                ProccessInitialTable();
        }

        private List<string> GetInitialTableHeaders()
        {
            var headerCellValues = _initialTable.GetRow(0);

            return headerCellValues.Select(h => h.Value).ToList();
        }

        private void FormatAccountTable()
        {
            var accountTableHeaders = new List<TableCell>();

            _accountTableHeaders = new Dictionary<string, int>();

            var columnIndex = 0;

            foreach (var header in _initialTableHeaders)
            {
                if (!IsManagerHeader(header))
                {
                    _accountTableHeaders.Add(header, columnIndex);
                    accountTableHeaders.Add(new TableCell(0, columnIndex, header));
                    columnIndex++;

                    //_accountTable.SetCellValue(0, columnIndex, header);
                }
            }

            _accountTable.AddRow(accountTableHeaders);
        }

        private void FormatManagerTable()
        {
            _managerSet = new HashSet<string>();

            var managerTableHeaders = new List<TableCell>();

            managerTableHeaders.Add(new TableCell(0, 0, _managerTableManagerHeader));

            _managerTable.AddRow(managerTableHeaders);

            //_managerTableRowIndex = 1;
        }

        private void FormatRelationTable()
        {
            var relationTableHeaders = new List<TableCell>();

            relationTableHeaders.Add(new TableCell(0, 0, _relationTableAccountHeader));
            relationTableHeaders.Add(new TableCell(0, 1, _relationTableManagerHeader));
            relationTableHeaders.Add(new TableCell(0, 2, _relationTableTypeHeader));

            _relationTable.AddRow(relationTableHeaders);

            //_relationTableRowIndex = 1;
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
                    string columnHeader = _initialTableHeaders[cell.ColumnIndex];

                    cell.Value = FormatCellValue(cell);

                    if (IsManagerHeader(columnHeader))
                    {
                        ProcessAsManagerCell(cell, columnHeader, accountId);
                    }
                    else
                    {
                        int accountTableColumnIndex = _accountTableHeaders[columnHeader];
                        accountCells.Add(new TableCell(rowIndex, accountTableColumnIndex, cell.Value));

                        //ProcessAsAccountCell(cell, columnHeader);
                    }

                    //ProccesInitialTableCell(cell, accountId);
                }
                _accountTable.AddRow(accountCells);
                rowIndex++;

                //var columnIndex = 0;

                //var tt = _initialTable.GetCellValues(rowIndex);

                //while (columnIndex < _initialTableColumnAmount)
                //{
                //    var cellContent = _initialTable.GetCellValue(rowIndex, columnIndex);

                //    //if (cellContent != null)
                //    //    ProccesInitialTableCell(rowIndex, columnIndex, accountId, cellContent);

                //    columnIndex++;
                //}
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

                InsertRowInRelationTable(accountId, manager, GetRelationType(columnHeader));
            }
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

        private bool IsManagerHeader(string titleValue)
        {
            return titleValue == _initialTableManagerColumnTitle
                   || titleValue == _initialTableRegionColumnTitle
                   || titleValue == _initialTableAreaColumnTitle
                   || titleValue == _initialTableDivisionColumnTitle;
        }

        private void InsertRowInManagerTable(string manager)
        {
            var rowCells = new List<TableCell>();
            rowCells.Add(new TableCell(0, 0, manager));
            _managerTable.AddRow(rowCells);

            //_managerTable.SetCellValue(_managerTableRowIndex, 0, manager);
            //_managerTableRowIndex++;
        }

        private void InsertRowInRelationTable(string accountId, string manager, string type)
        {
            var rowCells = new List<TableCell>();
            rowCells.Add(new TableCell(0, 0, accountId));
            rowCells.Add(new TableCell(0, 1, manager));
            rowCells.Add(new TableCell(0, 2, type));

            _relationTable.AddRow(rowCells);
            
            //_relationTable.SetCellValue(_relationTableRowIndex, 0, accountId);
            //_relationTable.SetCellValue(_relationTableRowIndex, 1, manager);
            //_relationTable.SetCellValue(_relationTableRowIndex, 2, type);

            //_relationTableRowIndex++;            
        }

        /*
        private void ProccesInitialTableCell(TableCell cell, string accountId)
        {
            string columnHeader = _initialTableHeaders[cell.ColumnIndex];

            var formattedCellValue = FormatCellValue(cell);

            if (IsManagerHeader(columnHeader))
            {
                ProcessAsManagerCell(cell, columnHeader, accountId);
            }
            else
            {
                ProcessAsAccountCell(cell, columnHeader);
            }
        }
        */

        /*
        private void ProcessAsAccountCell(TableCell cell, string columnHeader)
        {
            int columnIndex = _accountTableHeaders[columnHeader];

            _accountTable.SetCellValue(cell.RowIndex, columnIndex, cell.Value);
        }
        */
    }
}