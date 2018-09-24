﻿using System;
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
        
        private int _managerTableRowIndex;
        private int _relationTableRowIndex;

        public void Normalize(ITable initialTable, ITable accountTable, ITable managerTable, ITable relationTable)
        {
            _initialTable = initialTable;
            _accountTable = accountTable;
            _managerTable = managerTable;
            _relationTable = relationTable;
            
            FormatAccountTable(GetInitialColumnTitles());
            FormatManagerTable();
            FormatRelationTable();

            ProccessInitialTable();
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
                    _accountTable.SetValue(0, columnIndex, columnTitle);
                    columnIndex++;
                }
            }
        }

        private List<string> GetInitialColumnTitles()
        {
            var columnTitles = new List<string>();

            int columnIndex = 0;
            string columnTitle;
            while (!string.IsNullOrEmpty(columnTitle = _initialTable.GetValue(0, columnIndex)))
            {
                columnTitles.Add(columnTitle);
                columnIndex++;
            }

            return columnTitles;
        }

        private void FormatManagerTable()
        {
            _managerSet = new HashSet<string>();

            _managerTable.SetValue(0, 0, _managerTableManagerColumnTitle);

            _managerTableRowIndex = 1;
        }

        private void FormatRelationTable()
        {
            _relationTable.SetValue(0, 0, _relationTableAccountColumnTitle);
            _relationTable.SetValue(0, 1, _relationTableManagerColumnTitle);
            _relationTable.SetValue(0, 2, _relationTableTypeColumnTitle);

            _relationTableRowIndex = 1;
        }

        private void ProccessInitialTable()
        {
            int rowIndex = 1;
            string accountId;
            while (!string.IsNullOrEmpty(accountId = _initialTable.GetValue(rowIndex, 0)))
            {
                int columnIndex = 0;
                string cellContent;
                while ((cellContent = _initialTable.GetValue(rowIndex, columnIndex)) != null)
                {
                    ProccesInitialTableCell(rowIndex, columnIndex, accountId, cellContent);

                    columnIndex++;
                }
                rowIndex++;
            }
        }

        private void ProccesInitialTableCell(int rowIndex, int columnIndex, string accountId, string cellContent)
        {
            string columnTitle = _initialTable.GetValue(0, columnIndex);

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

        private void ProcessAsManagerCell(string rowId, string columnTitle, string cellContent)
        {
            var managers = cellContent.Split(_cellContentDividers, StringSplitOptions.RemoveEmptyEntries);

            foreach (var manager in managers)
            {
                if (!_managerSet.Contains(manager))
                {
                    AddRowIntoManagerTable(manager);
                    AddRowIntoRelationTable(rowId, manager, GetRelationType(columnTitle));
                    _managerSet.Add(manager);
                }
            }
        }

        private void ProcessAsAccountCell(int rowIndex, string columnTitle, string cellContent)
        {
            int columnIndex = _accountTableColumnTitles[columnTitle];
            _accountTable.SetValue(rowIndex, columnIndex, cellContent);
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

        private void AddRowIntoManagerTable(string manager)
        {
            _managerTable.SetValue(_managerTableRowIndex, 0, manager);
            _managerTableRowIndex++;
        }

        private void AddRowIntoRelationTable(string accountId, string manager, string type)
        {
            _relationTable.SetValue(_relationTableRowIndex, 0, accountId);
            _relationTable.SetValue(_relationTableRowIndex, 1, manager);

            _relationTableRowIndex++;
        }
    }
}