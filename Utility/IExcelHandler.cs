﻿using System.Collections.Generic;
using System.Drawing;

namespace Utility
{
    public interface IExcelHandler
    {
        List<string> GetWorkSheetNamesExceptFirst();

        void CreateColumnAhead(string columnTitle);

        Dictionary<int, string> GetSpecifiedColumnAllCellValue(int sheetIndex, string columnTitle);

        Dictionary<int, string> GetSpecifiedColumnAllCellValue(string sheetName, string v);

        string GetSpecifiedCellValue(string sheetName, int columnIndex, int rowIndex);

        string GetSpecifiedCellValue(int sheetIndex, int columnIndex, int rowIndex);

        string GetSpecifiedCellValue(string sheetName, string columnTitle, int rowIndex);

        void SetCellValue(int sheetIndex, int columnIndex, int rowIndex, string value);

        Dictionary<int, string> GetAllColumnTitle(int sheetIndex);

        bool ExistSheet(string sheetName);

        void Save();

        void AddConcatenateAheadColumn(int sheetIndex, int columnIndex, int rowIndex);

        int AddWorksheet(string oriSheetName, string newSheetName);

        string Name { get; }
        void SetCellColor(int sheetIndex, int columnIndex, int rowIndex, Color backColor);
    }
}