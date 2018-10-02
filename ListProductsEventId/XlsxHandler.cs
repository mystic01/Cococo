using System;
using System.Collections.Generic;
using ListProductsEventId.Tests;
using Microsoft.Office.Interop.Excel;

namespace ListProductsEventId
{
    internal class XlsxHandler : IExcelHandler
    {
        private Workbook _xlWorkBook;
        private Application _xlApp;

        public XlsxHandler(string filePath)
        {
            _xlApp = new Application();
            object missingValue = System.Reflection.Missing.Value;

            //for xls
            //_xlWorkBook = _xlApp.Workbooks.Open(filePath, CorruptLoad: true);

            //for xlsx
            _xlWorkBook = _xlApp.Workbooks.Open(filePath, missingValue, false, missingValue, missingValue,
                missingValue, true, missingValue, missingValue, true, missingValue, missingValue, missingValue);
        }

        ~XlsxHandler()
        {
            _xlWorkBook?.Close(false);
            _xlApp?.Quit();
        }

        public List<string> GetWorkSheetNamesExceptFirst()
        {
            var sheetNamesExceptFirst = new List<string>();
            try
            {
                var xlSheets = _xlWorkBook.Sheets;
                for (int i = 2; i <= xlSheets.Count; i++)
                {
                    var xlSheet = xlSheets[i] as Worksheet;
                    sheetNamesExceptFirst.Add(xlSheet.Name);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }

            return sheetNamesExceptFirst;
        }

        public void CreateColumnAhead(string columnTitle)
        {
            try
            {
                var firstSheet = _xlWorkBook.Sheets[1] as Worksheet;
                Range a1Range = firstSheet.Range["A1"];
                a1Range.EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight,
                    XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                firstSheet.Cells[1, 1] = columnTitle;
            }
            catch (Exception e)
            {
                Console.WriteLine($"ERROR!! {e}");
            }
        }

        public void Save()
        {
            _xlWorkBook.RefreshAll(); //xlsx 的檔案要加這行
            _xlApp.Calculate(); //xlsx 的檔案要加這行
            _xlWorkBook.Save();
        }

        public void AddConcatenateAheadColumn(int sheetIndex, int columnIndex, int rowIndex)
        {
            Worksheet worksheet = _xlWorkBook.Sheets[sheetIndex];
            Range insertRange = worksheet.Cells[1, columnIndex+1];
            insertRange.EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight);

            var formula = "=";
            for (int i = 1; i <= columnIndex; i++)
            {
                formula += $"RC[{-i}]";
                if (i < columnIndex)
                    formula += "&";
            }

            for (int i = 2; i <= rowIndex; i++)
            {
                (worksheet.Cells[i, columnIndex + 1] as Range).NumberFormat = "General";
                worksheet.Cells[i, columnIndex + 1] = formula;
            }
        }

        public Dictionary<int, string> GetSpecifiedColumnAllCellValue(string sheetName, string columnTitle)
        {
            var worksheet = FindWorkSheet(sheetName);
            return SpecifiedColumnAllCellValue(columnTitle, worksheet);
        }

        public Dictionary<int, string> GetSpecifiedColumnAllCellValue(int sheetIndex, string columnTitle)
        {
            var worksheet = _xlWorkBook.Sheets[sheetIndex] as Worksheet;
            return SpecifiedColumnAllCellValue(columnTitle, worksheet);
        }

        private Dictionary<int, string> SpecifiedColumnAllCellValue(string columnTitle, Worksheet worksheet)
        {
            var columnIndex = FindColumnByTitle(columnTitle, worksheet);

            var result = new Dictionary<int, string>();
            if (columnIndex != null)
            {
                var rowIndex = 2;
                while ((worksheet.Cells[rowIndex, columnIndex] as Range).Value != null)
                {
                    result[rowIndex] = (worksheet.Cells[rowIndex, columnIndex] as Range).Value.ToString();
                    rowIndex++;
                }
            }

            return result;
        }

        private int? FindColumnByTitle(string columnTitle, Worksheet worksheet)
        {
            var findColumn = false;
            var columnIndex = 1;
            while (worksheet.Cells[1, columnIndex] != null)
            {
                if ((worksheet.Cells[1, columnIndex] as Range).Value == columnTitle)
                {
                    findColumn = true;
                    break;
                }

                columnIndex++;
            }

            if (findColumn)
                return columnIndex;
            else
                return null;
        }

        private Worksheet FindWorkSheet(string sheetName)
        {
            Worksheet workSheet = null;
            for (int i = 1; i <= _xlWorkBook.Sheets.Count; i++)
            {
                if ((_xlWorkBook.Sheets[i] as Worksheet).Name == sheetName)
                    workSheet = _xlWorkBook.Sheets[i];
            }

            return workSheet;
        }

        public string GetSpecifiedCellValue(string sheetName, int columnIndex, int rowIndex)
        {
            var workSheet = FindWorkSheet(sheetName);
            return (workSheet.Cells[rowIndex, columnIndex] as Range)?.Value?.ToString();
        }

        public void SetCellValue(int sheetIndex, int columnIndex, int rowIndex, string value)
        {
            var workSheet = _xlWorkBook.Sheets[sheetIndex];
            workSheet.Cells[rowIndex, columnIndex] = value;
        }

        public Dictionary<int, string> GetAllColumnTitle(int sheetIndex)
        {
            Worksheet workSheet = _xlWorkBook.Sheets[sheetIndex];
            var result = new Dictionary<int, string>();

            var columnIndex = 1;
            while ((workSheet.Cells[1, columnIndex] as Range).Value != null)
            {
                result[columnIndex] = (workSheet.Cells[1, columnIndex] as Range).Value;

                columnIndex++;
            }

            return result;
        }

        public bool ExistSheet(string sheetName)
        {
            return (FindWorkSheet(sheetName) != null);
        }
    }
}