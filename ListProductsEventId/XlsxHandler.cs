using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace ListProductsEventId
{
    internal class XlsxHandler : ExcelBaseHandler
    {
        public XlsxHandler(string filePath)
        {
            _xlApp = new Application();
            object missingValue = System.Reflection.Missing.Value;

            //for xlsx
            _xlWorkBook = _xlApp.Workbooks.Open(filePath, missingValue, false, missingValue, missingValue,
                missingValue, true, missingValue, missingValue, true, missingValue, missingValue, missingValue);
        }

        ~XlsxHandler()
        {
            _xlWorkBook?.Close(false);
            _xlApp?.Quit();
        }

        public override List<string> GetWorkSheetNamesExceptFirst()
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

        public override void Save()
        {
            _xlWorkBook.RefreshAll(); //xlsx 的檔案要加這行
            _xlApp.Calculate(); //xlsx 的檔案要加這行
            _xlWorkBook.Save();
        }

        public override void AddConcatenateAheadColumn(int sheetIndex, int columnIndex, int rowIndex)
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

        public override Dictionary<int, string> GetSpecifiedColumnAllCellValue(int sheetIndex, string columnTitle)
        {
            var worksheet = _xlWorkBook.Sheets[sheetIndex] as Worksheet;
            return SpecifiedColumnAllCellValue(columnTitle, worksheet);
        }

        public override string GetSpecifiedCellValue(string sheetName, int columnIndex, int rowIndex)
        {
            var workSheet = FindWorkSheet(sheetName);
            return (workSheet.Cells[rowIndex, columnIndex] as Range)?.Value?.ToString();
        }

        public override Dictionary<int, string> GetAllColumnTitle(int sheetIndex)
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

        public override bool ExistSheet(string sheetName)
        {
            return (FindWorkSheet(sheetName) != null);
        }
    }
}