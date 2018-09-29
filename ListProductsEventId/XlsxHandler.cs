using System;
using System.Collections.Generic;
using ListProductsEventId.Tests;
using Microsoft.Office.Interop.Excel;

namespace ListProductsEventId
{
    internal class XlsxHandler : IExcelHandler
    {
        private string _filePath;
        private Workbook _xlWorkBook;
        private Application _xlApp;

        public XlsxHandler(string filePath)
        {
            _filePath = filePath;
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
                for (int i = 2; i < xlSheets.Count; i++)
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
                Console.WriteLine($"Column '{columnTitle}' created.");
                _xlWorkBook.RefreshAll();//xlsx 的檔案要加這行
                _xlApp.Calculate();//xlsx 的檔案要加這行
                _xlWorkBook.Save();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        public Dictionary<int, string> GetSpecifiedColumnAllCellValue(int sheetIndex, string columnTitle)
        {
            throw new NotImplementedException();
        }

        public bool ExistValueOnSheet(string sheetName, string value)
        {
            throw new NotImplementedException();
        }

        public string GetSpecifiedCellValue(string sheetName, int columnIndex, int rowIndex)
        {
            throw new NotImplementedException();
        }

        public void SetCellValue(int sheetIndex, int columnIndex, int rowIndex, string value)
        {
            throw new NotImplementedException();
        }

        public Dictionary<int, string> GetAllColumnTitle()
        {
            throw new NotImplementedException();
        }

        public bool ExistSheet(string sheetName)
        {
            throw new NotImplementedException();
        }
    }
}