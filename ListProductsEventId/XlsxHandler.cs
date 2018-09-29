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
                _xlWorkBook.RefreshAll();//xlsx ���ɮ׭n�[�o��
                _xlApp.Calculate();//xlsx ���ɮ׭n�[�o��
                _xlWorkBook.Save();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }
    }
}