using System;
using ListProductsEventId;
using Microsoft.Office.Interop.Excel;
using Utility;

namespace GroupTenItems
{
    public class XlsHandler : ExcelBaseHandler
    {
        public XlsHandler(string filePath)
        {
            _xlApp = new Application();
            object missingValue = System.Reflection.Missing.Value;

            //for xls
            _xlWorkBook = _xlApp.Workbooks.Open(filePath, CorruptLoad: true);
        }

        ~XlsHandler()
        {
            _xlWorkBook?.Close(false);
            _xlApp?.Quit();
        }

        public override void Save()
        {
            _xlApp.DisplayAlerts = false;
            _xlWorkBook.Save();
        }
    }
}