using Microsoft.Office.Interop.Excel;

namespace Utility
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

        public static void CreateNewFile(string filePath)
        {
            var xlsApp = new Application();
            object misValue = System.Reflection.Missing.Value;
            var workbook = xlsApp.Workbooks.Add(misValue);
            xlsApp.DisplayAlerts = false;
            workbook.SaveAs(filePath, XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            workbook.Close();
            xlsApp.Quit();
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