using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using Utility;

namespace ListProductsEventId
{
    public class ExcelBaseHandler : IExcelHandler
    {
        protected Workbook _xlWorkBook;
        protected Application _xlApp;

        public virtual List<string> GetWorkSheetNamesExceptFirst()
        {
            throw new NotImplementedException();
        }

        public virtual Dictionary<int, string> GetSpecifiedColumnAllCellValue(int sheetIndex, string columnTitle)
        {
            throw new NotImplementedException();
        }

        public virtual Dictionary<int, string> GetSpecifiedColumnAllCellValue(string sheetName, string columnTitle)
        {
            var worksheet = FindWorkSheet(sheetName);
            return SpecifiedColumnAllCellValue(columnTitle, worksheet);
        }

        protected Dictionary<int, string> SpecifiedColumnAllCellValue(string columnTitle, Worksheet worksheet)
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

        public virtual string GetSpecifiedCellValue(string sheetName, int columnIndex, int rowIndex)
        {
            throw new NotImplementedException();
        }

        public string GetSpecifiedCellValue(int sheetIndex, int columnIndex, int rowIndex)
        {
            var worksheet = _xlWorkBook.Sheets[sheetIndex] as Worksheet;
            return (worksheet.Cells[rowIndex, columnIndex] as Range)?.Value?.ToString();
        }

        public void SetCellValue(int sheetIndex, int columnIndex, int rowIndex, string value)
        {
            var workSheet = _xlWorkBook.Sheets[sheetIndex];
            workSheet.Cells[rowIndex, columnIndex] = value;
        }

        public virtual Dictionary<int, string> GetAllColumnTitle(int sheetIndex)
        {
            throw new NotImplementedException();
        }

        public virtual bool ExistSheet(string sheetName)
        {
            throw new NotImplementedException();
        }

        public virtual void Save()
        {
            throw new NotImplementedException();
        }

        public virtual void AddConcatenateAheadColumn(int sheetIndex, int columnIndex, int rowIndex)
        {
            throw new NotImplementedException();
        }

        protected Worksheet FindWorkSheet(string sheetName)
        {
            Worksheet workSheet = null;
            for (int i = 1; i <= _xlWorkBook.Sheets.Count; i++)
            {
                if ((_xlWorkBook.Sheets[i] as Worksheet).Name == sheetName)
                    workSheet = _xlWorkBook.Sheets[i];
            }

            return workSheet;
        }

        public virtual int AddWorksheet(string oriSheetName, string newSheetName)
        {
            var beforeSheet = FindWorkSheet(oriSheetName);
            var newWorksheet = _xlWorkBook.Sheets.Add(After:beforeSheet) as Worksheet;
            newWorksheet.Name = newSheetName;
            return newWorksheet.Index;
        }

        private int? FindColumnByTitle(string columnTitle, Worksheet worksheet)
        {
            var findColumn = false;
            var columnIndex = 1;

            try
            {

                while (worksheet.Cells[1, columnIndex] != null)
                {
                    if ((worksheet.Cells[1, columnIndex] as Range).Value == columnTitle)
                    {
                        findColumn = true;
                        break;
                    }

                    columnIndex++;
                }
            }
            catch (Exception)
            {
                return null;
            }

            if (findColumn)
                return columnIndex;
            else
                return null;
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
    }
}