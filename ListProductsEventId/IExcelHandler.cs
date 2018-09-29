using System.Collections.Generic;

namespace ListProductsEventId.Tests
{
    public interface IExcelHandler
    {
        List<string> GetWorkSheetNamesExceptFirst();
        void CreateColumnAhead(string columnTitle);
        Dictionary<int, string> GetSpecifiedColumnAllCellValue(int sheetIndex, string columnTitle);
        bool ExistValueOnSheet(string sheetName, string value);
        string GetSpecifiedCellValue(string sheetName, int columnIndex, int rowIndex);
        void SetCellValue(int sheetIndex, int columnIndex, int rowIndex, string value);
        Dictionary<int, string> GetAllColumnTitle();
        bool ExistSheet(string sheetName);
    }
}