using System.Collections.Generic;

namespace ListProductsEventId.Tests
{
    public interface IExcelHandler
    {
        List<string> GetWorkSheetNamesExceptFirst();
        void CreateColumnAhead(string columnTitle);
    }
}