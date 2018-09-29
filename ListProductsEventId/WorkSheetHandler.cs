using ListProductsEventId.Tests;

namespace ListProductsEventId
{
    public class WorkSheetHandler
    {
        private IExcelHandler excelHandler;

        public WorkSheetHandler(IExcelHandler excelHandler)
        {
            this.excelHandler = excelHandler;
        }

        public void GenerateColumnsViaSheetName()
        {
            var workSheetNamesExceptFirst = excelHandler.GetWorkSheetNamesExceptFirst();
            workSheetNamesExceptFirst.Reverse();
            foreach (var sheetName in workSheetNamesExceptFirst)
            {
                excelHandler.CreateColumnAhead(sheetName);
            }
        }
    }
}