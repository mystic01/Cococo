using System.Collections.Generic;
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

        public void LookupFromOtherSheetByProductId()
        {
            Dictionary<int, string> columns = excelHandler.GetAllColumnTitle();
            Dictionary<int, string> productIds = excelHandler.GetSpecifiedColumnAllCellValue(1, "商品主");
            foreach (var column in columns)
            {
                var columnIndex = column.Key;
                var columnTitle = column.Value;

                if (!excelHandler.ExistSheet(columnTitle))
                    break;

                foreach (var productId in productIds)
                {
                    if (excelHandler.ExistValueOnSheet(columnTitle, productId.Value))
                    {
                        var groupId = excelHandler.GetSpecifiedCellValue(columnTitle, 2, 1);
                        excelHandler.SetCellValue(1, columnIndex, productId.Key, groupId);
                    }
                }
            }

        }
    }
}