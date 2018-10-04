using System;
using System.Collections.Generic;
using System.Linq;
using Utility;

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
                excelHandler.Save();
                Console.WriteLine($"'{sheetName}' 已建立");
            }
        }

        public MaxIndexPair LookupFromOtherSheetByProductId()
        {
            var columns = excelHandler.GetAllColumnTitle(1);
            var productIds = excelHandler.GetSpecifiedColumnAllCellValue(1, "商品主號");
            int rowIndex = productIds.Count + 1;
            Console.WriteLine($"已取得所有商品主號，共 {productIds.Count} 個");

            int columnIndex = 0;
            foreach (var column in columns)
            {
                columnIndex = column.Key;
                var columnTitle = column.Value;

                if (!excelHandler.ExistSheet(columnTitle))
                {
                    columnIndex--;
                    break;
                }

                try
                {
                    var groupId = excelHandler.GetSpecifiedCellValue(columnTitle, 2, 1);
                    Dictionary<int, string> productIdsOnSheet = excelHandler.GetSpecifiedColumnAllCellValue(columnTitle, "商品主號");
                    var intersectIds = productIds.Values.Intersect(productIdsOnSheet.Values);
                    foreach (var intersectId in intersectIds)
                    {
                        var mappingRowIndex = productIds.FirstOrDefault(x => x.Value == intersectId).Key;
                        excelHandler.SetCellValue(1, columnIndex, mappingRowIndex, "," + groupId);
                    }
                }
                catch (Exception)
                {
                    Console.WriteLine($"ERROR!! {columnTitle} 那行查表時出錯！！");
                }

                excelHandler.Save();
                Console.WriteLine($"'{columnTitle}' 查表完成");
            }

            return new MaxIndexPair {MaxColumnIndex = columnIndex, MaxRowIndex = rowIndex};
        }

        public void AddConcatenateAheadColumn(int sheetIndex, int columnIndex, int rowIndex)
        {
            excelHandler.AddConcatenateAheadColumn(sheetIndex, columnIndex, rowIndex);
            excelHandler.Save();
            Console.WriteLine("建立串接行");
        }
    }
}