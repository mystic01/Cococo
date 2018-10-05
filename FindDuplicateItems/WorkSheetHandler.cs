using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using Utility;

namespace FindDuplicateItems
{
    public class WorkSheetHandler
    {
        private readonly IExcelHandler _xlsHandler1;
        private readonly IExcelHandler _xlsHandler2;

        public WorkSheetHandler(XlsHandler xlsHandler1, XlsHandler xlsHandler2)
        {
            _xlsHandler1 = xlsHandler1;
            _xlsHandler2 = xlsHandler2;
        }

        public void OutputDuplicateItems()
        {
            var collectionIds1 = _xlsHandler1.GetSpecifiedColumnAllCellValue("整理", "商品主號");
            var collectionIds2 = _xlsHandler2.GetSpecifiedColumnAllCellValue("整理", "商品主號");
            var intervalIds1 = _xlsHandler1.GetSpecifiedColumnAllCellValue("周間", "商品主號");
            var intervalIds2 = _xlsHandler2.GetSpecifiedColumnAllCellValue("周間", "商品主號");
            var groupId1 = _xlsHandler1.GetSpecifiedColumnAllCellValue("團購", "商品主號");
            var groupId2 = _xlsHandler2.GetSpecifiedColumnAllCellValue("團購", "商品主號");
            var allSourceAndIdLookupSet = new Dictionary<Tuple<IExcelHandler, string>, Dictionary<int, string>>
            {
                {new Tuple<IExcelHandler, string>(_xlsHandler1, "整理"),collectionIds1},
                {new Tuple<IExcelHandler, string>(_xlsHandler2,"整理"),collectionIds2},
                {new Tuple<IExcelHandler, string>(_xlsHandler1, "周間"),intervalIds1},
                {new Tuple<IExcelHandler, string>(_xlsHandler2, "周間"),intervalIds2},
                {new Tuple<IExcelHandler, string>(_xlsHandler1,"團購"),groupId1},
                {new Tuple<IExcelHandler, string>(_xlsHandler2, "團購"),groupId2}
            };

            var pids = new List<string>();
            pids.AddRange(collectionIds1.Values);
            pids.AddRange(collectionIds2.Values);
            pids.AddRange(intervalIds1.Values);
            pids.AddRange(intervalIds2.Values);
            pids.AddRange(groupId1.Values);
            pids.AddRange(groupId2.Values);

            var duplicatePids = pids.GroupBy(x => x).Where(g => g.Count() > 1).Select(y => y.Key).ToList();
            Console.WriteLine($"重複的共有 {duplicatePids.Count} 個品項");

            var exePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var newFilePath = Path.Combine(exePath, "重複總整理.xls");
            XlsHandler.CreateNewFile(newFilePath);

            IExcelHandler newXlsHandler = new XlsHandler(newFilePath);
            newXlsHandler.AddWorksheet("", "重複項目");
            newXlsHandler.CreateColumnAhead("來源");
            newXlsHandler.CreateColumnAhead("網路價");
            newXlsHandler.CreateColumnAhead("商品名稱");

            const int PRODUCT_COLUMNINDEX = 1;
            const int PRICE_COLUMNINDEX = 2;
            const int SOURCE_COLUMNINDEX = 3;

            var currRowIndex = 2;
            foreach (var pid in duplicatePids)
            {
                Console.Write(".");
                var minPrice = int.MaxValue;
                var minPriceRowIndex = 0;
                var minPriceIsMM = false;
                foreach (var sourceAndIdLookup in allSourceAndIdLookupSet)
                {
                    if (sourceAndIdLookup.Value.ContainsValue(pid))
                    {
                        var allIdLookup = sourceAndIdLookup.Value.Where(x => x.Value == pid);
                        foreach (var lookup in allIdLookup)
                        {
                            var rowIndex = lookup.Key;
                            var sourcePair = sourceAndIdLookup.Key;
                            var xlsHandler = sourcePair.Item1;
                            var sheetName = sourcePair.Item2;
                            var productName = xlsHandler.GetSpecifiedCellValue(sheetName, "商品名稱", rowIndex);
                            var price = xlsHandler.GetSpecifiedCellValue(sheetName, "網路價", rowIndex);
                            var sourceFileStr = xlsHandler.Name.Substring(0, 2);
                            var sourceStr = sourceFileStr + "-" + sheetName;

                            newXlsHandler.SetCellValue(1, PRODUCT_COLUMNINDEX, currRowIndex, productName);
                            newXlsHandler.SetCellValue(1, PRICE_COLUMNINDEX, currRowIndex, price);
                            newXlsHandler.SetCellValue(1, SOURCE_COLUMNINDEX, currRowIndex, sourceStr);
                            newXlsHandler.SetCellColor(1, PRODUCT_COLUMNINDEX, currRowIndex, Color.Silver);
                            newXlsHandler.SetCellColor(1, PRICE_COLUMNINDEX, currRowIndex, Color.Silver);
                            newXlsHandler.SetCellColor(1, SOURCE_COLUMNINDEX, currRowIndex, Color.Silver);

                            int priceInt = int.MaxValue;
                            try
                            {
                                priceInt = (int) double.Parse(price);
                            }
                            catch (Exception)
                            {
                                Console.WriteLine($"ERROR!! 存在有非數字的網路價錢: {xlsHandler.Name}.{sheetName}.{rowIndex}=>{price}");
                                newXlsHandler.SetCellColor(1, PRICE_COLUMNINDEX, currRowIndex, Color.Red);
                            }

                            if (sourceFileStr == "MM")
                            {
                                if (priceInt <= minPrice)
                                {
                                    minPrice = priceInt;
                                    minPriceRowIndex = currRowIndex;
                                    minPriceIsMM = true;
                                }
                            }
                            else//PF
                            {
                                if (!minPriceIsMM && priceInt < minPrice)
                                {
                                    minPrice = priceInt;
                                    minPriceRowIndex = currRowIndex;
                                    minPriceIsMM = false;
                                }
                            }

                            currRowIndex++;
                        }
                    }
                }

                if (minPriceRowIndex > 0)
                {
                    newXlsHandler.SetCellColor(1, PRODUCT_COLUMNINDEX, minPriceRowIndex, Color.White);
                    newXlsHandler.SetCellColor(1, PRICE_COLUMNINDEX, minPriceRowIndex, Color.White);
                    newXlsHandler.SetCellColor(1, SOURCE_COLUMNINDEX, minPriceRowIndex, Color.White);
                }
            }

            Console.WriteLine("");
            newXlsHandler.Save();
        }
    }
}