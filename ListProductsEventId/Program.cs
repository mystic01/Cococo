using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;

namespace ListProductsEventId
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            Console.WriteLine($"001 List Products Event Id v{Version}");
            Console.WriteLine("============================================");

            var exePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string[] files = Directory.GetFiles(exePath, "批次分類表*.xlsx");
            if (files.Length == 0)
                Console.WriteLine("找不到任何開頭檔名為「批次分類表」的 xlsx 檔。");
            else
            {
                Console.WriteLine($"開始處理 '{Path.GetFileName(files[0])}'");
                var workSheetHandler = new WorkSheetHandler(new XlsxHandler(files[0]));
                workSheetHandler.GenerateColumnsViaSheetName();
                Console.WriteLine("------");
                var lookupIndexResult = workSheetHandler.LookupFromOtherSheetByProductId();
                Console.WriteLine("------");
                workSheetHandler.AddConcatenateAheadColumn(1, lookupIndexResult.MaxColumnIndex, lookupIndexResult.MaxRowIndex);
            }

            Console.WriteLine("============== 完成，請按任意鍵關閉視窗 ==============");
            Console.ReadKey();
        }

        public static string Version
        {
            get
            {
                Assembly asm = Assembly.GetExecutingAssembly();
                FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(asm.Location);
                return String.Format("{0}.{1}", fvi.ProductMajorPart, fvi.ProductMinorPart);
            }
        }
    }
}