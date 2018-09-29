using System;
using System.IO;
using System.Reflection;


namespace ListProductsEventId
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            Console.WriteLine("001 List Products Event Id v0.1");
            Console.WriteLine("===========================================");

            var exePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string[] files = Directory.GetFiles(exePath, "批次分類表*.xlsx");
            if (files.Length == 0)
                Console.WriteLine("找不到任何開頭檔名為「批次分類表」的 xlsx 檔。");
            else
            {
                var filePath = Path.Combine(exePath, files[0]);
                var workSheetHandler = new WorkSheetHandler(new XlsxHandler(filePath));
                workSheetHandler.GenerateColumnsViaSheetName();
            }

            Console.WriteLine("DONE");
            Console.ReadKey();
        }
    }
}