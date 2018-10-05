using System;
using System.IO;
using System.Reflection;
using Utility;

namespace FindDuplicateItems
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            Console.WriteLine($"003 Find Duplicate Items v{Utility.Utility.Version}");
            Console.WriteLine("======================================================");

            var exePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string[] files1 = Directory.GetFiles(exePath, "MM整理*.xls");
            string[] files2 = Directory.GetFiles(exePath, "PF整理*.xls");
            if (files1.Length == 0)
                Console.WriteLine("找不到任何開頭檔名為「MM整理」的 xls 檔。");
            else if (files2.Length == 0)
                Console.WriteLine("找不到任何開頭檔名為「PF整理」的 xls 檔。");
            else
            {
                Console.WriteLine($"開始處理 '{Path.GetFileName(files1[0])} 及 '{Path.GetFileName(files2[0])}'");
                var workSheetHandler = new WorkSheetHandler(new XlsHandler(files1[0]), new XlsHandler(files2[0]));
                workSheetHandler.OutputDuplicateItems();
            }

            Console.WriteLine("============== 完成，請按任意鍵關閉視窗 ==============");
            Console.ReadKey();
        }
    }
}