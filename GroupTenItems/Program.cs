using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace GroupTenItems
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine($"002 Group Ten Items v{Version}");
            Console.WriteLine("======================================================");

            var exePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string[] files = Directory.GetFiles(exePath, "MM整理*.xls");
            if (files.Length == 0)
                Console.WriteLine("找不到任何開頭檔名為「MM整理」的 xls 檔。");
            else
            {
                Console.WriteLine($"開始處理 '{Path.GetFileName(files[0])}'");
                var workSheetHandler = new WorkSheetHandler(new XlsHandler(files[0]));
                workSheetHandler.GroupTenItemsViaPid("完整", "活動設定");
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
