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
            string[] files = Directory.GetFiles(exePath, "�妸������*.xlsx");
            if (files.Length == 0)
                Console.WriteLine("�䤣�����}�Y�ɦW���u�妸������v�� xlsx �ɡC");
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