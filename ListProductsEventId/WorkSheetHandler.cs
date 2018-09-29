using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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

        }
    }
}
