using Utility;

namespace GroupTenItems
{
    public class WorkSheetHandler
    {
        private readonly IExcelHandler _excelHandler;

        public WorkSheetHandler(IExcelHandler excelHandler)
        {
            _excelHandler = excelHandler;
        }

        public void GroupTenItemsViaPid(string oriSheetName, string newSheetName)
        {
            var newSheetIndex = _excelHandler.AddWorksheet(oriSheetName, newSheetName);
            var pidLookup = _excelHandler.GetSpecifiedColumnAllCellValue(oriSheetName, "pid");
            var eventLookup = _excelHandler.GetSpecifiedColumnAllCellValue(oriSheetName, "活動名稱");
            var productNameLookup = _excelHandler.GetSpecifiedColumnAllCellValue(oriSheetName, "商品名稱");

            const int PRODUCT_COLUMNINDEX = 1;
            const int EVENT_COLUMNINDEX = 2;
            const int PID_COLUMNINDEX = 3;
            const int PIDS_COLUMNINDEX = 4;

            _excelHandler.CreateColumnAhead("pid");
            _excelHandler.CreateColumnAhead("活動名稱");
            _excelHandler.CreateColumnAhead("商品名稱");
            _excelHandler.SetCellValue(newSheetIndex, PIDS_COLUMNINDEX, 1, "pids");


            var rowIndex = 2;
            while (pidLookup.ContainsKey(rowIndex))
            {
                _excelHandler.SetCellValue(newSheetIndex, PRODUCT_COLUMNINDEX, rowIndex, productNameLookup[rowIndex]);
                _excelHandler.SetCellValue(newSheetIndex, EVENT_COLUMNINDEX, rowIndex, eventLookup[rowIndex]);
                _excelHandler.SetCellValue(newSheetIndex, PID_COLUMNINDEX, rowIndex, pidLookup[rowIndex]);
                rowIndex++;
            }

            var groupRowIndex = 0;
            for (int i = 2; i < rowIndex; i++)
            {
                groupRowIndex++;
                var currEventCell = _excelHandler.GetSpecifiedCellValue(newSheetIndex, 2, i);
                var nextEventCell = _excelHandler.GetSpecifiedCellValue(newSheetIndex, 2, i + 1);
                var sameGroup = currEventCell == nextEventCell;

                if (sameGroup && groupRowIndex == 10)
                {
                    groupRowIndex = 0;
                    var groupListString = "";
                    for (int j = 9; j >= 0; j--)
                    {
                        groupListString += _excelHandler.GetSpecifiedCellValue(newSheetIndex, PID_COLUMNINDEX, i - j);
                        if (j > 0)
                            groupListString += ",";
                    }

                    _excelHandler.SetCellValue(newSheetIndex, PIDS_COLUMNINDEX, i, groupListString);
                }
                else if (!sameGroup || (i == rowIndex - 1))//Last One
                {
                    groupRowIndex = 0;
                    var groupListString = _excelHandler.GetSpecifiedCellValue(newSheetIndex, PID_COLUMNINDEX, i); ;
                    for (int j = 1; j <= 9; j++)
                    {
                        var prevPidsCell = _excelHandler.GetSpecifiedCellValue(newSheetIndex, PIDS_COLUMNINDEX, i - j);
                        if (prevPidsCell != null)
                            break;

                        groupListString = "," + groupListString;
                        groupListString = _excelHandler.GetSpecifiedCellValue(newSheetIndex, PID_COLUMNINDEX, i - j) + groupListString;
                    }

                    _excelHandler.SetCellValue(newSheetIndex, PIDS_COLUMNINDEX, i, groupListString);
                }
            }

            _excelHandler.Save();
        }
    }
}