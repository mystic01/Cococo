using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;
using Utility;

namespace GroupTenItems.Tests
{
    [TestClass()]
    public class WorkSheetHandlerTests
    {
        [TestMethod()]
        public void GroupTenItemsTest_OneGroup()
        {
            var excelHandler = Substitute.For<IExcelHandler>();
            excelHandler.AddWorksheet("完整", "活動設定").Returns(2);
            excelHandler.GetSpecifiedColumnAllCellValue("完整", "pid").Returns(new Dictionary<int, string> { { 2, "777" }, { 3, "888" } });
            excelHandler.GetSpecifiedColumnAllCellValue("完整", "活動名稱").Returns(new Dictionary<int, string> { { 2, "GOGOGOA" }, { 3, "GOGOGOA" } });
            excelHandler.GetSpecifiedColumnAllCellValue("完整", "商品名稱").Returns(new Dictionary<int, string> { { 2, "ProductA" }, { 3, "ProductB" } });

            excelHandler.GetSpecifiedCellValue(2, 3, 2).Returns("777");
            excelHandler.GetSpecifiedCellValue(2, 3, 3).Returns("888");
            excelHandler.GetSpecifiedCellValue(2, 4, 2).Returns((string)null);

            var target = new WorkSheetHandler(excelHandler);

            target.GroupTenItemsViaPid("完整", "活動設定");

            excelHandler.Received().SetCellValue(2, 1, 2, "ProductA");
            excelHandler.Received().SetCellValue(2, 1, 3, "ProductB");
            excelHandler.Received().SetCellValue(2, 2, 2, "GOGOGOA");
            excelHandler.Received().SetCellValue(2, 2, 3, "GOGOGOA");
            excelHandler.Received().SetCellValue(2, 3, 2, "777");
            excelHandler.Received().SetCellValue(2, 3, 3, "888");
            excelHandler.Received().SetCellValue(2, 4, 3, "777,888");
        }

        [TestMethod()]
        public void GroupTenItemsTest_TwoGroups()
        {
            var excelHandler = Substitute.For<IExcelHandler>();
            excelHandler.AddWorksheet("完整", "活動設定").Returns(2);
            excelHandler.GetSpecifiedColumnAllCellValue("完整", "pid").Returns(new Dictionary<int, string> { { 2, "777" }, { 3, "888" }, { 4, "999" } });
            excelHandler.GetSpecifiedColumnAllCellValue("完整", "活動名稱").Returns(new Dictionary<int, string> { { 2, "GOGOGOA" }, { 3, "GOGOGOA" }, { 4, "GOGOGOB" } });
            excelHandler.GetSpecifiedColumnAllCellValue("完整", "商品名稱").Returns(new Dictionary<int, string> { { 2, "ProductA" }, { 3, "ProductB" }, { 4, "ProductC" } });

            excelHandler.GetSpecifiedCellValue(2, 2, 2).Returns("GOGOGOA");
            excelHandler.GetSpecifiedCellValue(2, 2, 3).Returns("GOGOGOA");
            excelHandler.GetSpecifiedCellValue(2, 2, 4).Returns("GOGOGOB");
            excelHandler.GetSpecifiedCellValue(2, 3, 2).Returns("777");
            excelHandler.GetSpecifiedCellValue(2, 3, 3).Returns("888");
            excelHandler.GetSpecifiedCellValue(2, 3, 4).Returns("999");
            excelHandler.GetSpecifiedCellValue(2, 4, 2).Returns((string)null);
            excelHandler.GetSpecifiedCellValue(2, 4, 3).Returns("777,888");

            var target = new WorkSheetHandler(excelHandler);

            target.GroupTenItemsViaPid("完整", "活動設定");

            excelHandler.Received().SetCellValue(2, 1, 2, "ProductA");
            excelHandler.Received().SetCellValue(2, 1, 3, "ProductB");
            excelHandler.Received().SetCellValue(2, 1, 4, "ProductC");
            excelHandler.Received().SetCellValue(2, 2, 2, "GOGOGOA");
            excelHandler.Received().SetCellValue(2, 2, 3, "GOGOGOA");
            excelHandler.Received().SetCellValue(2, 2, 4, "GOGOGOB");
            excelHandler.Received().SetCellValue(2, 3, 2, "777");
            excelHandler.Received().SetCellValue(2, 3, 3, "888");
            excelHandler.Received().SetCellValue(2, 3, 4, "999");
            excelHandler.Received().SetCellValue(2, 4, 3, "777,888");
            excelHandler.Received().SetCellValue(2, 4, 4, "999");
        }

        [TestMethod()]
        public void GroupTenItemsTest_OneBigGroup()
        {
            var excelHandler = Substitute.For<IExcelHandler>();
            excelHandler.AddWorksheet("完整", "活動設定").Returns(2);
            excelHandler.GetSpecifiedColumnAllCellValue("完整", "pid").Returns(new Dictionary<int, string> { { 2, "1" }, { 3, "2" }, { 4, "3" }, { 5, "4" }, { 6, "5" }, { 7, "6" }, { 8, "7" }, { 9, "8" }, { 10, "9" }, { 11, "10" }, { 12, "11" } });
            excelHandler.GetSpecifiedColumnAllCellValue("完整", "活動名稱").Returns(new Dictionary<int, string> { { 2, "GOGOGOA" }, { 3, "GOGOGOA" }, { 4, "GOGOGOA" }, { 5, "GOGOGOA" }, { 6, "GOGOGOA" }, { 7, "GOGOGOA" }, { 8, "GOGOGOA" }, { 9, "GOGOGOA" }, { 10, "GOGOGOA" }, { 11, "GOGOGOA" }, { 12, "GOGOGOA" } });
            excelHandler.GetSpecifiedColumnAllCellValue("完整", "商品名稱").Returns(new Dictionary<int, string> { { 2, "GOGOGOA" }, { 3, "GOGOGOA" }, { 4, "GOGOGOA" }, { 5, "GOGOGOA" }, { 6, "GOGOGOA" }, { 7, "GOGOGOA" }, { 8, "GOGOGOA" }, { 9, "GOGOGOA" }, { 10, "GOGOGOA" }, { 11, "GOGOGOA" }, { 12, "GOGOGOA" } });

            excelHandler.GetSpecifiedCellValue(2, 2, 2).Returns("GOGOGOA");
            excelHandler.GetSpecifiedCellValue(2, 2, 3).Returns("GOGOGOA");
            excelHandler.GetSpecifiedCellValue(2, 2, 4).Returns("GOGOGOA");
            excelHandler.GetSpecifiedCellValue(2, 2, 5).Returns("GOGOGOA");
            excelHandler.GetSpecifiedCellValue(2, 2, 6).Returns("GOGOGOA");
            excelHandler.GetSpecifiedCellValue(2, 2, 7).Returns("GOGOGOA");
            excelHandler.GetSpecifiedCellValue(2, 2, 8).Returns("GOGOGOA");
            excelHandler.GetSpecifiedCellValue(2, 2, 9).Returns("GOGOGOA");
            excelHandler.GetSpecifiedCellValue(2, 2, 10).Returns("GOGOGOA");
            excelHandler.GetSpecifiedCellValue(2, 2, 11).Returns("GOGOGOA");
            excelHandler.GetSpecifiedCellValue(2, 2, 12).Returns("GOGOGOA");

            excelHandler.GetSpecifiedCellValue(2, 3, 2).Returns("1");
            excelHandler.GetSpecifiedCellValue(2, 3, 3).Returns("2");
            excelHandler.GetSpecifiedCellValue(2, 3, 4).Returns("3");
            excelHandler.GetSpecifiedCellValue(2, 3, 5).Returns("4");
            excelHandler.GetSpecifiedCellValue(2, 3, 6).Returns("5");
            excelHandler.GetSpecifiedCellValue(2, 3, 7).Returns("6");
            excelHandler.GetSpecifiedCellValue(2, 3, 8).Returns("7");
            excelHandler.GetSpecifiedCellValue(2, 3, 9).Returns("8");
            excelHandler.GetSpecifiedCellValue(2, 3, 10).Returns("9");
            excelHandler.GetSpecifiedCellValue(2, 3, 11).Returns("10");
            excelHandler.GetSpecifiedCellValue(2, 3, 12).Returns("11");

            excelHandler.GetSpecifiedCellValue(2, 4, 11).Returns("1,2,3,4,5,6,7,8,9,10");

            var target = new WorkSheetHandler(excelHandler);
            target.GroupTenItemsViaPid("完整", "活動設定");

            excelHandler.Received().SetCellValue(2, 4, 11, "1,2,3,4,5,6,7,8,9,10");
            excelHandler.Received().SetCellValue(2, 4, 12, "11");
        }
    }
}