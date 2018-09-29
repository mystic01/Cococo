using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;

namespace ListProductsEventId.Tests
{
    [TestClass()]
    public class WorkSheetHandlerTests
    {
        [TestMethod()]
        public void GenerateColumnsViaSheetNameTest_2Sheets()
        {
            var excelHandler = Substitute.For<IExcelHandler>();
            excelHandler.GetWorkSheetNamesExceptFirst().Returns(new List<string> { "AAA", "BBB" });
            var target = new WorkSheetHandler(excelHandler);

            target.GenerateColumnsViaSheetName();

            Received.InOrder(() =>
            {
                excelHandler.Received(1).CreateColumnAhead("BBB");
                excelHandler.Received(1).CreateColumnAhead("AAA");
            });
        }

        [TestMethod()]
        public void LookupFromOtherSheetByProductIdTest_SheetAAAWithA123_Gid888__SheetBBBWithB123_Gid999()
        {
            var excelHandler = Substitute.For<IExcelHandler>();
            excelHandler.GetAllColumnTitle().Returns(new Dictionary<int, string> { { 1, "AAA" }, { 2, "BBB" }, { 3, "CCC" } });
            excelHandler.GetSpecifiedColumnAllCellValue(1, "商品主")
                .Returns(new Dictionary<int, string> { { 2, "A123" }, { 3, "B123" } });
            excelHandler.ExistSheet("AAA").Returns(true);
            excelHandler.ExistSheet("BBB").Returns(true);
            excelHandler.ExistSheet("CCC").Returns(false);
            excelHandler.ExistValueOnSheet("AAA", "A123").Returns(true);
            excelHandler.ExistValueOnSheet("AAA", "B123").Returns(false);
            excelHandler.ExistValueOnSheet("BBB", "A123").Returns(false);
            excelHandler.ExistValueOnSheet("BBB", "B123").Returns(true);
            excelHandler.GetSpecifiedCellValue("AAA", 2, 1).Returns("888");
            excelHandler.GetSpecifiedCellValue("BBB", 2, 1).Returns("999");

            var target = new WorkSheetHandler(excelHandler);

            target.LookupFromOtherSheetByProductId();

            Received.InOrder(() =>
            {
                excelHandler.Received(1).SetCellValue(1, 1, 2, "888");
                excelHandler.Received(1).SetCellValue(1, 2, 3, "999");
            });
        }
    }
}