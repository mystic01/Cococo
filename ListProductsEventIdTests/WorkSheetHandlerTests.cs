using ListProductsEventId;
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
            excelHandler.GetAllColumnTitle(1).Returns(new Dictionary<int, string> { { 1, "AAA" }, { 2, "BBB" }, { 3, "CCC" } });
            excelHandler.GetSpecifiedColumnAllCellValue(1, "商品主號")
                .Returns(new Dictionary<int, string> { { 2, "A123" }, { 3, "B123" } });
            excelHandler.ExistSheet("AAA").Returns(true);
            excelHandler.ExistSheet("BBB").Returns(true);
            excelHandler.ExistSheet("CCC").Returns(false);
            excelHandler.GetSpecifiedColumnAllCellValue("AAA", "商品主號").Returns(new Dictionary<int, string> { { 2, "A123" } });
            excelHandler.GetSpecifiedColumnAllCellValue("BBB", "商品主號").Returns(new Dictionary<int, string> { { 2, "B123" } });
            excelHandler.GetSpecifiedCellValue("AAA", 2, 1).Returns("888");
            excelHandler.GetSpecifiedCellValue("BBB", 2, 1).Returns("999");

            var target = new WorkSheetHandler(excelHandler);

            target.LookupFromOtherSheetByProductId();

            Received.InOrder(() =>
            {
                excelHandler.Received(1).SetCellValue(1, 1, 2, ",888");
                excelHandler.Received(1).SetCellValue(1, 2, 3, ",999");
            });
        }

        [TestMethod()]
        public void AddConcatenateAheadColumnTest()
        {
            var excelHandler = Substitute.For<IExcelHandler>();
            var target = new WorkSheetHandler(excelHandler);
            target.AddConcatenateAheadColumn(0,0,0);
            Received.InOrder(() =>
            {
                excelHandler.Received(1).AddConcatenateAheadColumn(0,0,0);
                excelHandler.Received(1).Save();
            });
        }
    }
}