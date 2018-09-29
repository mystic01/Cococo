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

            excelHandler.Received(1).CreateColumnAhead("AAA");
            excelHandler.Received(1).CreateColumnAhead("BBB");
        }
    }
}