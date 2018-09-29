using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;
using NSubstitute.Core.SequenceChecking;

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
    }
}