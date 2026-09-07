using System.Collections.Generic;

using Automation.GenerateIgxl.Basic.Business.GenNwire.Business;
using Automation.Reader;

using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace Automation.Test.UT.Common
{
    [TestClass]
    public class NwireFlowTests
    {
        [TestMethod]
        public void PrintPattern_Should_Add_Print_FlowRow_With_Correct_Values()
        {
            // Arrange
            ExcelPackage package = new ExcelPackage();
            ExcelWorksheet ws = package.Workbook.Worksheets.Add("Flow");
            var flow = new SubFlowSheet(ws);

            var target = new NwireFlow();
            string pattern = "ABC";

            // Act
            target.PrintPattern(flow, pattern);

            // Assert
            Assert.AreEqual(1, flow.Rows.Count);

            IgxlLib.IgxlBase.FlowRow row = flow.Rows[0];
            Assert.AreEqual("Print", row.Opcode);
            Assert.AreEqual("Flow_ABC_Strat", row.Parameter);
        }

        [TestMethod]
        public void CreateSubFlow_WhenItemIsNull_ShouldNotThrow()
        {
            var sut = new NwireFlow();
            var dic = new Dictionary<string, string>();

            SubFlowSheet flow = sut.CreateSubFlow(null, dic);

            Assert.AreNotEqual(null, flow);
        }

        [TestMethod]
        public void CreateSubFlow_WhenItemInDictionary_ShouldUseValue()
        {
            var sut = new NwireFlow();
            var dic = new Dictionary<string, string>
            {
                { "ITEM1", "VALUE1" }
            };

            SubFlowSheet flow = sut.CreateSubFlow(" ITEM1 ", dic);

            Assert.AreEqual(NwireSetting.ConFlownWire + "VALUE1", flow.Name);
        }

        [TestMethod]
        public void CreateSubFlow_WhenItemNotInDictionary_ShouldUseItem()
        {
            var sut = new NwireFlow();
            var dic = new Dictionary<string, string>();

            SubFlowSheet flow = sut.CreateSubFlow("ITEM2", dic);

            Assert.AreEqual(NwireSetting.ConFlownWire + "ITEM2", flow.Name);
        }
    }
}
