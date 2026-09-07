using System.Collections.Generic;

using Automation.GenerateIgxl.BinCut.Base;
using Automation.GenerateIgxl.Scan.Harvest.Flow;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.BinCut.BinCutInstance;

namespace Automation.Test.UT.Scan
{
    [TestClass]
    public class NonBincutInstanceBaseTests
    {
        private class TestNonBinCutInstance : NonBinCutInstanceBase
        {
            public void TestAddCondition(string siteVar, ref FlowRow flowRow)
            {
                AddCondition(siteVar, ref flowRow);
            }
        }

        [TestMethod]
        public void AddCondition_Should_Set_DeviceCondition_FlagTrue_When_No_Exclamation()
        {
            // Arrange
            var flowRow = new FlowRow();
            var instance = new TestNonBinCutInstance();

            // Act
            instance.TestAddCondition("Site1", ref flowRow);

            // Assert
            Assert.AreEqual("Site1", flowRow.DeviceName);
            Assert.AreEqual("Flag-true", flowRow.DeviceCondition);
        }

        [TestMethod]
        public void AddCondition_Should_Set_DeviceCondition_FlagFalse_When_Exclamation()
        {
            var flowRow = new FlowRow();
            var instance = new TestNonBinCutInstance();

            instance.TestAddCondition("!Site2", ref flowRow);

            Assert.AreEqual("Site2", flowRow.DeviceName);
            Assert.AreEqual("Flag-false", flowRow.DeviceCondition);
        }

        [TestMethod]
        public void AddCondition_Should_Reverse_When_EndsWithFalse()
        {
            var flowRow = new FlowRow();
            var instance = new TestNonBinCutInstance();

            instance.TestAddCondition("Site3False", ref flowRow);
            Assert.AreEqual("Site3False", flowRow.DeviceName);
            Assert.AreEqual("Flag-false", flowRow.DeviceCondition);

            flowRow = new FlowRow();
            instance.TestAddCondition("!Site3False", ref flowRow);
            Assert.AreEqual("Site3False", flowRow.DeviceName);
            Assert.AreEqual("Flag-true", flowRow.DeviceCondition);
        }

        [TestMethod]
        public void GetIfFlowRows_Should_Return_Simple_List_When_No_SiteVar()
        {
            var instance = new TestNonBinCutInstance();
            var testFlow = new FlowRow { Label = "TestLabel", Job = "J1" };
            var flowRows = new List<FlowRow> { new() { Label = "R1" } };

            List<FlowRow> result = instance.GetIfFlowRows(null, [testFlow], flowRows);

            Assert.AreEqual(2, result.Count);
            Assert.AreEqual("TestLabel", result[0].Label);
            Assert.AreEqual("R1", result[1].Label);
        }

        [TestMethod]
        public void GetIfFlowRows_Should_Wrap_With_If_EndIf_When_SiteVar_Has_Logical_Operator()
        {
            var instance = new TestNonBinCutInstance();
            var dataRow = new BinCutFinalInstanceRow
            {
                BinCutInstanceRow = new BinCutInstanceRow { SiteVar = "A && B" }
            };
            var testFlow = new FlowRow { Label = "Test", Job = "J1" };

            List<FlowRow> result = instance.GetIfFlowRows(dataRow, [testFlow], null);

            Assert.AreEqual(3, result.Count);
            Assert.AreEqual("(A && B) || F_Debug_all", result[0].Parameter);
            Assert.AreEqual("Test", result[1].Label);
            Assert.AreEqual(OpCode.EndIf.ToLower(), result[2].Opcode.ToLower());
        }

        [TestMethod]
        public void GetIfFlowRows_Should_Set_DeviceCondition_When_SiteVar_Simple()
        {
            var instance = new TestNonBinCutInstance();
            var dataRow = new BinCutFinalInstanceRow
            {
                BinCutInstanceRow = new BinCutInstanceRow { SiteVar = "!FlagSite" }
            };
            var testFlow = new FlowRow { Label = "TestFlow" };

            List<FlowRow> result = instance.GetIfFlowRows(dataRow, [testFlow], null);

            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("FlagSite", testFlow.DeviceName);
            Assert.AreEqual("Flag-false", testFlow.DeviceCondition);
        }
    }
}
