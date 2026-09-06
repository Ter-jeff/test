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
    public class ScanNonBinCutInstanceTests
    {
        private static BinCutFinalInstanceRow MakeDataRow(string siteVar = "", string callFlow = "", string domain = "D1", string block = "B1")
        {
            return new BinCutFinalInstanceRow
            {
                BinCutInstanceRow = new BinCutInstanceRow
                {
                    SiteVar = siteVar,
                    CallFlow = callFlow
                },
                Domain = domain,
                Block = block
            };
        }

        [TestMethod]
        public void GetIfFlowRows_Should_Handle_Null_DataRow()
        {
            var instance = new ScanNonBinCutInstance();
            var testFlow = new FlowRow { Label = "MainFlow", Job = "J1" };
            var limitRows = new List<FlowRow> { new() { Label = "Limit1" } };

            List<FlowRow> result = instance.GetIfFlowRows(null, [testFlow], limitRows);

            Assert.AreEqual(2, result.Count);
            Assert.AreEqual("MainFlow", result[0].Label);
            Assert.AreEqual("Limit1", result[1].Label);
        }

        [TestMethod]
        public void GetIfFlowRows_Should_Insert_BinTable_When_IsSsn_True()
        {
            var instance = new ScanNonBinCutInstance();
            BinCutFinalInstanceRow dataRow = MakeDataRow(siteVar: "");
            var testFlow = new FlowRow { Label = "MainFlow", IsSsn = true, Job = "J1" };

            List<FlowRow> result = instance.GetIfFlowRows(dataRow, [testFlow], null);

            Assert.AreEqual(2, result.Count);
            Assert.AreEqual(OpCode.BinTable, result[1].Opcode);
            StringAssert.Contains(result[1].Parameter, "SSN_INIT_Fail");
        }

        [TestMethod]
        public void GetIfFlowRows_Should_Create_If_Block_When_SiteVar_Has_And_Operator()
        {
            var instance = new ScanNonBinCutInstance();
            BinCutFinalInstanceRow dataRow = MakeDataRow(siteVar: "A && B");
            var testFlow = new FlowRow { Label = "MainFlow", Job = "J1" };

            List<FlowRow> result = instance.GetIfFlowRows(dataRow, [testFlow], null);

            Assert.AreEqual(3, result.Count);
            Assert.AreEqual(OpCode.If.ToLower(), result[0].Opcode.ToLower());
            Assert.AreEqual("MainFlow", result[1].Label);
            Assert.AreEqual(OpCode.EndIf.ToLower(), result[2].Opcode.ToLower());
        }

        [TestMethod]
        public void GetIfFlowRows_Should_Create_If_Block_With_CallFlow_And_Ssn()
        {
            var instance = new ScanNonBinCutInstance();
            BinCutFinalInstanceRow dataRow = MakeDataRow(siteVar: "X || Y", callFlow: "Call_Flow_1");
            var testFlow = new FlowRow { Label = "Flow", Job = "JX", IsSsn = true };

            List<FlowRow> result = instance.GetIfFlowRows(dataRow, [testFlow], null);

            Assert.AreEqual(5, result.Count);
            Assert.AreEqual(OpCode.If.ToLower(), result[0].Opcode.ToLower());
            Assert.AreEqual("Flow", result[1].Label);
            Assert.AreEqual(OpCode.BinTable, result[2].Opcode);
            Assert.AreEqual(OpCode.Call, result[3].Opcode);
            Assert.AreEqual("Call_Flow_1", result[3].Parameter);
            Assert.AreEqual(OpCode.EndIf.ToLower(), result[4].Opcode.ToLower());
        }

        [TestMethod]
        public void GetIfFlowRows_Should_AddCondition_When_Simple_SiteVar()
        {
            var instance = new ScanNonBinCutInstance();
            BinCutFinalInstanceRow dataRow = MakeDataRow(siteVar: "!FlagSite");
            var testFlow = new FlowRow { Label = "MainFlow" };

            List<FlowRow> result = instance.GetIfFlowRows(dataRow, [testFlow], null);

            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("FlagSite", testFlow.DeviceName);
            Assert.AreEqual("Flag-false", testFlow.DeviceCondition);
        }

        [TestMethod]
        public void GetIfFlowRows_Should_Add_CallFlow_When_DeviceName_Empty()
        {
            var instance = new ScanNonBinCutInstance();
            BinCutFinalInstanceRow dataRow = MakeDataRow(siteVar: "", callFlow: "MyCall");
            var testFlow = new FlowRow { Label = "MainFlow" };

            List<FlowRow> result = instance.GetIfFlowRows(dataRow, [testFlow], null);

            Assert.AreEqual(2, result.Count);
            Assert.AreEqual("MainFlow", result[0].Label);
            Assert.AreEqual(OpCode.Call, result[1].Opcode);
            Assert.AreEqual("MyCall", result[1].Parameter);
        }

        [TestMethod]
        public void GetIfFlowRows_Should_Create_If_CallFlow_When_DeviceName_NotEmpty()
        {
            var instance = new ScanNonBinCutInstance();
            BinCutFinalInstanceRow dataRow = MakeDataRow(siteVar: "Flag1", callFlow: "MyCall");
            var testFlow = new FlowRow { Label = "MainFlow", DeviceName = "Flag1" };

            List<FlowRow> result = instance.GetIfFlowRows(dataRow, [testFlow], null);

            // Assert
            Assert.AreEqual(4, result.Count);
            Assert.AreEqual("", result[0].Label);
            Assert.AreEqual("", result[1].Opcode);
            Assert.AreEqual(OpCode.Call, result[2].Opcode);
            Assert.AreEqual(OpCode.EndIf.ToLower(), result[3].Opcode.ToLower());
        }
    }
}
