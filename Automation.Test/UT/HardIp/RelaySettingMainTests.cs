using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.SpecialSetting;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class RelaySettingMainTests : FunctionTestBase
    {

        [TestMethod]
        public void GenRelaySettingInJob_SingleSetting_ShouldGenerateCorrectFlowRow()
        {
            SubFlowSheet sheet = new SubFlowSheet("Test");

            List<FlowRow> rows = RelaySettingMain.GenRelaySettingInJob(
                sheet,
                "RelayOn:A",
                "JOB1",
                "ENA");

            Assert.AreEqual(1, rows.Count);
            Assert.AreEqual("ENA", rows[0].Enable);
            Assert.AreEqual("JOB1", rows[0].Job);
            Assert.AreEqual(OpCode.Test, rows[0].Opcode);
            Assert.AreEqual(
                HardIpConstData.PrefixAtgRelay + "RelayOn_A",
                rows[0].Parameter);
            Assert.AreEqual(1, sheet.Rows.Count);
        }

        [TestMethod]
        public void GenRelaySettingInJob_MultipleSetting_ShouldSplitCorrectly()
        {
            List<FlowRow> rows = RelaySettingMain.GenRelaySettingInJob(
                null,
                "RelayOn:A;RelayOff:B",
                "JOB1",
                "");

            Assert.AreEqual(2, rows.Count);
            Assert.AreEqual("AtgRelay_RelayOn_A", rows[0].Parameter);
            Assert.AreEqual("AtgRelay_RelayOff_B", rows[1].Parameter);
        }

        [TestMethod]
        public void GenRelaySettingInJob_ShouldReplaceSpecialChar()
        {
            List<FlowRow> rows = RelaySettingMain.GenRelaySettingInJob(
                null,
                "RelayOn:A&B:C",
                "JOB1",
                "");

            Assert.AreEqual(1, rows.Count);
            Assert.AreEqual(
                "AtgRelay_RelayOn_A_B_C",
                rows[0].Parameter);
        }

        [TestMethod]
        public void GenRelaySetting_ShouldHandleMultipleJobs()
        {
            Dictionary<string, string> input = new Dictionary<string, string>()
            {
                { "JOB1", "RelayOn:A" },
                { "JOB2", "RelayOff:B" }
            };

            List<FlowRow> rows = RelaySettingMain.GenRelaySetting(
                null,
                input,
                "EN",
                false);

            Assert.AreEqual(2, rows.Count);
            Assert.AreEqual("JOB1", rows[0].Job);
            Assert.AreEqual("JOB2", rows[1].Job);
        }

        [TestMethod]
        public void GenRelaySetting_WithReverse_ShouldReverseSetting()
        {
            Dictionary<string, string> input = new Dictionary<string, string>()
            {
                { "JOB1", "RelayOffA RelayOnB" }
            };

            List<FlowRow> rows = RelaySettingMain.GenRelaySetting(
                null,
                input,
                "",
                true);

            Assert.AreEqual(1, rows.Count);
            Assert.IsTrue(rows[0].Parameter.Contains("AtgRelay_"));
        }

        [TestMethod]
        public void GenRelaySettingInJob_NullFlowSheet_ShouldNotThrow()
        {
            List<FlowRow> rows = RelaySettingMain.GenRelaySettingInJob(
                null,
                "RelayOn:A",
                "JOB1",
                "");

            Assert.AreEqual(1, rows.Count);
        }

        [TestMethod]
        public void GenRelaySettingInJob_EmptySetting_ShouldStillGenerateRow()
        {
            List<FlowRow> rows = RelaySettingMain.GenRelaySettingInJob(
                null,
                "",
                "JOB1",
                "");

            Assert.AreEqual(1, rows.Count);
            Assert.AreEqual("AtgRelay_", rows[0].Parameter);
        }

    }
}
