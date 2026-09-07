using System.Collections.Generic;

using Automation.GenerateIgxl.BinCut.Base;
using Automation.Static;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.BinCut.Flow;
using TestPlanLib.BinCut.FlowNew;

namespace Automation.Test.UT.BinCut
{
    [TestClass]
    public class BinCutSourceItemTests : TestBase
    {
        private BinCutFlowSheetRow _flowRow = null!;
        private NewBinCutFlowSheetRow _newFlowRow = null!;

        [TestInitialize]
        public void Setup()
        {
            _flowRow = new BinCutFlowSheetRow("", [])
            {
                SheetName = "Sheet1",
                BinningDomain = "CPU",
                PerformanceMode = "MSX001",
                AllOther = "ALL",
                PinInfos = []
            };

            _newFlowRow = new NewBinCutFlowSheetRow("", "")
            {
                RowNum = 1,
                BinningDomain = "GPU",
                TableType = EnumBinCutTableType.Hv,
                TableBinType = EnumBinCutTableBinType.Bin1
            };

            LocalSpecs.TarFolder = OutputPath;
            PinMapSheet pinMapSheet = new PinMapSheet("");
            Pin pin = new Pin("VDD_A", "power");
            pinMapSheet.AddPin(pin);
            Pin pin1 = new Pin("VDD_B", "power");
            pinMapSheet.AddPin(pin1);
            TestProgram.IgxlWorkBk.PinMapPair = new KeyValuePair<string, PinMapSheet>(FolderStructure.DirPinMap, pinMapSheet);
        }

        [TestMethod]
        public void GetDomain_ShouldDetectCPU()
        {
            var item = new BinCutSourceItem(_flowRow, _newFlowRow, EnumColumnName.FUNC, "CPU something");
            string result = item.GetDomain("CPU TEST");
            Assert.AreEqual("Cpu", result);
        }

        [TestMethod]
        public void GetDomain_ShouldDetectGPU()
        {
            var item = new BinCutSourceItem(_flowRow, _newFlowRow, EnumColumnName.FUNC, "GPU something");
            string result = item.GetDomain("GPU TEST");
            Assert.AreEqual("Gfx", result);
        }

        [TestMethod]
        public void GetBinType_ShouldReturnHBV_WhenHvType()
        {
            var item = new BinCutSourceItem(_flowRow, _newFlowRow, EnumColumnName.FUNC, "content");
            string result = item.GetBinType();
            Assert.AreEqual("HBV", result);
        }

        [TestMethod]
        public void GetEnableWord_ShouldReturnConcatenatedJobs()
        {
            var item = new BinCutSourceItem(_flowRow, _newFlowRow, EnumColumnName.FUNC, "content")
            {
                JobCount = 3
            };
            string result = item.GetEnableWord(["CP1", "CP2"]);
            Assert.AreEqual("CP1||CP2", result);
        }

        [TestMethod]
        public void GetVbtFunction_ShouldReturn_HVCC_TMPS_VT_When_TMPS_Hv()
        {
            _newFlowRow.TableType = EnumBinCutTableType.Hv;
            var item = new BinCutSourceItem(_flowRow, _newFlowRow, EnumColumnName.FUNC, "TEMP SENSOR");

            string result = item.GetVbtFunction();
            Assert.AreEqual("GradeSearch_HVCC_TMPS_VT", result);
        }

        [TestMethod]
        public void IsTargetSpiPattern_ShouldMatchAllWords()
        {
            var item = new BinCutSourceItem(_flowRow, _newFlowRow, EnumColumnName.FUNC, "CPU SOC");
            bool result = item.IsTargetSpiPattern("CPU SOC", "SOMETHING_CPU_SOC_PATTERN");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void IsTargetSpiPattern_ShouldFail_WhenWordMissing()
        {
            var item = new BinCutSourceItem(_flowRow, _newFlowRow, EnumColumnName.FUNC, "CPU SOC");
            bool result = item.IsTargetSpiPattern("CPU SOC", "SOMETHING_GPU_PATTERN");
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void ShouldMatch_TMPS_InstanceName()
        {
            var item = new BinCutSourceItem(_flowRow, _newFlowRow, EnumColumnName.FUNC, "TEMP SENSOR");
            bool result = item.IsTargetInstanceName("MSX001_TMPS_CPU_HBV");
            Assert.IsTrue(result, "Expected TMPS instance name to match TEMP SENSOR content");
        }

        [TestMethod]
        public void ShouldMatch_ILB_InstanceName()
        {
            _flowRow.AllOther = "LV";
            var item = new BinCutSourceItem(_flowRow, _newFlowRow, EnumColumnName.FUNC, "CPU ILB");
            bool result = item.IsTargetInstanceName("MSX001_CPU_ILB_BV");
            Assert.IsTrue(result, "Expected ILB instance name to match CPU ILB content");
        }

        [TestMethod]
        public void ShouldNotMatch_ElbWhenDifferentMode()
        {
            var item = new BinCutSourceItem(_flowRow, _newFlowRow, EnumColumnName.FUNC, "SOC ELB");
            bool result = item.IsTargetInstanceName("MSX001_CPU_ELB_BV");
            Assert.IsFalse(result, "Expected non-matching SOC ELB vs CPU_ELB instance");
        }

        [TestMethod]
        public void ShouldMatch_PerformanceModeAndDomain()
        {
            _flowRow.AllOther = "LV";
            var item = new BinCutSourceItem(_flowRow, _newFlowRow, EnumColumnName.Mbist, "CPU BIST");
            bool result = item.IsTargetInstanceName("MSX001_CPUMBIST_BV");
            Assert.IsTrue(result, "Expected Mbist instance to match CPU MBIST content");
        }

        [TestMethod]
        public void ShouldFail_WhenModeMismatch()
        {
            var item = new BinCutSourceItem(_flowRow, _newFlowRow, EnumColumnName.Mbist, "GPU BIST");
            bool result = item.IsTargetInstanceName("MSX001_CPU_MBIST_BV");
            Assert.IsFalse(result, "Expected GPU BIST not to match CPU instance");
        }

        #region GetBinningDomain Tests
        [TestMethod]
        public void GetBinningDomain_SingleDomain_ReturnsFormattedString()
        {
            _newFlowRow.BinningDomain = "CORE";
            var service = new BinCutSourceItem(_flowRow, _newFlowRow, EnumColumnName.Mbist, "GPU BIST");

            string result = service.GetBinningDomain();

            Assert.AreEqual("VDD_CORE_MSX001", result);
        }

        [TestMethod]
        public void GetBinningDomain_MultipleDomains_ReturnsJoinedStringAndCallsAddPinGroup()
        {
            _newFlowRow.BinningDomain = "A,B";
            var service = new BinCutSourceItem(_flowRow, _newFlowRow, EnumColumnName.Mbist, "GPU BIST");

            string result = service.GetBinningDomain();

            Assert.IsTrue(result.Contains("VDD_A_B"));
        }
        #endregion

        #region JudgeIsTargetFlow Tests
        [TestMethod]
        public void JudgeIsTargetFlow_ExactMatch_ReturnsTrue()
        {
            var service = new BinCutSourceItem(_flowRow, _newFlowRow, EnumColumnName.Mbist, "MainFlow");
            BinCutFinalInstanceRow row = CreateMockRow("MainFlow", "All");

            bool result = service.JudgeIsTargetFlow(row);

            Assert.IsTrue(result);
        }

        [TestMethod]
        public void JudgeIsTargetFlow_RegexHashMatch_ReturnsTrue()
        {
            var service = new BinCutSourceItem(_flowRow, _newFlowRow, EnumColumnName.Mbist, "CPU#CORE");
            BinCutFinalInstanceRow row = CreateMockRow("CPU CORE", "All");

            bool result = service.JudgeIsTargetFlow(row);

            Assert.IsTrue(result);
        }

        [TestMethod]
        public void JudgeIsTargetFlow_ColonSplit_MatchesLastPart()
        {
            var service = new BinCutSourceItem(_flowRow, _newFlowRow, EnumColumnName.Mbist, "Prefix:TargetFlow");
            BinCutFinalInstanceRow row = CreateMockRow("TargetFlow", "All");

            bool result = service.JudgeIsTargetFlow(row);

            Assert.IsTrue(result);
        }

        [TestMethod]
        public void JudgeIsTargetFlow_SemicolonList_MatchesAny()
        {
            var service = new BinCutSourceItem(_flowRow, _newFlowRow, EnumColumnName.Mbist, "FlowA;FlowB;FlowC");
            BinCutFinalInstanceRow row = CreateMockRow("FlowB", "All");

            bool result = service.JudgeIsTargetFlow(row);

            Assert.IsTrue(result);
        }
        #endregion

        #region GetJobFlag (Logic via JudgeIsTargetFlow)
        [TestMethod]
        public void GetJobFlag_NegativeLogic_ReturnsCorrectBool()
        {
            var service = new BinCutSourceItem(_flowRow, _newFlowRow, EnumColumnName.Mbist, "FlowA");
            // !FT2 means "True if job is NOT FT2"
            BinCutFinalInstanceRow row = CreateMockRow("FlowA", "!FT2");

            bool result = service.JudgeIsTargetFlow(row);

            Assert.IsTrue(result);
        }
        #endregion

        // Helper to create the data structure
        private static BinCutFinalInstanceRow CreateMockRow(string flow, string stage)
        {
            return new BinCutFinalInstanceRow
            {
                BinCutInstanceRow = new BinCutInstanceRow
                {
                    FlowName = flow,
                    JobTestStage = stage
                }
            };
        }
    }
}
