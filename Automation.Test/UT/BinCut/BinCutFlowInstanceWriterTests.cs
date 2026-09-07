using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.BinCut.Base;
using Automation.GenerateIgxl.BinCut.Business;
using Automation.InputManager.Data;
using Automation.Static;
using Automation.Test.Static;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlBase.MultiRow;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.BinCut;
using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.BinCut.Binning;
using TestPlanLib.BinCut.Flow;

namespace Automation.Test.UT.BinCut
{
    [TestClass]
    public class BinCutFlowInstanceWriterTests
    {

        private BinCutInputData _binCutInputManager = null!;
        private List<BinCutFinalInstanceRow> _binCutFinalInstanceRows = null!;
        private BinCutFlowInstanceWriter _writer = null!;

        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestProgram.VbtFunctionLib.AddVbtFunctionRange(TestSuiteInitialize.Functions);
        }

        [TestInitialize]
        public void Setup()
        {
            _binCutInputManager = new BinCutInputData();
            _binCutFinalInstanceRows = [];

            _writer = new BinCutFlowInstanceWriter(
                _binCutInputManager,
                _binCutFinalInstanceRows,
                ["CP1", "FT1"]);
        }


        private const string TestMode = "MS001";
        private const string TestJob = "FT1";
        private static IdsDistributionTable CreateIdsDistributionTable(string polation)
        {
            var table = new IdsDistributionTable();
            table.AllIdsPowers.Add(new IdsPower
            {
                PowerName = TestMode,
                Job = TestJob,
                IdsInfos =
                {
                    new IdsInfo
                    {
                        Polation = { polation }
                    }
                }
            });

            return table;
        }

        private static BinningTables CreateBinningTables(string mode)
        {
            var tables = new BinningTables();

            tables.Add(new BinningTable
            {
                ModeIdx = 0,
                IntModeLIdx = 1,
                IntModeHIdx = 2,
                DomainIdx = 3,
                IntSkipTestIdx = 4,
                Job = "FT1",
                Rows =
                [
                    new()
                    {
                        RowData = [mode, mode, mode, "PCPU", "Yes"]
                    }
                ]
            });

            return tables;
        }

        private static List<BinCutFlowInstanceWriter.PendingInterpolate> CreatePendingInterpolate()
        {
            return
            [
                new BinCutFlowInstanceWriter.PendingInterpolate
                {
                    Interpolation = "",
                    Domain = "SOC",
                    Mode = "MS001",
                    IntModeL = "MS000",
                    IntModeH = "MS002",
                    SheetName = "Sheet1"
                }
            ];
        }

        private static List<string> CreatePerformanceList(string mode)
        {
            return
            [
                $"Flow_{mode}_TD_Mbist_BV"
            ];
        }

        private static BinningTable CreateInterpolationTable(string mode = "MS001", string domain = "PCPU")
        {
            return new BinningTable
            {
                ModeIdx = 0,
                IntModeLIdx = 1,
                IntModeHIdx = 2,
                DomainIdx = 3,
                IntSkipTestIdx = 4,
                CommentIdx = 5,
                Job = "FT1",
                Rows =
                [
                    new()
                    {
                        RowData =
                        [
                            mode,
                            "MS000",
                            "MS002",
                            domain,
                            "Yes",
                            "Max PV (MP00G/MP00H/MP00I/MP00J/MP00K)"
                        ]
                    }
                ]
            };
        }

        private static BinCutFlowTable CreateFlowTable(string performanceMode = "MS001", string binningDomain = "PCPU")
        {
            return new BinCutFlowTable
            {
                SheetName = "Flow_MS001",
                Rows =
                [
                    new BinCutFlowSheetRow(
                    "Flow_MS001",
                    new List<string>
                    {
                        performanceMode,
                        binningDomain
                    })
                    {
                        PerformanceMode = performanceMode,
                        BinningDomain = binningDomain,
                        TableType = EnumBinCutTableType.Lv
                    }
                ]
            };
        }
        private static BinCutFlowTable CreateMultiRailFlowTable(string performanceMode = "MS001", string binningDomain = "PCPU")
        {
            var row = new BinCutFlowSheetRow(
                "Flow_MS001",
                new List<string>
                {
            performanceMode,
            binningDomain
                })
            {
                PerformanceMode = performanceMode,
                BinningDomain = binningDomain,
                TableType = EnumBinCutTableType.Lv
            };

            row.PinInfos.Add(new TestPlanLib.BinCut.Flow.PinInfo
            {
                PinContext = "MS001 Evaluate Bin"
            });

            row.PinInfos.Add(new TestPlanLib.BinCut.Flow.PinInfo
            {
                PinContext = "MS002 Evaluate Bin"
            });

            return new BinCutFlowTable
            {
                SheetName = "Flow_MS001",
                Rows =
                [
                    row
                ]
            };
        }

        [TestMethod]
        public void GetCurrentJob_With_All()
        {
            string result = _writer.GetCurrentJob("All");
            Assert.AreEqual("CP1||FT1", result.ToUpper());
        }

        [TestMethod]
        public void GetCurrentJob_With_CP1()
        {
            string result = _writer.GetCurrentJob("CP1");
            Assert.AreEqual("CP1", result.ToUpper());
        }

        [TestMethod]
        public void GetCurrentJob_With_CP2()
        {
            string result = _writer.GetCurrentJob("CP2");
            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void GetCurrentJob_With_FT1()
        {
            string result = _writer.GetCurrentJob("WLFT1");
            Assert.AreEqual("WLFT1", result);
        }

        [TestMethod]
        public void GetCurrentJob_With_FT2()
        {
            string result = _writer.GetCurrentJob("WLFT2");
            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void HandleExtraBinTablesJobs_WithInterPolation_ShouldReturnFalse()
        {
            // Arrange
            IdsDistributionTable table = CreateIdsDistributionTable("inter");
            // Act
            (bool isExtrapolate, List<string> jobs) = _writer.HandleExtraBinTablesJobs(table, "MS001");
            // Assert
            Assert.IsFalse(isExtrapolate);
            Assert.AreEqual(0, jobs.Count);
        }

        [TestMethod]

        public void HandleExtraBinTablesJobs_WithExtraPolation_ShouldReturnTrue()
        {
            // Arrange
            IdsDistributionTable table = CreateIdsDistributionTable("extra");
            // Act
            (bool isExtrapolate, List<string> jobs) = _writer.HandleExtraBinTablesJobs(table, "MS001");
            // Assert
            Assert.IsTrue(isExtrapolate);
            Assert.AreEqual(1, jobs.Count);
            Assert.AreEqual("FT1", jobs[0]);
        }

        [TestMethod]
        public void HandleBinTablesJobs_IsInterpolation_ShouldReturnTrue()
        {
            // Arrange
            BinningTables tables = CreateBinningTables("MS001");
            List<BinCutFlowInstanceWriter.PendingInterpolate> pendingInterpolate = CreatePendingInterpolate();
            List<string> performanceList = CreatePerformanceList("MS001");
            _binCutInputManager.BinCutFlowTables = [CreateFlowTable()];
            // Act
            (bool isInterpolate, List<string> jobs) = _writer.HandleBinTablesJobs(0, tables, "MS001", "Sheet1", pendingInterpolate, performanceList);
            // Assert
            Assert.IsTrue(isInterpolate);
            Assert.AreEqual(1, jobs.Count);
            Assert.AreEqual(2, pendingInterpolate.Count);
        }

        [TestMethod]
        public void HandleBinTablesJobs_IsInterpolation_MultiRail_ShouldReturnTrue()
        {
            // Arrange
            BinningTables tables = CreateBinningTables("MS001");
            List<BinCutFlowInstanceWriter.PendingInterpolate> pendingInterpolate = CreatePendingInterpolate();
            List<string> performanceList = CreatePerformanceList("MS001");
            _binCutInputManager.BinCutFlowTables = [CreateMultiRailFlowTable()];
            // Act
            (bool isInterpolate, List<string> jobs) = _writer.HandleBinTablesJobs(0, tables, "MS001", "Sheet1", pendingInterpolate, performanceList);
            // Assert
            Assert.IsTrue(isInterpolate);
            Assert.AreEqual(1, jobs.Count);
            Assert.AreEqual(2, pendingInterpolate.Count);
        }

        [TestMethod]
        public void HandleBinTablesJobs_IsInterpolation_ForflowsFalse()
        {
            // Arrange
            BinningTables tables = CreateBinningTables("MS001");
            List<BinCutFlowInstanceWriter.PendingInterpolate> pendingInterpolate = CreatePendingInterpolate();
            List<string> performanceList = CreatePerformanceList("MA001");
            _binCutInputManager.BinCutFlowTables = [CreateFlowTable()];
            // Act
            (bool isInterpolate, List<string> jobs) = _writer.HandleBinTablesJobs(0, tables, "MS001", "Sheet1", pendingInterpolate, performanceList);
            // Assert
            Assert.IsTrue(isInterpolate);
            Assert.AreEqual(1, jobs.Count);
            Assert.AreEqual(2, pendingInterpolate.Count);
        }

        [TestMethod]
        public void HandleBinTablesJobs_IsNotInterpolation_ShouldReturnFalse()
        {
            // Arrange
            BinningTables tables = CreateBinningTables("MS002");
            List<BinCutFlowInstanceWriter.PendingInterpolate> pendingInterpolate = CreatePendingInterpolate();
            List<string> performanceList = CreatePerformanceList("MS002");
            _binCutInputManager.BinCutFlowTables = [CreateFlowTable()];
            // Act
            (bool isInterpolate, List<string> jobs) = _writer.HandleBinTablesJobs(0, tables, "MS001", "Sheet1", pendingInterpolate, performanceList);
            // Assert
            Assert.IsFalse(isInterpolate);
            Assert.AreEqual(0, jobs.Count);
            Assert.AreEqual(1, pendingInterpolate.Count);

        }

        [TestMethod]
        public void GenerateOtherInstanceRows_WithExtraPolation_ShouldGenerateEPLInstance()
        {
            // Arrange
            _writer.BinCutExtraPolationModes["MS001"]
                = new BinCutExtraPolationModeInfo
                {
                    IsExtraPolation = true
                };

            _binCutInputManager.BinningTables = [CreateInterpolationTable()];
            _binCutInputManager.BinCutFlowTables = [CreateFlowTable()];
            Dictionary<string, bool> flowNameList =
                new()
                {
                    { "Flow_MS001_TD_Mbist_BV", true }
                };
            // Act
            List<InstanceRow> result = _writer.GenerateOtherInstanceRows(flowNameList);
            // Assert
            Assert.IsTrue(result.Any(x => x.TestName.Contains("EPL_VDD_PCPU_MS001_BV")));
        }

        [TestMethod]
        public void GenerateOtherInstanceRows_WithNonExtraPolation_ShouldGenerateIPLInstance()
        {
            // Arrange
            _writer.BinCutExtraPolationModes["MS002"]
                = new BinCutExtraPolationModeInfo
                {
                    IsExtraPolation = true
                };

            _binCutInputManager.BinningTables = [CreateInterpolationTable()];
            _binCutInputManager.BinCutFlowTables = [CreateFlowTable()];
            Dictionary<string, bool> flowNameList =
                new()
                {
                    { "Flow_MS001_TD_Mbist_BV", true }
                };
            // Act
            List<InstanceRow> result = _writer.GenerateOtherInstanceRows(flowNameList);
            // Assert
            Assert.IsTrue(result.Any(x => x.TestName.Contains("IPL_VDD_PCPU_MS001_BV")));
        }

        [TestMethod]
        public void GetFlowSequence_WrapsListUnderAllKey()
        {
            // Arrange
            var sourceFlows = new List<string> { "Flow_A", "Flow_B" };

            // Act
            Dictionary<string, List<string>> result = _writer.GetFlowSequence(sourceFlows);

            // Assert
            Assert.AreEqual(1, result.Count);
            CollectionAssert.AreEqual(sourceFlows, result["All"]);
        }

        [TestMethod]
        public void GetHarvestFailFlag_CompoundSiteVarAndDeviceColumn_ExtractsFlagsInOrder()
        {
            // Arrange
            var rows = new List<BinCutFinalInstanceRow>
            {
                new() {
                    BinCutInstanceRow = new BinCutInstanceRow { SiteVar = "F_A && !F_B", EnableAndDevice = "X,!Device1@site" }
                }
            };

            // Act
            List<string> result = _writer.GetHarvestFailFlag(rows);

            // Assert
            CollectionAssert.AreEqual(new List<string> { "F_A", "F_B", "Device1" }, result);
        }

        [TestMethod]
        public void GetHarvestFailFlag_NoSiteVarOrDevice_ReturnsEmpty()
        {
            // Arrange
            var rows = new List<BinCutFinalInstanceRow> { new() };

            // Act
            List<string> result = _writer.GetHarvestFailFlag(rows);

            // Assert
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void SortHarvestFlags_NonHarvestCoreFlagsPassThroughUnchanged()
        {
            // Arrange
            var flags = new List<string> { "F_Other1", "F_Other2" };

            // Act
            List<string> result = _writer.SortHarvestFlags(flags);

            // Assert
            CollectionAssert.AreEqual(new List<string> { "F_Other1", "F_Other2" }, result);
        }

        [TestMethod]
        public void BubbleSortHarvestFlag_SortsNumericSuffixAscending()
        {
            // Arrange
            var flags = new List<string> { "F_HARV3", "F_HARV1", "F_HARV2" };

            // Act
            _writer.BubbleSortHarvestFlag(ref flags);

            // Assert
            CollectionAssert.AreEqual(new List<string> { "F_HARV1", "F_HARV2", "F_HARV3" }, flags);
        }

        [DataTestMethod]
        [DataRow("X,!Device1@site", "Device1", DisplayName = "NegatedDeviceAtSite")]
        [DataRow("Device2@SITE", "Device2", DisplayName = "CaseInsensitiveSiteSuffix")]
        [DataRow("X<#>,Y", "", DisplayName = "PlaceholderAndNoMatch")]
        [DataRow("", "", DisplayName = "EmptyInput")]
        public void GetDeviceNameFromEnableAndDeviceCol_ExtractsDeviceNameEndingInSite(string input, string expected)
        {
            // Act
            string result = _writer.GetDeviceNameFromEnableAndDeviceCol(input);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void ProcessJobsResult_SingleJob_ReturnsNegatedAndPositiveJob()
        {
            // Act
            (string, string) result = BinCutFlowInstanceWriter.ProcessJobsResult(["CP1"]);

            // Assert
            Assert.AreEqual("!CP1", result.Item1);
            Assert.AreEqual("CP1", result.Item2);
        }

        [TestMethod]
        public void ProcessJobsResult_MultipleJobs_ReturnsNegatedGroupAndJoinedJobs()
        {
            // Act
            (string, string) result = BinCutFlowInstanceWriter.ProcessJobsResult(["CP1", "FT1"]);

            // Assert
            Assert.AreEqual("!(CP1||FT1)", result.Item1);
            Assert.AreEqual("CP1||FT1", result.Item2);
        }

        [TestMethod]
        public void ProcessJobsResult_NoJobs_ReturnsEmptyStrings()
        {
            // Act
            (string, string) result = BinCutFlowInstanceWriter.ProcessJobsResult([]);

            // Assert
            Assert.AreEqual(string.Empty, result.Item1);
            Assert.AreEqual(string.Empty, result.Item2);
        }

        [TestMethod]
        public void GetFinalInterpoJobs_JobsMatchingGradeSearchJobs_ReturnsJoinedMatches()
        {
            // Act
            string result = _writer.GetFinalInterpoJobs(["CP1", "FT1", "FT2"]);

            // Assert
            Assert.AreEqual("CP1||FT1", result);
        }

        [TestMethod]
        public void GetFinalInterpoJobs_NoMatchingJobs_ReturnsEmpty()
        {
            // Act
            string result = _writer.GetFinalInterpoJobs(["FT2"]);

            // Assert
            Assert.AreEqual(string.Empty, result);
        }

        [TestMethod]
        public void GenHeaderFooterRows_StripsFlowPrefixAndAddsHeaderFooterPerEntry()
        {
            // Arrange
            var sourceRowDic = new Dictionary<string, List<BinCutSourceItem>> { { "Flow_ABC", new List<BinCutSourceItem>() } };

            // Act
            InstanceRows result = _writer.GenHeaderFooterRows(sourceRowDic);

            // Assert
            Assert.AreEqual(2, result.Count);
            Assert.AreEqual("ABC_Header_1", result[0].TestName);
            Assert.AreEqual("ABC_Footer_1", result[1].TestName);
        }

        [TestMethod]
        public void AddFlowItem_Static_T0TxParameter_SetsEnvAndAddsRow()
        {
            // Arrange
            var flow = new SubFlowSheet("Flow_T0TX_PreCall");

            // Act
            BinCutFlowInstanceWriter.AddFlowItem("EnableWd", "test", "Vddbinning_T0TX_Levels_Efuse_LV", ref flow, "FailBin", "ColA");

            // Assert
            FlowRow row = flow.Rows.Last();
            Assert.AreEqual("EnableWd", row.Enable);
            Assert.AreEqual("test", row.Opcode);
            Assert.AreEqual("Vddbinning_T0TX_Levels_Efuse_LV", row.Parameter);
            Assert.AreEqual("T0TX_Use", row.Env);
            Assert.AreEqual("FailBin", row.FailAction);
            Assert.AreEqual("ColA", row.ColumnA);
        }

        [TestMethod]
        public void AddFlowItem_Static_NonT0TxParameter_LeavesEnvAndFailActionEmpty()
        {
            // Arrange
            var flow = new SubFlowSheet("Flow_X");

            // Act
            BinCutFlowInstanceWriter.AddFlowItem("", "call", "Flow_Something", ref flow);

            // Assert
            FlowRow row = flow.Rows.Last();
            Assert.AreEqual(string.Empty, row.Env);
            Assert.AreEqual(string.Empty, row.FailAction);
        }

        [TestMethod]
        public void AddFlowItem_InstanceOverload_SetsAllFieldsDirectly()
        {
            // Arrange
            var flow = new SubFlowSheet("Flow_X");

            // Act
            _writer.AddFlowItem("EnableWd", "test", "Param1", "FailBin", ref flow);

            // Assert
            FlowRow row = flow.Rows.Last();
            Assert.AreEqual("EnableWd", row.Enable);
            Assert.AreEqual("test", row.Opcode);
            Assert.AreEqual("Param1", row.Parameter);
            Assert.AreEqual("FailBin", row.FailAction);
        }

        [TestMethod]
        public void AddInstanceItem_WithArgListAndArgs_SetsAllFields()
        {
            // Arrange
            var instanceRows = new List<InstanceRow>();

            // Act
            _writer.AddInstanceItem("TestA", "VBT", "VbtNameA", ref instanceRows, "ArgList1", ["a", "b"]);

            // Assert
            InstanceRow row = instanceRows.Single();
            Assert.AreEqual("TestA", row.TestName);
            Assert.AreEqual("VBT", row.VbtType);
            Assert.AreEqual("VbtNameA", row.VbtName);
            Assert.AreEqual("ArgList1", row.ArgList);
            CollectionAssert.AreEqual(new List<string> { "a", "b" }, row.Args);
        }

        [TestMethod]
        public void AddInstanceItem_NoArgListOrArgs_LeavesArgListEmpty()
        {
            // Arrange
            var instanceRows = new List<InstanceRow>();

            // Act
            _writer.AddInstanceItem("TestB", "VBT", "VbtNameB", ref instanceRows, "", null);

            // Assert
            InstanceRow row = instanceRows.Single();
            Assert.AreEqual(string.Empty, row.ArgList);
        }

        [TestMethod]
        public void GenPower_Binning_Calculation_ReturnsHardcodedVbtInstance()
        {
            // Act
            InstanceRow row = _writer.GenPower_Binning_Calculation();

            // Assert
            Assert.AreEqual("Power_Binning", row.TestName);
            Assert.AreEqual("VBT", row.VbtType);
            Assert.AreEqual("Power_Binning_Calculation", row.VbtName);
        }

        [TestMethod]
        public void GenCheck_IDS_ReturnsHardcodedVbtInstance()
        {
            // Act
            InstanceRow row = _writer.GenCheck_IDS();

            // Assert
            Assert.AreEqual("Judge_stored_IDS", row.TestName);
            Assert.AreEqual("VBT", row.VbtType);
            Assert.AreEqual("check_IDS", row.VbtName);
        }

        [TestMethod]
        public void GenPrintOutVddBinning_ReturnsHardcodedVbtInstance()
        {
            // Act
            InstanceRow row = _writer.GenPrintOutVddBinning();

            // Assert
            Assert.AreEqual("PrintOutVddBinning", row.TestName);
            Assert.AreEqual("VBT", row.VbtType);
            Assert.AreEqual("PrintOut_VDD_Bin", row.VbtName);
        }

        [TestMethod]
        public void GenSet_VBinResult_without_Test_ReturnsHardcodedVbtInstance()
        {
            // Act
            InstanceRow row = _writer.GenSet_VBinResult_without_Test([]);

            // Assert
            Assert.AreEqual("Set_VBinResult_without_Test", row.TestName);
            Assert.AreEqual("VBT", row.VbtType);
            Assert.AreEqual("Set_VBinResult_without_Test", row.VbtName);
        }

        [TestMethod]
        public void GenFuseBinnedProductVoltagesInstanceRow_ReturnsNull()
        {
            // Act
            InstanceRow? result = _writer.GenFuseBinnedProductVoltagesInstanceRow();

            // Assert
            Assert.IsNull(result);
        }

        [TestMethod]
        public void GenPrintConfigInstanceRow_UsesRealVbtFunction_SetsFlagArgument()
        {
            // Act
            InstanceRow row = _writer.GenPrintConfigInstanceRow(["F_A", "F_B"]);

            // Assert
            Assert.AreEqual("Print_BinCut_config", row.TestName);
            Assert.AreEqual("VBT", row.VbtType);
            Assert.AreEqual("F_A,F_B", row.GetArgument("str_flag_Group"));
        }
    }
}
