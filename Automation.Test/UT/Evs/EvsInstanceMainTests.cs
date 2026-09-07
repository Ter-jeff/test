using System.Collections.Generic;
using System.IO;

using Automation.GenerateIgxl.BinCut.Base;
using Automation.GenerateIgxl.EVS;
using Automation.Reader.ConfigFile.NamingRule.Base;
using Automation.Static;

using CommonLib.Enums;

using FileDiffLib;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.VbtLib;

namespace Automation.Test.UT.Evs
{
    [TestClass]
    public class EvsInstanceMainTests : FunctionTestBase
    {
        private EvsInstanceMainCs _evsWorkFlow = null!;
        private List<BinCutInstanceSheet> _dummyEvsSheets = null!;

        [TestInitialize]
        public void Setup()
        {
            _dummyEvsSheets = [];

            var config = new ScanConfig();

            _evsWorkFlow = new EvsInstanceMainCs(config, _dummyEvsSheets);
        }

        [TestMethod]
        public void GenEvsNormalInstance_ShouldReturnInstanceRow_WithCorrectTestName()
        {
            // Arrange
            var row = new BinCutFinalInstanceRow
            {
                BinCutInstanceRow = new BinCutInstanceRow
                {
                    RowNum = 1,
                    SheetName = "Sheet1",
                    UserFunction = "Func:Param",
                    DCcategory = "DC_Test"
                },
                PatternList = ["X_X_X_ADSSCB_X", "Pattern1"],
                PatSetName = "PatSet1"
            };

            // Act
            InstanceRow result = _evsWorkFlow.GenEvsNormalInstance(row);

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.IsTrue(result.TestName.Contains("Pattern1") || result.TestName.Contains("PatSet1"));
            Assert.AreEqual(row.BinCutInstanceRow.RowNum, result.RowNum);
        }

        [TestMethod]
        public void GenPowerUp_EVS_ShouldReturnInstanceRow_WithVbtName()
        {
            // Act
            InstanceRow result = _evsWorkFlow.GenPowerUp_EVS();

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.AreEqual("PowerUp_EVS", result.TestName);
            Assert.AreEqual("PowerUp_Parallel", result.VbtName);
            Assert.AreEqual("VBT", result.VbtType);
        }

        [TestMethod]
        public void GenEvsCurrentProfileStartInstance_ShouldAppendSuffix()
        {
            // Arrange
            var instanceRow = new InstanceRow { ColumnA = "TestColumn" };
            string subFlowName = "Flow1";
            string pinList = "PIN1,PIN2";
            string suffix = "Suffix1";

            // Act
            InstanceRow result = _evsWorkFlow.GenEvsCurrentProfileStartInstance(instanceRow, subFlowName, pinList, suffix);

            // Assert
            Assert.AreEqual("Start_Profile_AutoResolution", result.VbtName);
            Assert.AreEqual("VBT", result.VbtType);
        }

        [TestMethod]
        public void GenEvsVtrigInstance_ShouldSetArgs()
        {
            // Arrange
            var instanceRow = new InstanceRow { ColumnA = "ColA", TestName = "Test1" };
            var vbtFunctionBase = new Function();

            // Act
            InstanceRow result = _evsWorkFlow.GenEvsVTrigInstance(instanceRow, vbtFunctionBase, null);

            // Assert
            Assert.IsTrue(result.TestName.EndsWith("_Vtrig"));
            Assert.IsTrue(result.ColumnA!.StartsWith("Vtrig for "));
            Assert.IsTrue(result.Args.Count != 0);
        }

        [TestMethod]
        public void GenPowerDown_EVS_ShouldCreateExpectedInstanceRow()
        {
            LocalSpecs.Options.InstrumentType = EnumInstrument.UFlex;
            // Act
            InstanceRow result = _evsWorkFlow.GenPowerDown_EVS();

            // Assert
            Assert.AreEqual("Power down instance for power reset flow", result.ColumnA);
            Assert.AreEqual("PowerDown_EVS", result.TestName);
            Assert.AreEqual("VBT", result.VbtType);
            Assert.AreEqual("nWire_X_X_X", result.DcCategory);
            Assert.AreEqual("Typ", result.DcSelector);
            Assert.AreEqual("Common", result.AcCategory);
            Assert.AreEqual("Typ", result.AcSelector);
            Assert.AreEqual("TimeSet_nWire", result.TimeSets);
            Assert.AreEqual("Levels_nWire", result.PinLevels);
            Assert.AreEqual("PowerDown_Parallel", result.VbtName);
        }

        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void GenEvsFlowTest_IndividualForLoop()
        {
            string subName = "GenEvsFlow_True";
            string outputPath = Path.Combine(OutputPath, "Evs", subName);
            string expectPath = Path.Combine(ExpectPath, "Evs", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            LocalSpecs.TarFolder = outputPath;
            TestProgram.Clear();
            List<BinCutFinalInstanceRow> binCutFinalInstanceRows = GetBinCutFinalInstanceRows();

            // Act
            ScanConfig config = SettingStatic.ScanConfig;
            List<BinCutInstanceSheet> evsInstanceSheets = TestPlanStatic.EvsInstanceSheets;
            var main = new EvsInstanceMain(config, evsInstanceSheets);
            main.WorkFlow();
            List<SubFlowSheet> sheets = main.GenEvsFlows(binCutFinalInstanceRows);
            List<InstanceRow> instanceRows = main.GenEvsInstances(binCutFinalInstanceRows);
            TestProgram.Print();

            // Assert
            foreach (SubFlowSheet sheet in sheets)
            {
                sheet.Write(Path.Combine(outputPath, sheet.Name + ".txt"));
            }

            var instanceSheet = new InstanceSheet("Inst_Evs");
            instanceSheet.Rows.AddRange(instanceRows);
            instanceSheet.Write(Path.Combine(outputPath, instanceSheet.Name + ".txt"));

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        private static List<BinCutFinalInstanceRow> GetBinCutFinalInstanceRows()
        {
            return
            [
                new()
                {
                    Domain = "CPU",
                    BinCutInstanceRow = new BinCutInstanceRow
                    {
                        FlowName = "Flow_EVS_CPU_LV",
                        Instance = "INST_CPU_1",
                        EnableFlow = "YES",
                        DCcategory = "CPU_LV",
                        BinOutStage = "A",
                        Burst = "YES",
                        RowNum = 1,
                        SheetName = "EVS_CPU_SHEET",
                        SubFlow = "SCAN",
                        SiteVar = "F_SCAN && F_SCAN_1",
                        PatternList = ["CPU_PAT_A", "CPU_PAT_B"]
                    },
                    PatternList = ["CPU_PAT_A", "CPU_PAT_B"],
                    PayloadList = ["CPU_PAYLOAD_X"],
                    InitList = ["CPU_INIT_X"],
                    PatSetName = "EVS_CPU_PATSET_ROW1",
                    FinalInstName = "EVS_CPU_INST1",
                    IsUsed = true
                },
                new()
                {
                    Domain = "CPU",
                    BinCutInstanceRow = new BinCutInstanceRow
                    {
                        FlowName = "Flow_EVS_CPU_LV",
                        Instance = "INST_CPU_1",
                        EnableFlow = "YES",
                        DCcategory = "CPU_LV",
                        BinOutStage = "A",
                        Burst = "YES",
                        RowNum = 1,
                        SheetName = "EVS_CPU_SHEET",
                        SubFlow = "SCAN",
                        SiteVar = "F_SCAN",
                        Evs = new EvsRowData { EvsParallelSetting = "A:B;C:D", EvsPwrPin1 = "EvsPwrPin1", EvsPwrPin2 = "EvsPwrPin2" },
                        PatternList = ["CPU_PAT_A", "CPU_PAT_B"]
                    },
                    PatternList = ["CPU_PAT_A", "CPU_PAT_B"],
                    PayloadList = ["CPU_PAYLOAD_X"],
                    InitList = ["CPU_INIT_X"],
                    PatSetName = "EVS_CPU_PATSET_ROW1",
                    FinalInstName = "EVS_CPU_INST1",
                    IsUsed = true
                },
                new()
                {
                    Domain = "CPU",
                    BinCutInstanceRow = new BinCutInstanceRow
                    {
                        FlowName = "Flow_EVS_CPU_LV",
                        Instance = "INST_CPU_1",
                        EnableFlow = "YES",
                        DCcategory = "CPU_LV",
                        BinOutStage = "A",
                        Burst = "YES",
                        RowNum = 1,
                        SheetName = "EVS_CPU_SHEET",
                        SubFlow = "SCAN",
                        SiteVar = "F_SCAN",
                        Evs = new EvsRowData { EvsCategory = "EvsCategory" },
                        PatternList = ["CPU_PAT_A", "CPU_PAT_B"]
                    },
                    PatternList = ["CPU_PAT_A", "CPU_PAT_B"],
                    PayloadList = ["CPU_PAYLOAD_X"],
                    InitList = ["CPU_INIT_X"],
                    PatSetName = "EVS_CPU_PATSET_ROW1",
                    FinalInstName = "EVS_CPU_INST1",
                    IsUsed = true
                },
                new()
                {
                    Domain = "CPU",
                    BinCutInstanceRow = new BinCutInstanceRow
                    {
                        FlowName = "Flow_EVS_CPU_LV",
                        Instance = "INST_CPU_1",
                        EnableFlow = "YES",
                        DCcategory = "CPU_NV",
                        BinOutStage = "A",
                        Burst = "YES",
                        RowNum = 1,
                        SheetName = "EVS_CPU_SHEET",
                        SubFlow = "SCAN",
                        SiteVar = "F_SCAN",
                        Evs = new EvsRowData { EvsCategory = "EvsCategory" },
                        PatternList = ["CPU_PAT_A", "CPU_PAT_B"]
                    },
                    PatternList = ["CPU_PAT_A", "CPU_PAT_B"],
                    PayloadList = ["CPU_PAYLOAD_X"],
                    InitList = ["CPU_INIT_X"],
                    PatSetName = "EVS_CPU_PATSET_ROW1",
                    FinalInstName = "EVS_CPU_INST1",
                    IsUsed = true
                },
                new()
                {
                    Domain = "CPU",
                    BinCutInstanceRow = new BinCutInstanceRow
                    {
                        FlowName = "Flow_EVS_CPU_LV",
                        Instance = "INST_CPU_1",
                        EnableFlow = "YES",
                        DCcategory = "CPU_HV",
                        BinOutStage = "A",
                        Burst = "YES",
                        RowNum = 1,
                        SheetName = "EVS_CPU_SHEET",
                        SubFlow = "SCAN",
                        SiteVar = "F_SCAN",
                        Evs = new EvsRowData { EvsCategory = "EvsCategory" },
                        PatternList = ["CPU_PAT_A", "CPU_PAT_B"]
                    },
                    PatternList = ["CPU_PAT_A", "CPU_PAT_B"],
                    PayloadList = ["CPU_PAYLOAD_X"],
                    InitList = ["CPU_INIT_X"],
                    PatSetName = "EVS_CPU_PATSET_ROW1",
                    FinalInstName = "EVS_CPU_INST1",
                    IsUsed = true
                },
                new()
                {
                    Domain = "GFX",
                    BinCutInstanceRow = new BinCutInstanceRow
                    {
                        FlowName = "Flow_EVS_GFX_HV",
                        Instance = "INST_GFX_1",
                        EnableFlow = "YES",
                        DCcategory = "GFX_HV",
                        BinOutStage = "B",
                        Burst = "NO",
                        RowNum = 2,
                        SheetName = "EVS_GFX_SHEET",
                        SubFlow = "BIST",
                        SiteVar = "F_BIST",
                        PatternList = ["GFX_PAT_A"]
                    },
                    PatternList = ["GFX_PAT_A"],
                    PayloadList = ["GFX_PAYLOAD_X"],
                    PatSetName = "EVS_GFX_PATSET_ROW2",
                    FinalInstName = "EVS_GFX_INST1",
                    IsUsed = true
                },
                new()
                {
                    Domain = "GFX",
                    BinCutInstanceRow = new BinCutInstanceRow
                    {
                        FlowName = "Flow_EVS_GFX_HV",
                        Instance = "INST_GFX_1",
                        EnableFlow = "YES",
                        DCcategory = "GFX_HV",
                        BinOutStage = "B",
                        Burst = "NO",
                        RowNum = 2,
                        SheetName = "EVS_GFX_SHEET",
                        SubFlow = "BIST",
                        SiteVar = "F_BIST",
                        Evs = new EvsRowData { EvsParallelSetting = "A:B;C:D", EvsPwrPin1 = "EvsPwrPin1", EvsPwrPin2 = "EvsPwrPin2" },
                        PatternList = ["GFX_PAT_A"]
                    },
                    PatternList = ["GFX_PAT_A"],
                    PayloadList = ["GFX_PAYLOAD_X"],
                    PatSetName = "EVS_GFX_PATSET_ROW2",
                    FinalInstName = "EVS_GFX_INST1",
                    IsUsed = true
                },
                new()
                {
                    Domain = "GFX",
                    BinCutInstanceRow = new BinCutInstanceRow
                    {
                        FlowName = "Flow_EVS_GFX_HV",
                        Instance = "INST_GFX_1",
                        EnableFlow = "YES",
                        DCcategory = "GFX_LV",
                        BinOutStage = "B",
                        Burst = "NO",
                        RowNum = 2,
                        SheetName = "EVS_GFX_SHEET",
                        SubFlow = "BIST",
                        SiteVar = "F_BIST",
                        Evs = new EvsRowData { EvsCategory = "EvsCategory" },
                        PatternList = ["GFX_PAT_A"]
                    },
                    PatternList = ["GFX_PAT_A"],
                    PayloadList = ["GFX_PAYLOAD_X"],
                    PatSetName = "EVS_GFX_PATSET_ROW2",
                    FinalInstName = "EVS_GFX_INST1",
                    IsUsed = true
                },
                new()
                {
                    Domain = "GFX",
                    BinCutInstanceRow = new BinCutInstanceRow
                    {
                        FlowName = "Flow_EVS_GFX_HV",
                        Instance = "INST_GFX_1",
                        EnableFlow = "YES",
                        DCcategory = "GFX_NV",
                        BinOutStage = "B",
                        Burst = "NO",
                        RowNum = 2,
                        SheetName = "EVS_GFX_SHEET",
                        SubFlow = "BIST",
                        SiteVar = "F_BIST",
                        Evs = new EvsRowData { EvsCategory = "EvsCategory" },
                        PatternList = ["GFX_PAT_A"]
                    },
                    PatternList = ["GFX_PAT_A"],
                    PayloadList = ["GFX_PAYLOAD_X"],
                    PatSetName = "EVS_GFX_PATSET_ROW2",
                    FinalInstName = "EVS_GFX_INST1",
                    IsUsed = true
                },
                new()
                {
                    Domain = "GFX",
                    BinCutInstanceRow = new BinCutInstanceRow
                    {
                        FlowName = "Flow_EVS_GFX_HV",
                        Instance = "INST_GFX_1",
                        EnableFlow = "YES",
                        DCcategory = "GFX_HV",
                        BinOutStage = "B",
                        Burst = "NO",
                        RowNum = 2,
                        SheetName = "EVS_GFX_SHEET",
                        SubFlow = "BIST",
                        SiteVar = "F_BIST",
                        Evs = new EvsRowData { EvsCategory = "EvsCategory" },
                        PatternList = ["GFX_PAT_A"]
                    },
                    PatternList = ["GFX_PAT_A"],
                    PayloadList = ["GFX_PAYLOAD_X"],
                    PatSetName = "EVS_GFX_PATSET_ROW2",
                    FinalInstName = "EVS_GFX_INST1",
                    IsUsed = true
                }
            ];
        }

        [TestMethod]
        public void GenEvsIvCurveTestRow_AppendsIvSuffixAndRemovesTightFailAction()
        {
            // Arrange
            var flowRow = new FlowRow
            {
                ColumnA = "Base",
                Parameter = "Test1",
                FailAction = "F_A, F_EVS_Tight"
            };

            // Act
            FlowRow result = _evsWorkFlow.GenEvsIvCurveTestRow(flowRow);

            // Assert
            Assert.AreEqual("IV curve for Base", result.ColumnA);
            Assert.AreEqual("Test1_IV", result.Parameter);
            Assert.AreEqual("F_A", result.FailAction);
            Assert.AreEqual("IVCurve", result.Enable);
        }

        [TestMethod]
        public void GenEvsVTrigTestRow_AppendsVtrigSuffix()
        {
            // Arrange
            var flowRow = new FlowRow
            {
                ColumnA = "Base",
                Parameter = "Test1"
            };

            // Act
            FlowRow result = _evsWorkFlow.GenEvsVTrigTestRow(flowRow);

            // Assert
            Assert.AreEqual("Vtrig for Base", result.ColumnA);
            Assert.AreEqual("Test1_Vtrig", result.Parameter);
            Assert.AreEqual("Vtrig", result.Enable);
        }

    }
}
