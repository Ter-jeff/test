using System.Collections.Generic;
using System.IO;

using Automation.GenerateIgxl.BinCut.Base;
using Automation.GenerateIgxl.EVS;
using Automation.Reader.ConfigFile.NamingRule.Base;
using Automation.Static;
using Automation.Test.Static;

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
    public class EvsInstanceMainCsTests : FunctionTestBase
    {
        private EvsInstanceMainCs _evsWorkFlow = null!;
        private List<BinCutInstanceSheet> _dummyEvsSheets = null!;

        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            TestProgram.VbtFunctionLib.AddVbtFunctionRange(TestSuiteInitialize.Functions);
        }

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
                    DCcategory = "DC_Test",
                    PatternList = ["X_X_DSSC_X", "PP_BRNA0_L_INLP_SC_CFXX_SAA_COM_AUT_ALLFRV_DM_XORDSSC", "Pattern1"]
                },
                PatternList = ["X_X_DSSC_X", "PP_BRNA0_L_INLP_SC_CFXX_SAA_COM_AUT_ALLFRV_DM_XORDSSC", "Pattern1"],
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
        public void GenEVS_Static_Power_RampTest()
        {
            // Arrange
            var row = new BinCutFinalInstanceRow
            {
                BinCutInstanceRow = new BinCutInstanceRow
                {
                    RowNum = 1,
                    SheetName = "Sheet1",
                    UserFunction = "Func:Param",
                    DCcategory = "DC_Test",
                    Evs = new EvsRowData { EvsPwrPin1 = "EvsPwrPin1" }
                },
                PatternList = ["X_X_DSSC_X", "PP_BRNA0_L_INLP_SC_CFXX_SAA_COM_AUT_ALLFRV_DM_XORDSSC", "Pattern1"],
                PatSetName = "PatSet1"
            };

            // Act
            (InstanceRow instanceRow, string _, Function _) = _evsWorkFlow.GenEVS_Static_Power_Ramp(row);

            // Assert
            Assert.AreEqual("EVS_Static_Power_Ramp_Multi", instanceRow.TestName);
        }

        [TestMethod]
        public void GenEVS_Static_Power_RampTest1()
        {
            // Arrange
            var row = new BinCutFinalInstanceRow
            {
                BinCutInstanceRow = new BinCutInstanceRow
                {
                    RowNum = 1,
                    SheetName = "Sheet1",
                    UserFunction = "Func:Param",
                    DCcategory = "DC_Test",
                },
                PatternList = ["X_X_DSSC_X", "PP_BRNA0_L_INLP_SC_CFXX_SAA_COM_AUT_ALLFRV_DM_XORDSSC", "Pattern1"],
                PatSetName = "PatSet1"
            };

            // Act
            (InstanceRow instanceRow, string _, Function _) = _evsWorkFlow.GenEVS_Static_Power_Ramp(row);

            // Assert
            Assert.AreEqual("EVS_Static_Power_Ramp_Multi", instanceRow.TestName);
        }

        [TestMethod]
        public void GenPowerUp_EVS_ShouldReturnInstanceRow_WithVbtName()
        {
            TestProgram.VbtFunctionLib.AddVbtFunctionRange(TestSuiteInitialize.Functions);

            // Act
            InstanceRow result = _evsWorkFlow.GenPowerUp_EVS();

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.AreEqual("PowerUp_EVS", result.TestName);
            Assert.AreEqual("IgxlWrapper.CoreTestLibrary.DC.DCTest.PowerUp", result.VbtName);
            Assert.AreEqual(".NET", result.VbtType);
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
            Assert.AreEqual("IgxlWrapper.CoreTestLibrary.CurrentProfile.CurrentVoltageProfileMain.ProfileAutoResolution", result.VbtName);
            Assert.AreEqual(".NET", result.VbtType);
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
            TestProgram.VbtFunctionLib.AddVbtFunctionRange(TestSuiteInitialize.Functions);
            // Act
            InstanceRow result = _evsWorkFlow.GenPowerDown_EVS();

            // Assert
            Assert.AreEqual("Power down instance for power reset flow", result.ColumnA);
            Assert.AreEqual("PowerDown_EVS", result.TestName);
            Assert.AreEqual(".NET", result.VbtType);
            Assert.AreEqual("nWire_X_X_X", result.DcCategory);
            Assert.AreEqual("Typ", result.DcSelector);
            Assert.AreEqual("Common", result.AcCategory);
            Assert.AreEqual("Typ", result.AcSelector);
            Assert.AreEqual("TimeSet_nWire", result.TimeSets);
            Assert.AreEqual("Levels_nWire", result.PinLevels);
            Assert.AreEqual("IgxlWrapper.CoreTestLibrary.DC.DCTest.PowerDown", result.VbtName);
        }

        [TestMethod]
        public void GenEvsFlowTestCs_IndividualForLoop()
        {
            string subName = "GenEvsFlowCs";
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
            var main = new EvsInstanceMainCs(config, evsInstanceSheets);
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
                        SiteVar = "F_SCAN",
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
                        Evs = new EvsRowData { EvsParallelSetting = "A:B;C:D" },
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
                        Evs = new EvsRowData { EvsParallelSetting = "A:B;C:D" },
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
    }
}
