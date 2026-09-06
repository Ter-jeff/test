using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.GenerateIgxl.BinCut.Base;
using Automation.GenerateIgxl.HTOL;
using Automation.Reader.ConfigFile.NamingRule.Base;
using Automation.Static;
using Automation.Test.Static;

using CommonLib.Extension;

using FileDiffLib;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.VbtLib;

namespace Automation.Test.UT.Htol
{
    [TestClass]
    public class HtolWorkFlowMainTests : FunctionTestBase
    {
        private static List<Function> _functions = null!;

        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            _functions = TestSuiteInitialize.Functions;
        }

        private static List<BinCutInstanceRow> CreateSampleRows()
        {
            var rows = new List<BinCutInstanceRow>
            {
                new("HTOL_Sheet")
                {
                    RowNum = 1,
                    FlowNameOri = "HTOL_CPU_MAIN_FLOW",
                    Instance = "Flow_HTOL_01",
                    EnableAndDevice = "Enable_CPU",
                    SubFlow = "Sub1",
                    EnableFlow = "EnableFlow1",
                    JobTestStage = "CP1",
                    SiteVar = "Site1",
                    FailFlag = "F_HV,F_LV,F_NV,F_HV||F_LV,F_HV&&F_LV",
                    BinOutStage = "CP1",
                    DCcategory = "HTOL_X_X_X_Eqn",
                    TimeSet = "TS_HTOL:Param",
                    ShiftSpeed = "Normal",
                    PatternPinGroup = "PatternPinGroup(A)",
                    PinGroupBinoutFlag = "F_HV,F_LV,F_NV,F_HV||F_LV,F_HV&&F_LV",
                    PatternList =
                    [
                        "cz_brna0_c_fulp_an_aa00_dll_jtg_vix_allfrv_si_cpllds_t10pd",
                        "cz_brna0_c_fulp_an_aa00_dll_jtg_vix_allfrv_si_cpllds_t10",
                        "PP_BRNA0_L_INLP_SC_CFXX_SAA_COM_AUT_ALLFRV_DM_XORDSSC",
                        "PP_BRNA0_XORDSSCAA_X_X",
                        "dp_brna0_c_fulp_an_aa22_frq_jtg_dyc_allfrv_si_ciopll_t2p1r+dp_brna0_c_fulp_an_aa22_frq_jtg_dyc_allfrv_si_ciopll_t2p1r",
                        "PP_1_2_3_SC_X_TDF"
                    ],
                    InitList = ["INIT_1", "INIT_2"],
                    PayloadList = ["PP_1_2_3_SC_X_TDF"],
                    Type = BincutInstanceType.Pattern,
                    Char = "Char"
                },
                new("HTOL_Sheet")
                {
                    RowNum = 2,
                    FlowNameOri = "HTOL_CPU_MAIN_FLOW",
                    Instance = "Flow_HTOL_01",
                    EnableAndDevice = "Enable_CPU",
                    SubFlow = "Sub1",
                    EnableFlow = "EnableFlow1",
                    JobTestStage = "CP2",
                    SiteVar = "Site1",
                    FailFlag = "X",
                    BinOutStage = "CP2",
                    DCcategory = "HTOL_X_X_X_Eqn",
                    TimeSet = "TS_HTOL:Param",
                    ShiftSpeed = "Normal",
                    PatternPinGroup = "PatternPinGroup(A)",
                    PinGroupBinoutFlag = "F_HV,F_LV,F_NV,F_HV||F_LV,F_HV&&F_LV",
                    PatternList =
                    [
                        "cz_brna0_c_fulp_an_aa00_dll_jtg_vix_allfrv_si_cpllds_t10pd",
                        "cz_brna0_c_fulp_an_aa00_dll_jtg_vix_allfrv_si_cpllds_t10",
                        "PP_BRNA0_L_INLP_SC_CFXX_SAA_COM_AUT_ALLFRV_DM_XORDSSC",
                        "PP_BRNA0_XORDSSCAA_X_X",
                        "dp_brna0_c_fulp_an_aa22_frq_jtg_dyc_allfrv_si_ciopll_t2p1r+dp_brna0_c_fulp_an_aa22_frq_jtg_dyc_allfrv_si_ciopll_t2p1r",
                        "PP_1_2_3_SC_X_TDF",
                        "PP_1_2_3_SC_X_TDF_1"
                    ],
                    InitList = ["INIT_1", "INIT_2"],
                    PayloadList = ["PP_1_2_3_SC_X_TDF"],
                    Type = BincutInstanceType.Pattern,
                    Char = "Char"
                },
                new("HTOL_Sheet")
                {
                    RowNum = 3,
                    FlowNameOri = "MBIST_GPU_FLOW",
                    Instance = "Flow_MBIST_02",
                    EnableAndDevice = "Enable_GPU",
                    SubFlow = "Sub2",
                    EnableFlow = "EnableFlow2",
                    JobTestStage = "FT1",
                    SiteVar = "Site2",
                    FailFlag = "",
                    BinOutStage = "X",
                    DCcategory = "HTOL_X_X_X_Eqn",
                    TimeSet = "TS_MBIST",
                    ShiftSpeed = "Fast",
                    PatternPinGroup = "PatternPinGroup(B)",
                    PinGroupBinoutFlag = "F_A,F_A,F_B,F_A,",
                    PatternList =
                    [
                        "dp_brna0_c_fulp_an_aa22_frq_jtg_dyc_allfrv_si_ciopll_t2p1r"
                    ],
                    InitList = ["INIT_3"],
                    PayloadList = ["Payload_2"],
                    Type = BincutInstanceType.Pattern
                }
            };

            return rows;
        }

        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void HtolWorkFlowMainTest()
        {
            string subName = "HtolWorkFlowMain";
            string outputPath = Path.Combine(OutputPath, "Htol", subName);
            string expectPath = Path.Combine(ExpectPath, "Htol", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            LocalSpecs.TarFolder = outputPath;

            TestProgram.Clear();
            TestProgram.VbtFunctionLib.AddVbtFunctionRange([.. _functions.Where(x => x.Type.EqualsIgnoreCase(".NET"))]);
            var binCutFinalInstanceRows = new List<BinCutFinalInstanceRow>
            {
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
                        SiteVar = "F_SCAN"
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
                         SiteVar = "F_BIST"
                    },
                    PatternList = ["GFX_PAT_A"],
                    PayloadList = ["GFX_PAYLOAD_X"],
                    PatSetName = "EVS_GFX_PATSET_ROW2",
                    FinalInstName = "EVS_GFX_INST1",
                    IsUsed = true
                }
            };
            var instanceRows = new List<InstanceRow>();

            // Act
            ScanConfig config = SettingStatic.ScanConfig;
            var sheet = new BinCutInstanceSheet("SheetName")
            {
                Rows = CreateSampleRows()
            };
            var sheets = new List<BinCutInstanceSheet> { sheet };

            var main = new HtolInstanceMain(config, sheets);
            main.WorkFlow();
            List<InstanceRow> rows = main.GenInstances(binCutFinalInstanceRows, ref instanceRows);
            List<BinTableRow> binTableRows = main.GetBinTableRows(binCutFinalInstanceRows);

            // Assert
            var instanceSheet = new InstanceSheet("InstSheet");
            instanceSheet.Rows.AddRange(rows);
            instanceSheet.Write(Path.Combine(outputPath, instanceSheet.Name + ".txt"));
            var binTableSheet = new BinTableSheet("BinTable_Scan");
            binTableSheet.AddRows(binTableRows);
            binTableSheet.Write(Path.Combine(outputPath, binTableSheet.Name + ".txt"));

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
