using System.Collections.Generic;
using System.IO;

using Automation.GenerateIgxl.BinCut.Base;
using Automation.GenerateIgxl.BinCut.Business;
using Automation.InputManager.Data;
using Automation.Static;

using CommonLib.Enums;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Newtonsoft.Json;

using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.Static;

using BinCutFlowSheetRow = TestPlanLib.BinCut.Flow.BinCutFlowSheetRow;
using EnumColumnName = TestPlanLib.BinCut.BinCutInstance.EnumColumnName;
using NewBinCutFlowSheetRow = TestPlanLib.BinCut.FlowNew.NewBinCutFlowSheetRow;

namespace Automation.Test.UT.BinCut
{
    [TestClass]
    public class BinCutInstanceLvGeneratorTests : FunctionTestBase
    {
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            LocalSpecs.BasLibraryFolder = "";
            LocalSpecs.CsLibraryFolder = "";
        }

        [TestMethod]
        public void BinCutInstanceTest()
        {
            string subName = "BinCutInstance";
            string outputPath = Path.Combine(OutputPath, "BinCut", subName);
            string expectPath = Path.Combine(ExpectPath, "BinCut", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            LocalSpecs.ScghFileName = Path.Combine(InputPath, "borneo_documents", "A0_V04A", "borneo_A0_SCGH_X_X_X#208.xlsx");
            NeededSheets.InitSheetName(EnumDevice.AP, LocalSpecs.SettingFolder, LocalSpecs.CurrentProject);

            var binCutFlowSheetRow = new BinCutFlowSheetRow("sheetName", []);
            var newBinCutFlowSheetRow = new NewBinCutFlowSheetRow("sheetName", "job");
            var sources = new List<BinCutSourceItem>
            {
                new(binCutFlowSheetRow, newBinCutFlowSheetRow, EnumColumnName.FUNC, "TEMP_TEST")
                {
                    PerformanceMode = "TMPS_MODE"
                },
                new(binCutFlowSheetRow, newBinCutFlowSheetRow, EnumColumnName.FUNC, "TEMP SENSOR MONITOR"),
                new(binCutFlowSheetRow, newBinCutFlowSheetRow, EnumColumnName.FUNC, "ILB_FLOW"),
                new(binCutFlowSheetRow, newBinCutFlowSheetRow, EnumColumnName.FUNC, "ELB_FLOW"),
                new(binCutFlowSheetRow, newBinCutFlowSheetRow, EnumColumnName.TD, "TD_FLOW"),
                new(binCutFlowSheetRow, newBinCutFlowSheetRow, EnumColumnName.Mbist, "MBIST_FLOW"),
                new(binCutFlowSheetRow, newBinCutFlowSheetRow, EnumColumnName.FUNC, "FUNC_DDR_FLOW")
                {
                    PerformanceMode = "DDR"
                },
                new(binCutFlowSheetRow, newBinCutFlowSheetRow, EnumColumnName.CallNwireDisable, "UNKNOWN_FLOW")
            };

            string testName = "PP_KMDA0_P_FULP_FU_S001_PFF_RTS_FUN_ALLFV_SI_BBQOFF_fff221204";
            var binCutFinalInstanceRows = new List<BinCutFinalInstanceRow>
            {
                new()
                {
                    BinCutInstanceRow =
                    {
                        FlowName = "ILB_Flow",
                        Type = BincutInstanceType.Hardip
                    },
                    PatternList = { testName }
                },
                new()
                {
                    BinCutInstanceRow =
                    {
                        FlowName = "ELB_Flow",
                        Type = BincutInstanceType.Hardip
                    },
                    PatternList = { testName }
                },
                new()
                {
                    BinCutInstanceRow =
                    {
                        FlowName = "POST_Default",
                        Type = BincutInstanceType.Pattern
                    },
                    PatternList = { testName }
                },
                new()
                {
                    BinCutInstanceRow =
                    {
                        FlowName = "TMPS_Flow",
                        Type = BincutInstanceType.Pattern
                    },
                    PatternList = { testName }
                },
                new()
                {
                    BinCutInstanceRow =
                    {
                        FlowName = "TEMP_SENSOR",
                        Type = BincutInstanceType.Pattern
                    },
                    PatternList = { testName }
                },
                new()
                {
                    BinCutInstanceRow =
                    {
                        FlowName = "ILB",
                        Type = BincutInstanceType.Hardip
                    },
                    PatternList = { testName }
                },
                new()
                {
                    BinCutInstanceRow =
                    {
                        FlowName = "ELB",
                        Type = BincutInstanceType.Hardip
                    },
                    PatternList = { testName }
                },
                new()
                {
                    BinCutInstanceRow =
                    {
                        FlowName = "TD_Flow",
                        Type = BincutInstanceType.Pattern
                    },
                    PatternList = { testName }
                },
                new()
                {
                    BinCutInstanceRow =
                    {
                        FlowName = "MBIST_Flow",
                        Type = BincutInstanceType.Pattern
                    },
                    PatternList = { testName }
                },
                new()
                {
                    BinCutInstanceRow =
                    {
                        FlowName = "FUNC_DDR_Flow",
                        Type = BincutInstanceType.Pattern
                    },
                    PatternList = { testName }
                },
                new()
                {
                    BinCutInstanceRow =
                    {
                        FlowName = "UNKNOWN_Flow",
                        Type = BincutInstanceType.Pattern
                    },
                    PatternList = { testName }
                }
            };
            var binCutInputData = new BinCutInputData();

            var binCutInstanceLvGenerator = new BinCutInstanceLvGenerator(sources, binCutInputData, binCutFinalInstanceRows);
            binCutInstanceLvGenerator.GenInstanceRows(false, binCutInputData);

            string json = JsonConvert.SerializeObject(binCutInstanceLvGenerator.InstanceJobsMapping, Formatting.Indented);
            File.WriteAllText(Path.Combine(outputPath, "result.json"), json);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
