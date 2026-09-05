using System.Collections.Generic;
using System.IO;

using Automation.GenerateIgxl.BinCut.Base;
using Automation.GenerateIgxl.EVS;
using Automation.Reader.ConfigFile.NamingRule.Base;

using FileDiffLib;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.BinCut.BinCutInstance;

namespace Automation.Test.UT.Evs
{
    [TestClass]
    public class EvsMultiRampInstanceMainTests : FunctionTestBase
    {
        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void GenEvsInstances_ShouldGenerateRampInstances()
        {
            string subName = "EvsMultiRampMain";
            string outputPath = Path.Combine(OutputPath, "Evs", subName);
            string expectPath = Path.Combine(ExpectPath, "Evs", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            var config = new ScanConfig();
            var sheets = new List<BinCutInstanceSheet>();
            var evs = new EvsMultiRampInstanceMain(config, sheets);

            List<BinCutFinalInstanceRow> rows = GetBinCutFinalInstanceRows();

            // Act
            List<InstanceRow> instanceRows = evs.GenRampEvsInstances(rows);

            // Assert
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
                        PatternList = ["CPU_PAT_A", "CPU_PAT_B"],
                        Evs = new EvsRowData
                        {
                            EvsConditions = [
                                new("CP1", "A1")
                                {
                                    Voltage1 = "Voltage1:5",
                                    Voltage2 = "Voltage2:5"
                                }
                            ]
                        }
                    },
                    PatternList = ["CPU_PAT_A", "CPU_PAT_B"],
                    PayloadList = ["CPU_PAYLOAD_X"],
                    InitList = ["CPU_INIT_X"],
                    PatSetName = "EVS_CPU_PATSET_ROW1",
                    FinalInstName = "EVS_CPU_INST1",
                    IsUsed = true
                }
            ];
        }
    }
}
