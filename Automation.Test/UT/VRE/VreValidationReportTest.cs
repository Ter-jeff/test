using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.PostAction.VreValidationReport;
using Automation.Static;
using Automation.Test.Static;

using CommonLib.Extension;

using FileDiffLib;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

using TestPlanLib.Singleton;
using TestPlanLib.VbtLib;

namespace Automation.Test.UT.VRE
{
    [TestClass]
    public class VreReportTests : FunctionTestBase
    {
        private static List<Function> _functions = null!;
        private static HashSet<string> _bistHarvPattern = new HashSet<string>();

        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            _functions = TestSuiteInitialize.Functions;
        }

        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void VreValidationReportTest()
        {
            string subName = "VreValidationReport";
            string outputPath = Path.Combine(OutputPath, "Vre", subName);
            string expectPath = Path.Combine(ExpectPath, "Vre", subName);
            string binNumSheetPath = Path.Combine(InputPath, "borneo_documents", "A0_V04A", "TestPlan", "SOC_Xlsx", "BinNum.xlsx");
            HashSet<string> bistHarvPattern = new HashSet<string> { "" };
            HashSet<string> scanHarvPattern = new HashSet<string> { "PatA" };

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            LocalSpecs.TarFolder = outputPath;
            if (Directory.Exists(FolderStructure.DirIgLink))
            {
                Directory.Delete(FolderStructure.DirIgLink, true);
            }
            _ = Directory.CreateDirectory(FolderStructure.DirIgLink);
            FolderStructure.CreateFolder();
            TestProgram.VbtFunctionLib.AddVbtFunctionRange([.. _functions.Where(x => x.Type.EqualsIgnoreCase(".NET"))]);
            BlockStatus.GetAutomationBlockStatus(BlockConst.BinCut).Down = true;
            TestprogramInit();
            ExcelWorkbook binNumSheet = new ExcelPackage(new FileInfo(binNumSheetPath)).Workbook;

            BinNumberSingleton.Instance.Initialize(binNumSheet);
            new VreValidationReportMain(_bistHarvPattern, scanHarvPattern).WorkFlow();
            // Assert
            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
        private static void AddIntoPatSet(Dictionary<string, List<string>> patList, ref PatSetSheet patsetSheet, ref PatSetSheet patsetAllSheet)
        {
            foreach (KeyValuePair<string, List<string>> pattern in patList)
            {
                string patsetName = pattern.Key;
                var patset = new PatSet { PatSetName = patsetName };
                foreach (string pat in pattern.Value)
                {
                    patset.PatSetRows.Add(new PatSetRow() { PatternSet = patsetName, File = pat });
                    var patSetAllRow = new PatSet { PatSetName = pat };
                    patSetAllRow.PatSetRows.Add(new PatSetRow() { PatternSet = pat, File = $"{pat}:{pat}_V1.pat" });
                    patsetAllSheet.AddRow(patSetAllRow);
                }
                patsetSheet.AddRow(patset);
            }
        }
        private static InstanceRow CreateInstanceRow(string instanceName, bool isHarv, Function function)
        {
            var instanceRow = new InstanceRow
            {
                TestName = instanceName,
                VbtName = function.FullFunctionName,
                ArgList = function.Parameters,
                Args = function.ArgList
            };
            instanceRow.SetArgument("patterns", instanceName.Replace("Inst_", "PatSet_"));
            instanceRow.SetArgument("isHarvesting", (isHarv ? "TRUE" : "FALSE"));
            return instanceRow;
        }
        private static void TestprogramInit()
        {
            List<string> blockTypes = new List<string> { "Scan", "Mbist", "Harvest" };
            List<string> scanTypes = new List<string> { "Sa", "Td" };
            List<string> insertions = new List<string> { "CP1", "CP2" };
            PatSetSheet patsetSheet = new PatSetSheet("PatSets_All");
            var patsetSheetAtpg = new PatSetSheet("Patsets_Atpg");
            foreach (string insertion in insertions)
            {
                MainFlow mainFlow = new MainFlow($"MainFlow_{insertion}") { Rows = new List<FlowRow>(), JobNames = new List<string> { insertion } };
                foreach (string blockType in blockTypes)
                {
                    InstanceSheet instanceSheet = new InstanceSheet($"TestInst_{blockType}_{insertion}");
                    Function function;
                    instanceSheet.SourceInfo.Block = blockType;
                    mainFlow.AddRow(new FlowRow { Opcode = OpCode.Call, Parameter = $"Flow_{blockType}_{insertion}" });
                    SubFlowSheet subFlowSheet = new SubFlowSheet($"Flow_{blockType}_{insertion}", $"Instance_{blockType}_{insertion}") { Rows = new List<FlowRow>() };
                    if (blockType == "Scan")
                    {
                        foreach (string scanType in scanTypes)
                        {
                            foreach (bool isHarv in new List<bool> { true, false })
                            {
                                string harvStr = isHarv ? "Harv" : "NonHarv";
                                string instanceName = $"Inst_Soc{scanType}_{insertion}_{harvStr}";
                                function = TestProgram.VbtFunctionLib.GetFunctionByName("FuncTestMain", blockType, true);
                                subFlowSheet.Rows.Add(new FlowRow { Opcode = OpCode.Test, Parameter = instanceName });
                                instanceSheet.AddRow(CreateInstanceRow(instanceName, isHarv, function));
                                Dictionary<string, List<string>> patsetDic = new Dictionary<string, List<string>>();
                                patsetDic.Add(instanceName.Replace("Inst_", "PatSet_"), new List<string> { $"Pat_IN00_XX_{scanType}F_{insertion}_{harvStr}", $"Pat_PL00_XX_{scanType}F_{insertion}_{harvStr}" });
                                AddIntoPatSet(patsetDic, ref patsetSheetAtpg, ref patsetSheet);
                            }
                        }
                    }
                    else if (blockType == "Mbist")
                    {
                        foreach (bool isHarv in new List<bool> { true, false })
                        {
                            string harvStr = isHarv ? "Harv" : "NonHarv";
                            string instanceName = $"Inst_Soc{blockType}_{insertion}_{harvStr}";
                            subFlowSheet.Rows.Add(new FlowRow { Opcode = OpCode.Test, Parameter = instanceName });
                            function = TestProgram.VbtFunctionLib.GetFunctionByName("FuncTestMain", blockType, true);
                            instanceSheet.AddRow(CreateInstanceRow(instanceName, isHarv, function));
                            Dictionary<string, List<string>> patsetDic = new Dictionary<string, List<string>>();
                            patsetDic.Add(instanceName.Replace("Inst_", "PatSet_"), new List<string> { $"Pat_IN00_XX_{blockType}_{insertion}_{harvStr}", $"Pat_PL00_XX_{blockType}_{insertion}_{harvStr}" });
                            if (isHarv)
                            {
                                _bistHarvPattern.Add($"Pat_PL00_XX_{blockType}_{insertion}_{harvStr}");
                            }
                            AddIntoPatSet(patsetDic, ref patsetSheetAtpg, ref patsetSheet);
                        }
                    }
                    else if (blockType == "Harvest")
                    {
                        Function functionHarv = TestProgram.VbtFunctionLib.GetFunctionByName("Harvest_Summary", blockType, true);
                        string instanceName = $"Inst_Soc{blockType}_{insertion}";
                        subFlowSheet.Rows.Add(new FlowRow { Opcode = OpCode.Test, Parameter = instanceName });
                        instanceSheet.AddRow(CreateInstanceRow(instanceName, true, functionHarv));
                    }
                    TestProgram.IgxlWorkBk.InsSheets.Add(instanceSheet.Name, instanceSheet);
                    TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirNonBinCut, subFlowSheet);
                }
                TestProgram.IgxlWorkBk.AddMainFlowSheet(FolderStructure.DirMain, mainFlow);
            }
            TestProgram.IgxlWorkBk.PatSetSheets.Add("PatSets_All", patsetSheet);
            TestProgram.IgxlWorkBk.PatSetSheets.Add("Patsets_Atpg", patsetSheetAtpg);

        }
    }
}
