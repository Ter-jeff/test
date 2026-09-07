using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.GenerateIgxl.Scan.Harvest;
using Automation.Static;

using FileDiffLib;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Newtonsoft.Json;

using TestPlanLib.Harvest;
using TestPlanLib.Singleton;

namespace Automation.Test.UT.Scan
{
    [TestClass]
    public class ScanHarvestMainTests : FunctionTestBase
    {
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            TestPlanStatic.MainFlowSheet.CoreOverFailingFlows = new Dictionary<string, string> { { "KEY", "Block" } };
        }

        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestPlanStatic.MainFlowSheet.CoreOverFailingFlows = [];
        }

        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void GenHarvesting()
        {
            string subName = "GenHarvesting";
            string outputPath = Path.Combine(OutputPath, "BinCut", subName);
            string expectPath = Path.Combine(ExpectPath, "BinCut", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            LocalSpecs.TarFolder = outputPath;
            HarvestingTruthTableSheet harvestingTruthTableSheet = new HarvestingTruthTableSheet("")
            {
                MergeValueByFlag =
                {
                    { "F_A", new List<string> { "V1" } },
                    { "F_B", new List<string> { "V2" } },
                    { "F_C", new List<string> { "V3" } },
                    { "F_D", new List<string> { "V4" } },
                    { "Sum_A", new List<string> { "Sum_A" } }
                },
                Rows =
                [
                    new()
                    {
                        OutputBinTables = new Dictionary<string, string> { { "Key", "1" } },
                        Condition = "Condition_1",
                        ProposedBinName = "A_B_C_D"
                    }
                ]
            };
            BinNumberSingleton.Instance.Initialize(EpWorkbook.TestPlanWorkbook);
            List<SubFlowSheet> subFlowSheets = [];
            List<InstanceRow> instanceRows = [];

            // Act
            new ScanHarvestMain(harvestingTruthTableSheet).GenHarvesting(subFlowSheets, instanceRows);

            // Assert
            string json = JsonConvert.SerializeObject(subFlowSheets, Formatting.Indented);
            File.WriteAllText(Path.Combine(outputPath, "subFlowSheets.json"), json);
            string json2 = JsonConvert.SerializeObject(subFlowSheets, Formatting.Indented);
            File.WriteAllText(Path.Combine(outputPath, "instanceRows.json"), json2);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void GenBinTableRowsProposedBinNameTest()
        {
            string subName = "HarvestBinTableMain_ProposedBinName";
            string outputPath = Path.Combine(OutputPath, "Scan", subName);
            string expectPath = Path.Combine(ExpectPath, "Scan", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            LocalSpecs.TarFolder = outputPath;
            HarvestingTruthTableSheet harvestingTruthTableSheet = new HarvestingTruthTableSheet("")
            {
                MergeValueByFlag =
                {
                    { "F_A", new List<string> { "V1" } },
                    { "F_B", new List<string> { "V2" } },
                    { "F_C", new List<string> { "V3" } },
                    { "F_D", new List<string> { "V4" } },
                    { "Sum_A", new List<string> { "Sum_A" } }
                },
                Rows =
                [
                    new()
                    {
                        OutputBinTables = new Dictionary<string, string> { { "Key", "1" } },
                        Condition = "Condition_1",
                        ProposedBinName = "A_B_C_D"
                    }
                ]
            };
            BinNumberSingleton.Instance.Initialize(EpWorkbook.TestPlanWorkbook);

            // Act
            List<BinTableRow> binTableRows = new ScanHarvestMain(harvestingTruthTableSheet).GenBinTableRows();

            // Assert
            var binTableSheet = new BinTableSheet("BinTable_Scan");
            binTableSheet.AddRows(binTableRows);
            binTableSheet.Write(Path.Combine(outputPath, binTableSheet.Name + ".txt"));

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void GenBinTableRowsTrueTest()
        {
            string subName = "HarvestBinTableMainTrue";
            string outputPath = Path.Combine(OutputPath, "Scan", subName);
            string expectPath = Path.Combine(ExpectPath, "Scan", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            LocalSpecs.TarFolder = outputPath;
            HarvestingTruthTableSheet harvestingTruthTableSheet = new HarvestingTruthTableSheet("")
            {
                MergeValueByFlag =
                {
                    { "F_A", new List<string> { "V1" } },
                    { "F_B", new List<string> { "V2" } },
                    { "F_C", new List<string> { "V3" } },
                    { "F_D", new List<string> { "V4" } },
                    { "Sum_A", new List<string> { "Sum_A" } }
                },
                Rows =
                [
                    new()
                    {
                        OutputBinTables = new Dictionary<string, string> { { "Key", "1" } },
                        Condition = "Condition_1",
                        Flags =
                        [
                            new()
                            {
                                FlagName = "F_A",
                                Value = "1",
                                IsSum = true,
                                SumFlags = ["F_A"],
                                AllFlags = ["All_A"]
                            },
                            new()
                            {
                                FlagName = "F_B",
                                Value = "0",
                                IsSum = true,
                                SumFlags = ["F_B"],
                                AllFlags = ["All_A"]
                            },
                            new()
                            {
                                FlagName = "F_C",
                                Value = "1",
                                AllFlags = ["All_A"]
                            },
                            new()
                            {
                                FlagName = "F_D",
                                Value = "0",
                                AllFlags = ["All_A"]
                            },
                            new()
                            {
                                FlagName = "Sum_A",
                                Value = "1",
                                IsSum = true,
                                SumFlags = ["F_Sum_A"],
                                AllFlags = ["F_C", "F_D"]
                            }
                        ]
                    }
                ]
            };
            BinNumberSingleton.Instance.Initialize(EpWorkbook.TestPlanWorkbook);

            // Act
            List<BinTableRow> binTableRows = new ScanHarvestMain(harvestingTruthTableSheet).GenBinTableRows();

            // Assert
            var binTableSheet = new BinTableSheet("BinTable_Scan");
            binTableSheet.AddRows(binTableRows);
            binTableSheet.Write(Path.Combine(outputPath, binTableSheet.Name + ".txt"));

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        [DataTestMethod]
        [DataRow("[E]_[F]_{Sum(D)}", "01_WI_BinName", DisplayName = "01_WI_BinName")]
        [DataRow("", "02_WO_BinName", DisplayName = "02_WO_BinName")]
        [TestCategory("ExcludeFromMutationTest")]
        public void GenBinTableRowsFalseTest(string proposedBinName, string name)
        {
            string subName = "HarvestBinTableMain" + name;
            string outputPath = Path.Combine(OutputPath, "Scan", subName);
            string expectPath = Path.Combine(ExpectPath, "Scan", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            LocalSpecs.TarFolder = outputPath;
            HarvestingTruthTableSheet harvestingTruthTableSheet = new HarvestingTruthTableSheet("")
            {
                MergeValueByFlag =
                {
                    { "F_A", new List<string> { "V1" } },
                    { "F_B", new List<string> { "V2" } },
                    { "F_C", new List<string> { "V3" } },
                    { "F_D", new List<string> { "V4" } },
                    { "Sum_A", new List<string> { "Sum_A" } }
                },
                Rows =
                [
                    new()
                    {
                        OutputBinTables = new Dictionary<string, string> { { "Key", "1" } },
                        Condition = "Condition_1",
                        ProposedBinName = string.IsNullOrEmpty(proposedBinName) ? "" : proposedBinName,
                        Flags =
                        [
                            new()
                            {
                                Name = "A",
                                FlagName = "F_A",
                                Value = "1",
                                IsSum = true,
                                SumFlags = ["F_A"],
                                AllFlags = ["All_A"]
                            },
                            new()
                            {
                                Name = "B",
                                FlagName = "F_B",
                                Value = "0",
                                IsSum = true,
                                SumFlags = ["F_B"],
                                AllFlags = ["All_A"]
                            },
                            new()
                            {
                                Name = "C",
                                FlagName = "F_C",
                                Value = "1",
                                AllFlags = ["All_A"]
                            },
                            new()
                            {
                                Name = "Sum(D)",
                                FlagName = "F_D",
                                Value = "0",
                                AllFlags = ["All_A"]
                            },
                            new()
                            {
                                Name = "Sum",
                                FlagName = "Sum_A",
                                Value = "1",
                                IsSum = true,
                                SumFlags = ["F_Sum_A"],
                                AllFlags = ["F_C", "F_D"]
                            }
                        ]
                    }
                ]
            };
            BinNumberSingleton.Instance.Initialize(EpWorkbook.TestPlanWorkbook);

            // Act
            List<BinTableRow> binTableRows = new ScanHarvestMain(harvestingTruthTableSheet).GenBinTableRows();

            // Assert
            var binTableSheet = new BinTableSheet("BinTable_Scan");
            binTableSheet.AddRows(binTableRows);
            binTableSheet.Write(Path.Combine(outputPath, binTableSheet.Name + ".txt"));

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void GenHarvestPostCheckBinTable_ReturnsPopulatedRow()
        {
            // Act
            BinTableRow result = new ScanHarvestMain(null).GenHarvestPostCheckBinTable();

            // Assert
            Assert.AreNotEqual(null, result, "The returned BinTableRow should not be null.");
            Assert.AreEqual("Bin_Harvest_PostCheck_Fail", result.Name);
            Assert.AreEqual("AND", result.Op);

            // Verify ItemList and Items
            Assert.AreEqual("T", result.Items.First());

            Assert.IsFalse(string.IsNullOrEmpty(result.Sort), "Sort should be populated.");
            Assert.IsFalse(string.IsNullOrEmpty(result.Bin), "Bin should be populated.");
            Assert.IsFalse(string.IsNullOrEmpty(result.Result), "Result status should be assigned.");
        }

        [TestMethod]
        public void GenHarvPostCheckSubFlow_FiltersCorrectly_AndAddsMandatoryRows()
        {
            var instanceRows = new List<InstanceRow>
            {
                new() { VbtName = "Check_From_Fuse_To_Flag", TestName = "TestGroup_JobA" },
                new() { VbtName = "Other.Vbt", TestName = "TestGroup_JobB" }
            };

            // Act
            SubFlowSheet result = new ScanHarvestMain(null).GenHarvPostCheckSubFlow(instanceRows);

            // Assert: 1. Verify Sheet Name
            Assert.AreEqual("Flow_HarvestPostCheck", result.Name);

            // Assert: 2. Verify Filtering (Only JobA should be added as a Test opcode)
            var testRows = result.Rows.Where(r => r.Opcode == OpCode.Test).ToList();
            Assert.AreEqual(1, testRows.Count, "Only rows matching CheckFro should be added as Test rows.");
            Assert.AreEqual("JobA", testRows[0].Job);
            Assert.AreEqual("TestGroup_JobA", testRows[0].Parameter);

            // Assert: 3. Verify JobNames list (distinct jobs)
            Assert.IsTrue(result.JobNames.Contains("JobA"));
            Assert.IsFalse(result.JobNames.Contains("JobB"));

            // Assert: 4. Verify Mandatory Final Rows
            // Row count should be: 1 (Test) + 1 (BinTable) + 1 (Return) = 3
            Assert.AreEqual(3, result.Rows.Count);

            FlowRow binTableRow = result.Rows[^2];
            Assert.AreEqual(OpCode.BinTable, binTableRow.Opcode);
            Assert.AreEqual("Bin_Harvest_PostCheck_Fail", binTableRow.Parameter);

            FlowRow returnRow = result.Rows.Last();
            // Assuming your AddReturnRow adds a specific opcode/type
            Assert.AreNotEqual(null, returnRow);
        }
    }
}
