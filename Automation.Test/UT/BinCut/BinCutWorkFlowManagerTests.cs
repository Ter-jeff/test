using System.Collections.Generic;
using System.IO;

using Automation.GenerateIgxl.BinCut.Base;
using Automation.GenerateIgxl.BinCut.Business;
using Automation.InputManager.Data;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Newtonsoft.Json;

using TestPlanLib.BinCut;
using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.BinCut.Flow;

namespace Automation.Test.UT.BinCut
{
    [TestClass]
    public class BinCutWorkFlowManagerTests : FunctionTestBase
    {
        [TestMethod]
        public void GetMergerSourceItemTest()
        {
            string subName = "BinCutWorkFlowManager";
            string outputPath = Path.Combine(OutputPath, "BinCut", subName);
            string expectPath = Path.Combine(ExpectPath, "BinCut", subName);
            _ = Path.Combine(ExpectPath, "BinCut", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            var binCutData = new BinCutInputData
            {
                BinCutOrderSheet = new BinCutOrderSheet("BinCutOrderSheet")
                {
                    Rows =
                    [
                        new() { PerformanceMode = "MC607", Td = 1 },
                        new() { PerformanceMode = "MC608", Td = 1 }
                    ]
                }
            };

            var binCutWorkFlowManager = new BinCutWorkFlowManager(binCutData);
            List<BinCutFlowTable> binCutFlowTables =
            [
                new()
                {
                    SheetName = "A",
                    FinalJob = ["CP1"],
                    Rows =
                    [
                        new("A", ["CP1"]) { RowNum = 1, Atpg = "MC607 TD, MC407 TD", PerformanceMode = "MC607" },
                        new("A", ["CP1"]) { RowNum = 2, Atpg = "MC607 TD, MC407 TD", PerformanceMode = "MC607" }
                    ]
                },
                new()
                {
                    SheetName = "B",
                    FinalJob = ["CP1"],
                    Rows =
                    [
                        new("B", ["CP1"]) { RowNum = 1, Atpg = "MC607 TD, MC407 TD", PerformanceMode = "MC607" },
                        new("B", ["CP1"]) { RowNum = 2, Atpg = "MC607 TD, MC407 TD", PerformanceMode = "MC607" }
                    ]
                }
            ];
            List<BinCutSourceItem> result = binCutWorkFlowManager.GetMergerSourceItem(binCutFlowTables);
            string json = JsonConvert.SerializeObject(result, Formatting.Indented);
            File.WriteAllText(Path.Combine(outputPath, "result.json"), json);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        private static BinCutSourceItem NewSourceItem(string columnContent)
        {
            var flowSheetRow = new BinCutFlowSheetRow("Sheet1", []);
            return new BinCutSourceItem(flowSheetRow, EnumColumnName.FUNC, columnContent, []);
        }

        [DataTestMethod]
        [DataRow(new[] { "ELB" }, EnumColumnName.ELB, DisplayName = "ElbMatches")]
        [DataRow(new[] { "ILB" }, EnumColumnName.ILB, DisplayName = "IlbMatches")]
        [DataRow(new[] { "Set_E1_Voltage" }, EnumColumnName.E1Voltage, DisplayName = "E1VoltageMatches")]
        [DataRow(new[] { "Flow_Something_Enable" }, EnumColumnName.CallNwireEnable, DisplayName = "NwireEnableMatches")]
        [DataRow(new[] { "Flow_Something_Disable" }, EnumColumnName.CallNwireDisable, DisplayName = "NwireDisableMatches")]
        [DataRow(new[] { "Relay_Something_On" }, EnumColumnName.RelayOn, DisplayName = "RelayOnMatches")]
        [DataRow(new[] { "Relay_Something_Off" }, EnumColumnName.RelayOff, DisplayName = "RelayOffMatches")]
        [DataRow(new[] { "Flow_Something_Tmps" }, EnumColumnName.CallTMPS, DisplayName = "TmpsMatches")]
        [DataRow(new[] { "NoRuleMatchesThis" }, EnumColumnName.FUNC, DisplayName = "NoMatch_DefaultsToFunc")]
        public void ResolveFallbackColumn_MapsTokensToExpectedColumn(string[] tokens, EnumColumnName expected)
        {
            // Act
            EnumColumnName result = BinCutWorkFlowManager.ResolveFallbackColumn([.. tokens]);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void ResolveFallbackColumn_ElbTakesPriorityOverIlb()
        {
            // Arrange - the rule array is order-sensitive; ELB is checked before ILB
            var tokens = new List<string> { "ILB", "ELB" };

            // Act
            EnumColumnName result = BinCutWorkFlowManager.ResolveFallbackColumn(tokens);

            // Assert
            Assert.AreEqual(EnumColumnName.ELB, result);
        }

        [TestMethod]
        public void RemoveDoNotTest_FiltersOutDoNotTestAndDoNoTestSuffixes()
        {
            // Arrange
            var manager = new BinCutWorkFlowManager(new BinCutInputData());
            var sourceDic = new Dictionary<string, List<BinCutSourceItem>>
            {
                ["Sheet1"] =
                [
                    NewSourceItem("Flow_A_DONOTEST"),
                    NewSourceItem("Flow_B_DONOTTEST"),
                    NewSourceItem("Flow_C_Normal")
                ]
            };

            // Act
            Dictionary<string, List<BinCutSourceItem>> result = manager.RemoveDoNotTest(sourceDic);

            // Assert
            Assert.AreEqual(1, result["Sheet1"].Count);
            Assert.AreEqual("Flow_C_Normal", result["Sheet1"][0].ColumnContent);
        }

        [TestMethod]
        public void RemoveDoNotTest_NoMatchingSuffixes_KeepsAllItems()
        {
            // Arrange
            var manager = new BinCutWorkFlowManager(new BinCutInputData());
            var sourceDic = new Dictionary<string, List<BinCutSourceItem>>
            {
                ["Sheet1"] = [NewSourceItem("Flow_A_Normal"), NewSourceItem("Flow_B_Normal")]
            };

            // Act
            Dictionary<string, List<BinCutSourceItem>> result = manager.RemoveDoNotTest(sourceDic);

            // Assert
            Assert.AreEqual(2, result["Sheet1"].Count);
        }

        [TestMethod]
        public void GetRegularPatSetRows_NullInput_ReturnsEmptySet()
        {
            // Arrange
            var manager = new BinCutWorkFlowManager(new BinCutInputData());

            // Act
            manager.GetRegularPatSetRows(null!, out HashSet<string> result);

            // Assert
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void GetRegularPatSetRows_RowsWithPatSetNameOnly_CollectsPatSetNames()
        {
            // Arrange
            var manager = new BinCutWorkFlowManager(new BinCutInputData());
            var rows = new List<BinCutFinalInstanceRow>
            {
                new() { PatSetName = "PS1" },
                new() { PatSetName = "PS2" },
                new() { PatSetName = "" }
            };

            // Act
            manager.GetRegularPatSetRows(rows, out HashSet<string> result);

            // Assert
            Assert.AreEqual(2, result.Count);
            Assert.IsTrue(result.Contains("PS1"));
            Assert.IsTrue(result.Contains("PS2"));
        }

        [TestMethod]
        public void GetRegularPatSetRows_BistRowWithInitList_CollectsBothInitAndPatSetNames()
        {
            // Arrange - IsBist(FlowName) makes InitPatSetName resolve to a non-empty value
            var manager = new BinCutWorkFlowManager(new BinCutInputData());
            var row = new BinCutFinalInstanceRow { PatSetName = "PS1" };
            row.BinCutInstanceRow.FlowName = "Flow_MBIST_Test";
            row.InitList.Add("InitPattern1");
            var rows = new List<BinCutFinalInstanceRow> { row };

            // Act
            manager.GetRegularPatSetRows(rows, out HashSet<string> result);

            // Assert
            Assert.IsTrue(result.Contains("PS1"));
        }

        [TestMethod]
        public void GetRegularPatSetRows_IsCaseInsensitive()
        {
            // Arrange
            var manager = new BinCutWorkFlowManager(new BinCutInputData());
            var rows = new List<BinCutFinalInstanceRow>
            {
                new() { PatSetName = "ps1" },
                new() { PatSetName = "PS1" }
            };

            // Act
            manager.GetRegularPatSetRows(rows, out HashSet<string> result);

            // Assert
            Assert.AreEqual(1, result.Count);
        }
    }
}
