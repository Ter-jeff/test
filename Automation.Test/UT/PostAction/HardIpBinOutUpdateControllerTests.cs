using System.Collections.Generic;
using System.Linq;

using Automation.Utility.TpUpdate.HardIPBinoutTPUpdate;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.PostAction
{

    [TestClass]
    public class HardIpBinOutUpdateControllerTests : FunctionTestBase
    {
        [TestMethod]
        [DataRow("TestParam", "BinOut", "HeaderOnly", "HeaderOnly\tBinOut", true, DisplayName = "01_AddFailAction_WhenSingleColumn")]
        [DataRow("TestParam2", "Retry", "Header\tOldAction", "Header\tRetry", true, DisplayName = "02_ReplaceFailAction_WhenTwoColumns")]
        [DataRow("TestParam3", "", "Header\tOldAction", "Header\tOldAction", false, DisplayName = "03_NoChange_WhenFailActionEmpty")]
        public void UpdateNoBinOutInfo_ShouldUpdateCorrectly(string parameter, string failAction, string comment1, string expectedComment1, bool expectedReturn)
        {
            // Arrange
            var target = new HardIpBinOutUpdateController(true);
            var row = new FlowRow
            {
                Parameter = parameter,
                FailAction = failAction,
                Comment1 = comment1
            };

            // Act
            bool result = target.UpdateNoBinOutInfo(row);

            // Assert
            Assert.AreEqual(expectedReturn, result, "Return value mismatch.");
            Assert.AreEqual(expectedComment1, row.Comment1, "Comment1 mismatch.");
            if (expectedReturn)
            {
                Assert.AreEqual("", row.FailAction, "FailAction should be cleared.");
            }
        }

        [TestMethod]
        public void UpdateItemToProgram_ShouldUpdateHiLim_WhenValuesDiffer()
        {
            // Arrange
            var flowItem = new FlowRow { HiLim = "10.0", SheetName = "TestSheet" };
            var binoutItem = new BinOutStatusHipListRow("sourceSheetName");
            string jobpart = "PartA";
            binoutItem.UpdatedHiLimitDic[jobpart] = "15.0";
            // Keep empty to avoid LoLim logic
            binoutItem.UpdatedLoLimitDic[jobpart] = "";
            binoutItem.BinOutEnableDictionary[jobpart] = "binout";
            List<FlowRow> flowAllTestItems =
            [
                new() { HiLim = "10.0", SheetName = "TestSheet", FailAction = "FailAction" },
                new() { HiLim = "10.0", SheetName = "TestSheet", Comment1 = "Comment1 Comment2" }
            ];

            // Act
            var service = new HardIpBinOutUpdateController(true);
            bool result = service.UpdateItemToProgram(ref flowItem, binoutItem, jobpart, flowAllTestItems, "Sheet1");

            // Assert
            Assert.IsTrue(result, "Should return true because HiLim was updated.");
            Assert.AreEqual("15.0", flowItem.HiLim);
        }

        [TestMethod]
        public void UpdateItemToProgram_ShouldUpdateHiLim_WhenValuesDiffer_1()
        {
            // Arrange
            var flowItem = new FlowRow { HiLim = "10.0V", SheetName = "TestSheet", FailAction = "FailAction" };
            var binoutItem = new BinOutStatusHipListRow("sourceSheetName");
            string jobpart = "PartA";
            binoutItem.UpdatedHiLimitDic[jobpart] = "15.0";
            // Keep empty to avoid LoLim logic
            binoutItem.UpdatedLoLimitDic[jobpart] = "";
            binoutItem.BinOutEnableDictionary[jobpart] = "nonbinout";

            // Act
            var service = new HardIpBinOutUpdateController(true);
            bool result = service.UpdateItemToProgram(ref flowItem, binoutItem, jobpart, [], "Sheet1");

            // Assert
            Assert.IsTrue(result, "Should return true because HiLim was updated.");
            Assert.AreEqual("15.0", flowItem.HiLim);
        }

        [TestMethod]
        public void UpdateItemToProgram_ShouldHandleNA_BySettingEmptyString()
        {
            // Arrange
            var flowItem = new FlowRow { LoLim = "5.0" };
            var binoutItem = new BinOutStatusHipListRow("sourceSheetName");
            string jobpart = "PartA";
            binoutItem.UpdatedHiLimitDic[jobpart] = "";
            binoutItem.UpdatedLoLimitDic[jobpart] = "N/A";
            binoutItem.BinOutEnableDictionary[jobpart] = "binout";
            List<FlowRow> flowAllTestItems =
            [
                new() { HiLim = "10.0", SheetName = "TestSheet", Comment1 = "Comment1 Comment2" }
            ];

            // Act
            var service = new HardIpBinOutUpdateController(true);
            service.UpdateItemToProgram(ref flowItem, binoutItem, jobpart, flowAllTestItems, "Sheet1");

            // Assert
            Assert.AreEqual("", flowItem.LoLim, "N/A should be converted to an empty string.");
        }

        [TestMethod]
        public void UpdateItemToProgram_ShouldHandleNA_BySettingEmptyString_1()
        {
            // Arrange
            var flowItem = new FlowRow { LoLim = "5.0A", FailAction = "FailAction" };
            var binoutItem = new BinOutStatusHipListRow("sourceSheetName");
            string jobpart = "PartA";
            binoutItem.UpdatedHiLimitDic[jobpart] = "";
            binoutItem.UpdatedLoLimitDic[jobpart] = "N/A";
            binoutItem.BinOutEnableDictionary[jobpart] = "nonbinout";

            // Act
            var service = new HardIpBinOutUpdateController(true);
            service.UpdateItemToProgram(ref flowItem, binoutItem, jobpart, [], "Sheet1");

            // Assert
            Assert.AreEqual("", flowItem.LoLim, "N/A should be converted to an empty string.");
        }

        [TestMethod]
        public void UpdateItemToProgram_ShouldReturnFalse_WhenValuesAreSame()
        {
            // Arrange
            var flowItem = new FlowRow { HiLim = "10.0", LoLim = "5.0", Parameter = "MBIST_SRAM_PP001_HV", SheetName = "Flow_ABC" };
            var binoutItem = new BinOutStatusHipListRow("sourceSheetName");
            string jobpart = "PartA";
            binoutItem.UpdatedHiLimitDic[jobpart] = "10.0";
            binoutItem.UpdatedLoLimitDic[jobpart] = "5.0";
            binoutItem.BinOutEnableDictionary[jobpart] = "binout";

            // Act
            var service = new HardIpBinOutUpdateController(true);
            bool result = service.UpdateItemToProgram(ref flowItem, binoutItem, jobpart, [], "Sheet1");

            // Assert
            Assert.IsFalse(result, "Should return false if no data actually changed.");
        }

        [TestMethod]
        [DataRow("Enabled", "Enabled", true, DisplayName = "All jobs match")]
        [DataRow("Enabled", "Disabled", false, DisplayName = "Jobs do not match")]
        [DataRow("Enabled", "enabled", true, DisplayName = "Case insensitive match")]
        public void UpdateBinOutValueIsSame_ComparingRelevantJobs(string val1, string val2, bool expected)
        {
            // Arrange
            var service = new HardIpBinOutUpdateController(true);
            var binoutDic = new Dictionary<string, string>
            {
                { "JobA", val1 },
                { "JobB", val2 },
                { "JobC", "IgnoreMe" } // This job shouldn't be checked
            };
            var existJobs = new HashSet<string> { "JobA", "JobB" };

            // Act
            bool result = service.UpdateBinOutValueIsSame(binoutDic, existJobs);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void UpdateBinOutValueIsSame_ShouldReturnTrue_WhenOnlyOneJobExists()
        {
            // Arrange
            var service = new HardIpBinOutUpdateController(true);
            var binoutDic = new Dictionary<string, string> { { "JobA", "Enabled" } };
            var existJobs = new HashSet<string> { "JobA" };

            // Act & Assert
            Assert.IsTrue(service.UpdateBinOutValueIsSame(binoutDic, existJobs));
        }

        [TestMethod]
        public void UpdateBinOutValueIsSame_ShouldIgnoreJobsNotInHashSet()
        {
            // Arrange
            var service = new HardIpBinOutUpdateController(true);
            var binoutDic = new Dictionary<string, string>
            {
                { "JobA", "Enabled" },
                { "JobB", "Disabled" } // Differing value, but JobB is not in existJobs
            };
            var existJobs = new HashSet<string> { "JobA" };

            // Act
            bool result = service.UpdateBinOutValueIsSame(binoutDic, existJobs);

            // Assert
            Assert.IsTrue(result, "Should ignore JobB and return true because only JobA is relevant.");
        }

        [DataTestMethod]
        [DataRow("Flow_Hardip_ABC_DEF", "ABCDEF", DisplayName = "StripsPrefixAndUnderscores")]
        [DataRow("Flow_Other_ABC", "Flow_Other_ABC", DisplayName = "NonHardipPrefixUnchanged")]
        public void GetBlockFromSheetName_StripsHardipPrefix(string sheetpath, string expected)
        {
            // Arrange
            var target = new HardIpBinOutUpdateController(true);

            // Act
            string result = target.GetBlockFromSheetName(sheetpath);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void GetHardipBintableFromProgram_FindsMatchingSheetName()
        {
            // Arrange
            var target = new HardIpBinOutUpdateController(true);
            BinTableSheet hardipBintable = new BinTableSheet("Bin_Table_HardIP");
            hardipBintable.Rows.Add(new BinTableRow());

            List<BinTableSheet> sheets = new List<BinTableSheet> { new("Other"), hardipBintable };

            // Act
            BinTableSheet result = target.GetHardipBintableFromProgram(sheets);

            // Assert
            Assert.AreEqual("Bin_Table_HardIP", result.Name);
        }

        [TestMethod]
        public void GetHardipBintableFromProgram_EmptyList_ReturnsNull()
        {
            // Arrange
            var target = new HardIpBinOutUpdateController(true);

            // Act
            BinTableSheet result = target.GetHardipBintableFromProgram([]);

            // Assert
            Assert.AreEqual(null, result);
        }

        [TestMethod]
        public void GetHardipBintableRows_GroupsFlagsByBinName()
        {
            // Arrange
            var target = new HardIpBinOutUpdateController(true);
            var sheet = new BinTableSheet("Bin_Table_HardIP");
            sheet.AddRow(new BinTableRow { Name = "Bin_A", ItemList = "F_1,F_2" });
            sheet.AddRow(new BinTableRow { Name = "Bin_A", ItemList = "F_3" });
            sheet.AddRow(new BinTableRow { Name = "Bin_B", ItemList = "" });

            // Act
            Dictionary<string, List<string>> result = target.GetHardipBintableRows([sheet]);

            // Assert
            Assert.IsTrue(result.ContainsKey("Bin_A"));
            CollectionAssert.AreEquivalent(new List<string> { "F_1", "F_2", "F_3" }, result["Bin_A"]);
            Assert.IsFalse(result.ContainsKey("Bin_B"));
        }

        [TestMethod]
        public void GetHardipBintableRows_NoMatchingSheet_ReturnsEmpty()
        {
            // Arrange
            var target = new HardIpBinOutUpdateController(true);

            // Act
            Dictionary<string, List<string>> result = target.GetHardipBintableRows([]);

            // Assert
            Assert.AreEqual(0, result.Count);
        }

        [DataTestMethod]
        [DataRow("Block__TestName", "TestName_Block", DisplayName = "DoubleUnderscoreSwapsSegments")]
        [DataRow("Simple_Test", "Simple-Test_x", DisplayName = "NoDoubleUnderscoreAppendsX")]
        public void ProcessTpTestName_TransformsBasedOnDoubleUnderscore(string input, string expected)
        {
            // Arrange
            var target = new HardIpBinOutUpdateController(true);

            // Act
            string result = target.ProcessTpTestName(input);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [DataTestMethod]
        [DataRow("binout", BinOutStatus.Binout, DisplayName = "BinoutPrefix")]
        [DataRow("non binout", BinOutStatus.NonBinout, DisplayName = "NonBinoutPhrase")]
        [DataRow("something else", BinOutStatus.Unknown, DisplayName = "NoMatchIsUnknown")]
        public void GetBinoutStatusFormReport_MapsTextToStatus(string status, BinOutStatus expected)
        {
            // Arrange
            var target = new HardIpBinOutUpdateController(true);

            // Act
            BinOutStatus result = target.GetBinoutStatusFormReport(status);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void GetBinNameByFailFalg_FindsFirstMatchingBin()
        {
            // Arrange
            var target = new HardIpBinOutUpdateController(true);
            var bintable = new Dictionary<string, List<string>> { { "Bin_A", new List<string> { "F_1", "F_2" } }, { "Bin_B", new List<string> { "F_3" } } };
            var flaggroup = new HashSet<string> { "F_3" };

            // Act
            string result = target.GetBinNameByFailFalg(bintable, flaggroup);

            // Assert
            Assert.AreEqual("Bin_B", result);
        }

        [TestMethod]
        public void GetBinNameByFailFalg_NoMatch_ReturnsEmpty()
        {
            // Arrange
            var target = new HardIpBinOutUpdateController(true);
            var bintable = new Dictionary<string, List<string>> { { "Bin_A", new List<string> { "F_1" } } };
            var flaggroup = new HashSet<string> { "F_X" };

            // Act
            string result = target.GetBinNameByFailFalg(bintable, flaggroup);

            // Assert
            Assert.AreEqual(string.Empty, result);
        }

        [TestMethod]
        public void SaveBinoutItemsForUpdated_NewKey_AddsEntry()
        {
            // Arrange
            var target = new HardIpBinOutUpdateController(true);
            var allBinoutItems = new Dictionary<string, List<BinOutStatusHipListRow>>();
            var row = new BinOutStatusHipListRow("Sheet1");
            var binoutItems = new List<KeyValuePair<string, List<BinOutStatusHipListRow>>> { new("k1", [row]) };

            // Act
            target.SaveBinoutItemsForUpdated(ref allBinoutItems, binoutItems, "Inst1");

            // Assert
            Assert.IsTrue(allBinoutItems.ContainsKey("Inst1"));
            Assert.AreEqual(1, allBinoutItems["Inst1"].Count);
        }

        [TestMethod]
        public void SaveBinoutItemsForUpdated_ExistingKey_AppendsToList()
        {
            // Arrange
            var target = new HardIpBinOutUpdateController(true);
            var existingRow = new BinOutStatusHipListRow("Sheet0");
            var allBinoutItems = new Dictionary<string, List<BinOutStatusHipListRow>> { { "Inst1", new List<BinOutStatusHipListRow> { existingRow } } };
            var row = new BinOutStatusHipListRow("Sheet1");
            var binoutItems = new List<KeyValuePair<string, List<BinOutStatusHipListRow>>> { new("k1", [row]) };

            // Act
            target.SaveBinoutItemsForUpdated(ref allBinoutItems, binoutItems, "Inst1");

            // Assert
            Assert.AreEqual(2, allBinoutItems["Inst1"].Count);
        }

        [TestMethod]
        public void GetBinoutEnableDic_MergesAcrossRows_PrefersBinoutOverUnknown()
        {
            // Arrange
            var target = new HardIpBinOutUpdateController(true);
            var row1 = new BinOutStatusHipListRow("S1");
            row1.BinOutEnableDictionary["CP1_A"] = "SomethingElse";
            var row2 = new BinOutStatusHipListRow("S2");
            row2.BinOutEnableDictionary["CP1_A"] = "binout";

            // Act
            Dictionary<string, string> result = target.GetBinoutEnableDic([row1, row2]);

            // Assert
            Assert.AreEqual("binout", result["CP1_A"]);
        }

        [TestMethod]
        public void GetTestJobs_SingleEnableEntry_ReturnsFirstBinOutKey()
        {
            // Arrange
            var target = new HardIpBinOutUpdateController(true);
            var row = new BinOutStatusHipListRow("S1");
            row.EnableDictionary["CP1_A"] = "x";
            row.BinOutEnableDictionary["CP1_A"] = "binout";
            var kvp = new KeyValuePair<string, List<BinOutStatusHipListRow>>("k", [row]);

            // Act
            HashSet<string> result = target.GetTestJobs(kvp);

            // Assert
            CollectionAssert.AreEqual(new List<string> { "CP1_A" }, result.ToList());
        }

        [TestMethod]
        public void GetTestJobs_MultipleEnableEntries_BuildsJobPartsFromKeys()
        {
            // Arrange
            var target = new HardIpBinOutUpdateController(true);
            var row = new BinOutStatusHipListRow("S1");
            row.EnableDictionary["CP1_A_extra"] = "x";
            row.EnableDictionary["CP2_B_extra"] = "x";
            row.BinOutEnableDictionary["CP1_A"] = "binout";
            row.BinOutEnableDictionary["CP2_B"] = "binout";
            var kvp = new KeyValuePair<string, List<BinOutStatusHipListRow>>("k", [row]);

            // Act
            HashSet<string> result = target.GetTestJobs(kvp);

            // Assert
            CollectionAssert.AreEquivalent(new List<string> { "CP1_A", "CP2_B" }, result.ToList());
        }

        [TestMethod]
        public void HasNoTestItemNeedToUpdate_EnableJobNotInTestJobParts_ReturnsTrue()
        {
            // Arrange
            var target = new HardIpBinOutUpdateController(true);
            var item = new BinOutStatusHipListRow("S1");
            item.BinOutEnableDictionary["CP1_A"] = "binout";
            var testJobParts = new HashSet<string> { "CP2_B" };
            string noTestJob = "";

            // Act
            bool result = target.HasNoTestItemNeedToUpdate(item, testJobParts, ref noTestJob);

            // Assert
            Assert.IsTrue(result);
            Assert.AreEqual("CP1_A", noTestJob);
        }

        [TestMethod]
        public void HasNoTestItemNeedToUpdate_AllJobsCovered_ReturnsFalse()
        {
            // Arrange
            var target = new HardIpBinOutUpdateController(true);
            var item = new BinOutStatusHipListRow("S1");
            item.BinOutEnableDictionary["CP1_A"] = "binout";
            var testJobParts = new HashSet<string> { "CP1_A" };
            string noTestJob = "";

            // Act
            bool result = target.HasNoTestItemNeedToUpdate(item, testJobParts, ref noTestJob);

            // Assert
            Assert.IsFalse(result);
        }

        [DataTestMethod]
        [DataRow("", "CP1", true, DisplayName = "EmptyRowJobPartAlwaysMatches")]
        [DataRow("CP1", "CP1", true, DisplayName = "ExactMatch")]
        [DataRow("CP2", "CP1", false, DisplayName = "NoMatch")]
        public void CheckJobPartIsMatch_EvaluatesJobPartList(string rowJobPart, string currentJobPart, bool expected)
        {
            // Arrange
            var target = new HardIpBinOutUpdateController(true);

            // Act
            bool result = target.CheckJobPartIsMatch(rowJobPart, currentJobPart);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void GenerateFlagNameFromTool_ParameterDoesNotMatchPattern_ReturnsEmpty()
        {
            // Arrange
            var target = new HardIpBinOutUpdateController(true);
            var row = new FlowRow { Parameter = "NotMatchingAnything123", SheetName = "Flow_X" };

            // Act
            string result = target.GenerateFlagNameFromTool(row);

            // Assert
            Assert.AreEqual(string.Empty, result);
        }
    }
}
