using System.Collections.Generic;

using Automation.GenerateIgxl.Basic.Business;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.Harvest;

namespace Automation.Test.UT.Basic
{
    [TestClass]
    public class AutogenMainApTests
    {
        private static void InvokeAddNewMappingCoreRow(MappingCoreRow newRow, MappingCoreTable table, List<string> existRows)
        {
            HarvestCoreMappingChecker.AddNewMappingCoreRow(newRow, table, existRows);
        }

        private static MappingCoreRow NewRow(string initPattern = "Init1", string pattern = "Pat1", string coreName = "Core1", string harvestFlag = "F1", string powerSupply = "VDD1", string comment = "C1")
        {
            return new MappingCoreRow
            {
                InitPattern = initPattern,
                Pattern = pattern,
                CoreName = coreName,
                HarvestFlag = harvestFlag,
                PowerSupply = powerSupply,
                Comment = comment
            };
        }

        [TestMethod]
        public void AddNewMappingCoreRow_NewUniqueRow_AddsToTableAndExistRows()
        {
            // Arrange
            var table = new MappingCoreTable("Sheet1");
            var existRows = new List<string>();
            MappingCoreRow row = NewRow();

            // Act
            InvokeAddNewMappingCoreRow(row, table, existRows);

            // Assert
            Assert.AreEqual(1, table.Rows.Count);
            Assert.AreEqual(1, existRows.Count);
        }

        [TestMethod]
        public void AddNewMappingCoreRow_DuplicateRow_IsNotAddedAgain()
        {
            // Arrange - matches case-insensitively since the joined key is upper-cased
            var table = new MappingCoreTable("Sheet1");
            var existRows = new List<string> { "INIT1|PAT1|CORE1|F1|VDD1|C1" };
            MappingCoreRow row = NewRow();

            // Act
            InvokeAddNewMappingCoreRow(row, table, existRows);

            // Assert
            Assert.AreEqual(0, table.Rows.Count);
            Assert.AreEqual(1, existRows.Count);
        }

        [TestMethod]
        public void AddNewMappingCoreRow_DifferentRow_IsAddedAlongsideExisting()
        {
            // Arrange
            var table = new MappingCoreTable("Sheet1");
            var existRows = new List<string> { "INIT1|PAT1|CORE1|F1|VDD1|C1" };
            MappingCoreRow row = NewRow(pattern: "Pat2");

            // Act
            InvokeAddNewMappingCoreRow(row, table, existRows);

            // Assert
            Assert.AreEqual(1, table.Rows.Count);
            Assert.AreEqual(2, existRows.Count);
        }

        [TestMethod]
        public void ContentMap_Always_MapsEachHeaderToExpectedProperty()
        {
            MappingCoreRow row = NewRow();

            Assert.AreEqual(row.InitPattern, HarvestCoreMappingChecker._contentMap["inpattern"](row));
            Assert.AreEqual(row.Pattern, HarvestCoreMappingChecker._contentMap["plpattern"](row));
            Assert.AreEqual(row.CoreName, HarvestCoreMappingChecker._contentMap["corename/pingroup"](row));
            Assert.AreEqual(row.HarvestFlag, HarvestCoreMappingChecker._contentMap["harvestflag"](row));
            Assert.AreEqual(row.PowerSupply, HarvestCoreMappingChecker._contentMap["powersupply (multiple rails bincut search)"](row));
            Assert.AreEqual(row.Comment, HarvestCoreMappingChecker._contentMap["comment"](row));
        }

        [TestMethod]
        public void Regexes_LowerCaseInput_MatchCaseInsensitively()
        {
            Assert.IsTrue(HarvestCoreMappingChecker._regexPl.IsMatch("test_plab_test"));
            Assert.IsTrue(HarvestCoreMappingChecker._regex1.IsMatch("test_lpb_test"));
            Assert.IsTrue(HarvestCoreMappingChecker._regex2.IsMatch("test_ch_test"));
            Assert.IsTrue(HarvestCoreMappingChecker._regex3.IsMatch("test_ssc_test"));
        }
    }
}
