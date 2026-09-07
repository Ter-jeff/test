using System.Collections.Generic;

using Automation.GenerateIgxl.BinCut.Business;

using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.BinCut.Binning;

namespace Automation.Test.UT.BinCut
{
    [TestClass]
    public class BinCutBinTableWriterTests
    {
        [TestMethod]
        public void SetItemResultToTrue_AddsTForEachItemListEntry()
        {
            // Arrange
            var row = new BinTableRow { ItemList = "F_A,F_B,F_C" };

            // Act
            BinCutBinTableWriter.SetItemResultToTrue(ref row);

            // Assert
            CollectionAssert.AreEqual(new List<string> { "T", "T", "T" }, row.Items);
        }

        [TestMethod]
        public void SetItemResultToTrue_NullRow_DoesNotThrow()
        {
            // Act
            BinTableRow? row = null;
            BinCutBinTableWriter.SetItemResultToTrue(ref row);
        }

        [DataTestMethod]
        [DataRow("Rev:1.2\tRevInfo", "Rev:1P2RevInfo", DisplayName = "WithDotReplacesWithP")]
        [DataRow("Rev:1\tRevInfo", "Rev:1RevInfoP0", DisplayName = "NoDotAppendsP0")]
        [DataRow("NoRevKeywordHere\tOther", null, DisplayName = "NoRevKeyword_ReturnsNull")]
        public void GenBinningTableVersion_FormatsRevisionFromTitleRow(string titleRow1, string expected)
        {
            // Arrange
            var writer = new BinCutBinTableWriter();
            var table = new BinningTable { TitleRow1 = titleRow1 };

            // Act
            string? result = writer.GenBinningTableVersion(table);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void GenBinningTableVersion_NullBinningTable_ReturnsNull()
        {
            // Arrange
            var writer = new BinCutBinTableWriter();

            // Act
            string? result = writer.GenBinningTableVersion(null);

            // Assert
            Assert.IsNull(result);
        }

        [DataTestMethod]
        [DataRow("Flow_ELB_Something", "RTOS", DisplayName = "ElbMapsToRtos")]
        [DataRow("Flow_ILB_Something", "RTOS", DisplayName = "IlbMapsToRtos")]
        [DataRow("Flow_TMPS_Something", "TMPS", DisplayName = "TmpsMapsToTmps")]
        [DataRow("Flow_CPM_Something", "CPM", DisplayName = "CpmMapsToCpm")]
        [DataRow("Flow_Other", "RTOS", DisplayName = "DefaultMapsToRtos")]
        public void GetTypeFromFlowName_MapsContentToType(string content, string expected)
        {
            // Arrange
            var writer = new BinCutBinTableWriter();

            // Act
            string? result = writer.GetTypeFromFlowName(content);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void GetBinItems_RowNameBinX_AddsFTF()
        {
            // Arrange
            var writer = new BinCutBinTableWriter();
            var row = new BinTableRow { Name = "Bin_BinX" };

            // Act
            writer.GetBinItems("AnyBin", row);

            // Assert
            CollectionAssert.AreEqual(new List<string> { "F", "T", "F" }, row.Items);
        }

        [TestMethod]
        public void GetBinItems_RowNameBinY_AddsFFT()
        {
            // Arrange
            var writer = new BinCutBinTableWriter();
            var row = new BinTableRow { Name = "Bin_BinY" };

            // Act
            writer.GetBinItems("AnyBin", row);

            // Assert
            CollectionAssert.AreEqual(new List<string> { "F", "F", "T" }, row.Items);
        }

        [TestMethod]
        public void GetBinItems_OtherRowNameBin1_UsesBin1Pattern()
        {
            // Arrange
            var writer = new BinCutBinTableWriter();
            var row = new BinTableRow { Name = "SomeOtherName" };

            // Act
            writer.GetBinItems("Bin1", row);

            // Assert
            CollectionAssert.AreEqual(new List<string> { "T", "T", "F", "F" }, row.Items);
        }

        [TestMethod]
        public void GetBinItems_OtherRowNameBinX_UsesBinXPattern()
        {
            // Arrange
            var writer = new BinCutBinTableWriter();
            var row = new BinTableRow { Name = "SomeOtherName" };

            // Act
            writer.GetBinItems("BinX", row);

            // Assert
            CollectionAssert.AreEqual(new List<string> { "T", "F", "T", "F" }, row.Items);
        }

        [TestMethod]
        public void GetBinItems_OtherRowNameBinY_UsesBinYPattern()
        {
            // Arrange
            var writer = new BinCutBinTableWriter();
            var row = new BinTableRow { Name = "SomeOtherName" };

            // Act
            writer.GetBinItems("BinY", row);

            // Assert
            CollectionAssert.AreEqual(new List<string> { "T", "F", "F", "T" }, row.Items);
        }

        [TestMethod]
        public void GetIdsBinItems_Bin1_UsesBin1Pattern()
        {
            // Arrange
            var writer = new BinCutBinTableWriter();
            var row = new BinTableRow();

            // Act
            writer.GetIdsBinItems("Bin1", row);

            // Assert
            CollectionAssert.AreEqual(new List<string> { "T", "F", "F", "T" }, row.Items);
        }

        [TestMethod]
        public void GetIdsBinItems_BinX_UsesBinXPattern()
        {
            // Arrange
            var writer = new BinCutBinTableWriter();
            var row = new BinTableRow();

            // Act
            writer.GetIdsBinItems("BinX", row);

            // Assert
            CollectionAssert.AreEqual(new List<string> { "F", "T", "F", "T" }, row.Items);
        }

        [TestMethod]
        public void GetIdsBinItems_BinY_UsesBinYPattern()
        {
            // Arrange
            var writer = new BinCutBinTableWriter();
            var row = new BinTableRow();

            // Act
            writer.GetIdsBinItems("BinY", row);

            // Assert
            CollectionAssert.AreEqual(new List<string> { "F", "F", "T", "T" }, row.Items);
        }

        [TestMethod]
        public void GetPinDicFromBinning_NullBinningTable_ReturnsEmptyDictionary()
        {
            // Arrange
            var writer = new BinCutBinTableWriter();

            // Act
            Dictionary<string, string> result = writer.GetPinDicFromBinning(null, null, false);

            // Assert
            Assert.AreEqual(0, result?.Count);
        }
    }
}
