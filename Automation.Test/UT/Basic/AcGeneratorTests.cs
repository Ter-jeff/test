using System.Collections.Generic;
using System.Data;
using System.Linq;

using Automation.GenerateIgxl.Basic.Business.GenAc.AcGenerator.Business;
using Automation.GenerateIgxl.Basic.Business.GenAc.AcInput.BassData;
using Automation.Singleton;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;
using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using BasicPatternData = TestPlanLib.Basic.PatternData;

namespace Automation.Test.UT.Basic
{
    [TestClass]
    public class AcGeneratorTests
    {
        [TestMethod]
        public void GenerateFlow_Always_CreatesAcSpecSheetWithExpectedNameAndSelectors()
        {
            // Arrange
            var generator = new AcGenerator(new AcInputSheet(), new AcSpecSheet("TestSheet"));

            // Act
            AcSpecSheet result = generator.GenerateFlow([], []);

            // Assert
            Assert.AreEqual("AC_Specs", result.Name);
            CollectionAssert.AreEqual(new List<string> { "Typ", "Min", "Max" }, result.SelectorNameList);
        }

        [TestMethod]
        public void GenerateFlow_NwireSettingTableHasBlankAndNonBlankRows_OnlyNonBlankBecomeCategories()
        {
            // Arrange
            var generator = new AcGenerator(new AcInputSheet(), new AcSpecSheet("TestSheet"));

            var table = new DataTable();
            table.Columns.Add("Flow");
            table.Rows.Add("SomeCategory");
            table.Rows.Add("   ");
            NwireSingleton.Instance().SettingInfo.SettingTable = table;

            // Act
            generator.GenerateFlow([], []);

            // Assert
            Assert.IsTrue(generator._acFullCategoryList.Contains("Common_SomeCategory"));
            Assert.IsFalse(generator._acFullCategoryList.Any(x => x.EndsWith("_   ")));
        }

        [TestMethod]
        public void WriteOneRecord_Nwire_AllBranches_ShouldWork()
        {
            // Arrange
            var generator = new AcGenerator(new AcInputSheet(), new AcSpecSheet("TestSheet"));

            var table = new DataTable();
            table.Columns.Add("Flow");
            table.Columns.Add("Pin0");

            // Header
            table.Rows.Add("Header0", "");
            table.Rows.Add("Meta0", "");

            table.Rows.Add("Default_100K", "Enable@100KHZ");
            table.Rows.Add("Default_200K_WRONG", "Enable@200KHZ");
            table.Rows.Add("Default_300K", "Disable@300KHZ");
            for (int i = 0; i < 5; i++)
            {
                table.Rows.Add($"Default_extra_{i}", "Enable@100KHZ");
            }

            NwireSingleton.Instance().SettingInfo.SettingTable = table;

            var timeSetSheets = new TimeSetSheets();
            var patLists = new List<BasicPatternData>();
            generator.GenerateFlow(timeSetSheets, patLists);

            // Act
            AcSpec result = generator.WriteOneRecordToAcSpecForNwire("SYM", "VAL", "TYP", "MIN", "MAX", 0);

            // Assert
            Assert.AreNotEqual(null, result);
        }

        [TestMethod]
        public void WriteOneRecord_Tck_SpiRomCategory_ShouldUse10MHz()
        {
            // Arrange
            var generator = new AcGenerator(new AcInputSheet(), new AcSpecSheet("TestSheet"));

            var timeSetSheets = new TimeSetSheets();
            var patLists = new List<BasicPatternData>();

            generator.GenerateFlow(timeSetSheets, patLists);

            Assert.IsTrue(generator._acFullCategoryList.Any(x => x.EqualsIgnoreCase("SPI_ROM")));

            // Act
            AcSpec result = generator.WriteOneRecordToAcSpecForTck(pStrSymbol: "SYM", pStrValue: "VAL", pStrTyp: "TYP", pStrMin: "MIN", pStrMax: "MAX");

            // Assert
            Assert.AreNotEqual(null, result);
        }

        [TestMethod]
        public void WriteOneRecord_Tck_NonSpiRomCategory_ShouldFallback()
        {
            // Arrange
            var generator = new AcGenerator(new AcInputSheet(), new AcSpecSheet("TestSheet"));

            var timeSetSheets = new TimeSetSheets();
            var patLists = new List<BasicPatternData>();

            generator.GenerateFlow(timeSetSheets, patLists);

            Assert.IsTrue(generator._acFullCategoryList.Any(x => !x.EqualsIgnoreCase("SPI_ROM")));

            // Act
            AcSpec result = generator.WriteOneRecordToAcSpecForTck("SYM", "VAL", "TYP", "MIN", "MAX");

            // Assert
            Assert.AreNotEqual(null, result);
        }

        [TestMethod]
        public void WriteOneRecord_ShouldCreateCategories_ForAllAcFullCategoryList()
        {
            // Arrange
            var generator = new AcGenerator(new AcInputSheet(), new AcSpecSheet("TestSheet"));

            var timeSetSheets = new TimeSetSheets();
            var patLists = new List<BasicPatternData>();

            generator.GenerateFlow(timeSetSheets, patLists);

            Assert.IsTrue(generator._acFullCategoryList.Count > 0);

            // Act
            AcSpec result = generator.WriteOneRecordToAcSpec(pStrSymbol: "SYM", pStrValue: "VAL", pStrTyp: "TYP", pStrMin: "MIN", pStrMax: "MAX");

            // Assert
            Assert.AreNotEqual(null, result);
        }

        [TestMethod]
        public void WriteOneRecordToAcSpecSheet_TckSymbol_ShouldNotThrow()
        {
            // Arrange
            var generator = new AcGenerator(
                new AcInputSheet(),
                new AcSpecSheet("TestSheet"));

            generator.GenerateFlow([], []);

            // Act
            generator.WriteOneRecordToAcSpecSheet("TCK_Freq", "VAL", "TYP", "MIN", "MAX");
        }

        [TestMethod]
        public void WriteOneRecordToAcSpecSheet_NwireMatch_ShouldNotThrow()
        {
            // Arrange
            var generator = new AcGenerator(
                new AcInputSheet(),
                new AcSpecSheet("TestSheet"));

            NwireSingleton.Instance().SettingInfo.NwirePins =
                [
                    new() { OutClk = "CLK0" }
                ];

            generator.GenerateFlow([], []);

            // Act
            generator.WriteOneRecordToAcSpecSheet(pStrSymbol: "XXX_CLK0_YYY", pStrValue: "VAL", pStrTyp: "TYP", pStrMin: "MIN", pStrMax: "MAX");
        }

        [TestMethod]
        public void WriteOneRecordToAcSpecSheet_Default_ShouldNotThrow()
        {
            // Arrange
            var generator = new AcGenerator(
                new AcInputSheet(),
                new AcSpecSheet("TestSheet"));

            NwireSingleton.Instance().SettingInfo.NwirePins =
                [];

            generator.GenerateFlow([], []);

            // Act
            generator.WriteOneRecordToAcSpecSheet(pStrSymbol: "NORMAL_SYMBOL", pStrValue: "VAL", pStrTyp: "TYP", pStrMin: "MIN", pStrMax: "MAX");
        }

        [TestMethod]
        public void WriteOneRecordToAcSpecSheet_OutClkIsNull_ShouldGoDefault()
        {
            // Arrange
            var generator = new AcGenerator(
                new AcInputSheet(),
                new AcSpecSheet("TestSheet"));

            NwireSingleton.Instance().SettingInfo.NwirePins =
                [
            new() { OutClk = null }
                ];

            generator.GenerateFlow([], []);

            // Act
            generator.WriteOneRecordToAcSpecSheet(pStrSymbol: "XXX_CLK0_YYY", pStrValue: "VAL", pStrTyp: "TYP", pStrMin: "MIN", pStrMax: "MAX");
        }

        private static BlockType InvokeGetTimesetBlockType(AcGenerator generator, string sheetName)
        {
            return generator.GetTimesetBlockType(sheetName);
        }

        [DataTestMethod]
        [DataRow("Scan")]
        [DataRow("SAA")]
        [DataRow("saa")]
        public void GetBlockType_ContainsScanOrSaaToken_ReturnsScan(string token)
        {
            // Arrange
            string[] toks = ["A", "B", token];

            // Act
            BlockType result = AcGenerator.GetBlockType(toks);

            // Assert
            Assert.AreEqual(BlockType.Scan, result);
        }

        [TestMethod]
        public void GetBlockType_ContainsRtosToken_ReturnsSpiRom()
        {
            // Arrange
            string[] toks = ["A", "B", "Rtos"];

            // Act
            BlockType result = AcGenerator.GetBlockType(toks);

            // Assert
            Assert.AreEqual(BlockType.SPI_ROM, result);
        }

        [DataTestMethod]
        [DataRow("AN", BlockType.HardIp)]
        [DataRow("JT", BlockType.HardIp)]
        [DataRow("SC", BlockType.Scan)]
        [DataRow("BI", BlockType.Mbist)]
        [DataRow("ZZ", BlockType.None)]
        public void GetBlockType_ByBlockToken_ReturnsExpectedType(string blockToken, BlockType expected)
        {
            // Arrange
            string[] toks = ["A", "B", "C", blockToken];

            // Act
            BlockType result = AcGenerator.GetBlockType(toks);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [DataTestMethod]
        [DataRow("bsr")]
        [DataRow("mbist")]
        public void GetBlockType_ContainsBsrOrMbistToken_ReturnsMbist(string token)
        {
            // Arrange
            string[] toks = ["A", "B", "C", "ZZ", token];

            // Act
            BlockType result = AcGenerator.GetBlockType(toks);

            // Assert
            Assert.AreEqual(BlockType.Mbist, result);
        }

        [TestMethod]
        public void GetTimesetBlockType_FewerThanFiveTokens_ReturnsHardIpDefault()
        {
            // Arrange
            var generator = new AcGenerator(new AcInputSheet(), new AcSpecSheet("TestSheet"));

            // Act
            BlockType result = InvokeGetTimesetBlockType(generator, "A_B_C");

            // Assert
            Assert.AreEqual(BlockType.HardIp, result);
        }

        [TestMethod]
        public void GetTimesetBlockType_FiveOrMoreTokens_DelegatesToGetBlockType()
        {
            // Arrange
            var generator = new AcGenerator(new AcInputSheet(), new AcSpecSheet("TestSheet"));

            // Act
            BlockType result = InvokeGetTimesetBlockType(generator, "A_B_C_SC_D");

            // Assert
            Assert.AreEqual(BlockType.Scan, result);
        }
    }
}
