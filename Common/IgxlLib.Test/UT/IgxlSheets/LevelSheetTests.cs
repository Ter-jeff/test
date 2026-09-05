using System.Collections.Generic;

using IgxlLib.Const;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

namespace IgxlLib.Test.UT.IgxlSheets
{
    [TestClass]
    public class LevelSheetTests
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void LevelSheet_Constructor_WithSheetName()
        {
            // Arrange
            string sheetName = "Level";

            // Act
            var levelSheet = new LevelSheet(sheetName);

            // Assert
            Assert.IsNotNull(levelSheet);
            Assert.AreEqual(sheetName, levelSheet.Name);
            Assert.AreEqual("DTLevelSheet", levelSheet.SheetType);
            Assert.AreEqual(IgxlSheetNames.PinLevel, levelSheet.IgxlSheetName);
            Assert.AreEqual(0, levelSheet.Rows.Count);
        }

        [TestMethod]
        public void LevelSheet_AddRow()
        {
            // Arrange
            var levelSheet = new LevelSheet("Level");
            var levelRow = new LevelRow("Pin1", "vih", "1.5", "test comment");

            // Act
            levelSheet.AddRow(levelRow);

            // Assert
            Assert.AreEqual(1, levelSheet.Rows.Count);
            Assert.AreEqual("Pin1", levelSheet.Rows[0].PinName);
        }

        [TestMethod]
        public void LevelSheet_AddRows()
        {
            // Arrange
            var levelSheet = new LevelSheet("Level");
            var rows = new List<LevelRow>
            {
                new("Pin1", "vih", "1.5", "test"),
                new("Pin2", "vil", "0.5", "test"),
                new("Pin3", "vih", "1.8", "test")
            };

            // Act
            levelSheet.AddRows(rows);

            // Assert
            Assert.AreEqual(3, levelSheet.Rows.Count);
        }

        [TestMethod]
        public void LevelSheet_RemoveRow()
        {
            // Arrange
            var levelSheet = new LevelSheet("Level");
            var row1 = new LevelRow("Pin1", "vih", "1.5", "test");
            var row2 = new LevelRow("Pin2", "vil", "0.5", "test");
            levelSheet.AddRow(row1);
            levelSheet.AddRow(row2);

            // Act
            levelSheet.RemoveRow(row1);

            // Assert
            Assert.AreEqual(1, levelSheet.Rows.Count);
            Assert.AreEqual("Pin2", levelSheet.Rows[0].PinName);
        }

        [TestMethod]
        public void LevelSheet_InsertRow_AtIndex()
        {
            // Arrange
            var levelSheet = new LevelSheet("Level");
            var row1 = new LevelRow("Pin1", "vih", "1.5", "test");
            var row2 = new LevelRow("Pin2", "vil", "0.5", "test");
            var rowToInsert = new LevelRow("Pin3", "vih", "1.8", "test");
            levelSheet.AddRow(row1);
            levelSheet.AddRow(row2);

            // Act
            int index = levelSheet.InsertRow(1, rowToInsert);

            // Assert
            Assert.AreEqual(2, index);
            Assert.AreEqual(3, levelSheet.Rows.Count);
            Assert.AreEqual("Pin3", levelSheet.Rows[1].PinName);
        }

        [TestMethod]
        public void LevelSheet_InsertRow_AtEndWithNegativeIndex()
        {
            // Arrange
            var levelSheet = new LevelSheet("Level");
            var row1 = new LevelRow("Pin1", "vih", "1.5", "test");
            var rowToInsert = new LevelRow("Pin2", "vil", "0.5", "test");
            levelSheet.AddRow(row1);

            // Act
            int index = levelSheet.InsertRow(-1, rowToInsert);

            // Assert
            Assert.AreEqual(-1, index);
            Assert.AreEqual(2, levelSheet.Rows.Count);
        }

        [TestMethod]
        public void LevelSheet_Copy()
        {
            // Arrange
            var originalSheet = new LevelSheet("Level");
            originalSheet.AddRow(new LevelRow("Pin1", "vih", "1.5", "test"));
            originalSheet.AddRow(new LevelRow("Pin2", "vil", "0.5", "test"));

            // Act
            LevelSheet copiedSheet = originalSheet.Copy();

            // Assert
            Assert.AreNotSame(originalSheet, copiedSheet);
            Assert.AreEqual(originalSheet.Name, copiedSheet.Name);
            Assert.AreEqual(originalSheet.Rows.Count, copiedSheet.Rows.Count);
            Assert.AreEqual(originalSheet.Rows[0].PinName, copiedSheet.Rows[0].PinName);
        }

        [TestMethod]
        public void LevelSheet_AddPowerPinLevel()
        {
            // Arrange
            var levelSheet = new LevelSheet("Level");
            var powerLevel = new PowerLevel("VDD", "3.3", "1.8", "1.5", "0.1", "Power pin");

            // Act
            levelSheet.AddPowerPinLevel(powerLevel);

            // Assert
            Assert.AreEqual(4, levelSheet.Rows.Count);
            Assert.AreEqual("VDD", levelSheet.Rows[0].PinName);
            Assert.AreEqual("Vmain", levelSheet.Rows[0].Parameter);
            Assert.AreEqual("3.3", levelSheet.Rows[0].Value);
        }

        [TestMethod]
        public void LevelSheet_AddIoPinLevel()
        {
            // Arrange
            var levelSheet = new LevelSheet("Level");
            var ioLevel = new IoLevel("IO_1", "0.4", "1.4", "0.1", "2.9", "2.95", "2.97", "10mA", "20mA", "1.4", "0.5", "2.0", "0.05", "2.9", "CMOS");

            // Act
            levelSheet.AddIoPinLevel(ioLevel);

            // Assert
            Assert.IsTrue(levelSheet.Rows.Count >= 2);
            Assert.AreEqual("IO_1", levelSheet.Rows[0].PinName);
            Assert.AreEqual("vil", levelSheet.Rows[0].Parameter);
        }

        [TestMethod]
        public void LevelSheet_Errors_InitializedEmpty()
        {
            // Arrange & Act
            var levelSheet = new LevelSheet("Level");

            // Assert
            Assert.IsNotNull(levelSheet.GetErrors());
            Assert.AreEqual(0, levelSheet.GetErrors().Count);
        }

        [TestMethod]
        public void LevelSheet_SheetType_IsCorrect()
        {
            // Arrange & Act
            var levelSheet = new LevelSheet("Level");

            // Assert
            Assert.AreEqual("DTLevelSheet", levelSheet.SheetType);
        }

        [TestMethod]
        public void LevelSheet_AddDiffLevel()
        {
            // Arrange
            var levelSheet = new LevelSheet("Level");

            // Instantiate using your string-based constructor with unique mock values
            var diffLevel = new DiffLevel(
                "DIFF_PIN_1", // pinName
                "str_vicm",   // vicm
                "str_vid",    // vid
                "str_dvid0",  // dVid0
                "str_dvid1",  // dVid1
                "str_dvicm0", // dVicm0
                "str_dvicm1", // dVicm1
                "str_vod",    // vod
                "str_vod1",   // vodAlt1
                "str_vod2",   // vodAlt2
                "str_dvod0",  // dVod0
                "str_dvod1",  // dVod1
                "str_iol",    // iol
                "str_ioh",    // ioh
                "str_vodtyp", // vodTyp
                "str_vocm",   // vocmTyp
                "str_vt",     // vt
                "str_vcl",    // vcl
                "str_vch",    // vch
                "str_mode"    // driverMode
            );

            // Act
            levelSheet.AddDiffLevel(diffLevel);

            // Assert
            Assert.AreEqual(19, levelSheet.Rows.Count, "AddDiffLevel should create exactly 19 LevelRow objects.");

            // Validate each row's internal parameters and values sequentially
            AssertDiffRow(levelSheet.Rows[0], "DIFF_PIN_1", "Vicm", "str_vicm");
            AssertDiffRow(levelSheet.Rows[1], "DIFF_PIN_1", "Vid", "str_vid");
            AssertDiffRow(levelSheet.Rows[2], "DIFF_PIN_1", "dVid0", "str_dvid0");
            AssertDiffRow(levelSheet.Rows[3], "DIFF_PIN_1", "dVid1", "str_dvid1");
            AssertDiffRow(levelSheet.Rows[4], "DIFF_PIN_1", "dVicm0", "str_dvicm0");
            AssertDiffRow(levelSheet.Rows[5], "DIFF_PIN_1", "dVicm1", "str_dvicm1");
            AssertDiffRow(levelSheet.Rows[6], "DIFF_PIN_1", "Vod", "str_vod");
            AssertDiffRow(levelSheet.Rows[7], "DIFF_PIN_1", "Vod_Alt1", "str_vod1");
            AssertDiffRow(levelSheet.Rows[8], "DIFF_PIN_1", "Vod_Alt2", "str_vod2");
            AssertDiffRow(levelSheet.Rows[9], "DIFF_PIN_1", "dVod0", "str_dvod0");
            AssertDiffRow(levelSheet.Rows[10], "DIFF_PIN_1", "dVod1", "str_dvod1");
            AssertDiffRow(levelSheet.Rows[11], "DIFF_PIN_1", "Iol", "str_iol");
            AssertDiffRow(levelSheet.Rows[12], "DIFF_PIN_1", "Ioh", "str_ioh");
            AssertDiffRow(levelSheet.Rows[13], "DIFF_PIN_1", "VodTyp", "str_vodtyp");
            AssertDiffRow(levelSheet.Rows[14], "DIFF_PIN_1", "VocmTyp", "str_vocm");
            AssertDiffRow(levelSheet.Rows[15], "DIFF_PIN_1", "Vt", "str_vt");
            AssertDiffRow(levelSheet.Rows[16], "DIFF_PIN_1", "Vcl", "str_vcl");
            AssertDiffRow(levelSheet.Rows[17], "DIFF_PIN_1", "Vch", "str_vch");
            AssertDiffRow(levelSheet.Rows[18], "DIFF_PIN_1", "DriverMode", "str_mode");
        }

        private static void AssertDiffRow(LevelRow levelRow, string expectedPin, string expectedParameter, string expectedValue)
        {
            Assert.AreEqual(expectedPin, levelRow.PinName, $"PinName mismatch on '{expectedParameter}'.");
            Assert.AreEqual(expectedParameter, levelRow.Parameter, $"Parameter type string mismatch.");
            Assert.AreEqual(expectedValue, levelRow.Value, $"Value mapping mismatch on '{expectedParameter}'.");
            // Depending on your constructor signature, the 4th argument maps to Comment or Extra
            Assert.AreEqual("", levelRow.Comment, $"Comment field on '{expectedParameter}' must be empty.");
        }

        [TestMethod]
        public void LevelSheet_AddPowerPinVSlewLevel()
        {
            // Arrange
            var levelSheet = new LevelSheet("Level");

            // Create a mock PowerLevel object with custom values including VSlewRate and Comment
            // (Assuming constructor matches: pinName, vmain, etc., with your custom fields)
            var powerLevel = new PowerLevel("VDD_CORE", "3.3", "1.8", "1.5", "0.1", "Slew test comment")
            {
                VSlewRate = "5.5"
            };

            // Act
            levelSheet.AddPowerPinVSlewLevel(powerLevel);

            // Assert - AddPowerPinLevel adds 4 rows, plus 1 row for VSlewRate = 5 rows total
            Assert.AreEqual(5, levelSheet.Rows.Count, "Should add exactly 5 rows (4 from base power pin + 1 for VSlewRate).");

            // Verify the baseline power row property matches your existing tests
            Assert.AreEqual("VDD_CORE", levelSheet.Rows[0].PinName);
            Assert.AreEqual("Vmain", levelSheet.Rows[0].Parameter);
            Assert.AreEqual("3.3", levelSheet.Rows[0].Value);

            // Verify the final appended VSlewRate row explicitly
            LevelRow slewRow = levelSheet.Rows[4];
            Assert.AreEqual("VDD_CORE", slewRow.PinName);
            Assert.AreEqual("VSlewRate", slewRow.Parameter);
            Assert.AreEqual("5.5", slewRow.Value);
            Assert.AreEqual("Slew test comment", slewRow.Comment);
        }

        [TestMethod]
        public void LevelSheet_AddDcviPowerPinLevel()
        {
            // Arrange
            var levelSheet = new LevelSheet("Level");

            // Create mock data for the DCVI power level profile
            var dcviPowerLevel = new DcviPowerLevel(
                "VDD_DCVI",          // pinName
                "1.2",               // vps
                "0.05",              // isc
                "0.002",             // tDelay
                "DCVI tracking test" // comment
            );

            // Act
            levelSheet.AddDcviPowerPinLevel(dcviPowerLevel);

            // Assert - Method adds exactly 3 specific level mapping rows
            Assert.AreEqual(3, levelSheet.Rows.Count, "Should append exactly 3 target LevelRow items.");

            // Row 1 validation: Vps
            Assert.AreEqual("VDD_DCVI", levelSheet.Rows[0].PinName);
            Assert.AreEqual("Vps", levelSheet.Rows[0].Parameter);
            Assert.AreEqual("1.2", levelSheet.Rows[0].Value);
            Assert.AreEqual("DCVI tracking test", levelSheet.Rows[0].Comment);

            // Row 2 validation: Isc
            Assert.AreEqual("VDD_DCVI", levelSheet.Rows[1].PinName);
            Assert.AreEqual("Isc", levelSheet.Rows[1].Parameter);
            Assert.AreEqual("0.05", levelSheet.Rows[1].Value);
            Assert.AreEqual("DCVI tracking test", levelSheet.Rows[1].Comment);

            // Row 3 validation: tdelay
            Assert.AreEqual("VDD_DCVI", levelSheet.Rows[2].PinName);
            Assert.AreEqual("tdelay", levelSheet.Rows[2].Parameter);
            Assert.AreEqual("0.002", levelSheet.Rows[2].Value);
            Assert.AreEqual("DCVI tracking test", levelSheet.Rows[2].Comment);
        }
    }
}
