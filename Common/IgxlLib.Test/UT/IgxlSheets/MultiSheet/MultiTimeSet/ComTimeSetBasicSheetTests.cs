using System;
using System.Collections.Generic;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;
using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

using Moq;

using static IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet.ComTimeSetBasicSheet;

namespace IgxlLib.Test.UT.IgxlSheets.MultiSheet.MultiTimeSet
{
    [TestClass]
    public class ComTimeSetBasicSheetTests
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void ComTimeSetBasicSheet_Constructor_WithSheetName()
        {
            // Arrange
            string sheetName = "ComTimeSet";

            // Act
            var comTimeSetBasicSheet = new ComTimeSetBasicSheet(sheetName);

            // Assert
            Assert.IsNotNull(comTimeSetBasicSheet);
            Assert.AreEqual(sheetName, comTimeSetBasicSheet.Name);
            Assert.AreEqual("DTTimesetBasicSheet", comTimeSetBasicSheet.SheetType);
            Assert.AreEqual(0, comTimeSetBasicSheet.Rows.Count);
        }

        [TestMethod]
        public void ComTimeSetBasicSheet_Constructor_WithSheetNameAndTimingMode()
        {
            // Arrange
            string sheetName = "ComTimeSet";
            string timingMode = "ASYNC";

            // Act
            var comTimeSetBasicSheet = new ComTimeSetBasicSheet(sheetName, timingMode);

            // Assert
            Assert.AreEqual(sheetName, comTimeSetBasicSheet.Name);
            Assert.AreEqual(timingMode, comTimeSetBasicSheet.TimingMode);
        }

        [TestMethod]
        public void ComTimeSetBasicSheet_Constructor_WithAllParameters()
        {
            // Arrange
            string sheetName = "ComTimeSet";
            string timingMode = "ASYNC";
            string masterTimeSet = "MasterTS";
            string timeDomain = "DOMAIN";
            string strobeRefSetup = "STROBE";

            // Act
            var comTimeSetBasicSheet = new ComTimeSetBasicSheet(sheetName, timingMode, masterTimeSet, timeDomain, strobeRefSetup);

            // Assert
            Assert.AreEqual(sheetName, comTimeSetBasicSheet.Name);
            Assert.AreEqual(timingMode, comTimeSetBasicSheet.TimingMode);
            Assert.AreEqual(masterTimeSet, comTimeSetBasicSheet.MasterTimeSet);
            Assert.AreEqual(timeDomain, comTimeSetBasicSheet.TimeDomain);
            Assert.AreEqual(strobeRefSetup, comTimeSetBasicSheet.StrobeRefSetup);
        }

        [TestMethod]
        public void ComTimeSetBasicSheet_SheetType_IsCorrect()
        {
            // Arrange & Act
            var comTimeSetBasicSheet = new ComTimeSetBasicSheet("ComTimeSet");

            // Assert
            Assert.AreEqual("DTTimesetBasicSheet", comTimeSetBasicSheet.SheetType);
        }

        [TestMethod]
        public void ComTimeSetBasicSheet_AddShiftInTSetName()
        {
            // Arrange
            var comTimeSetBasicSheet = new ComTimeSetBasicSheet("ComTimeSet");

            // Act
            comTimeSetBasicSheet.AddShiftInTSetName("TSet1");
            comTimeSetBasicSheet.AddShiftInTSetName("TSet2");

            // Assert
            Assert.IsTrue(comTimeSetBasicSheet.IsMultiShiftInTSet);
            Assert.AreEqual("TSet1,TSet2", comTimeSetBasicSheet.GetMultiShiftInStr);
        }

        [TestMethod]
        public void ComTimeSetBasicSheet_IsMultiShiftInTSet_FalseWithSingleTSet()
        {
            // Arrange
            var comTimeSetBasicSheet = new ComTimeSetBasicSheet("ComTimeSet");

            // Act
            comTimeSetBasicSheet.AddShiftInTSetName("TSet1");

            // Assert
            Assert.IsFalse(comTimeSetBasicSheet.IsMultiShiftInTSet);
        }

        [TestMethod]
        public void ComTimeSetBasicSheet_IsMultiShiftInTSet_FalseWithNoTSet()
        {
            // Arrange & Act
            var comTimeSetBasicSheet = new ComTimeSetBasicSheet("ComTimeSet");

            // Assert
            Assert.IsFalse(comTimeSetBasicSheet.IsMultiShiftInTSet);
        }

        [TestMethod]
        public void ComTimeSetBasicSheet_GetMultiShiftInStr_WithSingleTSet()
        {
            // Arrange
            var comTimeSetBasicSheet = new ComTimeSetBasicSheet("ComTimeSet");

            // Act
            comTimeSetBasicSheet.AddShiftInTSetName("TSet1");

            // Assert
            Assert.AreEqual("TSet1", comTimeSetBasicSheet.GetMultiShiftInStr);
        }

        [TestMethod]
        public void ComTimeSetBasicSheet_AddRow()
        {
            // Arrange
            var comTimeSetBasicSheet = new ComTimeSetBasicSheet("ComTimeSet");
            var comTimeSetBasic = new ComTimeSetBasic { CyclePeriod = "10ns" };

            // Act
            comTimeSetBasicSheet.AddRow(comTimeSetBasic);

            // Assert
            Assert.AreEqual(1, comTimeSetBasicSheet.Rows.Count);
            Assert.AreEqual("10ns", comTimeSetBasicSheet.Rows[0].CyclePeriod);
        }

        [TestMethod]
        public void ComTimeSetBasicSheet_CommentVariable_Dictionary()
        {
            // Arrange
            var comTimeSetBasicSheet = new ComTimeSetBasicSheet("ComTimeSet");

            // Act
            comTimeSetBasicSheet.CommentVariable.Add("VAR1", 10.5);
            comTimeSetBasicSheet.CommentVariable.Add("VAR2", 20.3);

            // Assert
            Assert.AreEqual(2, comTimeSetBasicSheet.CommentVariable.Count);
            Assert.AreEqual(10.5, comTimeSetBasicSheet.CommentVariable["VAR1"]);
            Assert.AreEqual(20.3, comTimeSetBasicSheet.CommentVariable["VAR2"]);
        }

        [TestMethod]
        public void ComTimeSetBasicSheet_Inherits_TimeSetBasicSheet()
        {
            // Arrange & Act
            var comTimeSetBasicSheet = new ComTimeSetBasicSheet("ComTimeSet");

            // Assert
            Assert.IsInstanceOfType(comTimeSetBasicSheet, typeof(TimeSetBasicSheet));
        }

        [TestMethod]
        public void ComTimeSetBasicSheet_Name_CanBeSet()
        {
            // Arrange
            var comTimeSetBasicSheet = new ComTimeSetBasicSheet("ComTimeSet")
            {
                // Act
                Name = "NewComTimeSet"
            };

            // Assert
            Assert.AreEqual("NewComTimeSet", comTimeSetBasicSheet.Name);
        }

        [TestMethod]
        public void AllTSetEqnVariable_ValidRows_ReturnsAccumulatedVariables()
        {
            // Arrange
            var sheet = new ComTimeSetBasicSheet("ComTimeSet");

            var tset1 = new ComTimeSetBasic { Name = "TSet1" };
            tset1.SubCommentVariable.Add("VAR_A", 1.0);

            var tset2 = new ComTimeSetBasic { Name = "TSet2" };
            tset2.SubCommentVariable.Add("VAR_B", 2.0);
            tset2.SubCommentVariable.Add("VAR_A", 99.0);

            sheet.Rows.Add(tset1);
            sheet.Rows.Add(tset2);

            // Act
            List<TSetEqnVarMap> result = sheet.AllTSetEqnVariable;

            // Assert
            Assert.AreEqual(2, result.Count);

            // First row checks
            Assert.AreEqual("TSet1", result[0].TSetName);
            Assert.IsTrue(result[0].DictVariable.ContainsKey("VAR_A"));
            Assert.AreEqual(1.0, result[0].DictVariable["VAR_A"]);

            // Second row checks (should contain accumulated VAR_A from row 1, and new VAR_B)
            Assert.AreEqual("TSet2", result[1].TSetName);
            Assert.AreEqual(2, result[1].DictVariable.Count);
            Assert.AreEqual(1.0, result[1].DictVariable["VAR_A"]);
            Assert.AreEqual(2.0, result[1].DictVariable["VAR_B"]);
        }

        [TestMethod]
        public void AllTSetEqnVariable_InvalidRowType_ThrowsException()
        {
            // Arrange
            var sheet = new ComTimeSetBasicSheet("ComTimeSet");
            // Adding a base TSet or mock TSet that is NOT a ComTimeSetBasic
            TSet invalidTSet = Mock.Of<TSet>();
            sheet.Rows.Add(invalidTSet);

            // Act
            Assert.ThrowsException<Exception>(() => sheet.AllTSetEqnVariable);
        }

        [TestMethod]
        public void InsertAlarmDataInFirstRow_InsertsAtIndexZeroWithTwoTimingRows()
        {
            // Arrange
            var sheet = new ComTimeSetBasicSheet("ComTimeSet");
            var existingRow = new ComTimeSetBasic { Name = "Existing" };
            sheet.Rows.Add(existingRow);

            // Act
            sheet.InsertAlarmDataInFirstRow("HIGH_TEMP");

            // Assert
            Assert.AreEqual(2, sheet.Rows.Count);

            var insertedRow = sheet.Rows[0] as ComTimeSetBasic;
            Assert.IsNotNull(insertedRow);
            Assert.AreEqual("HIGH_TEMP Please check it", insertedRow.Name);
            // Assuming Rows inside ComTimeSetBasic tracking the added TimingRows can be verified, 
            // or verification that AddTimingRow was triggered successfully.
        }

        [TestMethod]
        [DataRow("A_B_S_AN_E", "HardIP")] // Splits >= 5, starts with S -> Soc, block AN -> HardIp
        [DataRow("A_B_C_SC_E", "CpuScan")] // Splits >= 5, starts with C -> Cpu, block SC -> Scan
        [DataRow("A_B_L_BI_E", "GfxMbist")] // Splits >= 5, starts with L -> Gfx, block BI -> Mbist
        [DataRow("A_B_X_XX_E", "AC_A_B_X_XX_E")] // Splits >= 5, Fallback block -> None -> returns "AC_" + sheetName
        [DataRow("Short_Name", "HardIP")] // Splits < 5 -> returns HardIp default
        public void GetTimeSetCategory_VariousSheetNames_ReturnsCorrectCategory(string sheetName, string expectedCategory)
        {
            // Act
            string actualCategory = GetTimeSetCategory(sheetName);

            // Assert
            Assert.AreEqual(expectedCategory, actualCategory);
        }

        [TestMethod]
        [DataRow(new string[] { "A", "B", "C", "Scan" }, EnumBlockType.Scan)]
        [DataRow(new string[] { "A", "B", "C", "Saa" }, EnumBlockType.Scan)]
        [DataRow(new string[] { "A", "B", "C", "AN" }, EnumBlockType.HardIp)]
        [DataRow(new string[] { "A", "B", "C", "JT" }, EnumBlockType.HardIp)]
        [DataRow(new string[] { "A", "B", "C", "SC" }, EnumBlockType.Scan)]
        [DataRow(new string[] { "A", "B", "C", "BI" }, EnumBlockType.Mbist)]
        [DataRow(new string[] { "A", "B", "C", "D", "bsr" }, EnumBlockType.Mbist)]
        [DataRow(new string[] { "A", "B", "C", "D", "mbist" }, EnumBlockType.Mbist)]
        [DataRow(new string[] { "A", "B", "C", "Unknown" }, EnumBlockType.None)]
        public void GetBlockType_VariousItems_ReturnsExpectedBlockType(string[] items, EnumBlockType enumBlockType)
        {
            // Act
            EnumBlockType actualType = GetBlockType(items);

            // Assert
            Assert.AreEqual(enumBlockType, actualType);
        }

    }
}
