using Automation.GenerateIgxl.HardIp.HardIpPreCheck;

using CommonLib.ErrorReport;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class WrongHardIpSheetNameCheckerTests
    {
        private readonly WrongHardIpSheetNameChecker _checker = new();

        [TestInitialize]
        public void Setup()
        {
            ErrorReportManager.ClearErrors();
        }

        [TestCleanup]
        public void Cleanup()
        {
            ErrorReportManager.ClearErrors();
        }

        private static void AddHeaderRow(ExcelWorksheet sheet, params string[] headers)
        {
            for (int i = 0; i < headers.Length; i++)
            {
                sheet.Cells[1, i + 1].Value = headers[i];
            }
        }

        [TestMethod]
        public void CheckWrongHardIpSheetName_HardIpPrefixedSheet_SkippedEvenWithManyHeaders()
        {
            // Arrange
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("HARDIP_Test1");
            AddHeaderRow(sheet, "TTR", "Part", "NoBinout", "FailFlag", "Step", "Description", "Condition", "Register", "Comment", "Test Name");

            // Act
            _checker.CheckWrongHardIpSheetName(package.Workbook);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckWrongHardIpSheetName_DcTestPrefixedSheet_Skipped()
        {
            // Arrange
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("DCTEST_Test1");
            AddHeaderRow(sheet, "TTR", "Part", "NoBinout", "FailFlag", "Step", "Description", "Condition", "Register", "Comment", "Test Name");

            // Act
            _checker.CheckWrongHardIpSheetName(package.Workbook);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckWrongHardIpSheetName_RtosPrefixedSheet_Skipped()
        {
            // Arrange
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("RTOS_Test1");
            AddHeaderRow(sheet, "TTR", "Part", "NoBinout", "FailFlag", "Step", "Description", "Condition", "Register", "Comment", "Test Name");

            // Act
            _checker.CheckWrongHardIpSheetName(package.Workbook);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckWrongHardIpSheetName_WirelessPrefixedSheet_Skipped()
        {
            // Arrange
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Wireless_Test1");
            AddHeaderRow(sheet, "TTR", "Part", "NoBinout", "FailFlag", "Step", "Description", "Condition", "Register", "Comment", "Test Name");

            // Act
            _checker.CheckWrongHardIpSheetName(package.Workbook);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckWrongHardIpSheetName_LcdPrefixedSheet_Skipped()
        {
            // Arrange
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("LCD_Test1");
            AddHeaderRow(sheet, "TTR", "Part", "NoBinout", "FailFlag", "Step", "Description", "Condition", "Register", "Comment", "Test Name");

            // Act
            _checker.CheckWrongHardIpSheetName(package.Workbook);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckWrongHardIpSheetName_NonPrefixedSheetWithManyKnownHeaders_AddsError()
        {
            // Arrange - 10 of the 18 known headers match (10*2=20 > 18), and the sheet name
            // doesn't start with any of the recognized HardIP prefixes
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("SomeSheet");
            AddHeaderRow(sheet, "TTR", "Part", "NoBinout", "FailFlag", "Step", "Description", "Condition", "Register", "Comment", "Test Name");

            // Act
            _checker.CheckWrongHardIpSheetName(package.Workbook);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckWrongHardIpSheetName_NonPrefixedSheetWithFewKnownHeaders_NoError()
        {
            // Arrange - no headers match any known pattern
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("SomeSheet");
            AddHeaderRow(sheet, "Foo1", "Foo2", "Foo3");

            // Act
            _checker.CheckWrongHardIpSheetName(package.Workbook);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckWrongHardIpSheetName_ExactlyHalfHeadersMatch_NoErrorAtBoundary()
        {
            // Arrange - 9 of 18 known headers match (9*2=18, not strictly greater than 18)
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("SomeSheet");
            AddHeaderRow(sheet, "TTR", "Part", "NoBinout", "FailFlag", "Step", "Description", "Condition", "Register", "Comment");

            // Act
            _checker.CheckWrongHardIpSheetName(package.Workbook);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckWrongHardIpSheetName_EmptySheetWithNoDimension_NoError()
        {
            // Arrange
            using var package = new ExcelPackage();
            package.Workbook.Worksheets.Add("SomeSheet");

            // Act
            _checker.CheckWrongHardIpSheetName(package.Workbook);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }
    }
}
