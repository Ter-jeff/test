using Automation.PreCheck.PreChecks.Basic;

using CommonLib.ErrorReport;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace Automation.Test.UT.Basic
{
    [TestClass]
    public class BasicPreCheckPowerInfoTests
    {
        [TestInitialize]
        public void Init()
        {
            ErrorReportManager.ClearErrors();
        }

        // CheckExist() populates StartRow/StartColumn/etc.; CheckHeaders()/CheckFormat() are
        // exercised directly to avoid CheckBusiness(), which reaches the
        // MultiTestSettingSheetsSingleton and requires a fully configured real workbook.
        private static void RunCheckExistThenHeaders(BasicPreCheckPowerInfo preCheck, out bool existResult, out bool headersResult)
        {
            existResult = preCheck.CheckExist();
            headersResult = preCheck.CheckHeaders();
        }

        [TestMethod]
        public void CheckHeaders_MissingPowerSequenceColumn_ReturnsFalseWithExistentialError()
        {
            // Arrange - only "PinName" is present, no "PowerSequence"/"PowerUpSequence" column
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("PowerInfo");
            sheet.Cells[1, 1].Value = "PinName";
            sheet.Cells[2, 1].Value = "PIN_A";

            var preCheck = new BasicPreCheckPowerInfo(package.Workbook, "PowerInfo");

            // Act
            RunCheckExistThenHeaders(preCheck, out bool existResult, out bool headersResult);

            // Assert
            Assert.IsTrue(existResult);
            Assert.IsFalse(headersResult);
            Assert.IsTrue(ErrorReportManager.GetErrorList().Exists(e => e.Message.Contains("Omit column \"PowerSequence\"")));
        }

        [TestMethod]
        public void CheckHeaders_ValidPowerSequenceColumn_ReturnsTrueWithoutErrors()
        {
            // Arrange
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("PowerInfo");
            sheet.Cells[1, 1].Value = "PinName";
            sheet.Cells[1, 2].Value = "PowerSequence";
            sheet.Cells[2, 1].Value = "PIN_A";
            sheet.Cells[2, 2].Value = "1";

            var preCheck = new BasicPreCheckPowerInfo(package.Workbook, "PowerInfo");

            // Act
            RunCheckExistThenHeaders(preCheck, out bool existResult, out bool headersResult);

            // Assert
            Assert.IsTrue(existResult);
            Assert.IsTrue(headersResult);
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckHeaders_PowerUpSequenceHeaderVariant_SatisfiesHeaderCheck()
        {
            // Arrange - "PowerUpSequence" is an accepted alternative to "PowerSequence"
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("PowerInfo");
            sheet.Cells[1, 1].Value = "PinName";
            sheet.Cells[1, 2].Value = "PowerUpSequence";
            sheet.Cells[2, 1].Value = "PIN_A";
            sheet.Cells[2, 2].Value = "1";

            var preCheck = new BasicPreCheckPowerInfo(package.Workbook, "PowerInfo");

            // Act
            RunCheckExistThenHeaders(preCheck, out bool existResult, out bool headersResult);

            // Assert
            Assert.IsTrue(existResult);
            Assert.IsTrue(headersResult);
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckFormat_MissingRequiredColumnValue_AddsWarningAndReturnsFalse()
        {
            // Arrange - PowerSequence column present in the header but blank for the data row
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("PowerInfo");
            sheet.Cells[1, 1].Value = "PinName";
            sheet.Cells[1, 2].Value = "PowerSequence";
            sheet.Cells[2, 1].Value = "PIN_A";

            var preCheck = new BasicPreCheckPowerInfo(package.Workbook, "PowerInfo");
            RunCheckExistThenHeaders(preCheck, out _, out bool headersResult);
            Assert.IsTrue(headersResult);

            // Act
            bool formatResult = preCheck.CheckFormat();

            // Assert
            Assert.IsFalse(formatResult);
            Assert.IsTrue(ErrorReportManager.GetErrorList().Exists(e => e.Message.Contains("Omit \"PowerSequence\" value for pin: PIN_A")));
        }

        [TestMethod]
        public void CheckFormat_AllRequiredColumnValuesPresent_ReturnsTrueWithoutWarning()
        {
            // Arrange
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("PowerInfo");
            sheet.Cells[1, 1].Value = "PinName";
            sheet.Cells[1, 2].Value = "PowerSequence";
            sheet.Cells[2, 1].Value = "PIN_A";
            sheet.Cells[2, 2].Value = "1";

            var preCheck = new BasicPreCheckPowerInfo(package.Workbook, "PowerInfo");
            RunCheckExistThenHeaders(preCheck, out _, out bool headersResult);
            Assert.IsTrue(headersResult);

            // Act
            bool formatResult = preCheck.CheckFormat();

            // Assert
            Assert.IsTrue(formatResult);
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }
    }
}
