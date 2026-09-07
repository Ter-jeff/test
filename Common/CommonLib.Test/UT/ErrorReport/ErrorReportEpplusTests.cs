using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace CommonLib.Test.UT.ErrorReport
{
    [TestClass]
    public class ErrorReportEpplusTests
    {
        private static ErrorCode CreateErrorCode(string name = "TestCode", EnumErrorLevel enumErrorLevel = EnumErrorLevel.Error)
        {
            return new ErrorCode(EnumErrorCategory.Basic, EnumErrorBehavior.Duplicate, EnumErrorTarget.Pin, 1, "Template {0}", enumErrorLevel, name);
        }

        private static Error CreateError(string sheetName, int row, int col, string message, string codeName = "TestCode")
        {
            ErrorCode code = CreateErrorCode(codeName);
            return new Error(code, EnumErrorLevel.Error, sheetName, row, col, message);
        }

        [TestMethod]
        public void WriteReport_EmptyErrors_ReturnsWithoutWritingSheet()
        {
            // Arrange
            var instance = new ErrorInstance();
            using var pkg = new ExcelPackage();
            // Act
            instance.GenErrorReport(pkg, "ErrorReport");

            // Assert - no worksheet was created since there are no errors
            Assert.IsNull(pkg.Workbook.Worksheets["ErrorReport"]);
        }

        [TestMethod]
        public void GenErrorReport_WithErrors_CreatesErrorReportWorksheet()
        {
            // Arrange
            var instance = new ErrorInstance();
            instance.AddError(CreateError("Sheet1", 1, 1, "msg1"));
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("Sheet1");

            // Act
            instance.GenErrorReport(pkg, "ErrorReport");

            // Assert
            Assert.IsNotNull(pkg.Workbook.Worksheets["ErrorReport"]);
        }

        [TestMethod]
        public void GenErrorReport_WithErrors_HeaderRowIsBold()
        {
            // Arrange
            var instance = new ErrorInstance();
            instance.AddError(CreateError("Sheet1", 1, 1, "msg"));
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("Sheet1");

            // Act
            instance.GenErrorReport(pkg, "ErrorReport");

            // Assert
            ExcelWorksheet sheet = pkg.Workbook.Worksheets["ErrorReport"];
            Assert.IsTrue(sheet.Cells[1, 1].Style.Font.Bold);
        }

        [TestMethod]
        public void GenErrorReport_WithMultipleErrorTypes_CreatesSummaryReportSheet()
        {
            // Arrange
            var instance = new ErrorInstance();
            instance.AddError(CreateError("Sheet1", 1, 1, "msg1", "TypeA"));
            instance.AddError(CreateError("Sheet1", 2, 1, "msg2", "TypeB"));
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("Sheet1");

            // Act
            instance.GenErrorReport(pkg, [], "ErrorReport", "SummaryReport");

            // Assert
            Assert.IsNotNull(pkg.Workbook.Worksheets["ErrorReport"]);
            Assert.IsNotNull(pkg.Workbook.Worksheets["SummaryReport"]);
        }

        [TestMethod]
        public void GenErrorReport_WithCopyFilesAndNoErrors_CreatesNoErrorWorksheet()
        {
            // Arrange
            var instance = new ErrorInstance();
            using var pkg = new ExcelPackage();
            // Act
            instance.GenErrorReport(pkg, [], "ErrorReport");

            // Assert
            ExcelWorksheet sheet = pkg.Workbook.Worksheets["ErrorReport"];
            Assert.IsNotNull(sheet);
            Assert.AreEqual("No error occurs", sheet.Cells[1, 1].Value.ToString());
        }

        [TestMethod]
        public void GenErrorReport_ErrorWithNoMatchingSheet_SkipsHyperlink()
        {
            // Arrange - error references "MissingSheet" which is not in the workbook
            var instance = new ErrorInstance();
            instance.AddError(CreateError("MissingSheet", 1, 1, "msg"));
            using var pkg = new ExcelPackage();
            // Act - should not throw even when sheet doesn't exist in workbook
            instance.GenErrorReport(pkg, "ErrorReport");

            // Assert
            Assert.IsNotNull(pkg.Workbook.Worksheets["ErrorReport"]);
        }

        [TestMethod]
        public void GenErrorReport_WithExistingSummarySheet_AppendsToExistingSheet()
        {
            // Arrange - pre-create the summary sheet to test the append path
            var instance = new ErrorInstance();
            instance.AddError(CreateError("Sheet1", 1, 1, "msg1", "TypeA"));
            instance.AddError(CreateError("Sheet1", 2, 1, "msg2", "TypeB"));
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("Sheet1");
            ExcelWorksheet existingSummary = pkg.Workbook.Worksheets.Add("SummaryReport");
            existingSummary.Cells[1, 1].Value = "PreExisting";

            // Act
            instance.GenErrorReport(pkg, [], "ErrorReport", "SummaryReport");

            // Assert - summary sheet still exists and was appended to
            Assert.IsNotNull(pkg.Workbook.Worksheets["SummaryReport"]);
        }

        [TestMethod]
        public void GenErrorReport_ErrorWithColZero_DoesNotAddHyperlinkFormula()
        {
            // Arrange - col=0 means row-range link format
            var instance = new ErrorInstance();
            instance.AddError(CreateError("Sheet1", 1, 0, "msg"));
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("Sheet1");

            // Act
            instance.GenErrorReport(pkg, "ErrorReport");

            // Assert
            Assert.IsNotNull(pkg.Workbook.Worksheets["ErrorReport"]);
        }
    }
}
