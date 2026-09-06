using System.Collections.Generic;
using System.IO;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Static;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace CommonLib.Test.UT.ErrorReport
{
    [TestClass]
    public class ErrorReportManagerTests
    {
        private bool _originalStopByCritical;

        [TestInitialize]
        public void Setup()
        {
            _originalStopByCritical = ErrorReportStatic.StopByCritical;
            ErrorReportStatic.StopByCritical = false;
            ErrorReportManager.ClearErrors();
        }

        [TestCleanup]
        public void Teardown()
        {
            ErrorReportStatic.StopByCritical = _originalStopByCritical;
            ErrorReportManager.ClearErrors();
        }

        [TestMethod]
        public void GetErrorCount_AfterClear_ReturnsZero()
        {
            // Arrange & Act
            int count = ErrorReportManager.GetErrorCount();

            // Assert
            Assert.AreEqual(0, count);
        }

        [TestMethod]
        public void AddError_WithMessage_IncreasesCount()
        {
            // Arrange
            ErrorCode code = BasicErrorType.E_MissingPin_01;

            // Act
            ErrorReportManager.AddError(code, "Sheet1", 1, 1, "test");

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void AddError_WithMessageArgs_IncreasesCount()
        {
            // Arrange
            ErrorCode code = BasicErrorType.E_MissingPin_01;

            // Act
            ErrorReportManager.AddError(code, "Sheet1", 1, 1, "msg", ["ArgValue"]);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void AddError_NegativeColumnNum_NormalizesToZero()
        {
            // Arrange
            ErrorCode code = BasicErrorType.E_MissingPin_01;

            // Act
            ErrorReportManager.AddError(code, "Sheet1", 1, -5, "msg");
            List<Error> errors = ErrorReportManager.GetErrorList();

            // Assert
            Assert.AreEqual(0, errors[0].ColNum);
        }

        [TestMethod]
        public void Contains_MessageAdded_ReturnsTrue()
        {
            // Arrange
            ErrorReportManager.AddError(BasicErrorType.E_MissingPin_01, "Sheet1", 1, 1, "unique message");

            // Act
            bool result = ErrorReportManager.Contains("unique message");

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void Contains_MessageAddedDifferentCase_ReturnsTrue()
        {
            // Arrange
            ErrorReportManager.AddError(BasicErrorType.E_MissingPin_01, "Sheet1", 1, 1, "Case Message");

            // Act
            bool result = ErrorReportManager.Contains("case message");

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void Contains_MessageNotAdded_ReturnsFalse()
        {
            // Arrange & Act
            bool result = ErrorReportManager.Contains("not added");

            // Assert
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void ClearErrors_RemovesAllErrors()
        {
            // Arrange
            ErrorReportManager.AddError(BasicErrorType.E_MissingPin_01, "Sheet1", 1, 1, "msg");

            // Act
            ErrorReportManager.ClearErrors();

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void GetErrorCountByCategory_MatchingCategory_ReturnsCount()
        {
            // Arrange - E_Business_01 and E_MissingPatternFile_01 have different names
            ErrorReportManager.AddError(BasicErrorType.E_MissingPin_01, "Sheet1", 1, 1, "msg1");
            ErrorReportManager.AddError(BasicErrorType.E_MissingPin_01, "Sheet1", 2, 1, "msg2");
            ErrorReportManager.AddError(PatternMissing.E_MissingPatternFile_01, "Sheet1", 3, 1, "msg3");

            // Act
            int count = ErrorReportManager.GetErrorCountByCategory(EnumErrorCategory.Basic);

            // Assert
            Assert.AreEqual(3, count);
        }

        [TestMethod]
        public void GetErrorCountByName_NoMatch_ReturnsZero()
        {
            // Arrange
            ErrorReportManager.AddError(BasicErrorType.E_MissingPin_01, "Sheet1", 1, 1, "msg");

            // Act
            int count = ErrorReportManager.GetErrorCountByCategory(EnumErrorCategory.PreAction);

            // Assert
            Assert.AreEqual(0, count);
        }

        [TestMethod]
        public void GetErrorList_ReturnsAllAddedErrors()
        {
            // Arrange
            ErrorReportManager.AddError(BasicErrorType.E_MissingPin_01, "Sheet1", 1, 1, "msg1");
            ErrorReportManager.AddError(BasicErrorType.E_MissingPin_02, "Sheet2", 2, 2, "msg2");

            // Act
            List<Error> list = ErrorReportManager.GetErrorList();

            // Assert
            Assert.AreEqual(2, list.Count);
        }

        [TestMethod]
        public void GetSortedErrors_MultipleErrors_ReturnsSortedList()
        {
            // Arrange
            ErrorReportManager.AddError(BasicErrorType.E_MissingPin_01, "SheetZ", 1, 1, "msg");
            ErrorReportManager.AddError(BasicErrorType.E_MissingPin_01, "SheetA", 1, 1, "msg2");

            // Act
            List<Error> sorted = ErrorReportManager.GetSortedErrors();

            // Assert
            Assert.AreEqual("SheetA", sorted[0].SheetName);
            Assert.AreEqual("SheetZ", sorted[1].SheetName);
        }

        [TestMethod]
        public void PrintErrors_WritesHeaderAndErrorLines()
        {
            // Arrange
            string tempFile = Path.GetTempFileName();
            ErrorReportManager.AddError(BasicErrorType.E_MissingPin_01, "Sheet1", 5, 2, "printed msg");

            try
            {
                // Act
                ErrorReportManager.PrintErrors(tempFile);
                string[] lines = File.ReadAllLines(tempFile);

                // Assert
                Assert.IsTrue(lines.Length >= 2, "Expected header + at least one data line");
                StringAssert.Contains(lines[0], "SheetName");
                StringAssert.Contains(lines[1], "Sheet1");
                StringAssert.Contains(lines[1], "printed msg");
            }
            finally
            {
                File.Delete(tempFile);
            }
        }

        [TestMethod]
        public void RemoveNonUsedPattern_NullSet_DoesNotRemoveErrors()
        {
            // Arrange
            ErrorReportManager.AddError(BasicErrorType.E_MissingPin_01, "Sheet1", 1, 1, "keep me");

            // Act
            ErrorReportManager.RemoveNonUsedPattern(null);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void RemoveNonUsedPattern_EmptySet_DoesNotRemoveErrors()
        {
            // Arrange
            ErrorReportManager.AddError(BasicErrorType.E_MissingPin_01, "Sheet1", 1, 1, "keep me");

            // Act
            ErrorReportManager.RemoveNonUsedPattern([]);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void RemoveNonUsedPattern_NonPatternErrorsAlwaysKept()
        {
            // Arrange - non-pattern errors are always kept regardless of usedPatternList
            ErrorReportManager.AddError(BasicErrorType.E_MissingPin_01, "Sheet1", 1, 1, "keep me");

            // Act
            ErrorReportManager.RemoveNonUsedPattern(["some_pattern"]);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void RemoveNonUsedPattern_PatternErrorNotInUsedSet_IsRemoved()
        {
            // Arrange - add a MissingPatternFile error with pattern "unused_pat"
            var errorInfo = new ErrorInfo { Pattern = "unused_pat" };
            ErrorReportManager.AddError(PatternMissing.E_MissingPatternFile_01, "Sheet1", 1, 1, "missing", errorInfo);

            // Act
            ErrorReportManager.RemoveNonUsedPattern(["other_pat"]);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void RemoveNonUsedPattern_PatternErrorInUsedSet_IsKept()
        {
            // Arrange - add a MissingPatternFile error with pattern "used_pat"
            var errorInfo = new ErrorInfo { Pattern = "used_pat" };
            ErrorReportManager.AddError(PatternMissing.E_MissingPatternFile_01, "Sheet1", 1, 1, "missing", errorInfo);

            // Act
            ErrorReportManager.RemoveNonUsedPattern(["used_pat"]);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void RemoveNonUsedPattern_MixedErrors_KeepsNonPatternAndUsedPatternOnly()
        {
            // Arrange
            var usedInfo = new ErrorInfo { Pattern = "used_pat" };
            var unusedInfo = new ErrorInfo { Pattern = "unused_pat" };
            ErrorReportManager.AddError(BasicErrorType.E_MissingPin_01, "Sheet1", 1, 1, "business error");
            ErrorReportManager.AddError(PatternMissing.E_MissingPatternFile_01, "Sheet1", 2, 1, "used", usedInfo);
            ErrorReportManager.AddError(PatternMissing.E_MissingPatternFile_01, "Sheet1", 3, 1, "unused", unusedInfo);

            // Act
            ErrorReportManager.RemoveNonUsedPattern(["used_pat"]);

            // Assert - business error + used pattern error remain; unused pattern error is removed
            Assert.AreEqual(2, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void GenErrorReport_WithExcelAndNoErrors_DoesNotCreateSheet()
        {
            // Arrange
            using var pkg = new ExcelPackage();
            // Act
            ErrorReportManager.GenErrorReport(pkg, "ErrorReport");

            // Assert
            Assert.IsNull(pkg.Workbook.Worksheets["ErrorReport"]);
        }

        [TestMethod]
        public void GenErrorReport_WithExcelAndErrors_CreatesErrorReportSheet()
        {
            // Arrange
            ErrorReportManager.AddError(BasicErrorType.E_MissingPin_01, "Sheet1", 1, 1, "msg");
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("Sheet1");

            // Act
            ErrorReportManager.GenErrorReport(pkg, "ErrorReport");

            // Assert
            Assert.IsNotNull(pkg.Workbook.Worksheets["ErrorReport"]);
        }

        [TestMethod]
        public void GenErrorReport_WithCopyFilesAndNoErrors_CreatesNoErrorSheet()
        {
            // Arrange
            using var pkg = new ExcelPackage();
            // Act
            ErrorReportManager.GenErrorReport(pkg, [], "ErrorReport");

            // Assert
            ExcelWorksheet sheet = pkg.Workbook.Worksheets["ErrorReport"];
            Assert.IsNotNull(sheet);
            Assert.AreEqual("No error occurs", sheet.Cells[1, 1].Value.ToString());
        }

        [TestMethod]
        public void GenErrorReport_WithCopyFilesAndErrors_CreatesErrorReportSheet()
        {
            // Arrange
            ErrorReportManager.AddError(BasicErrorType.E_MissingPin_01, "Sheet1", 1, 1, "msg");
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("Sheet1");

            // Act
            ErrorReportManager.GenErrorReport(pkg, [], "ErrorReport");

            // Assert
            Assert.IsNotNull(pkg.Workbook.Worksheets["ErrorReport"]);
        }
    }
}
