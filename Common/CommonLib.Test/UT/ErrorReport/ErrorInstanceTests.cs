using System.Collections.Generic;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test.UT.ErrorReport
{
    [TestClass]
    public class ErrorInstanceTests
    {
        private ErrorInstance _instance;
        private ErrorCode _errorCode;

        [TestInitialize]
        public void Setup()
        {
            _instance = new ErrorInstance();
            _errorCode = new ErrorCode(EnumErrorCategory.Basic, EnumErrorBehavior.Duplicate, EnumErrorTarget.Pin, 1, "Template {0}", EnumErrorLevel.Error);
        }

        private Error CreateError(string sheetName = "Sheet1", int row = 1, int col = 1, string message = "test message", string codeName = null)
        {
            ErrorCode code = string.IsNullOrEmpty(codeName) ? _errorCode : new ErrorCode(EnumErrorCategory.Basic, EnumErrorBehavior.Duplicate, EnumErrorTarget.Pin, 1, "T", EnumErrorLevel.Error);
            return new Error(code, EnumErrorLevel.Error, sheetName, row, col, message);
        }

        [TestMethod]
        public void GetErrorCount_EmptyInstance_ReturnsZero()
        {
            // Arrange & Act
            int count = _instance.GetErrorCount();

            // Assert
            Assert.AreEqual(0, count);
        }

        [TestMethod]
        public void AddError_SingleError_CountBecomesOne()
        {
            // Arrange
            Error error = CreateError();

            // Act
            _instance.AddError(error);

            // Assert
            Assert.AreEqual(1, _instance.GetErrorCount());
        }

        [TestMethod]
        public void AddErrors_MultipleErrors_CountMatchesListSize()
        {
            // Arrange
            var errors = new List<Error>
            {
                CreateError(row: 1),
                CreateError(row: 2),
                CreateError(row: 3)
            };

            // Act
            _instance.AddErrors(errors);

            // Assert
            Assert.AreEqual(3, _instance.GetErrorCount());
        }

        [TestMethod]
        public void AddErrors_EmptyList_CountRemainsZero()
        {
            // Arrange & Act
            _instance.AddErrors([]);

            // Assert
            Assert.AreEqual(0, _instance.GetErrorCount());
        }

        [TestMethod]
        public void ClearErrors_AfterAddingErrors_CountBecomesZero()
        {
            // Arrange
            _instance.AddError(CreateError());
            _instance.AddError(CreateError());

            // Act
            _instance.ClearErrors();

            // Assert
            Assert.AreEqual(0, _instance.GetErrorCount());
        }

        [TestMethod]
        public void GetErrorList_ReturnsAllAddedErrors()
        {
            // Arrange
            Error error1 = CreateError(row: 1, message: "msg1");
            Error error2 = CreateError(row: 2, message: "msg2");
            _instance.AddError(error1);
            _instance.AddError(error2);

            // Act
            List<Error> list = _instance.GetErrorList();

            // Assert
            Assert.AreEqual(2, list.Count);
        }

        [TestMethod]
        public void Exist_MessageExistsExactCase_ReturnsTrue()
        {
            // Arrange
            _instance.AddError(CreateError(message: "hello world"));

            // Act
            bool result = _instance.Exist("hello world");

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void Exist_MessageExistsDifferentCase_ReturnsTrue()
        {
            // Arrange
            _instance.AddError(CreateError(message: "Hello World"));

            // Act
            bool result = _instance.Exist("hello world");

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void Exist_MessageNotPresent_ReturnsFalse()
        {
            // Arrange
            _instance.AddError(CreateError(message: "some message"));

            // Act
            bool result = _instance.Exist("different message");

            // Assert
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void Exist_EmptyInstance_ReturnsFalse()
        {
            // Arrange & Act
            bool result = _instance.Exist("anything");

            // Assert
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void GetSortedErrors_DuplicateKeyErrors_ReturnsOnlyOne()
        {
            // Arrange - same sheet, message, row, col, code name → duplicate
            _instance.AddError(CreateError("Sheet1", 1, 1, "dup msg"));
            _instance.AddError(CreateError("Sheet1", 1, 1, "dup msg"));

            // Act
            List<Error> sorted = _instance.GetSortedErrors();

            // Assert
            Assert.AreEqual(1, sorted.Count);
        }

        [TestMethod]
        public void GetSortedErrors_DifferentMessages_ReturnsBoth()
        {
            // Arrange
            _instance.AddError(CreateError("Sheet1", 1, 1, "msg A"));
            _instance.AddError(CreateError("Sheet1", 1, 1, "msg B"));

            // Act
            List<Error> sorted = _instance.GetSortedErrors();

            // Assert
            Assert.AreEqual(2, sorted.Count);
        }

        [TestMethod]
        public void GetSortedErrors_MultipleSheets_SortedBySheetName()
        {
            // Arrange
            _instance.AddError(CreateError("SheetZ", 1, 1, "msg"));
            _instance.AddError(CreateError("SheetA", 1, 1, "msg2"));

            // Act
            List<Error> sorted = _instance.GetSortedErrors();

            // Assert
            Assert.AreEqual("SheetA", sorted[0].SheetName);
            Assert.AreEqual("SheetZ", sorted[1].SheetName);
        }

        [TestMethod]
        public void GetSortedErrors_SameSheetDifferentRows_SortedByRowNum()
        {
            // Arrange
            _instance.AddError(CreateError("Sheet1", 10, 1, "msg1"));
            _instance.AddError(CreateError("Sheet1", 2, 1, "msg2"));

            // Act
            List<Error> sorted = _instance.GetSortedErrors();

            // Assert
            Assert.AreEqual(2, sorted[0].RowNum);
            Assert.AreEqual(10, sorted[1].RowNum);
        }

        [TestMethod]
        public void GetSortedErrors_EmptyInstance_ReturnsEmptyList()
        {
            // Arrange & Act
            List<Error> sorted = _instance.GetSortedErrors();

            // Assert
            Assert.AreEqual(0, sorted.Count);
        }
    }
}
