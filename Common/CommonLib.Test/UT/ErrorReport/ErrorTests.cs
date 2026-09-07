using CommonLib.ErrorReport.Base;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test.UT.ErrorReport
{
    [TestClass]
    public class ErrorTests
    {
        private static ErrorCode CreateErrorCode(EnumErrorLevel enumErrorLevel = EnumErrorLevel.Error)
        {
            return new ErrorCode(EnumErrorCategory.Basic, EnumErrorBehavior.Duplicate, EnumErrorTarget.Pin, 1, "Template {0}", enumErrorLevel);
        }

        [TestMethod]
        public void Constructor_WithMessage_SetsPropertiesCorrectly()
        {
            // Arrange
            ErrorCode errorCode = CreateErrorCode();

            // Act
            var error = new Error(errorCode, EnumErrorLevel.Error, "Sheet1", 5, 3, "my message");

            // Assert
            Assert.AreEqual("Sheet1", error.SheetName);
            Assert.AreEqual(5, error.RowNum);
            Assert.AreEqual(3, error.ColNum);
            Assert.AreEqual("my message", error.Message);
            Assert.AreEqual(errorCode, error.ErrorCode);
            Assert.AreEqual(EnumErrorLevel.Error, error.ErrorLevel);
        }

        [TestMethod]
        public void Constructor_WithMessageArgs_FormatsMessageFromTemplate()
        {
            // Arrange
            ErrorCode errorCode = CreateErrorCode();

            // Act
            var error = new Error(errorCode, "Sheet1", 1, 1, ["ArgValue"]);

            // Assert
            Assert.AreEqual("Template ArgValue", error.Message);
        }

        [TestMethod]
        public void Constructor_WithMessageArgs_UsesErrorCodeLevel()
        {
            // Arrange
            ErrorCode errorCode = CreateErrorCode(enumErrorLevel: EnumErrorLevel.Warning);

            // Act
            var error = new Error(errorCode, "Sheet1", 1, 1, ["X"]);

            // Assert
            Assert.AreEqual(EnumErrorLevel.Warning, error.ErrorLevel);
        }

        [TestMethod]
        public void Constructor_WithErrorInfo_SetsPatternAndComments()
        {
            // Arrange
            string[] expectedComments = ["comment1", "comment2"];
            ErrorCode errorCode = CreateErrorCode();
            var errorInfo = new ErrorInfo
            {
                Pattern = "pat001",
                Comments = [.. expectedComments]
            };

            // Act
            var error = new Error(errorCode, EnumErrorLevel.Error, "Sheet1", 1, 1, "msg", errorInfo);

            // Assert
            Assert.AreEqual("pat001", error.Pattern);
            CollectionAssert.AreEqual(expectedComments, error.Comments);
        }

        [TestMethod]
        public void Constructor_NullErrorInfo_PatternIsNullAndCommentsIsEmpty()
        {
            // Arrange
            ErrorCode errorCode = CreateErrorCode();

            // Act
            var error = new Error(errorCode, EnumErrorLevel.Error, "Sheet1", 1, 1, "msg", null);

            // Assert
            Assert.IsNull(error.Pattern);
            Assert.AreEqual(0, error.Comments.Count);
        }

        [TestMethod]
        public void CopyConstructor_NullOther_DoesNotThrow()
        {
            // Arrange & Act
            var error = new Error(null);

            // Assert - default values, no exception
            Assert.AreEqual("", error.SheetName);
        }

        [TestMethod]
        public void CopyConstructor_ValidOther_CopiesAllFields()
        {
            // Arrange
            ErrorCode errorCode = CreateErrorCode();
            var original = new Error(errorCode, EnumErrorLevel.Error, "Sheet1", 5, 3, "msg")
            {
                Pattern = "pat"
            };
            original.Comments.Add("c1");

            // Act
            var copy = new Error(original);

            // Assert
            Assert.AreEqual(original.SheetName, copy.SheetName);
            Assert.AreEqual(original.RowNum, copy.RowNum);
            Assert.AreEqual(original.ColNum, copy.ColNum);
            Assert.AreEqual(original.Message, copy.Message);
            Assert.AreEqual(original.Pattern, copy.Pattern);
            CollectionAssert.AreEqual(original.Comments, copy.Comments);
        }

        [TestMethod]
        public void Copy_ReturnsNewInstanceWithSameValues()
        {
            // Arrange
            ErrorCode errorCode = CreateErrorCode();
            var original = new Error(errorCode, EnumErrorLevel.Error, "Sheet1", 2, 4, "msg");

            // Act
            Error copy = original.Copy();

            // Assert
            Assert.AreNotSame(original, copy);
            Assert.AreEqual(original.SheetName, copy.SheetName);
            Assert.AreEqual(original.RowNum, copy.RowNum);
            Assert.AreEqual(original.ColNum, copy.ColNum);
            Assert.AreEqual(original.Message, copy.Message);
        }

        [TestMethod]
        public void SheetName_NullValue_ReturnsEmptyString()
        {
            // Arrange
            ErrorCode errorCode = CreateErrorCode();
            var error = new Error(errorCode, EnumErrorLevel.Error, null, 1, 1, "msg");

            // Act & Assert
            Assert.AreEqual("", error.SheetName);
        }

        [TestMethod]
        public void SheetName_EmptyValue_ReturnsEmptyString()
        {
            // Arrange
            ErrorCode errorCode = CreateErrorCode();
            var error = new Error(errorCode, EnumErrorLevel.Error, "", 1, 1, "msg");

            // Act & Assert
            Assert.AreEqual("", error.SheetName);
        }

        [TestMethod]
        public void GetAddress_ValidRowAndCol_ReturnsLetterPlusRow()
        {
            // Arrange
            _ = new Error(CreateErrorCode(), EnumErrorLevel.Error, "S", 5, 1, "msg");

            // Act
            string address = Error.GetAddress(5, 1);

            // Assert
            Assert.AreEqual("A5", address);
        }

        [TestMethod]
        public void GetAddress_RowZero_ReturnsRefError()
        {
            // Arrange
            _ = new Error(CreateErrorCode(), EnumErrorLevel.Error, "S", 1, 1, "msg");

            // Act
            string address = Error.GetAddress(0, 1);

            // Assert
            Assert.AreEqual("#REF!", address);
        }

        [TestMethod]
        public void GetAddress_ColZero_ReturnsRefError()
        {
            // Arrange
            _ = new Error(CreateErrorCode(), EnumErrorLevel.Error, "S", 1, 1, "msg");

            // Act
            string address = Error.GetAddress(1, 0);

            // Assert
            Assert.AreEqual("#REF!", address);
        }

        [TestMethod]
        public void GetAddress_AbsoluteTrue_ReturnsDollarPrefixedAddress()
        {
            // Arrange
            _ = new Error(CreateErrorCode(), EnumErrorLevel.Error, "S", 1, 1, "msg");

            // Act
            string address = Error.GetAddress(3, 2, absolute: true);

            // Assert
            Assert.AreEqual("$B$3", address);
        }

        [TestMethod]
        public void GetAddress_NoParams_RowOrColZero_ReturnsRefError()
        {
            // Arrange
            var error = new Error(CreateErrorCode(), EnumErrorLevel.Error, "S", 0, 0, "msg");

            // Act
            string address = error.GetAddress();

            // Assert
            Assert.AreEqual("#REF!", address);
        }

        [TestMethod]
        public void GetAddress_NoParams_ValidValues_ReturnsCorrectAddress()
        {
            // Arrange
            var error = new Error(CreateErrorCode(), EnumErrorLevel.Error, "S", 3, 2, "msg");

            // Act
            string address = error.GetAddress();

            // Assert
            Assert.AreEqual("B3", address);
        }

        [TestMethod]
        public void ColLetter_Col1_ReturnsA()
        {
            // Arrange
            var error = new Error(CreateErrorCode(), EnumErrorLevel.Error, "S", 1, 1, "msg");

            // Act & Assert
            Assert.AreEqual("A", error.ColLetter);
        }

        [TestMethod]
        public void ColLetter_Col26_ReturnsZ()
        {
            // Arrange
            var error = new Error(CreateErrorCode(), EnumErrorLevel.Error, "S", 1, 26, "msg");

            // Act & Assert
            Assert.AreEqual("Z", error.ColLetter);
        }

        [TestMethod]
        public void ColLetter_Col27_ReturnsAA()
        {
            // Arrange
            var error = new Error(CreateErrorCode(), EnumErrorLevel.Error, "S", 1, 27, "msg");

            // Act & Assert
            Assert.AreEqual("AA", error.ColLetter);
        }

        [TestMethod]
        public void ColLetter_Col0_ReturnsEmptyString()
        {
            // Arrange
            var error = new Error(CreateErrorCode(), EnumErrorLevel.Error, "S", 1, 0, "msg");

            // Act & Assert
            Assert.AreEqual("", error.ColLetter);
        }

        [TestMethod]
        public void GetHyperlink_ColZero_ReturnsRowRangeLink()
        {
            // Arrange
            var error = new Error(CreateErrorCode(), EnumErrorLevel.Error, "Sheet1", 5, 0, "msg");

            // Act
            string link = error.GetHyperlink();

            // Assert
            Assert.AreEqual("Sheet1!5:5", link);
        }

        [TestMethod]
        public void GetHyperlink_ColPositive_ReturnsCellLink()
        {
            // Arrange
            var error = new Error(CreateErrorCode(), EnumErrorLevel.Error, "Sheet1", 3, 2, "msg");

            // Act
            string link = error.GetHyperlink();

            // Assert
            Assert.AreEqual("Sheet1!B3", link);
        }

        [TestMethod]
        public void Link_ColZero_ReturnsHyperlinkFormulaWithRowRange()
        {
            // Arrange
            var error = new Error(CreateErrorCode(), EnumErrorLevel.Error, "Sheet1", 5, 0, "msg");

            // Act
            string link = error.Link;

            // Assert
            StringAssert.Contains(link, "Sheet1");
            StringAssert.Contains(link, "5:5");
            StringAssert.StartsWith(link, "=HYPERLINK(");
        }

        [TestMethod]
        public void Link_ColPositive_ReturnsHyperlinkFormulaWithCellAddress()
        {
            // Arrange
            var error = new Error(CreateErrorCode(), EnumErrorLevel.Error, "Sheet1", 1, 1, "msg");

            // Act
            string link = error.Link;

            // Assert
            StringAssert.Contains(link, "A1");
            StringAssert.StartsWith(link, "=HYPERLINK(");
        }

        [TestMethod]
        public void Print_ValidError_ReturnsFormattedString()
        {
            // Arrange
            ErrorCode errorCode = CreateErrorCode();
            var error = new Error(errorCode, EnumErrorLevel.Warning, "Sheet1", 2, 3, "test message");

            // Act
            string result = error.Print();

            // Assert
            StringAssert.Contains(result, "Sheet1");
            StringAssert.Contains(result, "2");
            StringAssert.Contains(result, "Warning");
            StringAssert.Contains(result, "test message");
        }
    }
}
