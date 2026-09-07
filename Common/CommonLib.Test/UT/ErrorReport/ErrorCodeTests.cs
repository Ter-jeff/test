using CommonLib.ErrorReport.Base;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test.UT.ErrorReport
{
    [TestClass]
    public class ErrorCodeTests
    {
        [TestMethod]
        public void Constructor_AllParameters_SetsPropertiesCorrectly()
        {
            // Arrange & Act
            var errorCode = new ErrorCode(EnumErrorCategory.Basic, EnumErrorBehavior.Duplicate, EnumErrorTarget.Pin, 42, "Template {0}", EnumErrorLevel.Warning, "MyGuidance");

            // Assert
            Assert.AreEqual(EnumErrorCategory.Basic, errorCode.EnumErrorCategory);
            Assert.AreEqual(EnumErrorBehavior.Duplicate, errorCode.EnumErrorBehavior);
            Assert.AreEqual(EnumErrorTarget.Pin, errorCode.EnumErrorTarget);
            Assert.AreEqual(42, errorCode.Code);
            Assert.AreEqual("Template {0}", errorCode.MessageTemplate);
            Assert.AreEqual(EnumErrorLevel.Warning, errorCode.ErrorLevel);
            Assert.AreEqual("MyGuidance", errorCode.Guidance);
        }

        [TestMethod]
        public void Constructor_NullTemplate_MessageTemplateIsEmptyString()
        {
            // Arrange & Act
            var errorCode = new ErrorCode(EnumErrorCategory.Basic, EnumErrorBehavior.Duplicate, EnumErrorTarget.Pin, 1, null, EnumErrorLevel.Error);

            // Assert
            Assert.AreEqual(string.Empty, errorCode.MessageTemplate);
        }

        [TestMethod]
        public void Constructor_NullGuidance_GuidanceIsEmptyString()
        {
            // Arrange & Act
            var errorCode = new ErrorCode(EnumErrorCategory.Basic, EnumErrorBehavior.Duplicate, EnumErrorTarget.Pin, 1, "T", EnumErrorLevel.Error, null);

            // Assert
            Assert.AreEqual(string.Empty, errorCode.Guidance);
        }

        [TestMethod]
        public void FullCode_ErrorLevel_StartsWithE()
        {
            // Arrange
            var errorCode = new ErrorCode(EnumErrorCategory.Basic, EnumErrorBehavior.Duplicate, EnumErrorTarget.Pin, 1, "T", EnumErrorLevel.Error);

            // Act
            string fullCode = errorCode.FullCode;

            // Assert
            Assert.AreEqual('E', fullCode[0]);
        }

        [TestMethod]
        public void FullCode_WarningLevel_StartsWithW()
        {
            // Arrange
            var errorCode = new ErrorCode(EnumErrorCategory.Basic, EnumErrorBehavior.Duplicate, EnumErrorTarget.Pin, 1, "T", EnumErrorLevel.Warning);

            // Act
            string fullCode = errorCode.FullCode;

            // Assert
            Assert.AreEqual('W', fullCode[0]);
        }

        [TestMethod]
        public void FullCode_InfoLevel_StartsWithI()
        {
            // Arrange
            var errorCode = new ErrorCode(EnumErrorCategory.Basic, EnumErrorBehavior.Duplicate, EnumErrorTarget.Pin, 1, "T", EnumErrorLevel.Info);

            // Act
            string fullCode = errorCode.FullCode;

            // Assert
            Assert.AreEqual('I', fullCode[0]);
        }

        [TestMethod]
        public void FullCode_BasicCategory_ContainsBaDescriptionAtPosition1()
        {
            // Arrange
            var errorCode = new ErrorCode(EnumErrorCategory.Basic, EnumErrorBehavior.Duplicate, EnumErrorTarget.Pin, 1, "T", EnumErrorLevel.Error);

            // Act
            string fullCode = errorCode.FullCode;

            // Assert
            Assert.AreEqual("BA", fullCode.Substring(1, 2));
        }

        [TestMethod]
        public void FullCode_BinCutCategory_ContainsBcDescriptionAtPosition1()
        {
            // Arrange
            var errorCode = new ErrorCode(EnumErrorCategory.BinCut, EnumErrorBehavior.Duplicate, EnumErrorTarget.Pin, 1, "T", EnumErrorLevel.Error);

            // Act
            string fullCode = errorCode.FullCode;

            // Assert
            Assert.AreEqual("BC", fullCode.Substring(1, 2));
        }

        [TestMethod]
        public void FullCode_ValidValues_IsExactlyNineCharacters()
        {
            // Arrange
            var errorCode = new ErrorCode(EnumErrorCategory.Basic, EnumErrorBehavior.Duplicate, EnumErrorTarget.Pin, 1, "T", EnumErrorLevel.Error);

            // Act
            string fullCode = errorCode.FullCode;

            // Assert
            Assert.AreEqual(9, fullCode.Length, $"Expected 9-char FullCode but got '{fullCode}' ({fullCode.Length} chars)");
        }

        [TestMethod]
        public void FullCode_Code1_NnnPortionIsZeroZeroOne()
        {
            // Arrange
            var errorCode = new ErrorCode(EnumErrorCategory.Basic, EnumErrorBehavior.Duplicate, EnumErrorTarget.Pin, 1, "T", EnumErrorLevel.Error);

            // Act
            string fullCode = errorCode.FullCode;

            // Assert
            Assert.AreEqual("001", fullCode.Substring(6, 3));
        }

        [TestMethod]
        public void FullCode_Code999_NnnPortionIsNineNineNine()
        {
            // Arrange
            var errorCode = new ErrorCode(EnumErrorCategory.Basic, EnumErrorBehavior.Duplicate, EnumErrorTarget.Pin, 999, "T", EnumErrorLevel.Error);

            // Act
            string fullCode = errorCode.FullCode;

            // Assert
            Assert.AreEqual("999", fullCode.Substring(6, 3));
        }

        [TestMethod]
        public void FullCode_DuplicateBehaviorAndPinTarget_SssPortionIsOneFiveFive()
        {
            // Arrange - subCode = (int)Behavior * 100 + (int)Target = 1 * 100 + 55 = 155
            var errorCode = new ErrorCode(EnumErrorCategory.Basic, EnumErrorBehavior.Duplicate, EnumErrorTarget.Pin, 1, "T", EnumErrorLevel.Error);

            // Act
            string fullCode = errorCode.FullCode;

            // Assert
            Assert.AreEqual("155", fullCode.Substring(3, 3));
        }

        [TestMethod]
        public void FormatMessage_NullArgs_ReturnsTemplate()
        {
            // Arrange
            var errorCode = new ErrorCode(EnumErrorCategory.Basic, EnumErrorBehavior.Duplicate, EnumErrorTarget.Pin, 1, "Template", EnumErrorLevel.Error);

            // Act
            string result = errorCode.FormatMessage(null);

            // Assert
            Assert.AreEqual("Template", result);
        }

        [TestMethod]
        public void FormatMessage_EmptyArgs_ReturnsTemplate()
        {
            // Arrange
            var errorCode = new ErrorCode(EnumErrorCategory.Basic, EnumErrorBehavior.Duplicate, EnumErrorTarget.Pin, 1, "Template", EnumErrorLevel.Error);

            // Act
            string result = errorCode.FormatMessage();

            // Assert
            Assert.AreEqual("Template", result);
        }

        [TestMethod]
        public void FormatMessage_SingleArg_ReturnsFormattedString()
        {
            // Arrange
            var errorCode = new ErrorCode(EnumErrorCategory.Basic, EnumErrorBehavior.Duplicate, EnumErrorTarget.Pin, 1, "Hello {0}!", EnumErrorLevel.Error);

            // Act
            string result = errorCode.FormatMessage("World");

            // Assert
            Assert.AreEqual("Hello World!", result);
        }

        [TestMethod]
        public void FormatMessage_MultipleArgs_ReturnsFormattedString()
        {
            // Arrange
            var errorCode = new ErrorCode(EnumErrorCategory.Basic, EnumErrorBehavior.Duplicate, EnumErrorTarget.Pin, 1, "{0} and {1}", EnumErrorLevel.Error);

            // Act
            string result = errorCode.FormatMessage("A", "B");

            // Assert
            Assert.AreEqual("A and B", result);
        }

        [TestMethod]
        public void FormatMessage_TooFewArgsForTemplate_ReturnsRawTemplate()
        {
            // Arrange - template expects {0} and {1} but only one arg → FormatException swallowed
            var errorCode = new ErrorCode(EnumErrorCategory.Basic, EnumErrorBehavior.Duplicate, EnumErrorTarget.Pin, 1, "{0} and {1}", EnumErrorLevel.Error);

            // Act
            string result = errorCode.FormatMessage("OnlyOne");

            // Assert
            Assert.AreEqual("{0} and {1}", result);
        }
    }
}
