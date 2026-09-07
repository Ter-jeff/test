using System.IO;
using System.Text;

using IgxlLib.Enums;
using IgxlLib.Utility;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.Utility
{
    [TestClass]
    public class IgxlLoaderHelpersTests
    {
        // Helper method to create a Stream from a string
        private static MemoryStream CreateStreamFromString(string content) => new(Encoding.UTF8.GetBytes(content));

        #region GetIgxlSheetType Tests

        [TestMethod]
        public void GetIgxlSheetType_ExtensionIsBas_ReturnsBas()
        {
            // Arrange
            using Stream stream = CreateStreamFromString("Some random text");
            // Act
            EnumSheetType result = IgxlLoaderHelpers.GetIgxlSheetType(stream, "file.bas");

            // Assert
            Assert.AreEqual(EnumSheetType.Bas, result);
        }

        [TestMethod]
        public void GetIgxlSheetType_ExtensionIsCls_ReturnsCls()
        {
            // Arrange
            using Stream stream = CreateStreamFromString("Some random text");
            // Act
            EnumSheetType result = IgxlLoaderHelpers.GetIgxlSheetType(stream, "file.CLS");

            // Assert
            Assert.AreEqual(EnumSheetType.Cls, result);
        }

        [TestMethod]
        public void GetIgxlSheetType_OtherExtension_ReadsFirstLineAndEvaluates()
        {
            // Arrange
            using Stream stream = CreateStreamFromString("DTFlowtableSheet version 1.0\nSecond line data");
            // Act
            EnumSheetType result = IgxlLoaderHelpers.GetIgxlSheetType(stream, "file.txt");

            // Assert
            Assert.AreEqual(EnumSheetType.DTFlowtableSheet, result);
        }
        #endregion

        #region GetIgxlType Tests
        [TestMethod]
        public void GetIgxlType_NullOrEmpty_ReturnsDTUnknown()
        {
            Assert.AreEqual(EnumSheetType.DTUnknown, IgxlLoaderHelpers.GetIgxlSheetType(null));
            Assert.AreEqual(EnumSheetType.DTUnknown, IgxlLoaderHelpers.GetIgxlSheetType(""));
        }

        [TestMethod]
        [DataRow("This is a DTFlowtableSheet string", EnumSheetType.DTFlowtableSheet)]
        [DataRow("DTFlowtableSheet mixed case", EnumSheetType.DTFlowtableSheet)]
        [DataRow("Contains DTTestInstancesSheet inside", EnumSheetType.DTTestInstancesSheet)]
        [DataRow("Contains DTDCSpecSheet inside", EnumSheetType.DTDCSpecSheet)]
        [DataRow("Contains DTACSpecSheet inside", EnumSheetType.DTACSpecSheet)]
        [DataRow("Contains DTLevelSheet inside", EnumSheetType.DTLevelSheet)]
        [DataRow("Contains DTGlobalSpecSheet inside", EnumSheetType.DTGlobalSpecSheet)]
        [DataRow("Contains DTTimesetBasicSheet inside", EnumSheetType.DTTimesetBasicSheet)]
        [DataRow("Contains DTBintablesSheet inside", EnumSheetType.DTBintablesSheet)]
        [DataRow("Contains DTChanMap inside", EnumSheetType.DTChanMap)]
        [DataRow("Contains DTCharacterizationSheet inside", EnumSheetType.DTCharacterizationSheet)]
        [DataRow("Contains DTJobListSheet inside", EnumSheetType.DTJobListSheet)]
        [DataRow("Contains DTPatternSetSheet inside", EnumSheetType.DTPatternSetSheet)]
        [DataRow("Contains DTPatternSubroutineSheet inside", EnumSheetType.DTPatternSubroutineSheet)]
        [DataRow("Contains DTReferencesSheet inside", EnumSheetType.DTReferencesSheet)]
        [DataRow("Contains DTPinMap inside", EnumSheetType.DTPinMap)]
        [DataRow("Contains DTPortMapSheet inside", EnumSheetType.DTPortMapSheet)]
        public void GetIgxlType_MatchingKeywords_ReturnsCorrectSheetType(string testInput, EnumSheetType enumSheetType)
        {
            // Act
            EnumSheetType result = IgxlLoaderHelpers.GetIgxlSheetType(testInput);

            // Assert
            Assert.AreEqual(enumSheetType, result);
        }

        [TestMethod]
        public void GetIgxlType_StartsWithRevAndContainsBinCut_ReturnsBinCutEqn()
        {
            // Arrange
            string input = "Rev: 1.2 - Bin Cut Equations";

            // Act
            EnumSheetType result = IgxlLoaderHelpers.GetIgxlSheetType(input);

            // Assert
            Assert.AreEqual(EnumSheetType.BinCut_eqn, result);
        }

        [TestMethod]
        public void GetIgxlType_ContainsBinCutButDoesNotStartWithRev_ReturnsDTUnknown()
        {
            // Arrange
            string input = "Prefix Rev: 1.2 Bin Cut";

            // Act
            EnumSheetType result = IgxlLoaderHelpers.GetIgxlSheetType(input);

            // Assert
            Assert.AreEqual(EnumSheetType.DTUnknown, result);
        }

        [TestMethod]
        public void GetIgxlType_UnrecognizedText_ReturnsDTUnknown()
        {
            // Arrange
            string input = "Some completely unrelated text string";

            // Act
            EnumSheetType result = IgxlLoaderHelpers.GetIgxlSheetType(input);

            // Assert
            Assert.AreEqual(EnumSheetType.DTUnknown, result);
        }
        #endregion

        [TestMethod]
        public void GetIgxlSheetType_WithNullText_ReturnsUnknown()
        {
            // Arrange & Act
            EnumSheetType result = IgxlLoaderHelpers.GetIgxlSheetType(null);

            // Assert
            Assert.AreEqual(EnumSheetType.DTUnknown, result);
        }

        [TestMethod]
        public void GetIgxlSheetType_WithEmptyText_ReturnsUnknown()
        {
            // Arrange & Act
            EnumSheetType result = IgxlLoaderHelpers.GetIgxlSheetType("");

            // Assert
            Assert.AreEqual(EnumSheetType.DTUnknown, result);
        }

        [TestMethod]
        public void GetIgxlSheetType_WithFlowtableText_ReturnsFlowtableSheet()
        {
            // Arrange
            string text = "DTFlowtableSheet";

            // Act
            EnumSheetType result = IgxlLoaderHelpers.GetIgxlSheetType(text);

            // Assert
            Assert.AreEqual(EnumSheetType.DTFlowtableSheet, result);
        }

        [TestMethod]
        public void GetIgxlSheetType_WithTestInstancesText_ReturnsTestInstancesSheet()
        {
            // Arrange
            string text = "DTTestInstancesSheet";

            // Act
            EnumSheetType result = IgxlLoaderHelpers.GetIgxlSheetType(text);

            // Assert
            Assert.AreEqual(EnumSheetType.DTTestInstancesSheet, result);
        }

        [TestMethod]
        public void GetIgxlSheetType_WithLevelText_ReturnsLevelSheet()
        {
            // Arrange
            string text = "DTLevelSheet";

            // Act
            EnumSheetType result = IgxlLoaderHelpers.GetIgxlSheetType(text);

            // Assert
            Assert.AreEqual(EnumSheetType.DTLevelSheet, result);
        }

        [TestMethod]
        public void GetIgxlSheetType_WithBinTableText_ReturnsBintablesSheet()
        {
            // Arrange
            string text = "DTBintablesSheet";

            // Act
            EnumSheetType result = IgxlLoaderHelpers.GetIgxlSheetType(text);

            // Assert
            Assert.AreEqual(EnumSheetType.DTBintablesSheet, result);
        }

        [TestMethod]
        public void GetIgxlSheetType_CaseInsensitive_ReturnsCorrectType()
        {
            // Arrange
            string text = "dtflowtablesheet";

            // Act
            EnumSheetType result = IgxlLoaderHelpers.GetIgxlSheetType(text);

            // Assert
            Assert.AreEqual(EnumSheetType.DTFlowtableSheet, result);
        }

        [TestMethod]
        public void GetIgxlSheetType_WithBinCutText_ReturnsBinCutEqn()
        {
            // Arrange
            string text = "Rev: 1.0 Bin Cut";

            // Act
            EnumSheetType result = IgxlLoaderHelpers.GetIgxlSheetType(text);

            // Assert
            Assert.AreEqual(EnumSheetType.BinCut_eqn, result);
        }

        [TestMethod]
        public void GetIgxlSheetType_WithInvalidText_ReturnsUnknown()
        {
            // Arrange
            string text = "InvalidSheetType";

            // Act
            EnumSheetType result = IgxlLoaderHelpers.GetIgxlSheetType(text);

            // Assert
            Assert.AreEqual(EnumSheetType.DTUnknown, result);
        }

        [TestMethod]
        public void GetIgxlSheetType_WithPartialMatch_ReturnsCorrectType()
        {
            // Arrange
            string text = "This is DTFlowtableSheet information";

            // Act
            EnumSheetType result = IgxlLoaderHelpers.GetIgxlSheetType(text);

            // Assert
            Assert.AreEqual(EnumSheetType.DTFlowtableSheet, result);
        }
    }
}
