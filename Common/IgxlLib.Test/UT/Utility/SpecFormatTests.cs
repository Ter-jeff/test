using IgxlLib.Utility;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.Utility
{
    [TestClass]
    public class SpecFormatTests
    {
        [TestMethod]
        public void GenGlbSpecSymbol_ValidPinName_ReturnsPinWithGlbSuffix()
        {
            // Arrange
            string pinName = "VDD";
            string expected = "VDD_GLB";

            // Act
            string result = SpecFormat.GenGlbSpecSymbol(pinName);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void GenAcSpecSymbol_ValidPinName_ReturnsPinWithVarSuffix()
        {
            // Arrange
            string pinName = "CLK";
            string expected = "CLK_VAR";

            // Act
            string result = SpecFormat.GenAcSpecSymbol(pinName);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void GenGlbRatio_ValidInputs_ReturnsFormattedRatioString()
        {
            // Arrange
            string global = "Vdd";
            string ratio = "FACTOR";
            string expected = "=_Vdd*_FACTOR";

            // Act
            string result = SpecFormat.GenGlbRatio(global, ratio);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void CreateSpecSymbol_ValidSuffix_AppendsSuffixWithUnderscore()
        {
            // Arrange
            string src = " Pin1 ";
            string suffix = "TEST";
            string expected = "Pin1_TEST";

            // Act
            string result = SpecFormat.CreateSpecSymbol(src, suffix);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void CreateSpecSymbol_NullOrEmptySuffix_ReturnsTrimmedSource()
        {
            // Arrange
            string src = " Pin1 ";
            string expected = "Pin1";

            // Act
            string resultWithNull = SpecFormat.CreateSpecSymbol(src, null);
            string resultWithEmpty = SpecFormat.CreateSpecSymbol(src, "");

            // Assert
            Assert.AreEqual(expected, resultWithNull);
            Assert.AreEqual(expected, resultWithEmpty);
        }

        [TestMethod]
        public void GenGlbSpecSymbolAtLevelSheet_WithBlock_IncludesBlockAndGlbSuffix()
        {
            // Arrange
            string pinName = "IO_PAD";
            string parameterName = "VOH";
            string block = "B1";
            string expected = "_IO_PAD_VOH_B1_GLB";

            // Act
            string result = SpecFormat.GenGlbSpecSymbolAtLevelSheet(pinName, parameterName, block);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void GenGlbSpecSymbolAtLevelSheet_EmptyBlock_ExcludesBlock()
        {
            // Arrange
            string pinName = "IO_PAD";
            string parameterName = "VOH";
            string expected = "_IO_PAD_VOH_GLB";

            // Act
            string result = SpecFormat.GenGlbSpecSymbolAtLevelSheet(pinName, parameterName);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void GenDcSpecSymbolAtLevelSheet_AllParametersProvided_ReturnsFullString()
        {
            // Arrange
            string pinName = "ANA";
            string parameterName = "VMax";
            string parameterSyntax = "SYNTAX";
            string blockType = "BL_A";
            string chiplet = "CHIP1";
            string expected = "_ANA_VMax_SYNTAX_VAR_BL_A_CHIP1";

            // Act
            string result = SpecFormat.GenDcSpecSymbolAtLevelSheet(pinName, parameterName, parameterSyntax, blockType, chiplet);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void GenDcSpecSymbolAtLevelSheet_EmptyOptionalParameters_ExcludesThem()
        {
            // Arrange
            string pinName = "ANA";
            string expected = "_ANA_VAR";

            // Act
            string result = SpecFormat.GenDcSpecSymbolAtLevelSheet(pinName, "", null, "", null);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void GenSpecValueSingleValue_ValidString_PrependsEqualSign()
        {
            // Arrange
            string input = "5.0";
            string expected = "=5.0";

            // Act
            string result = SpecFormat.GenSpecValueSingleValue(input);

            // Assert
            Assert.AreEqual(expected, result);
        }
    }
}
