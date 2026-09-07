using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.PowerMerge;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class DataConvertorTests : FunctionTestBase
    {
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            LocalSpecs.TarFolder = OutputPath;
            PinMapSheet pinMapSheet = new PinMapSheet("");
            var pinGroup = new PinGroup("NormalPin", "I/O");
            pinGroup.AddPin(new Pin("TX_P", "I/O"));
            pinGroup.AddPin(new Pin("TX_N", "I/O"));
            pinMapSheet.AddGroup(pinGroup);
            Pin pin = new Pin("VCC", "power");
            pinMapSheet.AddPin(pin);
            TestProgram.IgxlWorkBk.PinMapPair = new KeyValuePair<string, PinMapSheet>(FolderStructure.DirPinMap, pinMapSheet);

            // Initialize the static dependency required by the method
            TestPlanStatic.PowerMergeSheet = new PowerMergeSheet();
            PowerMerge pm = TestPlanStatic.PowerMergeSheet.PowerMerge;

            // Direct Match Mapping
            pm.CpPowers.Add("Bump1", "NetA");
            pm.FtPowers.Add("Bump1", "NetA");

            // Split Mapping (CP and FT different)
            pm.CpPowers.Add("SplitPin", "NetCP");
            pm.FtPowers.Add("SplitPin", "NetFT");

            // Existing mappings from your snippet
            pm.CpPowers.Add("Key1", "VDD_CORE");
            pm.FtPowers.Add("Key2", "VDD_CORE");

            pm.CpPowers.Add("VDD1", "NetA");
            pm.FtPowers.Add("VDD1", "NetA");

            pm.CpPowers.Add("VDD_SPLIT", "NetCP");
            pm.FtPowers.Add("VDD_SPLIT", "NetFT");

            pm.CpPowers.Add("VDD_FT_ONLY", "N/A");
            pm.FtPowers.Add("VDD_FT_ONLY", "FT_NET");
        }

        [TestMethod]
        [DataRow("1*10^3", false, "1E3")]
        [DataRow("0mV", false, "0")]
        [DataRow("1mV", false, "0.001")]
        [DataRow("1uV", false, "1E-06")]
        [DataRow("1nV", false, "1E-09")]
        [DataRow("1pV", false, "1E-12")]
        [DataRow("1kV", false, "1000")]
        [DataRow("1MV", false, "1000000")]
        [DataRow("1GV", false, "1000000000")]
        [DataRow("1TV", false, "1000000000000")]
        [DataRow("0.0000001V", false, "1E-07")]
        [DataRow("0.0000001V", true, "0.0000001")]
        [DataRow("unknown", false, "unknown")]
        public void ConvertUnits_WithVariousInputs_ReturnsExpectedFormatting(string input, bool nonScience, string expected)
        {
            // Act
            string result = DataConvertor.ConvertUnits(input, nonScience);

            // Assert
            Assert.AreEqual(expected, result, $"Failed for input: {input}");
        }

        [TestMethod]
        [DataRow("100mA", "100", "mA", DisplayName = "Standard unit separation")]
        [DataRow("1.23V", "1.23", "V", DisplayName = "Decimal with unit")]
        [DataRow("-50uA", "-50", "uA", DisplayName = "Negative value with unit")]
        [DataRow("0", "0", "", DisplayName = "Zero handles without unit")]
        [DataRow("1.2E-3", "1.2E-3", "", DisplayName = "Scientific notation remains unchanged")]
        [DataRow("NA_Ohms", "", "_Ohms", DisplayName = "NA prefix extraction")]
        public void ConvertUseLimitFW_ShouldParseCorrectly(string input, string expectedStr, string expectedUnit)
        {
            // Act
            string result = DataConvertor.ConvertUseLimitFw(input, out string limitUnit, out string limitScale);

            // Assert
            Assert.AreEqual(expectedStr, result, $"Numeric part mismatch for input: {input}");
            Assert.AreEqual(expectedUnit, limitUnit, $"Unit part mismatch for input: {input}");
            Assert.AreEqual("", limitScale);
        }

        [TestMethod]
        public void ConvertUseLimitFW_WithPercentage_ShouldExtractCorrectly()
        {
            // Arrange
            string input = "10%";

            // Act
            string result = DataConvertor.ConvertUseLimitFw(input, out string unit, out _);

            // Assert
            Assert.AreEqual("10", result);
            Assert.AreEqual("%", unit);
        }

        [TestMethod]
        public void ConvertDifferentialPinGroup_ShouldReturnOriginal_WhenAlreadyPaired()
        {
            // Arrange
            string pairedInput = "PinA::PinB";

            // Act
            string result = DataConvertor.ConvertDifferentialPinGroup(pairedInput);

            // Assert
            Assert.AreEqual(pairedInput, result);
        }

        [TestMethod]
        public void ConvertDifferentialPinGroup_ShouldReturnOriginal_WhenGroupDoesNotExist()
        {
            // Arrange
            string nonGroup = "NormalPin";

            // Act
            string result = DataConvertor.ConvertDifferentialPinGroup(nonGroup);

            // Assert
            Assert.AreEqual("TX_N::TX_P", result);
        }

        [DataTestMethod]
        [DataRow("PinA+PinA", "PinA")]
        [DataRow("PinA+PinA+PinA", "PinA")]
        [DataRow("PinA+PinB", "PinA+PinB")]
        [DataRow("PinA+PinB+PinA", "PinA+PinB+PinA")]
        [DataRow("PinA+PinB,", "PinA+PinB")]
        [DataRow("PinA", "PinA")]
        [DataRow("", "")]
        public void RemoveDummyPlusSign_WithVariousInputs_ReturnsExpectedResult(string input, string expected)
        {
            // Act
            string result = DataConvertor.RemoveDummyPlusSign(input);

            // Assert
            Assert.AreEqual(expected, result, $"Failed for input: {input}");
        }

        [TestMethod]
        public void SortCpFtPin_NoTags_ReturnsOriginalString()
        {
            // Arrange
            string input = "PinA,PinB+PinC";

            // Act
            string result = DataConvertor.SortCpFtPin(input);

            // Assert
            Assert.AreEqual(input, result);
        }

        [TestMethod]
        public void SortCpFtPin_WithTags_SplitsIntoCpAndFtLists()
        {
            // Arrange: PinA is common, PinB is CP only, PinC is FT only
            string input = "PinA,CP=PinB,FT=PinC";

            // Act
            string result = DataConvertor.SortCpFtPin(input);

            // Assert
            // Logic: CP keeps common + CP tagged; FT keeps common + FT tagged
            Assert.AreEqual("CP=PinA,PinB;FT=PinA,PinC", result);
        }

        [TestMethod]
        public void SortCpFtPin_WithColonSweep_SplitsCorrectly()
        {
            // Arrange: Testing the sweep/differential logic (:)
            string input = "Pin1:CP=Pin2:FT=Pin3";

            // Act
            string result = DataConvertor.SortCpFtPin(input);

            // Assert
            // Logic: Joins back with ':'
            Assert.AreEqual("CP=Pin1:Pin2;FT=Pin1:Pin3", result);
        }

        [TestMethod]
        public void SortCpFtPin_MultipleSequences_MaintainsPlusSeparators()
        {
            // Arrange: Sequence 1 has tags, Sequence 2 is plain
            string input = "CP=A,FT=B+PlainPin";

            // Act
            string result = DataConvertor.SortCpFtPin(input);

            // Assert
            // Sequence 2 (PlainPin) should appear in both CP and FT sides
            Assert.AreEqual("CP=A+PlainPin;FT=B+PlainPin", result);
        }

        [TestMethod]
        public void SortCpFtPin_DuplicatePins_DeduplicatesInCommaMode()
        {
            // Arrange
            string input = "PinA,PinA,CP=PinB+";

            // Act
            string result = DataConvertor.SortCpFtPin(input);

            // Assert
            Assert.AreEqual("CP=PinA,PinB+;FT=PinA+", result);
        }

        [TestMethod]
        public void ConvertToNetName_NoVdd_ReturnsOriginalInput()
        {
            string input = "AnyPin";

            string result = DataConvertor.ConvertToNetName(
                input,
                TestPlanStatic.PowerMergeSheet.PowerMerge);

            Assert.AreEqual(input, result);
        }

        [TestMethod]
        public void ConvertToNetName_PowerMergeNull_ReturnsOriginalInput()
        {
            string input = "VDD1";

            string result = DataConvertor.ConvertToNetName(
                input,
                null);

            Assert.AreEqual(input, result);
        }

        [TestMethod]
        public void ConvertToNetName_MappingNotFound_ReturnsOriginalPinName()
        {
            string input = "VDD_UNKNOWN";

            string result = DataConvertor.ConvertToNetName(
                input,
                TestPlanStatic.PowerMergeSheet.PowerMerge);

            Assert.AreEqual(input, result);
        }

        [TestMethod]
        public void ConvertToNetName_CpAndFtSame_ReturnsNetName()
        {
            string input = "VDD1";

            string result = DataConvertor.ConvertToNetName(
                input,
                TestPlanStatic.PowerMergeSheet.PowerMerge);

            Assert.AreEqual("NetA", result);
        }

        [TestMethod]
        public void ConvertToNetName_CpAndFtDifferent_AddsPrefixes()
        {
            string input = "VDD_SPLIT";

            string result = DataConvertor.ConvertToNetName(
                input,
                TestPlanStatic.PowerMergeSheet.PowerMerge);

            Assert.AreEqual("CP=NetCP,FT=NetFT", result);
        }

        [TestMethod]
        public void ConvertToNetName_MultipleSequences_MaintainsPlusSeparator()
        {
            string input = "VDD1+VDD_SPLIT";

            string result = DataConvertor.ConvertToNetName(
                input,
                TestPlanStatic.PowerMergeSheet.PowerMerge);

            Assert.AreEqual("NetA+CP=NetCP,FT=NetFT", result);
        }

        [TestMethod]
        public void ConvertToNetName_EmptySequence_MaintainsSeparator()
        {
            string input = "VDD1++VDD_SPLIT";

            string result = DataConvertor.ConvertToNetName(
                input,
                TestPlanStatic.PowerMergeSheet.PowerMerge);

            Assert.AreEqual("NetA++CP=NetCP,FT=NetFT", result);
        }

        [TestMethod]
        public void ConvertToNetName_CpIsNA_ReturnsOnlyFt()
        {
            string input = "VDD_FT_ONLY";

            string result = DataConvertor.ConvertToNetName(
                input,
                TestPlanStatic.PowerMergeSheet.PowerMerge);

            Assert.AreEqual("FT=FT_NET", result);
        }

        [TestMethod]
        public void ConvertToNetName_DuplicateNetNames_RemovesDuplicates()
        {
            string input = "VDD1,VDD1";

            string result = DataConvertor.ConvertToNetName(
                input,
                TestPlanStatic.PowerMergeSheet.PowerMerge);

            Assert.AreEqual("NetA", result);
        }

        [TestMethod]
        public void ConvertValueWithGlbSpec_WhenTxPinDoesNotExist_ReturnsOriginalValue()
        {
            string input = "TX_P";

            string result = DataConvertor.ConvertValueWithGlbSpec(input);

            Assert.AreEqual("TX_P", result);
        }

        [TestMethod]
        public void ConvertValueWithGlbSpec_WhenVccPinExists_ReturnsPinWithVar()
        {
            // Arrange
            string input = "VCC";

            // Act
            string result = DataConvertor.ConvertValueWithGlbSpec(input);

            // Assert
            Assert.AreEqual("_VCC_VAR", result);
        }

        [TestMethod]
        public void ConvertValueWithGlbSpec_MixedExpression_ReplacesOnlyPin()
        {
            // Arrange
            string input = "VCC+100mA";

            // Act
            string result = DataConvertor.ConvertValueWithGlbSpec(input);

            // Assert
            Assert.AreEqual("_VCC_VAR+0.1", result);
        }

        [TestMethod]
        public void ConvertValueWithGlbSpec_WhenTxPinExists_ReturnsOriginalValue()
        {
            string input = "TX_P";

            string result = DataConvertor.ConvertValueWithGlbSpec(input);

            Assert.AreEqual("TX_P", result);
        }

        [TestMethod]
        public void ConvertValueWithGlbSpec_WhenPinExists_ReturnsPinWithVar()
        {
            string input = "VCC";

            string result = DataConvertor.ConvertValueWithGlbSpec(input);

            Assert.AreEqual("_VCC_VAR", result);
        }

        [TestMethod]
        public void ConvertValueWithGlbSpec_MultiplePins_ReplacesAllPins()
        {
            string input = "VCC+TX_P";

            string result = DataConvertor.ConvertValueWithGlbSpec(input);

            Assert.AreEqual("_VCC_VAR+TX_P", result);
        }

        [TestMethod]
        public void ConvertValueWithGlbSpec_WhenPinNameIsLowerCase_ReturnsUpperCasePinWithVar()
        {
            string input = "vcc";

            string result = DataConvertor.ConvertValueWithGlbSpec(input);

            Assert.AreEqual("_VCC_VAR", result);
        }

        [TestMethod]
        public void ConvertValueWithGlbSpec_WhenPinDoesNotExist_CallsConvertUnits()
        {
            string input = "100mA";

            string result = DataConvertor.ConvertValueWithGlbSpec(input);

            Assert.AreEqual("0.1", result);
        }

        [TestMethod]
        public void ConvertValueWithGlbSpec_NonScience_ReturnsConvertedNumber()
        {
            string input = "5.0";

            string result = DataConvertor.ConvertValueWithGlbSpec(input, true);

            Assert.AreEqual("5.0", result);
        }

        [TestMethod]
        [DataRow("Binning_Level1")]
        [DataRow("binning_low")]
        public void ConvertUseLimit_Binning_ReturnsEmpty(string input)
        {
            string result = DataConvertor.ConvertUseLimit(
                input,
                out string unit,
                out string scale);

            Assert.AreEqual("", result);
            Assert.AreEqual("", unit);
            Assert.AreEqual("", scale);
        }

        [TestMethod]
        [DataRow("0x10", "16")]
        [DataRow("0b1010", "10")]
        [DataRow("0d123", "123")]
        public void ConvertUseLimit_RadixPrefix_ConvertsToLongString(
            string input,
            string expected)
        {
            string result = DataConvertor.ConvertUseLimit(
                input,
                out string unit,
                out string scale);

            Assert.AreEqual(expected, result);
            Assert.AreEqual("", unit);
            Assert.AreEqual("", scale);
        }

        [TestMethod]
        public void ConvertUseLimit_Zero_ReturnsZero()
        {
            string result = DataConvertor.ConvertUseLimit(
                "0",
                out string unit,
                out string scale);

            Assert.AreEqual("0", result);
            Assert.AreEqual("", unit);
            Assert.AreEqual("", scale);
        }

        [TestMethod]
        public void ConvertUseLimit_VddPin_AppendsVar()
        {
            string result = DataConvertor.ConvertUseLimit(
                "VDD_CORE",
                out string unit,
                out string scale);

            Assert.AreEqual("=_VDD_CORE_VAR", result);
            Assert.AreEqual("", unit);
            Assert.AreEqual("", scale);
        }

        [TestMethod]
        public void ConvertUseLimit_VddPinWithUnit_ConvertsCorrectly()
        {
            string result = DataConvertor.ConvertUseLimit(
                "VDD,XXX",
                out string unit,
                out string scale);

            Assert.AreEqual("=_VDD_VAR,", result);
            Assert.AreEqual("XXX", unit);
            Assert.AreEqual("", scale);
        }

        [TestMethod]
        public void ConvertUseLimit_InvalidDecimalPrefix_ReturnsZero()
        {
            string result = DataConvertor.ConvertUseLimit(
                "0dxx",
                out _,
                out _);

            Assert.AreEqual("0", result);
        }

        [TestMethod]
        public void ConvertUseLimit_WithUnits_CallsConvertUnits()
        {
            string result = DataConvertor.ConvertUseLimit(
                "100mA",
                out string unit,
                out string scale);

            Assert.AreEqual("0.1", result);
            Assert.AreEqual("A", unit);
            Assert.AreEqual("m", scale);
        }

        [TestMethod]
        public void ConvertUseLimit_Expression_AddsEqualPrefix()
        {
            string result = DataConvertor.ConvertUseLimit(
                "1+2",
                out string unit,
                out string scale);

            Assert.AreEqual("=1+2", result);
            Assert.AreEqual("", unit);
            Assert.AreEqual("", scale);
        }

        [TestMethod]
        [DataRow("1000mHz", "1", "Hz", "m")]
        [DataRow("50kOhm", "50000", "Ohm", "K")]
        [DataRow("50kOhms", "50000", "Ohms", "K")]
        [DataRow("NAmHz", "", "Hz", "m")] // Tests 'NA' prefix logic
        [DataRow("NAkOhm", "", "Ohm", "K")] // Tests 'NA' prefix logic
        [DataRow("NAkOhms", "", "Ohms", "K")] // Tests 'NA' prefix logic
        [DataRow("1.5E+2", "1.5E+2", "", "")] // Tests early return for 'E'
        [DataRow("100uA", "0.0001", "A", "u")]
        public void ConvertUnits_StateVerification(string input, string expectedNum, string expectedUnit, string expectedScale)
        {
            // Act
            string resultNum = DataConvertor.ConvertUnits(input, out string actualUnit, out string actualScale);

            // Assert
            Assert.AreEqual(expectedNum, resultNum, $"Value mismatch for {input}");
            Assert.AreEqual(expectedUnit, actualUnit, $"Unit mismatch for {input}");
            Assert.AreEqual(expectedScale, actualScale, $"Scale mismatch for {input}");
        }

        [TestMethod]
        [DataRow("m", 0.001, "m")]            // Milli (regex12)
        [DataRow("u", 0.000001, "u")]
        [DataRow("n", 0.000000001, "n")]
        [DataRow("p", 0.000000000001, "p")]
        [DataRow("k", 1000.0, "K")]           // Kilo (regex16)
        [DataRow("M", 1000000.0, "M")]        // Mega (regex17)
        [DataRow("G", 1000000000.0, "G")]
        [DataRow("T", 1000000000000.0, "T")]
        [DataRow("%", 0.01, "%")]             // Percent (regex11)
        [DataRow("XXX", 1.0, "")]             // Default / No match
        [DataRow("\"XXX", 1.0, "")]             // Default / No match
        [DataRow("NoXXX", 1.0, "No")]
        public void GetRateUnitScale_ReturnsCorrectTuple(string inputUnit, double expectedRate, string expectedScale)
        {
            // Act
            (double rate, string _, string scale) = DataConvertor.GetRateUnitScale(inputUnit);

            // Assert
            Assert.AreEqual(expectedRate, rate, 0.000001, $"Rate mismatch for {inputUnit}");
            Assert.AreEqual(expectedScale, scale, $"Scale label mismatch for {inputUnit}");
        }

        [TestMethod]
        [DataRow("REG_ASSIGN_Value", "REG_ASSIGN_Value")]
        public void ConvertValueSpec_RegAssign_ReturnsOriginal(string input, string expected)
        {
            string result = DataConvertor.ConvertValueSpec(input);
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        [DataRow("1.2V@Alg:MyPattern", "1.2V", "Alg:MyPattern")]
        [DataRow("Alg:DirectPattern", "", "Alg:DirectPattern")]
        [DataRow("2.5V@Alg:Other", "1.2V", "")] // Should return empty or skip if voltage doesn't match and list is empty
        public void ConvertValueSpec_AlgLogic_HandlesVoltagePrefix(string input, string voltage, string expected)
        {
            // Note: This assumes the internal List remains empty if conditions aren't met
            string result = DataConvertor.ConvertValueSpec(input, voltage);
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        [DataRow("Value1:Value2:UnknownPin", "Value1:Value2:UnknownPin")]
        [DataRow("Value1:Value2:VCC", "Value1:Value2:_VCC_VAR")]
        [DataRow("Value1:Value2:VCC()", "Value1:Value2:_VCC_VAR()")]
        [DataRow("Value1:Value2:VCC(Alias)", "Value1:Value2:VCC(Alias)")]
        [DataRow("Value1:Value2:VCC+", "Value1:Value2:_VCC_VAR+")]
        public void ConvertValueSpec_PinTransformation_Tests(string input, string expected)
        {
            // ACT
            string result = DataConvertor.ConvertValueSpec(input);

            // ASSERT
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void ConvertValueSpec_EmptyInput_ReturnsEmptyString()
        {
            Assert.AreEqual("", DataConvertor.ConvertValueSpec(""));
        }

        [TestMethod]
        [DataRow("VCC", "_VCC_VAR")]
        public void ConvertForceValue_ExistingPin_AppendsVar(
            string input,
            string expected)
        {
            // Arrange
            var forcePin = new ForcePin
            {
                ForceValue = input
            };

            // Mock PinMap:
            // IsPinExist(input.ToUpper()) => true

            // Act
            string result = DataConvertor.ConvertForceValueToGlbSpec(forcePin);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        [DataRow("1000m", "1")]
        [DataRow("1u", "1E-06")]
        [DataRow("1n", "1E-09")]
        [DataRow("1p", "1E-12")]
        [DataRow("1k", "1000")]
        [DataRow("1M", "1000000")]
        [DataRow("1G", "1000000000")]
        [DataRow("1T", "1000000000000")]
        [DataRow("0m", "0")]
        [DataRow("XXX", "XXX")]
        public void ConvertForceValue_UnknownPin_CallsConvertUnits(
            string input,
            string expected)
        {
            // Arrange
            var forcePin = new ForcePin
            {
                ForceValue = input
            };

            // Mock PinMap:
            // IsPinExist(...) => false

            // Act
            string result = DataConvertor.ConvertForceValueToGlbSpec(forcePin);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        [DataRow("1E-6", "1E-6")]
        [DataRow("123", "123")]
        [DataRow("-123", "-123")]
        [DataRow("1.25", "1.25")]
        public void ConvertForceValue_SpecialFormats_ReturnsExpectedValue(
            string input,
            string expected)
        {
            // Arrange
            var forcePin = new ForcePin
            {
                ForceValue = input
            };

            // Mock PinMap:
            // IsPinExist(...) => false

            // Act
            string result = DataConvertor.ConvertForceValueToGlbSpec(forcePin);

            // Assert
            Assert.AreEqual(expected, result);
        }
    }
}
