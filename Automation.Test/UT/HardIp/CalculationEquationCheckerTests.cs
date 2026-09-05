using System.Collections.Generic;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.HardIpPreCheck;
using Automation.GenerateIgxl.HardIp.InputObject;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class CalculationEquationCheckerTests : FunctionTestBase
    {
        private HardIpSheet _sheet = null!;
        private HardIpPattern _pattern = null!;
        private Dictionary<string, string> _storeNameDir = null!;
        private CalculationEquationChecker _checker = null!;

        [TestInitialize]
        public void Setup()
        {
            _sheet = new HardIpSheet
            {
                PlanHeaderIdx = new Dictionary<string, int> { { "miscInfoIndex", 1 } }
            };

            _pattern = new HardIpPattern
            {
                CalcEqn = "PIN1:1:Calc_ADC_Convert_Average(A,B,C)",
                MeasPins =
                [
                    new() { CusStr = "TEST1", PinName = "P1", MeasType = MeasType.MeasLimit },
                    new() { CusStr = "TEST2", PinName = "P2", MeasType = MeasType.MeasCalc }
                ],
                SheetName = "TestSheet",
                RowNum = 10,
                MiscInfo = "calc;TestCalcFunc"
            };

            _storeNameDir = new Dictionary<string, string>
            {
                { "EXISTING_PARA", "P1" }
            };

            _checker = new CalculationEquationChecker(_sheet, _pattern, _storeNameDir);

            ErrorReportManager.ClearErrors();
        }

        [TestMethod]
        public void Check_ShouldReportMissingDictionaryKey()
        {
            // Act
            _checker.Check();

            // Assert
            List<Error> errors = ErrorReportManager.GetErrorList();
            Assert.IsTrue(errors.Count > 0, "Checker should report errors when missing dictionary keys exist.");
            Assert.IsTrue(errors.Exists(e => e.Message.Contains("Missing Dictionary key")), "Error message should contain 'Missing Dictionary key'");
        }

        [TestMethod]
        public void SingleInput_ShouldMoveToDspAndReturnCorrectString()
        {
            // Arrange
            var calc = new List<string> { "CalcFunc(A,B,C,D,E,F)" };
            bool isAdcMoveToDsp = false;

            // Act
            string result = _checker.ProcessAdcConvertAverage(calc, ref isAdcMoveToDsp);

            // Assert
            Assert.IsTrue(isAdcMoveToDsp, "Single input should set isAdcMoveToDsp = true");
            StringAssert.StartsWith(result, "Calc_ADC_Convert_Average", "Result should start with proper function name");
            StringAssert.Contains(result, "A,B,C,D", "First 4 args should be preserved");
            StringAssert.Contains(result, "E&F", "Remaining args should be combined with &");
        }

        [TestMethod]
        public void MultipleInput_SameLength_ShouldReturnProperString()
        {
            // Arrange
            var calc = new List<string>
            {
                "CalcFunc(A,B,C,D,E1,F1)",
                "CalcFunc(A,B,C,D,E2,F2)"
            };
            bool isAdcMoveToDsp = false;

            // Act
            string result = _checker.ProcessAdcConvertAverage(calc, ref isAdcMoveToDsp);

            // Assert
            Assert.IsTrue(isAdcMoveToDsp, "All inputs same length should set isAdcMoveToDsp = true");
            StringAssert.StartsWith(result, "Calc_ADC_Convert_Average", "Result should start with proper function name");
            StringAssert.Contains(result, "A,B,C,D", "First 4 args should be preserved");
            StringAssert.Contains(result, "E1&E2&F1&F2", "Remaining args should be combined with &");
        }

        [TestMethod]
        public void SingleInput_EmptyOrInvalid_ShouldReturnEmptyString()
        {
            // Arrange
            var calc = new List<string> { "CalcFunc()" };
            bool isAdcMoveToDsp = false;

            // Act
            string result = _checker.ProcessAdcConvertAverage(calc, ref isAdcMoveToDsp);

            // Assert
            Assert.IsTrue(isAdcMoveToDsp, "Single input should set isAdcMoveToDsp = true even if empty args");
            Assert.AreEqual("Calc_ADC_Convert_Average()", result, "Result string should handle empty args");
        }

        [TestMethod]
        public void MultipleInput_DifferentLengths_ShouldReturnEmptyString()
        {
            // Arrange
            var calc = new List<string>
            {
                "CalcFunc(A,B,C,D,E)",
                "CalcFunc(A,B,C,D)"
            };
            bool isAdcMoveToDsp = false;

            // Act
            string result = _checker.ProcessAdcConvertAverage(calc, ref isAdcMoveToDsp);

            // Assert
            Assert.IsFalse(isAdcMoveToDsp, "Different length inputs should NOT move to DSP");
            Assert.AreEqual("", result, "Result should be empty string for unequal length inputs");
        }

        [TestMethod]
        public void ProcessGrayCode_ShouldReturnSignedGray_WhenFalseArgument()
        {
            var calc = new List<string> { "Calc_GrayCode(False,REG1)" };
            var captures = new List<string> { "REG1" };

            string result = _checker.ProcessGrayCode(calc, captures);

            Assert.AreEqual("SignedGray:REG1", result);
        }

        [TestMethod]
        public void ProcessGrayCode_ShouldReturnUnSignedGray_WhenTrueArgument()
        {
            var calc = new List<string> { "Calc_GrayCode(True,REG_A)" };
            var captures = new List<string> { "REG_A" };

            string result = _checker.ProcessGrayCode(calc, captures);

            Assert.AreEqual("UnSignedGray:REG_A", result);
        }

        [TestMethod]
        public void ProcessGrayCode_ShouldMatchCaseInsensitive_Arguments()
        {
            var calc = new List<string> { "Calc_GrayCode(false,regX)" };
            var captures = new List<string> { "REGX" };

            string result = _checker.ProcessGrayCode(calc, captures);

            Assert.AreEqual("SignedGray:REGX", result);
        }

        [TestMethod]
        public void ProcessGrayCode_ShouldReturnMultipleGrayCodes_WhenMultipleCalcInputs()
        {
            var calc = new List<string>
            {
                "Calc_GrayCode(False,REG1)",
                "Calc_GrayCode(True,REG2)"
            };
            var captures = new List<string> { "REG1", "REG2" };

            string result = _checker.ProcessGrayCode(calc, captures);
            List<string> parts = [.. result.Split(';')];

            Assert.AreEqual(2, parts.Count);
            Assert.IsTrue(parts.Contains("SignedGray:REG1"));
            Assert.IsTrue(parts.Contains("UnSignedGray:REG2"));
        }

        [TestMethod]
        public void ProcessGrayCode_ShouldReturnEmpty_WhenNoMatchingCapture()
        {
            var calc = new List<string> { "Calc_GrayCode(False,REG_UNKNOWN)" };
            var captures = new List<string> { "REG1" };

            string result = _checker.ProcessGrayCode(calc, captures);

            Assert.AreEqual("SignedGray:", result);
        }

        [TestMethod]
        public void ProcessGrayCode_ShouldReturnEmpty_WhenCalcEmpty()
        {
            string result = _checker.ProcessGrayCode([], []);
            Assert.AreEqual("", result);
        }
    }
}
