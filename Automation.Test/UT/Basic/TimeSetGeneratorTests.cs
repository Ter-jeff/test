using System.Collections.Generic;

using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Business;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.BinCut.BinCutInstance;

namespace Automation.Test.UT.Basic
{
    [TestClass]
    public class TimeSetGeneratorTests
    {
        private TimeSetGenerator _generator = null!;
        private TimeSetSheets _multiTimeSetSheets = null!;

        [TestInitialize]
        public void Setup()
        {
            _generator = new TimeSetGenerator();
            _multiTimeSetSheets = [];

            var timeSetBasic = new ComTimeSetBasic
            {
                Name = "TS1",
                CyclePeriod = "1.0",
                TimingRows =
                [
                    new()
                    {
                        DriveOn = "1E-6",
                        DriveData = "2E-6",
                        DriveReturn = "3E-6",
                        DriveOff = "0",
                        CompareOpen = "0",
                        DataFmt = "RL"
                    }
                ],
                SubContextVariable = ["FREQ1"],
                SubCommentVariable = []
            };

            var sheet = new ComTimeSetBasicSheet("TS1");
            sheet.Rows.Add(timeSetBasic);
            _multiTimeSetSheets.Add(sheet);
        }

        [TestMethod]
        public void ChangeTimeSetValueDetail_ShouldReturnFormulaString()
        {
            // Arrange
            var allCyclePeriod = new List<string> { "1.0", "2.0", "3.0", "0", "4.0" };
            string shiftInFreq = "ShiftInFreq";

            // Act
            string result = _generator.ChangeTimeSetValueDetail("1.0", allCyclePeriod, shiftInFreq, 1);

            // Assert
            Assert.IsTrue(result.Contains('='));
            Assert.IsTrue(result.Contains(shiftInFreq));
        }

        [TestMethod]
        public void JudgeTimeSetIsEqBased_ShouldReturnTrueIfRegexMatches()
        {
            // Arrange
            var timeSet = new ComTimeSetBasic
            {
                CyclePeriod = "1.0",
                TimingRows =
                [
                    new()
                    {
                        DriveOn = "some/path/Pattern.atp.gz"
                    }
                ]
            };

            // Act
            bool result = _generator.JudgeTimeSetIsEqBased(timeSet);

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void ChangeTimeSetValueForEqnBase_ShouldUpdateTimingRows()
        {
            var timeSetBasic = (ComTimeSetBasic)_multiTimeSetSheets[0].Rows[0];

            // Act
            _generator.ChangeTimeSetValueForEqnBase(timeSetBasic, false);

            // Assert
            foreach (TimingRow row in timeSetBasic.TimingRows)
            {
                Assert.IsTrue(row.DriveOn.StartsWith('=') || row.DriveOn == "0");
                Assert.IsTrue(row.DriveData.StartsWith('=') || row.DriveData == "0");
                Assert.IsTrue(row.DriveReturn.StartsWith('=') || row.DriveReturn == "0");
                Assert.IsTrue(row.DriveOff == "0");
                Assert.IsTrue(row.CompareOpen.StartsWith('=') || row.CompareOpen == "0");
            }
        }

        [DataTestMethod]
        [DataRow("=1+1", "2", DisplayName = "SimpleAddition")]
        [DataRow("=1E-6", "0.000001", DisplayName = "ScientificNotation")]
        public void GetEgValueInDecimal_ValidExpression_ReturnsParsedDecimal(string input, string expected)
        {
            // Act
            decimal result = _generator.GetEgValueInDecimal(input);

            // Assert
            Assert.AreEqual(decimal.Parse(expected), result);
        }

        [TestMethod]
        public void GetEgValueInDecimal_ExpressionYieldsNonDecimalResult_ThrowsException()
        {
            // Arrange - DataTable.Compute evaluates "1=1" to the boolean "True", which is not
            // decimal-parseable, triggering the method's own format-validation exception.
            Assert.ThrowsException<System.Exception>(() => _generator.GetEgValueInDecimal("=1=1"));
        }

        [DataTestMethod]
        [DataRow("=_TCK_Freq_VAR", true, DisplayName = "SingleVariable")]
        [DataRow("123", false, DisplayName = "NoUnderscore")]
        public void IsContextVariable_DetectsUnderscorePrefixedToken(string cell, bool expected)
        {
            // Act
            bool result = _generator.IsContextVariable(cell);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void GetContextVariable_SingleToken_AddsFullUnderscoreJoinedVariable()
        {
            // Arrange
            var subContextVar = new List<string>();

            // Act
            _generator.GetContextVariable("=_TCK_Freq_VAR", ref subContextVar);

            // Assert
            Assert.AreEqual(1, subContextVar.Count);
            Assert.AreEqual("TCK_Freq_VAR", subContextVar[0]);
        }

        [TestMethod]
        public void GetContextVariable_AlreadyPresent_DoesNotAddDuplicate()
        {
            // Arrange
            var subContextVar = new List<string> { "TCK_Freq_VAR" };

            // Act
            _generator.GetContextVariable("=_TCK_Freq_VAR", ref subContextVar);

            // Assert
            Assert.AreEqual(1, subContextVar.Count);
        }

        [DataTestMethod]
        [DataRow("N/A", true, DisplayName = "SlashNA")]
        [DataRow("na", true, DisplayName = "NaCaseInsensitive")]
        [DataRow("SomeFile.csv", false, DisplayName = "RealFileName")]
        public void IgnoredFileName_DetectsNaSentinelValues(string fileName, bool expected)
        {
            // Act
            bool result = _generator.IgnoredFileName(fileName);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void Converter_HeaderMatchesV23_ReturnsTimeRow2P3Converter()
        {
            // Act
            TimeRow1P4Converter result = TimeSetGenerator.Converter("DTTimesetBasicSheet,version=2.3");

            // Assert
            Assert.IsInstanceOfType(result, typeof(TimeRow2P3Converter));
        }

        [TestMethod]
        public void Converter_HeaderDoesNotMatch_ReturnsBaseConverter()
        {
            // Act
            TimeRow1P4Converter result = TimeSetGenerator.Converter("SomeOtherHeader");

            // Assert
            Assert.IsFalse(result is TimeRow2P3Converter);
        }

        [TestMethod]
        public void GetOriTimeSet_TimeSetWithColon_ReturnsPartBeforeColon()
        {
            // Arrange
            var row = new BinCutInstanceRow { TimeSet = "TS1:5" };

            // Act
            string result = _generator.GetOriTimeSet(row);

            // Assert
            Assert.AreEqual("TS1", result);
        }

        [TestMethod]
        public void GetOriTimeSet_TimeSetWithoutColon_ReturnsUnchanged()
        {
            // Arrange
            var row = new BinCutInstanceRow { TimeSet = "TS2" };

            // Act
            string result = _generator.GetOriTimeSet(row);

            // Assert
            Assert.AreEqual("TS2", result);
        }
    }
}
