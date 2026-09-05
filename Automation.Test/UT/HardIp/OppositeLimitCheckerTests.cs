using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.HardIpPreCheck;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;

using CommonLib.ErrorReport;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class OppositeLimitCheckerTests
    {
        private OppositeLimitChecker _checker = null!;

        [TestInitialize]
        public void Setup()
        {
            var pattern = new HardIpPattern
            {
                SheetName = "TestSheet",
                RowNum = 1,
                Pattern = new PatternClass("")
            };
            var hardIpSheet = new HardIpSheet
            {
                PlanHeaderIdx = new Dictionary<string, int> { { "registerIndex", 1 }, { "miscInfoIndex", 2 } }
            };
            _checker = new OppositeLimitChecker(hardIpSheet, pattern);
            ErrorReportManager.ClearErrors();
        }

        [TestCleanup]
        public void Cleanup()
        {
            ErrorReportManager.ClearErrors();
        }

        private void InitLists(string unit = "")
        {
            _checker._oppositeLimitValue = [];
            _checker._wrongSyntax = [];
            _checker._unitMismatch = [];
            _checker._noUnit = [];
            _checker._noLimit = [];
            _checker._clearLimit = [];
            _checker._unit = unit;
        }

        private static MeasPin NewPin(string measType)
        {
            return new MeasPin { MeasType = measType };
        }

        private static MeasLimit NewLimit(string loLimit, string hiLimit, int loHeaderIndex = 1, int hiHeaderIndex = 2)
        {
            return new MeasLimit("job") { LoLimit = loLimit, HiLimit = hiLimit, LoHeaderIndex = loHeaderIndex, HiHeaderIndex = hiHeaderIndex };
        }

        #region GetMeasurePinUnit

        [DataTestMethod]
        [DataRow("MeasV", "V")]
        [DataRow("MeasI", "A")]
        [DataRow("MeasF", "Hz")]
        [DataRow("MeasR", "ohms?")]
        [DataRow("SomethingElse", "")]
        public void GetMeasurePinUnit_VariousMeasTypes_ReturnsExpectedUnit(string measType, string expected)
        {
            // Act
            string result = _checker.GetMeasurePinUnit(NewPin(measType));

            // Assert
            Assert.AreEqual(expected, result);
        }

        #endregion

        #region CheckOneUseLimit

        [TestMethod]
        public void CheckOneUseLimit_HighLessThanLow_AddsBothHeaderIndicesToOppositeLimit()
        {
            // Arrange
            InitLists();
            MeasLimit limit = NewLimit("10", "5", loHeaderIndex: 3, hiHeaderIndex: 4);

            // Act
            _checker.CheckOneUseLimit(limit, NewPin("MeasV"));

            // Assert
            List<int> oppositeLimitValue = _checker._oppositeLimitValue;
            Assert.AreEqual(2, oppositeLimitValue.Count);
            CollectionAssert.Contains(oppositeLimitValue, 3);
            CollectionAssert.Contains(oppositeLimitValue, 4);
        }

        [TestMethod]
        public void CheckOneUseLimit_HighGreaterThanLow_NoOppositeLimitAdded()
        {
            // Arrange
            InitLists();
            MeasLimit limit = NewLimit("5", "10");

            // Act
            _checker.CheckOneUseLimit(limit, NewPin("MeasV"));

            // Assert
            List<int> oppositeLimitValue = _checker._oppositeLimitValue;
            Assert.AreEqual(0, oppositeLimitValue.Count);
        }

        [TestMethod]
        public void CheckOneUseLimit_NonNumericLimits_SkipsOppositeCheck()
        {
            // Arrange - non-numeric limits can't be compared, so the opposite-limit check is skipped
            InitLists();
            MeasLimit limit = NewLimit("abc", "def");

            // Act
            _checker.CheckOneUseLimit(limit, NewPin("MeasV"));

            // Assert
            List<int> oppositeLimitValue = _checker._oppositeLimitValue;
            Assert.AreEqual(0, oppositeLimitValue.Count);
        }

        #endregion

        #region CheckOneSideUseLimit

        [TestMethod]
        public void CheckOneSideUseLimit_StartsWithBinning_ReturnsEarlyWithNoListPopulated()
        {
            // Arrange
            InitLists();
            MeasLimit limit = NewLimit("Binning123", "Binning456");

            // Act
            _checker.CheckOneSideUseLimit(limit, "Low", "MeasV");

            // Assert
            Assert.AreEqual(0, _checker._noLimit.Count);
            Assert.AreEqual(0, _checker._wrongSyntax.Count);
        }

        [TestMethod]
        public void CheckOneSideUseLimit_EmptyLimit_AddsToNoLimit()
        {
            // Arrange
            InitLists();
            MeasLimit limit = NewLimit("", "", loHeaderIndex: 7);

            // Act
            _checker.CheckOneSideUseLimit(limit, "Low", "MeasV");

            // Assert
            List<int> noLimit = _checker._noLimit;
            Assert.AreEqual(1, noLimit.Count);
            Assert.AreEqual(7, noLimit[0]);
        }

        [TestMethod]
        public void CheckOneSideUseLimit_VariableKeywordLowSide_ClearsLoLimitAndAddsToClearLimit()
        {
            // Arrange
            InitLists();
            MeasLimit limit = NewLimit("VARIABLE_X", "10", loHeaderIndex: 9);

            // Act
            _checker.CheckOneSideUseLimit(limit, "Low", "MeasV");

            // Assert
            Assert.AreEqual("", limit.LoLimit);
            List<int> clearLimit = _checker._clearLimit;
            Assert.AreEqual(1, clearLimit.Count);
            Assert.AreEqual(9, clearLimit[0]);
        }

        [TestMethod]
        public void CheckOneSideUseLimit_VariableKeywordHighSide_ClearsHiLimitAndAddsToClearLimit()
        {
            // Arrange
            InitLists();
            MeasLimit limit = NewLimit("10", "VARIABLE_X", hiHeaderIndex: 11);

            // Act
            _checker.CheckOneSideUseLimit(limit, "High", "MeasV");

            // Assert
            Assert.AreEqual("", limit.HiLimit);
            List<int> clearLimit = _checker._clearLimit;
            Assert.AreEqual(1, clearLimit.Count);
            Assert.AreEqual(11, clearLimit[0]);
        }

        [TestMethod]
        public void CheckOneSideUseLimit_UnrecognizedSyntax_AddsToWrongSyntax()
        {
            // Arrange - starts with a letter (not digit/./-), no VDD, no 0X/0B/0D prefix
            InitLists();
            MeasLimit limit = NewLimit("ABC", "", loHeaderIndex: 13);

            // Act
            _checker.CheckOneSideUseLimit(limit, "Low", "MeasV");

            // Assert
            List<int> wrongSyntax = _checker._wrongSyntax;
            Assert.AreEqual(1, wrongSyntax.Count);
            Assert.AreEqual(13, wrongSyntax[0]);
        }

        [TestMethod]
        public void CheckOneSideUseLimit_MatchingUnit_NoMismatchNoMissingUnit()
        {
            // Arrange
            InitLists(unit: "V");
            MeasLimit limit = NewLimit("5V", "", loHeaderIndex: 15);

            // Act
            _checker.CheckOneSideUseLimit(limit, "Low", "MeasV");

            // Assert
            Assert.AreEqual(0, _checker._unitMismatch.Count);
            Assert.AreEqual(0, _checker._noUnit.Count);
        }

        [TestMethod]
        public void CheckOneSideUseLimit_MismatchedUnit_AddsToUnitMismatch()
        {
            // Arrange - expects "V" unit but limit carries "A"
            InitLists(unit: "V");
            MeasLimit limit = NewLimit("5A", "", loHeaderIndex: 17);

            // Act
            _checker.CheckOneSideUseLimit(limit, "Low", "MeasV");

            // Assert
            List<int> unitMismatch = _checker._unitMismatch;
            Assert.AreEqual(1, unitMismatch.Count);
            Assert.AreEqual(17, unitMismatch[0]);
        }

        [TestMethod]
        public void CheckOneSideUseLimit_MissingUnitOnMeasurableType_AddsToNoUnit()
        {
            // Arrange - numeric limit with no unit suffix, on a MeasV-recognized type
            InitLists(unit: "V");
            MeasLimit limit = NewLimit("5", "", loHeaderIndex: 19);

            // Act
            _checker.CheckOneSideUseLimit(limit, "Low", "MeasV");

            // Assert
            List<int> noUnit = _checker._noUnit;
            Assert.AreEqual(1, noUnit.Count);
            Assert.AreEqual(19, noUnit[0]);
        }

        [TestMethod]
        public void CheckOneSideUseLimit_VddContainingLimit_NoNoUnitDespiteMissingUnit()
        {
            // Arrange - VDD-referencing limits are exempt from both wrong-syntax and no-unit checks
            InitLists(unit: "V");
            MeasLimit limit = NewLimit("VDD_CORE", "", loHeaderIndex: 21);

            // Act
            _checker.CheckOneSideUseLimit(limit, "Low", "MeasV");

            // Assert
            Assert.AreEqual(0, _checker._wrongSyntax.Count);
            Assert.AreEqual(0, _checker._noUnit.Count);
        }

        [TestMethod]
        public void CheckOneSideUseLimit_NonMeasurableType_NoNoUnitEvenWithoutUnit()
        {
            // Arrange - measType doesn't match MeasV|MeasI|MeasF|MeasR, so the no-unit check never applies
            InitLists(unit: "");
            MeasLimit limit = NewLimit("5", "", loHeaderIndex: 23);

            // Act
            _checker.CheckOneSideUseLimit(limit, "Low", "NotMeasurable");

            // Assert
            Assert.AreEqual(0, _checker._noUnit.Count);
        }

        #endregion

        #region Check integration

        [TestMethod]
        public void Check_TrimInstanceWithNonZeroSequenceAndNoIgnoreFlag_AddsMissingPinNameError()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                SheetName = "TestSheet",
                RowNum = 1,
                Pattern = new PatternClass(""),
                MiscInfoDict = new Dictionary<string, string> { { "trimCategory", "x" } },
                MeasPins =
                [
                    new MeasPin { MeasType = "MeasV", SequenceIndex = 1, MiscInfo = "" }
                ]
            };
            var hardIpSheet = new HardIpSheet
            {
                PlanHeaderIdx = new Dictionary<string, int> { { "registerIndex", 1 }, { "miscInfoIndex", 2 } }
            };
            var checker = new OppositeLimitChecker(hardIpSheet, pattern);

            // Act
            checker.Check();

            // Assert
            Assert.IsTrue(ErrorReportManager.GetErrorCount() >= 1);
        }

        [TestMethod]
        public void Check_NoTrimInstance_NoMissingPinNameError()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                SheetName = "TestSheet",
                RowNum = 1,
                Pattern = new PatternClass(""),
                MeasPins =
                [
                    new MeasPin { MeasType = "MeasV", SequenceIndex = 1, MiscInfo = "" }
                ]
            };
            var hardIpSheet = new HardIpSheet
            {
                PlanHeaderIdx = new Dictionary<string, int> { { "registerIndex", 1 }, { "miscInfoIndex", 2 } }
            };
            var checker = new OppositeLimitChecker(hardIpSheet, pattern);

            // Act
            checker.Check();

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        #endregion
    }
}
