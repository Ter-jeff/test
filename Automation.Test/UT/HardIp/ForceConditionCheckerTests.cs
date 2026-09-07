using System.Linq;

using Automation.GenerateIgxl.HardIp.HardIpPreCheck;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class ForceConditionCheckerTests
    {
        private HardIpPattern _pattern = null!;
        private ForceConditionChecker _checker = null!;

        [TestInitialize]
        public void Setup()
        {
            _pattern = new HardIpPattern
            {
                SheetName = "TestSheet",
                RowNum = 1,
                Pattern = new PatternClass("")
            };
            var hardIpSheet = new HardIpSheet { ForceIndex = 3 };
            _checker = new ForceConditionChecker(hardIpSheet, _pattern);
            ErrorReportManager.ClearErrors();
        }

        [TestCleanup]
        public void Cleanup()
        {
            ErrorReportManager.ClearErrors();
        }

        private static ForcePin NewForcePin(ForceConditionType type, string forceType = "", string forceValue = "", string pinName = "PIN1")
        {
            return new ForcePin { Type = type, ForceType = forceType, ForceValue = forceValue, PinName = pinName };
        }

        #region CheckForcePin

        [TestMethod]
        public void CheckForcePin_NormalType_InvalidForceType_AddsError()
        {
            // Act
            _checker.CheckForcePin(_pattern, NewForcePin(ForceConditionType.Normal, forceType: "INVALID"), null);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckForcePin_NormalType_ValidForceType_NoError()
        {
            // Act
            _checker.CheckForcePin(_pattern, NewForcePin(ForceConditionType.Normal, forceType: "VT"), null);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckForcePin_OthersType_InvalidForceMethod_AddsError()
        {
            // Act
            _checker.CheckForcePin(_pattern, NewForcePin(ForceConditionType.Others, forceValue: "INVALID"), null);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckForcePin_OthersType_ValidForceMethod_NoError()
        {
            // Act
            _checker.CheckForcePin(_pattern, NewForcePin(ForceConditionType.Others, forceValue: "RELAYON"), null);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckForcePin_MeasPinProvided_UsesMeasPinRowNum()
        {
            // Arrange
            var measPin = new MeasPin { RowNum = 99 };

            // Act
            _checker.CheckForcePin(_pattern, NewForcePin(ForceConditionType.Normal, forceType: "INVALID"), measPin);

            // Assert
            Error error = ErrorReportManager.GetErrorList().Last();
            Assert.AreEqual(99, error.RowNum);
        }

        [TestMethod]
        public void CheckForcePin_NoMeasPin_UsesPatternRowNum()
        {
            // Act
            _checker.CheckForcePin(_pattern, NewForcePin(ForceConditionType.Normal, forceType: "INVALID"), null);

            // Assert
            Error error = ErrorReportManager.GetErrorList().Last();
            Assert.AreEqual(1, error.RowNum);
        }

        #endregion

        #region CheckRelay

        [TestMethod]
        public void CheckRelay_OthersType_RelayOnValue_AddsWarning()
        {
            // Act
            _checker.CheckRelay(_pattern, NewForcePin(ForceConditionType.Others, forceValue: "RelayOn"));

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckRelay_OthersType_RelayOffValue_AddsWarning()
        {
            // Act
            _checker.CheckRelay(_pattern, NewForcePin(ForceConditionType.Others, forceValue: "RelayOff"));

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckRelay_OthersType_BareOffValue_MatchesRegexQuirkAndWarns()
        {
            // Arrange - the regex "relay_*On|Off" alternates loosely as (relay_*On)|(Off), so a
            // bare "Off" with no "relay" prefix still matches the second alternative.
            // Act
            _checker.CheckRelay(_pattern, NewForcePin(ForceConditionType.Others, forceValue: "Off"));

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckRelay_OthersType_UnrelatedValue_NoWarning()
        {
            // Act
            _checker.CheckRelay(_pattern, NewForcePin(ForceConditionType.Others, forceValue: "CONNECTPPMU"));

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckRelay_NormalType_NoWarningRegardlessOfValue()
        {
            // Act
            _checker.CheckRelay(_pattern, NewForcePin(ForceConditionType.Normal, forceValue: "RelayOn"));

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        #endregion

        #region CheckAcSelector

        [TestMethod]
        public void CheckAcSelector_ArgCountNotThree_AddsWrongFormatError()
        {
            // Arrange - the extra trailing segment slips past GetAcSelector()'s unanchored regex
            _pattern.ForceCondition = new ForceClass { ForceCondition = "ACSelector:A&B:TYP:EXTRA" };

            // Act
            _checker.CheckAcSelector();

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckAcSelector_InvalidVoltage_AddsError()
        {
            // Arrange
            _pattern.ForceCondition = new ForceClass { ForceCondition = "ACSelector:INVALID:Typ" };

            // Act
            _checker.CheckAcSelector();

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckAcSelector_InvalidType_AddsError()
        {
            // Arrange
            _pattern.ForceCondition = new ForceClass { ForceCondition = "ACSelector:NV:INVALID" };

            // Act
            _checker.CheckAcSelector();

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckAcSelector_ValidVoltageAndType_NoError()
        {
            // Arrange
            _pattern.ForceCondition = new ForceClass { ForceCondition = "ACSelector:NV:Typ" };

            // Act
            _checker.CheckAcSelector();

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        #endregion

        #region CheckMeasPinForceCondition

        [TestMethod]
        public void CheckMeasPinForceCondition_NoForceConditions_ReturnsEarlyNoError()
        {
            // Arrange
            var pin = new MeasPin { ForceConditions = [] };

            // Act
            _checker.CheckMeasPinForceCondition(pin);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckMeasPinForceCondition_EmptyMeasTypeWithVddPin_AddsWrongForceConditionError()
        {
            // Arrange - "VDD"-prefixed pin name resolves to "power" type, safely bypassing the
            // shared TestProgram-backed pin-map lookup used by the I/O mismatch check
            var pin = new MeasPin
            {
                MeasType = "",
                PinName = "VDD_CORE",
                RowNum = 5,
                ForceConditions =
                [
                    new() { ForcePins = [NewForcePin(ForceConditionType.Normal, forceType: "VT")] }
                ]
            };

            // Act
            _checker.CheckMeasPinForceCondition(pin);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckMeasPinForceCondition_MeasCMeasType_AddsError()
        {
            // Arrange
            var pin = new MeasPin
            {
                MeasType = "MeasC",
                PinName = "VDD_CORE",
                RowNum = 5,
                ForceConditions =
                [
                    new() { ForcePins = [NewForcePin(ForceConditionType.Normal, forceType: "VT")] }
                ]
            };

            // Act
            _checker.CheckMeasPinForceCondition(pin);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckMeasPinForceCondition_ValidMeasTypeWithVddPin_NoError()
        {
            // Arrange
            var pin = new MeasPin
            {
                MeasType = "MeasV",
                PinName = "VDD_CORE",
                RowNum = 5,
                ForceConditions =
                [
                    new() { ForcePins = [NewForcePin(ForceConditionType.Normal, forceType: "VT")] }
                ]
            };

            // Act
            _checker.CheckMeasPinForceCondition(pin);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        #endregion

        #region Check integration

        [TestMethod]
        public void Check_ForceConditionListNonEmpty_InvokesCheckForcePin()
        {
            // Arrange
            _pattern.ForceConditionList =
            [
                new() { ForcePins = [NewForcePin(ForceConditionType.Normal, forceType: "INVALID")] }
            ];

            // Act
            _checker.Check();

            // Assert
            Assert.IsTrue(ErrorReportManager.GetErrorCount() >= 1);
        }

        [TestMethod]
        public void Check_EmptyTimeSetUsedAndAcSelector_NoThrowNoExtraChecks()
        {
            // Arrange - TimeSetUsed.McgSetting and AcSelectorUsed both default to empty/unset
            _pattern.TimeSetUsed.McgSetting = "";

            // Act
            _checker.Check();

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void Check_MeasPinsWithEmptyMeasType_AddsErrorViaMeasPinLoop()
        {
            // Arrange
            _pattern.MeasPins.Add(new MeasPin { MeasType = "", PinName = "VDD_CORE", RowNum = 7, ForceConditions = [new()] });

            // Act
            _checker.Check();

            // Assert
            Assert.IsTrue(ErrorReportManager.GetErrorCount() >= 1);
        }

        #endregion
    }
}
