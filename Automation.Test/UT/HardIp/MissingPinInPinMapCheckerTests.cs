using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.HardIpPreCheck;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Moq;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class MissingPinInPinMapCheckerTests : TestBase
    {
        private HardIpSheet _sheet = null!;
        private HardIpPattern _pattern = null!;
        private Mock<PinMapSheet> _mockPinMapSheet = null!;

        [TestInitialize]
        public void Setup()
        {
            ErrorReportManager.ClearErrors();

            _sheet = new HardIpSheet
            {
                PlanHeaderIdx = new Dictionary<string, int> { { "forceIndex", 3 }, { "measIndex", 5 } }
            };

            _pattern = new HardIpPattern { SheetName = "Sheet1", RowNum = 1 };

            _mockPinMapSheet = new Mock<PinMapSheet>("SomeName");
            var kvp = new KeyValuePair<string, PinMapSheet>("TestKey", _mockPinMapSheet.Object);
            TestProgram.IgxlWorkBk.Clear();
            TestProgram.IgxlWorkBk.PinMapPair = kvp;
        }

        private static List<Error> Errors()
        {
            return ErrorReportManager.GetErrorList();
        }

        #region CheckOneLimit (via Check -> CheckMeasPin's limit region)

        [TestMethod]
        public void Check_HiLimitReferencesMissingVddPin_AddsHiLimitError()
        {
            var pin = new MeasPin { PinName = "", MeasType = "Calc", RowNum = 10 };
            pin.MeasLimitsH.Add(new MeasLimit("job") { HiLimit = "VDD_CORE * 1.1" });
            _pattern.MeasPins.Add(pin);

            var checker = new MissingPinInPinMapChecker(_sheet, _pattern);
            checker.Check();

            Assert.IsTrue(Errors().Any(e => e.ErrorCode.FullCode == HardIpErrorType.E_MissingPinName_01.FullCode));
        }

        [TestMethod]
        public void Check_LoLimitReferencesMissingVddPin_AddsLoLimitError()
        {
            var pin = new MeasPin { PinName = "", MeasType = "Calc", RowNum = 10 };
            pin.MeasLimitsL.Add(new MeasLimit("job") { LoLimit = "VDD_CORE * 0.9" });
            _pattern.MeasPins.Add(pin);

            var checker = new MissingPinInPinMapChecker(_sheet, _pattern);
            checker.Check();

            Assert.IsTrue(Errors().Any(e => e.ErrorCode.FullCode == HardIpErrorType.E_MissingPinName_02.FullCode));
        }

        [TestMethod]
        public void Check_LimitDoesNotReferenceVddPin_NoLimitError()
        {
            var pin = new MeasPin { PinName = "", MeasType = "Calc", RowNum = 10 };
            pin.MeasLimitsH.Add(new MeasLimit("job") { HiLimit = "1.5" });
            _pattern.MeasPins.Add(pin);

            var checker = new MissingPinInPinMapChecker(_sheet, _pattern);
            checker.Check();

            Assert.IsFalse(Errors().Any(e =>
                e.ErrorCode.FullCode == HardIpErrorType.E_MissingPinName_01.FullCode ||
                e.ErrorCode.FullCode == HardIpErrorType.E_MissingPinName_02.FullCode));
        }

        [TestMethod]
        public void Check_HiLimitReferencesExistingVddPin_NoLimitError()
        {
            _mockPinMapSheet.Setup(x => x.IsPinExist("VDD_CORE")).Returns(true);
            var pin = new MeasPin { PinName = "", MeasType = "Calc", RowNum = 10 };
            pin.MeasLimitsH.Add(new MeasLimit("job") { HiLimit = "VDD_CORE" });
            _pattern.MeasPins.Add(pin);

            var checker = new MissingPinInPinMapChecker(_sheet, _pattern);
            checker.Check();

            Assert.IsFalse(Errors().Any(e => e.ErrorCode.FullCode == HardIpErrorType.E_MissingPinName_01.FullCode));
        }

        [TestMethod]
        public void Check_MoreThan500MissingHiLimitPins_CapsHiLimitErrorsAt500()
        {
            // Guards both the "_errorCnt < 500" boundary and the "_errorCnt++" increment for the Hi-limit
            // branch: without a real increment or with a "<= 500" boundary, the cap would not hold at 500.
            var pin = new MeasPin { PinName = "", MeasType = "Calc", RowNum = 10 };
            for (int i = 0; i < 501; i++)
            {
                pin.MeasLimitsH.Add(new MeasLimit("job") { HiLimit = $"VDD_PIN{i}" });
            }
            _pattern.MeasPins.Add(pin);

            var checker = new MissingPinInPinMapChecker(_sheet, _pattern);
            checker.Check();

            int count = Errors().Count(e => e.ErrorCode.FullCode == HardIpErrorType.E_MissingPinName_01.FullCode);
            Assert.AreEqual(500, count);
        }

        [TestMethod]
        public void Check_MoreThan500MissingLoLimitPins_CapsLoLimitErrorsAt500()
        {
            var pin = new MeasPin { PinName = "", MeasType = "Calc", RowNum = 10 };
            for (int i = 0; i < 501; i++)
            {
                pin.MeasLimitsL.Add(new MeasLimit("job") { LoLimit = $"VDD_PIN{i}" });
            }
            _pattern.MeasPins.Add(pin);

            var checker = new MissingPinInPinMapChecker(_sheet, _pattern);
            checker.Check();

            int count = Errors().Count(e => e.ErrorCode.FullCode == HardIpErrorType.E_MissingPinName_02.FullCode);
            Assert.AreEqual(500, count);
        }

        #endregion

        #region CheckPatternForcePins

        [TestMethod]
        public void Check_ForcePinWithDoubleColon_UsesTextBeforeColonAndAddsErrorWhenMissing()
        {
            var condition = new ForceCondition();
            condition.ForcePins.Add(new ForcePin { PinName = "PINA::extra" });
            _pattern.ForceConditionList.Add(condition);

            var checker = new MissingPinInPinMapChecker(_sheet, _pattern);
            checker.Check();

            Error error = Errors().Single(e => e.ErrorCode.FullCode == HardIpErrorType.E_MissingPinName_03.FullCode);
            CollectionAssert.Contains(error.MessageArgs!, "PINA");
        }

        [TestMethod]
        public void Check_ForcePinWithEquals_UsesTextAfterEqualsAndAddsErrorWhenMissing()
        {
            var condition = new ForceCondition();
            condition.ForcePins.Add(new ForcePin { PinName = "prefix=PINB" });
            _pattern.ForceConditionList.Add(condition);

            var checker = new MissingPinInPinMapChecker(_sheet, _pattern);
            checker.Check();

            Error error = Errors().Single(e => e.ErrorCode.FullCode == HardIpErrorType.E_MissingPinName_03.FullCode);
            CollectionAssert.Contains(error.MessageArgs!, "PINB");
        }

        [TestMethod]
        public void Check_ForcePinExists_NoForcePinError()
        {
            _mockPinMapSheet.Setup(x => x.IsPinExist("PINC")).Returns(true);
            var condition = new ForceCondition();
            condition.ForcePins.Add(new ForcePin { PinName = "PINC" });
            _pattern.ForceConditionList.Add(condition);

            var checker = new MissingPinInPinMapChecker(_sheet, _pattern);
            checker.Check();

            Assert.IsFalse(Errors().Any(e => e.ErrorCode.FullCode == HardIpErrorType.E_MissingPinName_03.FullCode));
        }

        [TestMethod]
        public void Check_MoreThan500MissingForcePins_CapsForcePinErrorsAt500()
        {
            for (int i = 0; i < 501; i++)
            {
                var condition = new ForceCondition();
                condition.ForcePins.Add(new ForcePin { PinName = $"FORCEPIN{i}" });
                _pattern.ForceConditionList.Add(condition);
            }

            var checker = new MissingPinInPinMapChecker(_sheet, _pattern);
            checker.Check();

            int count = Errors().Count(e => e.ErrorCode.FullCode == HardIpErrorType.E_MissingPinName_03.FullCode);
            Assert.AreEqual(500, count);
        }

        #endregion

        #region CheckSweepPins

        [TestMethod]
        public void Check_SweepPinMissing_AddsSweepPinError()
        {
            _pattern.SweepVoltage["X"] =
            [
                new("SWEEPPIN1,SWEEPPIN2", "0,1", "", "1", "")
            ];

            var checker = new MissingPinInPinMapChecker(_sheet, _pattern);
            checker.Check();

            List<Error> errors = [.. Errors().Where(e => e.ErrorCode.FullCode == HardIpErrorType.E_MissingPinName_04.FullCode)];
            Assert.AreEqual(2, errors.Count);
        }

        [TestMethod]
        public void Check_SweepPinsExist_NoSweepPinError()
        {
            _mockPinMapSheet.Setup(x => x.IsPinExist("SWEEPPIN1")).Returns(true);
            _pattern.SweepVoltage["X"] =
            [
                new("SWEEPPIN1", "0,1", "", "1", "")
            ];

            var checker = new MissingPinInPinMapChecker(_sheet, _pattern);
            checker.Check();

            Assert.IsFalse(Errors().Any(e => e.ErrorCode.FullCode == HardIpErrorType.E_MissingPinName_04.FullCode));
        }

        [TestMethod]
        public void Check_MoreThan500MissingSweepPins_CapsSweepPinErrorsAt500()
        {
            string pinNames = string.Join(",", Enumerable.Range(0, 501).Select(i => $"SWEEPPIN{i}"));
            _pattern.SweepVoltage["X"] =
            [
                new(pinNames, "0,1", "", "1", "")
            ];

            var checker = new MissingPinInPinMapChecker(_sheet, _pattern);
            checker.Check();

            int count = Errors().Count(e => e.ErrorCode.FullCode == HardIpErrorType.E_MissingPinName_04.FullCode);
            Assert.AreEqual(500, count);
        }

        #endregion

        #region CheckMeasPin - meas pin existence

        [TestMethod]
        public void Check_MeasPinMissingAndMeasTypeNormal_AddsMissingMeasPinError()
        {
            _pattern.MeasPins.Add(new MeasPin { PinName = "MEASPIN1", MeasType = "V", RowNum = 20 });

            var checker = new MissingPinInPinMapChecker(_sheet, _pattern);
            checker.Check();

            Assert.IsTrue(Errors().Any(e => e.ErrorCode.FullCode == HardIpErrorType.E_MissingPinName_05.FullCode));
        }

        [TestMethod]
        public void Check_MeasPinNameEmpty_SkipsExistenceCheckWithoutAddingError()
        {
            // An empty tmpPin segment must be skipped entirely (never even checked for existence);
            // guards against a mutation to the empty-string literal it's compared against.
            _pattern.MeasPins.Add(new MeasPin { PinName = "", MeasType = "V", RowNum = 20 });

            var checker = new MissingPinInPinMapChecker(_sheet, _pattern);
            checker.Check();

            Assert.IsFalse(Errors().Any(e => e.ErrorCode.FullCode == HardIpErrorType.E_MissingPinName_05.FullCode));
        }

        [TestMethod]
        public void Check_MeasPinMissingButMeasTypeIsCalc_NoMissingMeasPinError()
        {
            _pattern.MeasPins.Add(new MeasPin { PinName = "MEASPIN1", MeasType = "Calc", RowNum = 20 });

            var checker = new MissingPinInPinMapChecker(_sheet, _pattern);
            checker.Check();

            Assert.IsFalse(Errors().Any(e => e.ErrorCode.FullCode == HardIpErrorType.E_MissingPinName_05.FullCode));
        }

        [TestMethod]
        public void Check_MeasPinMissingButMeasTypeIsLimit_NoMissingMeasPinError()
        {
            _pattern.MeasPins.Add(new MeasPin { PinName = "MEASPIN1", MeasType = "Limit", RowNum = 20 });

            var checker = new MissingPinInPinMapChecker(_sheet, _pattern);
            checker.Check();

            Assert.IsFalse(Errors().Any(e => e.ErrorCode.FullCode == HardIpErrorType.E_MissingPinName_05.FullCode));
        }

        [TestMethod]
        public void Check_MeasPinExists_NoMissingMeasPinError()
        {
            _mockPinMapSheet.Setup(x => x.IsPinExist("MEASPIN1")).Returns(true);
            _pattern.MeasPins.Add(new MeasPin { PinName = "MEASPIN1", MeasType = "V", RowNum = 20 });

            var checker = new MissingPinInPinMapChecker(_sheet, _pattern);
            checker.Check();

            Assert.IsFalse(Errors().Any(e => e.ErrorCode.FullCode == HardIpErrorType.E_MissingPinName_05.FullCode));
        }

        [TestMethod]
        public void Check_MeasPinNameWithDoubleColon_ParsesTextBeforeColon()
        {
            _mockPinMapSheet.Setup(x => x.IsPinExist("MEASPINA")).Returns(true);
            _pattern.MeasPins.Add(new MeasPin { PinName = "MEASPINA::extra", MeasType = "V", RowNum = 20 });

            var checker = new MissingPinInPinMapChecker(_sheet, _pattern);
            checker.Check();

            Assert.IsFalse(Errors().Any(e => e.ErrorCode.FullCode == HardIpErrorType.E_MissingPinName_05.FullCode));
        }

        [TestMethod]
        public void Check_MeasPinNameWithEquals_ParsesTextAfterEquals()
        {
            _mockPinMapSheet.Setup(x => x.IsPinExist("MEASPINB")).Returns(true);
            _pattern.MeasPins.Add(new MeasPin { PinName = "prefix=MEASPINB", MeasType = "V", RowNum = 20 });

            var checker = new MissingPinInPinMapChecker(_sheet, _pattern);
            checker.Check();

            Assert.IsFalse(Errors().Any(e => e.ErrorCode.FullCode == HardIpErrorType.E_MissingPinName_05.FullCode));
        }

        #endregion

        #region CheckMeasPin - force pins of meas pin

        [TestMethod]
        public void Check_ForcePinOfMeasPinMissingAndNotInGroupList_AddsForcePinOfMeasPinError()
        {
            var pin = new MeasPin { PinName = "MEASPIN1", MeasType = "V", RowNum = 20 };
            _mockPinMapSheet.Setup(x => x.IsPinExist("MEASPIN1")).Returns(true);
            var condition = new ForceCondition();
            condition.ForcePins.Add(new ForcePin { PinName = "FORCEPINX" });
            pin.ForceConditions.Add(condition);
            _pattern.MeasPins.Add(pin);

            var checker = new MissingPinInPinMapChecker(_sheet, _pattern);
            checker.Check();

            Assert.IsTrue(Errors().Any(e => e.ErrorCode.FullCode == HardIpErrorType.E_MissingPinName_06.FullCode));
        }

        [TestMethod]
        public void Check_ForcePinOfMeasPinMissingButPresentInGroupList_NoForcePinOfMeasPinError()
        {
            var pinGroup = new PinGroup("FORCEPINY");
            pinGroup.AddPin("MEMBERPIN1");
            _mockPinMapSheet.Object.AddGroup(pinGroup);
            var pin = new MeasPin { PinName = "MEASPIN1", MeasType = "V", RowNum = 20 };
            _mockPinMapSheet.Setup(x => x.IsPinExist("MEASPIN1")).Returns(true);
            var condition = new ForceCondition();
            condition.ForcePins.Add(new ForcePin { PinName = "FORCEPINY" });
            pin.ForceConditions.Add(condition);
            _pattern.MeasPins.Add(pin);

            var checker = new MissingPinInPinMapChecker(_sheet, _pattern);
            checker.Check();

            Assert.IsFalse(Errors().Any(e => e.ErrorCode.FullCode == HardIpErrorType.E_MissingPinName_06.FullCode));
        }

        [TestMethod]
        public void Check_ForcePinOfMeasPinExists_NoForcePinOfMeasPinError()
        {
            var pin = new MeasPin { PinName = "MEASPIN1", MeasType = "V", RowNum = 20 };
            _mockPinMapSheet.Setup(x => x.IsPinExist("MEASPIN1")).Returns(true);
            _mockPinMapSheet.Setup(x => x.IsPinExist("FORCEPINZ")).Returns(true);
            var condition = new ForceCondition();
            condition.ForcePins.Add(new ForcePin { PinName = "FORCEPINZ" });
            pin.ForceConditions.Add(condition);
            _pattern.MeasPins.Add(pin);

            var checker = new MissingPinInPinMapChecker(_sheet, _pattern);
            checker.Check();

            Assert.IsFalse(Errors().Any(e => e.ErrorCode.FullCode == HardIpErrorType.E_MissingPinName_06.FullCode));
        }

        [TestMethod]
        public void Check_ManyMissingForcePinsOfMeasPin_ForcePinOfMeasPinErrorHasNoCap()
        {
            // Unlike the other regions, this branch increments _errorCnt but never checks it against the
            // 500 cap, so every missing force-pin-of-meas-pin still gets reported.
            _mockPinMapSheet.Setup(x => x.IsPinExist("MEASPINSHARED")).Returns(true);
            for (int i = 0; i < 501; i++)
            {
                var pin = new MeasPin { PinName = "MEASPINSHARED", MeasType = "V", RowNum = 20 + i };
                var condition = new ForceCondition();
                condition.ForcePins.Add(new ForcePin { PinName = $"FORCEPINOFMEAS{i}" });
                pin.ForceConditions.Add(condition);
                _pattern.MeasPins.Add(pin);
            }

            var checker = new MissingPinInPinMapChecker(_sheet, _pattern);
            checker.Check();

            int count = Errors().Count(e => e.ErrorCode.FullCode == HardIpErrorType.E_MissingPinName_06.FullCode);
            Assert.AreEqual(501, count);
        }

        #endregion
    }
}
