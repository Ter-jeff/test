using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.HardIpPreCheck;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class PinSeqCheckerTests
    {
        private HardIpSheet _sheet = null!;
        private PinSeqChecker _checker = null!;

        private HardIpPattern _patternItem = null!;
        private HardIpInfo _patInfo = null!;

        [TestInitialize]
        public void Setup()
        {
            _sheet = new HardIpSheet
            {
                PlanHeaderIdx = { ["measIndex"] = 7 }
            };
            var patternInner = new PatternClass("dp_brna0_c_fulp_an_aa22_frq_jtg_dyc_allfrv_si_ciopll_t2p1r")
            {
                RealPatternName = "dp_brna0_c_fulp_an_aa22_frq_jtg_dyc_allfrv_si_ciopll_t2p1r"
            };

            _patternItem = new HardIpPattern
            {
                Pattern = patternInner,
                MiscInfo = "TEST_SEQUENCE:A,B,C",
                MeasPins = [new() { MeasType = "measv", PinCount = 1 }]
            };

            _checker = new PinSeqChecker(_sheet, _patternItem);

            _patInfo = new HardIpInfo
            {
                SeqInfo = [],
                Payload = "TestPayload"
            };

            ErrorReportManager.ClearErrors();
        }

        [TestMethod]
        public void PinsNum_PatInfo_Should_Skip_Sequence_When_Marked_N()
        {
            // Arrange
            var seqList = new List<string> { "Y", "N", "Y" };
            var infoList = new List<HardIpSeqInfo>
            {
                new() { PinList = ["A", "B"] },
                new() { PinList = ["C"] },
                new() { PinList = ["D"] }
            };

            // Act
            int result = _checker.PinsNum_PatInfo(infoList, seqList);

            // Assert
            Assert.AreEqual(3, result);
        }

        [TestMethod]
        public void PinsNum_PatInfo_Should_Count_All_When_No_Misc_List()
        {
            // Arrange
            var infoList = new List<HardIpSeqInfo>
            {
                new() { PinList = ["A", "B"] },
                new() { PinList = ["C"] }
            };

            // Act
            int result = _checker.PinsNum_PatInfo(infoList, null);

            // Assert
            Assert.AreEqual(3, result);
        }

        [TestMethod]
        public void CheckSeq_ShouldHandleDifferentialPins_WithReversedNames()
        {
            // Arrange
            _patternItem.MeasPins.Add(new MeasPin { PinName = "P::N", MeasType = "MeasVdiff", RowNum = 1 });

            _patInfo.SeqInfo.Add(new HardIpSeqInfo
            {
                SeqName = "Vdiff",
                PinList = ["N::P"]
            });

            // Act
            _checker.CheckSeq(_patternItem, _patInfo);

            // Assert
            Assert.IsTrue(ErrorReportManager.GetErrorCount() == 1, "Should match differential pins even if names are reversed.");
        }

        [TestMethod]
        public void CheckSeq_ShouldSkipSequence_WhenMiscInfoIsN()
        {
            // Arrange
            _checker._seqlstInMisc = ["N"];

            _patternItem.MeasPins.Add(new MeasPin { PinName = "Pin1", MeasType = "MeasV" });
            _patInfo.SeqInfo.Add(new HardIpSeqInfo
            {
                SeqName = "V",
                PinList = ["Pin1"]
            });

            // Act
            _checker.CheckSeq(_patternItem, _patInfo);

            // Assert
            Assert.IsTrue(ErrorReportManager.GetErrorList().Any(e => e.ErrorCode.FullCode == HardIpErrorType.W_MissingTestplanPin_01.FullCode));
        }

        [TestMethod]
        public void CheckSeq_ShouldReportError_WhenPinInPatternInfoMissingInPlan()
        {
            // Arrange
            _patInfo.SeqInfo.Add(new HardIpSeqInfo
            {
                SeqName = "V",
                PinList = ["PinX"]
            });

            // Act
            _checker.CheckSeq(_patternItem, _patInfo);

            // Assert
            Assert.IsTrue(ErrorReportManager.GetErrorList().Any(e => e.ErrorCode.FullCode == HardIpErrorType.E_WrongMeasPinInPatInfo_02.FullCode), "Should report error when info pin cannot match test plan.");
        }

        [TestMethod]
        public void PinsNum_Plan_StandardMeasurements_ReturnsTotalCount()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                MeasPins =
                [
                    new() { MeasType = "measv", PinCount = 2 },
                    new() { MeasType = "measi", PinCount = 3 },
                    new() { MeasType = "meascalc", PinCount = 5 }
                ]
            };

            // Act
            int result = _checker.PinsNum_Plan(pattern);

            // Assert
            Assert.AreEqual(5, result);
        }

        [TestMethod]
        public void PinsNum_Plan_DifferentialMeasurements_IncludesVdiffAndExcludesDuplicates()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                MeasPins =
                [
                    new() { MeasType = "measv", PinName = "pin_a", PinCount = 1 },
                    new() { MeasType = "measvdiff", PinName = "pin_a::pin_b" }
                ]
            };

            // Act
            int result = _checker.PinsNum_Plan(pattern);

            // Assert
            Assert.AreEqual(3, result);
        }

        [TestMethod]
        public void PinsNum_Plan_CaseInsensitivity_CountsCorrectly()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                MeasPins =
            [
                new() { MeasType = "MEASV", PinCount = 1 },
                new() { MeasType = "measVDIFF2", PinCount = 4 }
            ]
            };

            // Act
            int result = _checker.PinsNum_Plan(pattern);

            // Assert
            Assert.AreEqual(5, result);
        }

        [TestMethod]
        public void PinsNum_Plan_EmptyPins_ReturnsZero()
        {
            // Arrange
            var pattern = new HardIpPattern { MeasPins = [] };

            // Act
            int result = _checker.PinsNum_Plan(pattern);

            // Assert
            Assert.AreEqual(0, result);
        }
    }
}
