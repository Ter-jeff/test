using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.HardIpPreCheck;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class MeasCCaptureCheckerTests
    {
        [TestInitialize]
        public void Setup()
        {
            ErrorReportManager.ClearErrors();
        }

        [TestCleanup]
        public void Cleanup()
        {
            ErrorReportManager.ClearErrors();
        }

        private static MeasCCaptureChecker NewChecker(HardIpPattern pattern)
        {
            var hardIpSheet = new HardIpSheet { PlanHeaderIdx = new Dictionary<string, int> { { "measIndex", 3 } } };
            return new MeasCCaptureChecker(hardIpSheet, pattern);
        }

        private static HardIpPattern NewPattern(params MeasPin[] pins)
        {
            var pattern = new HardIpPattern { SheetName = "TestSheet", RowNum = 1, Pattern = new PatternClass("") };
            pattern.MeasPins.AddRange(pins);
            return pattern;
        }

        [TestMethod]
        public void Check_NoDuplicateCusStr_NoError()
        {
            // Arrange
            HardIpPattern pattern = NewPattern(
                new MeasPin { CusStr = "A", SequenceIndex = 0 },
                new MeasPin { CusStr = "B", SequenceIndex = 0 }
            );

            // Act
            NewChecker(pattern).Check();

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void Check_DuplicateCusStrInSameGroup_AddsErrorPerRepeatWithIncrementingCount()
        {
            // Arrange - three pins share the same CusStr within the same SequenceIndex group
            HardIpPattern pattern = NewPattern(
                new MeasPin { CusStr = "DUP", SequenceIndex = 0, RowNum = 1 },
                new MeasPin { CusStr = "DUP", SequenceIndex = 0, RowNum = 2 },
                new MeasPin { CusStr = "DUP", SequenceIndex = 0, RowNum = 3 }
            );

            // Act
            NewChecker(pattern).Check();

            // Assert
            Assert.AreEqual(2, ErrorReportManager.GetErrorCount());
            Assert.IsTrue(ErrorReportManager.Contains(HardIpErrorType.W_WrongMeasC_01.FormatMessage("DUP", "1")));
            Assert.IsTrue(ErrorReportManager.Contains(HardIpErrorType.W_WrongMeasC_01.FormatMessage("DUP", "2")));
        }

        [TestMethod]
        public void Check_DuplicateCusStrInDifferentGroups_NoError()
        {
            // Arrange - same CusStr but in different SequenceIndex groups, so no collision
            HardIpPattern pattern = NewPattern(
                new MeasPin { CusStr = "DUP", SequenceIndex = 0 },
                new MeasPin { CusStr = "DUP", SequenceIndex = 1 }
            );

            // Act
            NewChecker(pattern).Check();

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void Check_EmptyCusStrPins_ExcludedFromGroupingNoError()
        {
            // Arrange
            HardIpPattern pattern = NewPattern(
                new MeasPin { CusStr = "", SequenceIndex = 0 },
                new MeasPin { CusStr = "", SequenceIndex = 0 }
            );

            // Act
            NewChecker(pattern).Check();

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void Check_EmptyMeasPinsList_NoErrorNoThrow()
        {
            // Arrange
            HardIpPattern pattern = NewPattern();

            // Act
            NewChecker(pattern).Check();

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }
    }
}
