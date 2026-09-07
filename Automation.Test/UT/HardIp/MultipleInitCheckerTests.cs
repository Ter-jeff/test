using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.HardIpPreCheck;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.Static;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class MultipleInitCheckerTests
    {
        [TestInitialize]
        public void Setup()
        {
            ErrorReportManager.ClearErrors();
            LocalSpecs.HardIpInfos = [];
        }

        [TestCleanup]
        public void Cleanup()
        {
            ErrorReportManager.ClearErrors();
            LocalSpecs.HardIpInfos = [];
        }

        private static MultipleInitChecker NewChecker(string realPatternName)
        {
            var pattern = new HardIpPattern
            {
                SheetName = "TestSheet",
                RowNum = 1,
                Pattern = new PatternClass(realPatternName)
            };
            var hardIpSheet = new HardIpSheet { PlanHeaderIdx = new Dictionary<string, int> { { "registerIndex", 1 } } };
            return new MultipleInitChecker(hardIpSheet, pattern);
        }

        [TestMethod]
        public void Check_SingleNonMultiplePattern_NoOp()
        {
            // Arrange
            MultipleInitChecker checker = NewChecker("PAT1");

            // Act
            checker.Check();

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void Check_MultiplePatternsSameTimeSet_NoError()
        {
            // Arrange
            LocalSpecs.HardIpInfos = new HardIpInfos(
                new HardIpInfo { Payload = "PAT1", TimeSet = "TS1" },
                new HardIpInfo { Payload = "PAT2", TimeSet = "TS1" }
            );
            MultipleInitChecker checker = NewChecker("PAT1+PAT2");

            // Act
            checker.Check();

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void Check_MultiplePatternsDifferentTimeSet_AddsError()
        {
            // Arrange - the reference timeset comes from the LAST pattern in the group (PAT2),
            // so only PAT1 (the mismatching one) should be reported as the offending init
            LocalSpecs.HardIpInfos = new HardIpInfos(
                new HardIpInfo { Payload = "PAT1", TimeSet = "TS_DIFFERENT" },
                new HardIpInfo { Payload = "PAT2", TimeSet = "TS1" }
            );
            MultipleInitChecker checker = NewChecker("PAT1+PAT2");

            // Act
            checker.Check();

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
            Assert.IsTrue(ErrorReportManager.Contains(HardIpErrorType.E_WrongTimeSet_01.FormatMessage("PAT2", "PAT1")));
        }

        [TestMethod]
        public void Check_PatternsNotFoundInHardIpInfos_ExceptionCaughtNoError()
        {
            // Arrange - LocalSpecs.HardIpInfos is empty, so infos ends up empty and
            // infos.Last() throws; the checker swallows the exception silently
            MultipleInitChecker checker = NewChecker("PAT1+PAT2");

            // Act
            checker.Check();

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }
    }
}
