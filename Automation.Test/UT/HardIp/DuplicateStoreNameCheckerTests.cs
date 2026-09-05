using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.HardIpPreCheck;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;

using CommonLib.ErrorReport;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class DuplicateStoreNameCheckerTests
    {
        private readonly DuplicateStoreNameChecker _checker = new();

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

        private static HardIpPattern NewSinglePattern(string realPatternName, string sheetName, int rowNum, params string[] cusStrs)
        {
            var pattern = new HardIpPattern
            {
                SheetName = sheetName,
                RowNum = rowNum,
                Pattern = new PatternClass(realPatternName)
            };
            foreach (string cusStr in cusStrs)
            {
                pattern.MeasPins.Add(new MeasPin { CusStr = cusStr });
            }
            return pattern;
        }

        [TestMethod]
        public void CheckStoreName_NoDuplicateCusStr_NoErrors()
        {
            // Arrange
            var patterns = new List<HardIpPattern>
            {
                NewSinglePattern("PAT1", "S1", 0, "A"),
                NewSinglePattern("PAT2", "S2", 0, "B")
            };

            // Act
            _checker.CheckStoreName(patterns);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckStoreName_SinglePatternsShareCusStr_AddsErrorPerDuplicateItem()
        {
            // Arrange - two independent single (non "+"/"#") patterns share the same CusStr
            var patterns = new List<HardIpPattern>
            {
                NewSinglePattern("PAT1", "S1", 0, "DUPSTORE"),
                NewSinglePattern("PAT2", "S2", 0, "DUPSTORE")
            };

            // Act
            _checker.CheckStoreName(patterns);

            // Assert - one error per DupItem sharing the duplicated store name
            Assert.AreEqual(2, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckStoreName_EmptyCusStr_IgnoredEvenIfRepeated()
        {
            // Arrange
            var patterns = new List<HardIpPattern>
            {
                NewSinglePattern("PAT1", "S1", 0, ""),
                NewSinglePattern("PAT2", "S2", 0, "")
            };

            // Act
            _checker.CheckStoreName(patterns);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckStoreName_RepeatedCusStrWithinSamePattern_DedupedNotFlagged()
        {
            // Arrange - a single pattern with two MeasPins sharing the same CusStr; the per-pattern
            // HashSet dedup means only the first occurrence is recorded, so no cross-pattern
            // duplicate is ever formed from this alone.
            var patterns = new List<HardIpPattern>
            {
                NewSinglePattern("PAT1", "S1", 0, "REPEAT", "REPEAT")
            };

            // Act
            _checker.CheckStoreName(patterns);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckStoreName_MultiBlockConstituentPatternsShareCusStr_AddsError()
        {
            // Arrange - a multi-pattern ("PAT1+PAT2") seeds the multi-block set; the two
            // constituent single patterns "PAT1" and "PAT2" match into its group via Contains().
            // Because those single patterns also remain in the independent single-pattern pass,
            // a shared CusStr between them is flagged TWICE: once per DupItem (2) in the
            // multi-block pass, and again once per DupItem (2) in the single-pattern pass.
            var patterns = new List<HardIpPattern>
            {
                NewSinglePattern("PAT1+PAT2", "SMulti", 0),
                NewSinglePattern("PAT1", "S1", 0, "DUPX"),
                NewSinglePattern("PAT2", "S2", 0, "DUPX")
            };

            // Act
            _checker.CheckStoreName(patterns);

            // Assert
            Assert.AreEqual(4, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckStoreName_MultiBlockConstituentPatternsDistinctCusStr_NoError()
        {
            // Arrange
            var patterns = new List<HardIpPattern>
            {
                NewSinglePattern("PAT1+PAT2", "SMulti", 0),
                NewSinglePattern("PAT1", "S1", 0, "AAA"),
                NewSinglePattern("PAT2", "S2", 0, "BBB")
            };

            // Act
            _checker.CheckStoreName(patterns);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckStoreName_EmptyPatternList_NoErrorNoThrow()
        {
            // Act
            _checker.CheckStoreName([]);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }
    }
}
