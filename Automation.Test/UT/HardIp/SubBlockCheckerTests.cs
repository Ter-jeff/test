using System.Collections.Generic;
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
    public class SubBlockCheckerTests : FunctionTestBase
    {
        [TestInitialize]
        public void Setup()
        {
            // Clear errors before each test to ensure a clean state
            ErrorReportManager.ClearErrors();
        }

        [TestMethod]
        public void CheckNoBurstItem_LocalPatternsAllBurst_ReturnsWithoutError()
        {
            // Arrange
            KeyValuePair<string, HardIpSheet> dic = CreateDummyDic("SheetA");
            var refPatterns = new Dictionary<string, HardIpPattern>
            {
                { "P1", new HardIpPattern { IsBurst = true } }
            };
            var refPatternsGlobal = new Dictionary<string, HardIpPattern>();

            // Act
            new SubBlockChecker([]).CheckNoBurstItem(dic, true, refPatterns, refPatternsGlobal);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount(), "Should exit early if all local patterns are burst.");
        }

        [TestMethod]
        public void CheckNoBurstItem_LocalNotBurst_GlobalMismatch_AddsError()
        {
            // Arrange
            string sheetName = "SheetA";
            KeyValuePair<string, HardIpSheet> dic = CreateDummyDic(sheetName);

            // Local fails (needBurst=true but local is false)
            var refPatterns = new Dictionary<string, HardIpPattern>
            {
                { "Local1", new HardIpPattern { IsBurst = false } }
            };

            // Global matches sheet but is not burst
            var refPatternsGlobal = new Dictionary<string, HardIpPattern>
            {
                { "Global1", new HardIpPattern
                    {
                        SheetName = sheetName,
                        IsBurst = false,
                        RowNum = 10,
                        Pattern = new PatternClass("Path/To/Payload"),
                        MiscInfo = "TestInfo"
                    }
                }
            };

            // Act
            new SubBlockChecker([]).CheckNoBurstItem(dic, true, refPatterns, refPatternsGlobal);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
            Error error = ErrorReportManager.GetErrorList().Last();
            Assert.AreEqual(10, error.RowNum);
            Assert.IsTrue(error.Message.Contains("not be referenced by burst item"));
        }

        [TestMethod]
        public void CheckNoBurstItem_GlobalIsBurst_NoErrorReported()
        {
            // Arrange
            string sheetName = "SheetA";
            KeyValuePair<string, HardIpSheet> dic = CreateDummyDic(sheetName);
            var refPatterns = new Dictionary<string, HardIpPattern> { { "L", new HardIpPattern { IsBurst = false } } };

            // Global matches sheet AND IS BURST (should not trigger error)
            var refPatternsGlobal = new Dictionary<string, HardIpPattern>
            {
                { "G", new HardIpPattern { SheetName = sheetName, IsBurst = true } }
            };

            // Act
            new SubBlockChecker([]).CheckNoBurstItem(dic, true, refPatterns, refPatternsGlobal);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        private static KeyValuePair<string, HardIpSheet> CreateDummyDic(string sheetName)
        {
            var sheet = new HardIpSheet();
            sheet.PlanHeaderIdx["miscInfoIndex"] = 5;
            return new KeyValuePair<string, HardIpSheet>(sheetName, sheet);
        }

        private static HardIpSheet NewSheet(string sheetName, params HardIpPattern[] rows)
        {
            var sheet = new HardIpSheet { SheetName = sheetName, Rows = [.. rows] };
            sheet.PlanHeaderIdx["miscInfoIndex"] = 5;
            return sheet;
        }

        private static HardIpPattern NewPattern(string sheetName, string miscInfo, int rowNum = 0)
        {
            return new HardIpPattern
            {
                SheetName = sheetName,
                RowNum = rowNum,
                MiscInfo = miscInfo,
                Pattern = new PatternClass(""),
                SheetSubBlockName = $"{sheetName}_unused_{rowNum}"
            };
        }

        [TestMethod]
        public void Check_RefSubBlockFoundInGlobal_SetsIsBurstTrue()
        {
            // Arrange
            var globalPattern = new HardIpPattern { SheetName = "SheetA", SheetSubBlockName = "SheetA_PatA", IsBurst = false };
            var planDic = new Dictionary<string, HardIpSheet> { { "SheetA", NewSheet("SheetA", globalPattern) } };
            var checker = new SubBlockChecker(planDic);
            HardIpPattern current = NewPattern("SheetA", "ref_subblock:PatA");

            // Act
            checker.Check(planDic["SheetA"], current);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
            Assert.IsTrue(globalPattern.IsBurst);
        }

        [TestMethod]
        public void Check_RefSubBlockNotFoundInGlobal_AddsError()
        {
            // Arrange
            var planDic = new Dictionary<string, HardIpSheet> { { "SheetA", NewSheet("SheetA") } };
            var checker = new SubBlockChecker(planDic);
            HardIpPattern current = NewPattern("SheetA", "ref_subblock:MissingPat");

            // Act
            checker.Check(planDic["SheetA"], current);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void Check_MiscInfoWithoutRefSubBlock_NoErrorFromFirstLoop()
        {
            // Arrange
            var planDic = new Dictionary<string, HardIpSheet> { { "SheetA", NewSheet("SheetA") } };
            var checker = new SubBlockChecker(planDic);
            HardIpPattern current = NewPattern("SheetA", "NoColonHere");

            // Act
            checker.Check(planDic["SheetA"], current);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void Check_SubBlockNameParam_ValidUniqueUsage_NoError()
        {
            // Arrange
            var planDic = new Dictionary<string, HardIpSheet> { { "SheetA", NewSheet("SheetA") } };
            var checker = new SubBlockChecker(planDic);
            HardIpPattern current = NewPattern("SheetA", "SubBlock:Name1");

            // Act
            checker.Check(planDic["SheetA"], current);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void Check_SubBlockNameParam_ArgCountNotTwo_AddsUsageError()
        {
            // Arrange
            var planDic = new Dictionary<string, HardIpSheet> { { "SheetA", NewSheet("SheetA") } };
            var checker = new SubBlockChecker(planDic);
            HardIpPattern current = NewPattern("SheetA", "SubBlock:A:B");

            // Act
            checker.Check(planDic["SheetA"], current);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void Check_SubBlockNameParam_RepeatedAcrossCalls_AddsRepeatSubBlockError()
        {
            // Arrange - the local sub-block list accumulates across Check() calls on the same
            // checker instance until ResetSubBlockList() is invoked
            var planDic = new Dictionary<string, HardIpSheet> { { "SheetA", NewSheet("SheetA") } };
            var checker = new SubBlockChecker(planDic);
            checker.Check(planDic["SheetA"], NewPattern("SheetA", "SubBlock:Dup1"));
            ErrorReportManager.ClearErrors();

            // Act
            checker.Check(planDic["SheetA"], NewPattern("SheetA", "SubBlock:Dup1"));

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void Check_SubBlockNameParam_ExistsInCurrentAfterReset_AddsExistedSubBlockWarning()
        {
            // Arrange - ResetSubBlockList() clears only the local list, not the current-sheet
            // accumulator, so re-using the same name after a reset should warn instead of error
            var planDic = new Dictionary<string, HardIpSheet> { { "SheetA", NewSheet("SheetA") } };
            var checker = new SubBlockChecker(planDic);
            checker.Check(planDic["SheetA"], NewPattern("SheetA", "SubBlock:Rep1"));
            checker.ResetSubBlockList();
            ErrorReportManager.ClearErrors();

            // Act
            checker.Check(planDic["SheetA"], NewPattern("SheetA", "SubBlock:Rep1"));

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void Check_SubBlockValueIsHyphen_AddsRepeatSubBlockWarning()
        {
            // Arrange - Regex.IsMatch("-", paramValue) treats paramValue as the pattern, so a
            // literal "-" value matches against the fixed "-" input and triggers the warning
            var planDic = new Dictionary<string, HardIpSheet> { { "SheetA", NewSheet("SheetA") } };
            var checker = new SubBlockChecker(planDic);
            HardIpPattern current = NewPattern("SheetA", "SubBlock:-");

            // Act
            checker.Check(planDic["SheetA"], current);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }
    }
}
