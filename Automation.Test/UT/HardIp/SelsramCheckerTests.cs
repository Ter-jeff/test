using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.HardIpPreCheck;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.BinCut;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class SelsramCheckerTests : FunctionTestBase
    {
        [TestMethod]
        public void SelsramCheckerTest()
        {
            // Act
            ErrorReportManager.ClearErrors();
            var hardIpSheet = new HardIpSheet
            {
                PlanHeaderIdx =
                {
                    ["patternIndex"] = 1
                }
            };
            var pattern = new HardIpPattern
            {
                MiscInfo = "CheckSelsram",
                RegisterAssignment = "DSA1=1;DSA2=2",
                Pattern = new PatternClass("_SRMDSSC")
            };
            new SelsramChecker(hardIpSheet, pattern).Check();

            // Assert
            List<Error> errors = ErrorReportManager.GetErrorList();
            Assert.AreEqual(10, errors.Count);
            Assert.AreEqual(HardIpErrorType.E_SelsramDigSrcAssignmentNotDefineInTable_01.FullCode, errors[0].ErrorCode.FullCode);
            Assert.AreEqual("Block: * Pattern: *_SRMDSSC* Alpha: S of DigSrc_Assignment is not defined in SELSRM_Mapping_Table.", errors[0].Message);
        }

        private static SelsramChecker NewChecker(HardIpPattern pattern)
        {
            var hardIpSheet = new HardIpSheet
            {
                PlanHeaderIdx = new Dictionary<string, int> { { "patternIndex", 1 }, { "registerIndex", 2 } }
            };
            return new SelsramChecker(hardIpSheet, pattern);
        }

        private static HardIpPattern NewPattern(string sheetName, string registerAssignment, string patternName)
        {
            return new HardIpPattern
            {
                SheetName = sheetName,
                RowNum = 1,
                RegisterAssignment = registerAssignment,
                Pattern = new PatternClass(patternName)
            };
        }

        private static SelsrmMappingTableRow NewRow(string block, string pattern, string digSrcAssignment, string logicPins = "", string sramPins = "", int rowNum = 1)
        {
            return new SelsrmMappingTableRow
            {
                Block = block,
                Pattern = pattern,
                DigSrcAssignment = digSrcAssignment,
                LogicPins = logicPins,
                SramPins = sramPins,
                Alpha = "A",
                RowNum = rowNum
            };
        }

        private static void SetMappingSheet(SelsramChecker checker, SelsrmMappingSheet sheet)
        {
            checker._selsrmMappingSheet = sheet;
        }

        private static void InvokeCheckItem(SelsramChecker checker, HardIpPattern pattern)
        {
            checker.CheckItem(pattern);
        }

        private static void InvokeCheckHeader(SelsramChecker checker)
        {
            checker.CheckHeader();
        }

        [TestMethod]
        public void Check_MiscInfoWithoutCheckSelsram_ReturnsEarlyWithoutTouchingMappingTable()
        {
            // Arrange - the guard should short-circuit before ever needing a mapping sheet
            ErrorReportManager.ClearErrors();
            HardIpPattern pattern = NewPattern("HARDIP_S1", "", "PAT1");
            pattern.MiscInfo = "Something";
            SelsramChecker checker = NewChecker(pattern);

            // Act
            checker.Check();

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckItem_NoMatchingRow_AddsCanNotGetSelsramSettingError()
        {
            // Arrange
            ErrorReportManager.ClearErrors();
            HardIpPattern pattern = NewPattern("HARDIP_BlockA", "DSA1=1", "PatX");
            SelsramChecker checker = NewChecker(pattern);
            var sheet = new SelsrmMappingSheet("SELSRM_Mapping_Table");
            sheet.Rows.Add(NewRow("NoMatchBlock", "NoMatchPattern", "DSA1"));
            SetMappingSheet(checker, sheet);

            // Act
            InvokeCheckItem(checker, pattern);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
            Assert.AreEqual(HardIpErrorType.E_CanNotGetSelsramSetting_01.FullCode, ErrorReportManager.GetErrorList()[0].ErrorCode.FullCode);
        }

        [TestMethod]
        public void CheckItem_MatchingRowWithBlankDigSrcAssignment_AddsNotDefinedInTableError()
        {
            // Arrange
            ErrorReportManager.ClearErrors();
            HardIpPattern pattern = NewPattern("HARDIP_BlockA", "DSA1=1", "PatX");
            SelsramChecker checker = NewChecker(pattern);
            var sheet = new SelsrmMappingSheet("SELSRM_Mapping_Table");
            sheet.HeaderIndex["DigSrc_Assignment"] = 5;
            sheet.Rows.Add(NewRow("BlockA", "PatX", ""));
            SetMappingSheet(checker, sheet);

            // Act
            InvokeCheckItem(checker, pattern);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
            Assert.AreEqual(HardIpErrorType.E_SelsramDigSrcAssignmentNotDefineInTable_01.FullCode, ErrorReportManager.GetErrorList()[0].ErrorCode.FullCode);
        }

        [TestMethod]
        public void CheckItem_MatchingRowWithNaDigSrcAssignment_AddsNotDefinedInTableError()
        {
            // Arrange
            ErrorReportManager.ClearErrors();
            HardIpPattern pattern = NewPattern("HARDIP_BlockA", "DSA1=1", "PatX");
            SelsramChecker checker = NewChecker(pattern);
            var sheet = new SelsrmMappingSheet("SELSRM_Mapping_Table");
            sheet.HeaderIndex["DigSrc_Assignment"] = 5;
            sheet.Rows.Add(NewRow("BlockA", "PatX", "N/A"));
            SetMappingSheet(checker, sheet);

            // Act
            InvokeCheckItem(checker, pattern);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
            Assert.AreEqual(HardIpErrorType.E_SelsramDigSrcAssignmentNotDefineInTable_01.FullCode, ErrorReportManager.GetErrorList()[0].ErrorCode.FullCode);
        }

        [TestMethod]
        public void CheckItem_ValidDigSrcAssignmentFoundInRegisterAssignment_NoError()
        {
            // Arrange
            ErrorReportManager.ClearErrors();
            HardIpPattern pattern = NewPattern("HARDIP_BlockA", "DSA1=1;DSA2=2", "PatX");
            SelsramChecker checker = NewChecker(pattern);
            var sheet = new SelsrmMappingSheet("SELSRM_Mapping_Table");
            sheet.Rows.Add(NewRow("BlockA", "PatX", "DSA1"));
            SetMappingSheet(checker, sheet);

            // Act
            InvokeCheckItem(checker, pattern);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckItem_ValidDigSrcAssignmentNotFoundInRegisterAssignment_AddsNotDefinedInInstanceError()
        {
            // Arrange
            ErrorReportManager.ClearErrors();
            HardIpPattern pattern = NewPattern("HARDIP_BlockA", "DSA1=1", "PatX");
            SelsramChecker checker = NewChecker(pattern);
            var sheet = new SelsrmMappingSheet("SELSRM_Mapping_Table");
            sheet.Rows.Add(NewRow("BlockA", "PatX", "DSA3"));
            SetMappingSheet(checker, sheet);

            // Act
            InvokeCheckItem(checker, pattern);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
            Assert.AreEqual(HardIpErrorType.E_SelsramDigSrcAssignmentNotDefineInInstance_01.FullCode, ErrorReportManager.GetErrorList()[0].ErrorCode.FullCode);
        }

        [TestMethod]
        public void CheckItem_RowWithPreservedLogicPins_ExcludedFromMatching()
        {
            // Arrange - a "PRESERVED" logic-pins row is filtered out of the candidate match set,
            // so the pattern behaves as if no matching row exists at all
            ErrorReportManager.ClearErrors();
            HardIpPattern pattern = NewPattern("HARDIP_BlockA", "DSA1=1", "PatX");
            SelsramChecker checker = NewChecker(pattern);
            var sheet = new SelsrmMappingSheet("SELSRM_Mapping_Table");
            sheet.Rows.Add(NewRow("BlockA", "PatX", "DSA1", logicPins: "PRESERVED"));
            SetMappingSheet(checker, sheet);

            // Act
            InvokeCheckItem(checker, pattern);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
            Assert.AreEqual(HardIpErrorType.E_CanNotGetSelsramSetting_01.FullCode, ErrorReportManager.GetErrorList()[0].ErrorCode.FullCode);
        }

        [TestMethod]
        public void CheckItem_RowWithPreservedSramPins_ExcludedFromMatching()
        {
            // Arrange
            ErrorReportManager.ClearErrors();
            HardIpPattern pattern = NewPattern("HARDIP_BlockA", "DSA1=1", "PatX");
            SelsramChecker checker = NewChecker(pattern);
            var sheet = new SelsrmMappingSheet("SELSRM_Mapping_Table");
            sheet.Rows.Add(NewRow("BlockA", "PatX", "DSA1", sramPins: "PRESERVED"));
            SetMappingSheet(checker, sheet);

            // Act
            InvokeCheckItem(checker, pattern);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
            Assert.AreEqual(HardIpErrorType.E_CanNotGetSelsramSetting_01.FullCode, ErrorReportManager.GetErrorList()[0].ErrorCode.FullCode);
        }

        [TestMethod]
        public void CheckItem_WildcardBlockAndPatternMatchAnything()
        {
            // Arrange
            ErrorReportManager.ClearErrors();
            HardIpPattern pattern = NewPattern("HARDIP_AnyBlock", "DSA1=1", "AnyPattern");
            SelsramChecker checker = NewChecker(pattern);
            var sheet = new SelsrmMappingSheet("SELSRM_Mapping_Table");
            sheet.Rows.Add(NewRow("*", "*", "DSA1"));
            SetMappingSheet(checker, sheet);

            // Act
            InvokeCheckItem(checker, pattern);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckHeader_MissingDigSrcAssignmentKey_AddsError()
        {
            // Arrange
            ErrorReportManager.ClearErrors();
            SelsramChecker checker = NewChecker(NewPattern("HARDIP_BlockA", "", "PatX"));
            var sheet = new SelsrmMappingSheet("SELSRM_Mapping_Table");
            SetMappingSheet(checker, sheet);

            // Act
            InvokeCheckHeader(checker);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
            Assert.AreEqual(HardIpErrorType.E_SelsramMappingTableError_02.FullCode, ErrorReportManager.GetErrorList()[0].ErrorCode.FullCode);
        }

        [TestMethod]
        public void CheckHeader_HasDigSrcAssignmentKey_NoError()
        {
            // Arrange
            ErrorReportManager.ClearErrors();
            SelsramChecker checker = NewChecker(NewPattern("HARDIP_BlockA", "", "PatX"));
            var sheet = new SelsrmMappingSheet("SELSRM_Mapping_Table");
            sheet.HeaderIndex["DigSrc_Assignment"] = 5;
            SetMappingSheet(checker, sheet);

            // Act
            InvokeCheckHeader(checker);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }
    }
}
