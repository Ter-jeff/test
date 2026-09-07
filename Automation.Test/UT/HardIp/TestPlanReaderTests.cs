using System.Linq;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class TestPlanReaderTests : FunctionTestBase
    {
        private static ExcelWorksheet CreateValidHardIpSheet()
        {
            ExcelPackage excelPackage = new ExcelPackage();
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("DCTEST_CPU");

            worksheet.Cells[1, 1].Value = "Pattern";
            worksheet.Cells[1, 2].Value = "Description";
            worksheet.Cells[1, 3].Value = "Condition";
            worksheet.Cells[1, 4].Value = "Enable (all sites)";
            worksheet.Cells[1, 5].Value = "Meas";
            worksheet.Cells[1, 6].Value = "TTR";
            worksheet.Cells[1, 7].Value = "RegisterAssignment";

            return worksheet;
        }

        [TestMethod]
        public void ReadSheet_OnlyHeader_ReturnEmptyPatternRows()
        {
            ExcelWorksheet worksheet = CreateValidHardIpSheet();
            TestPlanReader reader = new TestPlanReader();

            TestPlanSheet result = reader.ReadSheet(worksheet);

            Assert.AreNotEqual(null, result);
            Assert.AreEqual("DCTEST_CPU", result.SheetName);
            Assert.AreEqual(0, result.PatternRows.Count);

            Assert.IsTrue(result.PlanHeaderIdx.ContainsKey("patternIndex"));
            Assert.IsTrue(result.PlanHeaderIdx.ContainsKey("forceIndex"));
            Assert.IsTrue(result.PlanHeaderIdx.ContainsKey("measIndex"));
        }

        [TestMethod]
        public void ReadSheet_SinglePatternRow_CreateOnePatternRow()
        {
            ExcelWorksheet worksheet = CreateValidHardIpSheet();

            worksheet.Cells[2, 1].Value = "dd_test_pattern";
            worksheet.Cells[2, 4].Value = "Y";
            worksheet.Cells[2, 6].Value = "CP1";
            worksheet.Cells[3, 7].Value = "REG1";

            TestPlanReader reader = new TestPlanReader();
            TestPlanSheet result = reader.ReadSheet(worksheet);

            Assert.AreEqual(1, result.PatternRows.Count);

            PatternRow patternRow = result.PatternRows.First();
            Assert.AreEqual(2, patternRow.RowNum);
            Assert.AreEqual("DCTEST_CPU", patternRow.SheetName);
            Assert.AreEqual("Y", patternRow.Enable);
        }

        [TestMethod]
        public void ReadSheet_PatternAndMeasRow_DoesNotCrashAndKeepsPatternRow()
        {
            ExcelWorksheet worksheet = CreateValidHardIpSheet();

            worksheet.Cells[2, 1].Value = "dd_test_pattern";
            worksheet.Cells[3, 5].Value = "MeasV";
            worksheet.Cells[3, 7].Value = "REG1";

            TestPlanReader reader = new TestPlanReader();
            TestPlanSheet result = reader.ReadSheet(worksheet);

            Assert.AreEqual(1, result.PatternRows.Count);
        }

        [TestMethod]
        public void ReadSheet_SheetLevelForceCondition_BeforePattern_IsCollected()
        {
            ExcelWorksheet worksheet = CreateValidHardIpSheet();

            worksheet.Cells[2, 3].Value = "FORCE1";
            worksheet.Cells[3, 3].Value = "FORCE2";
            worksheet.Cells[4, 1].Value = "PAT1";
            worksheet.Cells[3, 7].Value = "REG1";

            TestPlanReader reader = new TestPlanReader();
            TestPlanSheet result = reader.ReadSheet(worksheet);

            Assert.AreEqual("FORCE1;FORCE2", result.ForceStr);
        }

        [TestMethod]
        public void ReadSheet_PatternRowWithMeas_AddsPatternExistMeasurementError()
        {
            ErrorReportManager.ClearErrors();
            ExcelWorksheet worksheet = CreateValidHardIpSheet();
            worksheet.Cells[2, 1].Value = "PAT1";
            worksheet.Cells[2, 5].Value = "Meas1";
            worksheet.Cells[3, 7].Value = "REG1";

            TestPlanReader reader = new TestPlanReader();
            reader.ReadSheet(worksheet);

            Assert.IsTrue(ErrorReportManager.GetErrorList().Any(e => e.ErrorCode == HardIpErrorType.E_PatternExistMeasument_01));
        }

        [TestMethod]
        public void ReadSheet_MeasWithoutPattern_AddsMissingPatternForMeasurementError()
        {
            ErrorReportManager.ClearErrors();
            ExcelWorksheet worksheet = CreateValidHardIpSheet();
            worksheet.Cells[2, 5].Value = "MeasOnly";
            worksheet.Cells[3, 7].Value = "REG1";

            TestPlanReader reader = new TestPlanReader();
            reader.ReadSheet(worksheet);

            Assert.IsTrue(ErrorReportManager.GetErrorList().Any(e => e.ErrorCode == HardIpErrorType.E_MisPatternForMeasument_01));
        }

        [TestMethod]
        public void ReadSheet_RunPatternWithoutPattern_AddsMissingPatternError()
        {
            ErrorReportManager.ClearErrors();
            ExcelWorksheet worksheet = CreateValidHardIpSheet();
            worksheet.Cells[2, 2].Value = "Run the pattern ABC";
            worksheet.Cells[3, 7].Value = "REG1";

            TestPlanReader reader = new TestPlanReader();
            reader.ReadSheet(worksheet);

            Assert.IsTrue(ErrorReportManager.GetErrorList().Any(e => e.ErrorCode == HardIpErrorType.E_MissingPatternInTestPlan_01));
        }

        [TestMethod]
        public void ReadSheet_RegisterAssignmentAfterPattern_AppendedToPatternRow()
        {
            ExcelWorksheet worksheet = CreateValidHardIpSheet();
            worksheet.Cells[2, 1].Value = "PAT1";
            worksheet.Cells[3, 7].Value = "REG1";

            TestPlanReader reader = new TestPlanReader();
            TestPlanSheet result = reader.ReadSheet(worksheet);

            Assert.AreEqual("REG1", result.PatternRows[0].RegisterAssignment);
        }

        [TestMethod]
        public void ReadSheet_NoSheetLevelForceCondition_StartRowIsDefault()
        {
            ExcelWorksheet worksheet = CreateValidHardIpSheet();
            worksheet.Cells[2, 1].Value = "PAT1";
            worksheet.Cells[3, 7].Value = "REG1";

            TestPlanReader reader = new TestPlanReader();
            TestPlanSheet result = reader.ReadSheet(worksheet);

            Assert.AreEqual(string.Empty, result.ForceStr);
            Assert.AreEqual(1, result.PatternRows.Count);
        }

        [TestMethod]
        public void ReadSheet_PatternFollowedByEmptyRow_PatternStillExists()
        {
            ExcelWorksheet worksheet = CreateValidHardIpSheet();
            worksheet.Cells[2, 1].Value = "PAT1";
            worksheet.Cells[3, 1].Value = "";
            worksheet.Cells[3, 7].Value = "REG1";

            TestPlanReader reader = new TestPlanReader();
            TestPlanSheet result = reader.ReadSheet(worksheet);

            Assert.AreEqual(1, result.PatternRows.Count);
        }

        [TestMethod]
        public void ReadSheet_MultiplePatternRows_CreateMultiplePatternRows()
        {
            ExcelWorksheet worksheet = CreateValidHardIpSheet();
            worksheet.Cells[2, 1].Value = "PAT1";
            worksheet.Cells[3, 1].Value = "PAT2";
            worksheet.Cells[3, 7].Value = "REG1";

            TestPlanReader reader = new TestPlanReader();
            TestPlanSheet result = reader.ReadSheet(worksheet);

            Assert.AreEqual(2, result.PatternRows.Count);
        }
    }
}
