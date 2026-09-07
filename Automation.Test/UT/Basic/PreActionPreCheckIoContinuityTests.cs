using System.Collections.Generic;

using Automation.PreCheck.PreChecks.PreAction;

using CommonLib.ErrorReport;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace Automation.Test.UT.Basic
{
    [TestClass]
    public class PreActionPreCheckIoContinuityTests
    {
        [TestInitialize]
        public void Init()
        {
            ErrorReportManager.ClearErrors();
        }

        private static List<string> GetGrpList(PreActionPreCheckIoContinuity instance)
        {
            return instance._grpList;
        }

        [TestMethod]
        public void CheckMain_MissingBumpNameHeader_ReturnsFalse()
        {
            // Arrange - "Bump Name" is the FirstHeader; without it CheckExist() fails outright
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("IoConti");
            sheet.Cells[1, 1].Value = "SomethingElse";

            var preCheck = new PreActionPreCheckIoContinuity(package.Workbook, "IoConti");

            // Act
            bool result = preCheck.CheckMain();

            // Assert
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void CheckMain_ValidSheetWithFsAndDdValues_ReturnsTrueNoErrors()
        {
            // Arrange
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("IoConti");
            sheet.Cells[1, 1].Value = "Bump Name";
            sheet.Cells[1, 2].Value = "I/O Voltage";
            sheet.Cells[1, 3].Value = "FS/DD";
            sheet.Cells[2, 1].Value = "PIN_A";
            sheet.Cells[2, 2].Value = "1.8V";
            sheet.Cells[2, 3].Value = "FS";
            sheet.Cells[3, 1].Value = "PIN_B";
            sheet.Cells[3, 2].Value = "1.2V";
            sheet.Cells[3, 3].Value = "DD";

            var preCheck = new PreActionPreCheckIoContinuity(package.Workbook, "IoConti");

            // Act
            bool result = preCheck.CheckMain();

            // Assert
            Assert.IsTrue(result);
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckMain_InvalidFsDdValue_AddsErrorAndReturnsFalse()
        {
            // Arrange
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("IoConti");
            sheet.Cells[1, 1].Value = "Bump Name";
            sheet.Cells[1, 2].Value = "I/O Voltage";
            sheet.Cells[1, 3].Value = "FS/DD";
            sheet.Cells[2, 1].Value = "PIN_A";
            sheet.Cells[2, 2].Value = "1.8V";
            sheet.Cells[2, 3].Value = "MAYBE";

            var preCheck = new PreActionPreCheckIoContinuity(package.Workbook, "IoConti");

            // Act
            bool result = preCheck.CheckMain();

            // Assert
            Assert.IsFalse(result);
            Assert.IsTrue(ErrorReportManager.GetErrorList().Exists(e => e.Message.Contains("Needs to fill FS or DD")));
        }

        [TestMethod]
        public void CheckBusiness_VoltageWithTrailingZeroBeforeV_StripsZeroInGroupName()
        {
            // Arrange - "1.0V" becomes "1p0v" then the trailing "0v" is stripped to "v"
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("IoConti");
            sheet.Cells[1, 1].Value = "Bump Name";
            sheet.Cells[1, 2].Value = "I/O Voltage";
            sheet.Cells[1, 3].Value = "FS/DD";
            sheet.Cells[2, 1].Value = "PIN_A";
            sheet.Cells[2, 2].Value = "1.0V";
            sheet.Cells[2, 3].Value = "FS";

            var preCheck = new PreActionPreCheckIoContinuity(package.Workbook, "IoConti");

            // Act
            preCheck.CheckMain();

            // Assert
            List<string> grpList = GetGrpList(preCheck);
            Assert.AreEqual(1, grpList.Count);
            Assert.AreEqual("Pins_1pv", grpList[0]);
        }

        [TestMethod]
        public void CheckBusiness_DuplicateVoltageValues_OnlyAddedOnceToGroupList()
        {
            // Arrange
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("IoConti");
            sheet.Cells[1, 1].Value = "Bump Name";
            sheet.Cells[1, 2].Value = "I/O Voltage";
            sheet.Cells[1, 3].Value = "FS/DD";
            sheet.Cells[2, 1].Value = "PIN_A";
            sheet.Cells[2, 2].Value = "1.8V";
            sheet.Cells[2, 3].Value = "FS";
            sheet.Cells[3, 1].Value = "PIN_B";
            sheet.Cells[3, 2].Value = "1.8V";
            sheet.Cells[3, 3].Value = "FS";

            var preCheck = new PreActionPreCheckIoContinuity(package.Workbook, "IoConti");

            // Act
            preCheck.CheckMain();

            // Assert
            List<string> grpList = GetGrpList(preCheck);
            Assert.AreEqual(1, grpList.Count);
        }

        [TestMethod]
        public void CheckBusiness_BlankVoltage_NotAddedToGroupList()
        {
            // Arrange
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("IoConti");
            sheet.Cells[1, 1].Value = "Bump Name";
            sheet.Cells[1, 2].Value = "I/O Voltage";
            sheet.Cells[1, 3].Value = "FS/DD";
            sheet.Cells[2, 1].Value = "PIN_A";
            sheet.Cells[2, 3].Value = "FS";

            var preCheck = new PreActionPreCheckIoContinuity(package.Workbook, "IoConti");

            // Act
            preCheck.CheckMain();

            // Assert
            List<string> grpList = GetGrpList(preCheck);
            Assert.AreEqual(0, grpList.Count);
        }
    }
}
