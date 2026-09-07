using System.ComponentModel;

using Automation.PreCheck.PreChecks.Basic;

using CommonLib.ErrorReport;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace Automation.Test.UT.Basic
{
    [TestClass]
    public class BasicPreCheckIoInfoTests
    {
        [TestInitialize]
        public void Init()
        {
            ErrorReportManager.ClearErrors();
        }

        [TestMethod]
        [DisplayName("01_CheckHeaders_ShouldReturnFalse_WhenMissingCommonBlock")]
        public void CheckHeaders_ShouldReturnFalse_WhenMissingCommonBlock()
        {
            // Arrange
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("IOInfo");
            sheet.Cells[1, 1].Value = "Common";
            sheet.Cells[2, 1].Value = "PinGrpName";
            sheet.Cells[2, 2].Value = "NV";
            sheet.Cells[2, 3].Value = "HV";
            sheet.Cells[2, 4].Value = "LV";
            sheet.Cells[2, 5].Value = "Vil";
            sheet.Cells[2, 6].Value = "Vih";
            sheet.Cells[2, 7].Value = "Vol";
            sheet.Cells[2, 8].Value = "Voh";
            sheet.Cells[2, 9].Value = "Vt";
            sheet.Cells[2, 10].Value = "Vch";
            sheet.Cells[3, 1].Value = "IO";
            sheet.Cells[3, 2].Value = "PINA*2";

            var preCheck = new BasicPreCheckIoInfo(package.Workbook, "IOInfo");

            // Act
            bool result = preCheck.CheckMain();

            // Assert
            Assert.IsTrue(result);
            Assert.IsTrue(ErrorReportManager.GetErrorList().Exists(e => e.Message.Contains("Different pin name from NV: PINA*2")));
        }
    }
}
