using System.Collections.Generic;
using System.IO;

using Automation.Utility.Atpg;

using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace Automation.Test.UT.Scan
{
    [TestClass]
    public class HarvestGroupPinsCheckerTests : FunctionTestBase
    {
        [TestInitialize]
        public void Setup() { }

        [TestMethod]
        public void CheckPatternPins_ShouldReportMissingAndExtraPins()
        {
            AssertOnlyWindowsOS("requires patinfo.exe via IGXLROOT environment variable");

            // Arrange
            var pinsFromGroup = new List<string> { "A1", "A2", "A3" };
            var pinMap = new PinMapSheet("");
            var patternList = new List<string> { "CZ_BRNA0_C_FULP_AN_AA00_DLL_JTG_VIX_ALLFRV_SI_CPLLDS_T6PD" };

            var file = new FileInfo(Path.GetTempFileName());
            var excelPackage = new ExcelPackage(file);
            ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.Add("Extra harvest pins");
            sheet.Cells[1, 1].Value = "Pins";
            ExcelWorksheet sheet1 = excelPackage.Workbook.Worksheets.Add("Miss harvest pins");
            sheet1.Cells[1, 1].Value = "Pins";

            // Act
            AtpgService.CheckPatternPins(pinsFromGroup, patternList, pinMap, excelPackage, []);

            // Assert
            ExcelWorksheet extraSheet = excelPackage.Workbook.Worksheets["Extra harvest pins"];
            ExcelWorksheet missSheet = excelPackage.Workbook.Worksheets["Miss harvest pins"];

            Assert.AreEqual("CZ_BRNA0_C_FULP_AN_AA00_DLL_JTG_VIX_ALLFRV_SI_CPLLDS_T6PD", missSheet.Cells[2, 1].Text);
            Assert.AreEqual("A1", missSheet.Cells[2, 2].Text);
            Assert.AreEqual(1, extraSheet.Dimension.Rows);
            Assert.AreEqual(2, missSheet.Dimension.Rows);
        }
    }
}
