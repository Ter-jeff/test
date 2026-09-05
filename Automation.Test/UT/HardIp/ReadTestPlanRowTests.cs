using System.Linq;

using Automation.GenerateIgxl.Wireless.DVDC.InputObject;
using Automation.GenerateIgxl.Wireless.DVDC.InputReader;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class ReadTestPlanRowTests : FunctionTestBase
    {
        [TestMethod]
        public void ReadSheet_Should_Create_TestPlanRow_With_Meas_And_TestName()
        {
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("WS");

            sheet.Cells[1, 1].Value = "Pattern";
            sheet.Cells[1, 2].Value = "Condition";
            sheet.Cells[1, 3].Value = "Misc Info";
            sheet.Cells[1, 4].Value = "Meas";
            sheet.Cells[1, 5].Value = "Test Name";
            sheet.Cells[1, 20].Value = "";

            sheet.Cells[2, 1].Value = "PATTERN_1";
            sheet.Cells[2, 2].Value = "FC1";

            sheet.Cells[3, 1].Value = "";
            sheet.Cells[3, 2].Value = "FC1";
            sheet.Cells[3, 3].Value = "K=V";
            sheet.Cells[3, 4].Value = "MEAS_1";
            sheet.Cells[3, 5].Value = "TEST_1";

            var reader = new WirelessTestPlanReader();

            GenerateIgxl.HardIp.InputObject.TestPlanSheet result = reader.ReadSheet(sheet);
            var patRow = result.PatternRows.Single() as WirelessPatternRow;
            System.Collections.Generic.List<GenerateIgxl.HardIp.InputObject.TestPlanRow> tpRows = patRow!.GetTestPlanRows();

            Assert.AreEqual(1, tpRows.Count);
            Assert.AreEqual("MEAS_1", tpRows[0].Meas);
            Assert.AreEqual("TEST_1", tpRows[0].TestName);
            Assert.AreEqual("K=V", tpRows[0].MiscInfo);
        }
    }
}
