using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.Wireless.DVDC.InputReader;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class ReadNonPatternRowNewTests : FunctionTestBase
    {
        [TestMethod]
        public void ReadSheet_When_MiscInfo_Header_Exists_Should_Parse_MiscInfo()
        {
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("WS");

            sheet.Cells[1, 1].Value = "Pattern";
            sheet.Cells[1, 2].Value = "Condition";
            sheet.Cells[1, 3].Value = "Misc Info";
            sheet.Cells[1, 4].Value = "Meas";
            sheet.Cells[1, 5].Value = "Test Name";
            sheet.Cells[1, 20].Value = "";

            sheet.Cells[2, 1].Value = "PAT_1";
            sheet.Cells[2, 2].Value = "FC1";

            sheet.Cells[3, 1].Value = "";
            sheet.Cells[3, 2].Value = "FC1";
            sheet.Cells[3, 3].Value = "KEY=VALUE";
            sheet.Cells[3, 4].Value = "MEAS1";
            sheet.Cells[3, 5].Value = "TEST1";

            var reader = new WirelessTestPlanReader();

            TestPlanSheet result = reader.ReadSheet(sheet);
            PatternRow patRow = result.PatternRows.Single();
            var subRow = patRow.PatChildRows.Single() as PatSubChildRow;
            TestPlanRow tpRow = subRow!.TpRows.Single();

            Assert.AreEqual("KEY=VALUE", tpRow.MiscInfo);
            Assert.AreEqual("MEAS1", tpRow.Meas);
            Assert.AreEqual("TEST1", tpRow.TestName);
        }
        [TestMethod]
        public void ReadSheet_When_RfHeaders_Exist_Should_Read_RfFields()
        {
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("WS");

            sheet.Cells[1, 1].Value = "Pattern";
            sheet.Cells[1, 2].Value = "Condition";
            sheet.Cells[1, 3].Value = "Misc Info";
            sheet.Cells[1, 4].Value = "RF Test Type";
            sheet.Cells[1, 5].Value = "RF Interpose";
            sheet.Cells[1, 6].Value = "RF Instrument Setup";
            sheet.Cells[1, 7].Value = "BB Instrument Setup";
            sheet.Cells[1, 8].Value = "BB Device Setup";
            sheet.Cells[1, 9].Value = "BB Calc";
            sheet.Cells[1, 10].Value = "Meas";
            sheet.Cells[1, 11].Value = "Test Name";

            sheet.Cells[1, 20].Value = "";

            sheet.Cells[2, 1].Value = "PAT_1";
            sheet.Cells[2, 2].Value = "FC1";

            sheet.Cells[3, 1].Value = "";
            sheet.Cells[3, 2].Value = "FC1";
            sheet.Cells[3, 3].Value = "KEY=VALUE";
            sheet.Cells[3, 4].Value = "RF_TYPE";
            sheet.Cells[3, 5].Value = "INTERPOSE";
            sheet.Cells[3, 6].Value = "RF_SETUP";
            sheet.Cells[3, 7].Value = "BB_SETUP";
            sheet.Cells[3, 8].Value = "BB_DEVICE";
            sheet.Cells[3, 9].Value = "BB_CALC";
            sheet.Cells[3, 10].Value = "MEAS1";
            sheet.Cells[3, 11].Value = "TEST1";

            var reader = new WirelessTestPlanReader();

            TestPlanSheet result = reader.ReadSheet(sheet);

            PatternRow patRow = result.PatternRows.Single();
            var subRow = patRow.PatChildRows.Single() as PatSubChildRow;
            TestPlanRow tpRow = subRow!.TpRows.Single();

            Assert.AreEqual("RF_TYPE", tpRow.InterposeFunc);
            Assert.AreEqual("INTERPOSE", tpRow.RfInterpose);
            Assert.AreEqual("KEY=VALUE", tpRow.MiscInfo);
            Assert.AreEqual("MEAS1", tpRow.Meas);
            Assert.AreEqual("TEST1", tpRow.TestName);

            Assert.AreEqual(
                "RF_SETUP;BB_SETUP;BB_DEVICE;BB_CALC",
                tpRow.RfIntrumentSetup);
        }
        [TestMethod]
        public void ReadSheet_When_Meas_Is_Empty_Should_Default_To_MeasLimit()
        {
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("WS");

            sheet.Cells[1, 1].Value = "Pattern";
            sheet.Cells[1, 2].Value = "Condition";
            sheet.Cells[1, 3].Value = "Misc Info";
            sheet.Cells[1, 4].Value = "Meas";
            sheet.Cells[1, 5].Value = "Test Name";
            sheet.Cells[1, 20].Value = "";

            sheet.Cells[2, 1].Value = "PAT_1";

            sheet.Cells[3, 1].Value = "";
            sheet.Cells[3, 2].Value = "";         // ForceCondition
            sheet.Cells[3, 3].Value = "";         // MiscInfo
            sheet.Cells[3, 4].Value = "";         // Meas
            sheet.Cells[3, 5].Value = "TEST1";    // TestName

            var reader = new WirelessTestPlanReader();

            TestPlanSheet result = reader.ReadSheet(sheet);

            PatternRow patRow = result.PatternRows.Single();
            var subRow = (PatSubChildRow)patRow.PatChildRows.Single();
            TestPlanRow tpRow = subRow.TpRows.Single();

            Assert.AreEqual(MeasType.MeasLimit, tpRow.Meas);
        }
        [TestMethod]
        public void ReadSheet_When_RfSetup_Is_Merged_Should_Read_From_Merged_Row()
        {
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("WS");

            sheet.Cells[1, 1].Value = "Pattern";
            sheet.Cells[1, 2].Value = "Condition";
            sheet.Cells[1, 3].Value = "Misc Info";
            sheet.Cells[1, 4].Value = "RF Instrument Setup";
            sheet.Cells[1, 5].Value = "Meas";
            sheet.Cells[1, 6].Value = "Test Name";
            sheet.Cells[1, 20].Value = "";

            sheet.Cells[2, 1].Value = "PAT_1";

            // Merge RF Setup
            sheet.Cells[3, 4].Value = "RF_SETUP";
            sheet.Cells[3, 4, 4, 4].Merge = true;

            // Row3
            sheet.Cells[3, 2].Value = "";
            sheet.Cells[3, 5].Value = "MEAS1";
            sheet.Cells[3, 6].Value = "TEST1";

            // Row4
            sheet.Cells[4, 2].Value = "";
            sheet.Cells[4, 5].Value = "MEAS2";
            sheet.Cells[4, 6].Value = "TEST2";

            var reader = new WirelessTestPlanReader();

            TestPlanSheet result = reader.ReadSheet(sheet);

            PatternRow patRow = result.PatternRows.Single();
            var subRow = (PatSubChildRow)patRow.PatChildRows.Last();

            TestPlanRow tpRow = subRow.TpRows.Single();

            Assert.AreEqual("RF_SETUP", tpRow.RfIntrumentSetup);
            Assert.AreEqual(3, tpRow.MergeRowNumForMeas);
        }
    }
}
