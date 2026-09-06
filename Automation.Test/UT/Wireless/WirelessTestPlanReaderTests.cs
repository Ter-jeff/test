using Automation.GenerateIgxl.Wireless.DVDC.InputReader;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace Automation.Test.UT.Wireless
{
    [TestClass]
    public class WirelessTestPlanReaderTests
    {
        private WirelessTestPlanReader _reader = null!;

        [TestInitialize]
        public void Setup()
        {
            _reader = new WirelessTestPlanReader();
        }

        [TestMethod]
        public void GetWirelessHeaderIndexs_AllHeadersPresent_SetsAllIndexes()
        {
            // Arrange
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells[1, 1].Value = "TrimFuseName";
            sheet.Cells[1, 2].Value = "TrimTarget";
            sheet.Cells[1, 3].Value = "TrimMeas";
            sheet.Cells[1, 4].Value = "BestCodeCalcFunc";
            sheet.Cells[1, 5].Value = "TrimType";
            sheet.Cells[1, 6].Value = "RF Instrument Setup";
            sheet.Cells[1, 7].Value = "RF Test Type";
            sheet.Cells[1, 8].Value = "RF Interpose";
            sheet.Cells[1, 9].Value = "BB Instrument Setup";
            sheet.Cells[1, 10].Value = "BB Device Setup";
            sheet.Cells[1, 11].Value = "BB Calc";
            _reader.ExcelWorksheet = sheet;
            _reader._endColNumber = 11;

            // Act
            _reader.GetWirelessHeaderIndexs();

            // Assert
            Assert.AreEqual(1, _reader._oTpName);
            Assert.AreEqual(2, _reader._trimTarget);
            Assert.AreEqual(3, _reader._trimMeas);
            Assert.AreEqual(4, _reader._trimCalcEqn);
            Assert.AreEqual(5, _reader._trimType);
            Assert.AreEqual(6, _reader._rfSetup);
            Assert.AreEqual(7, _reader._rfCalc);
            Assert.AreEqual(8, _reader._rfInterpose);
            Assert.AreEqual(9, _reader._bbSetup);
            Assert.AreEqual(10, _reader._bbDeviceSetup);
            Assert.AreEqual(11, _reader._bbCalc);
        }

        [TestMethod]
        public void GetWirelessHeaderIndexs_TrimRegNameAlias_SetsOTpNameIndex()
        {
            // Arrange
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells[1, 1].Value = "TrimRegName";
            _reader.ExcelWorksheet = sheet;
            _reader._endColNumber = 1;

            // Act
            _reader.GetWirelessHeaderIndexs();

            // Assert
            Assert.AreEqual(1, _reader._oTpName);
        }

        [TestMethod]
        public void GetWirelessHeaderIndexs_TrimCalcEqnUnderscoredAlias_SetsTrimCalcEqnIndex()
        {
            // Arrange
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells[1, 1].Value = "Trim_Calc_Eqn";
            _reader.ExcelWorksheet = sheet;
            _reader._endColNumber = 1;

            // Act
            _reader.GetWirelessHeaderIndexs();

            // Assert
            Assert.AreEqual(1, _reader._trimCalcEqn);
        }

        [TestMethod]
        public void GetWirelessHeaderIndexs_RfInstrumentSetupWithUnderscores_MatchesAfterStripping()
        {
            // Arrange - underscores are stripped from the header before the regex match
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells[1, 1].Value = "RF_Instrument_Setup";
            _reader.ExcelWorksheet = sheet;
            _reader._endColNumber = 1;

            // Act
            _reader.GetWirelessHeaderIndexs();

            // Assert
            Assert.AreEqual(1, _reader._rfSetup);
        }

        [TestMethod]
        public void GetWirelessHeaderIndexs_UnrecognizedHeader_NoIndexSet()
        {
            // Arrange
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells[1, 1].Value = "SomethingElse";
            _reader.ExcelWorksheet = sheet;
            _reader._endColNumber = 1;

            // Act
            _reader.GetWirelessHeaderIndexs();

            // Assert
            Assert.AreEqual(0, _reader._oTpName);
            Assert.AreEqual(0, _reader._trimTarget);
            Assert.AreEqual(0, _reader._rfSetup);
        }

        [TestMethod]
        public void GetDimensions_WorksheetHasDimension_ReturnsTrueAndSetsEndColumn()
        {
            // Arrange
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells[1, 1].Value = "Header";
            sheet.Cells[2, 3].Value = "Data";
            _reader.ExcelWorksheet = sheet;

            // Act
            bool result = _reader.GetDimensions();

            // Assert
            Assert.IsTrue(result);
            Assert.AreEqual(3, _reader._endColNumber);
        }

        [TestMethod]
        public void GetDimensions_EmptyWorksheetWithNoDimension_ReturnsFalse()
        {
            // Arrange
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");
            _reader.ExcelWorksheet = sheet;

            // Act
            bool result = _reader.GetDimensions();

            // Assert
            Assert.IsFalse(result);
        }
    }
}
