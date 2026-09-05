using System.Collections.Generic;
using System.Linq;

using Automation.Utility.TpUpdate.HardIPBinoutTPUpdate;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace Automation.Test.UT.PostAction
{
    [TestClass]
    public class BinoutStatusHipListTests
    {
        private static readonly string[] _headers =
        [
            "TestInstance",
            "TestNum",
            "TestName",
            "Type",
            "Pattern/Pin",
            "LoLimit",
            "HiLimit",
            "T/P Unit",
            "TestResult",
            "MeterIRangeSetting",
            "CP1_BinOut",
            "CP1_Enable",
            "T/P UseLimit",
            "Comment",
            "CP1_BinOut Update",
            "CP1_LSL Update",
            "CP1_USL Update",
            "Notes",
            "SwBin",
            "T/P LoLimit",
            "T/P HiLimit",
            "ExecutionProfileTestTime",
            "By Stage"
        ];

        private static readonly string[] _dataRow =
        [
            "Inst1",
            "1",
            "Test1",
            "Cont",
            "Pin1",
            "1.0",
            "2.0",
            "V",
            "Pass",
            "AutoRange",
            "1",
            "Y",
            "Yes",
            "Cmt",
            "U1",
            "L1",
            "H1",
            "Note1",
            "5",
            "0.5",
            "2.5",
            "10",
            "CP1"
        ];

        private static ExcelWorksheet BuildWorksheet(ExcelPackage package, string[] headerRow, string[]? dataRow)
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");
            for (int col = 0; col < headerRow.Length; col++)
            {
                worksheet.Cells[1, col + 1].Value = headerRow[col];
            }

            if (dataRow != null)
            {
                for (int col = 0; col < dataRow.Length; col++)
                {
                    worksheet.Cells[2, col + 1].Value = dataRow[col];
                }
            }

            return worksheet;
        }

        [TestMethod]
        public void ReadSheet_NullWorksheet_ReturnsNull()
        {
            // Arrange
            var reader = new BinOutStatusHipListReader(["CP1"]);

            // Act
            BinOutStatusHipListSheet? result = reader.ReadSheet(null!);

            // Assert
            Assert.IsNull(result);
        }

        [TestMethod]
        public void ReadSheet_EmptyWorksheetHasNoDimension_ReturnsNull()
        {
            // Arrange
            using var package = new ExcelPackage();
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Empty");
            var reader = new BinOutStatusHipListReader(["CP1"]);

            // Act
            BinOutStatusHipListSheet? result = reader.ReadSheet(worksheet);

            // Assert
            Assert.IsNull(result);
        }

        [TestMethod]
        public void ReadSheet_NoTestInstanceHeaderFound_ReturnsNull()
        {
            // Arrange
            using var package = new ExcelPackage();
            ExcelWorksheet worksheet = BuildWorksheet(package, ["SomeOtherHeader", "AnotherHeader"], null);
            var reader = new BinOutStatusHipListReader(["CP1"]);

            // Act
            BinOutStatusHipListSheet? result = reader.ReadSheet(worksheet);

            // Assert
            Assert.IsNull(result);
        }

        [TestMethod]
        public void ReadSheet_FullyPopulatedWorksheet_ParsesAllColumnsIntoRow()
        {
            // Arrange
            using var package = new ExcelPackage();
            ExcelWorksheet worksheet = BuildWorksheet(package, _headers, _dataRow);
            var jobs = new HashSet<string> { "CP1" };
            var reader = new BinOutStatusHipListReader(jobs);

            // Act
            BinOutStatusHipListSheet? result = reader.ReadSheet(worksheet);

            // Assert
            Assert.IsNotNull(result);
            Assert.AreEqual("Sheet1", result.SheetName);
            Assert.AreEqual(1, result.Rows.Count);

            BinOutStatusHipListRow row = result.Rows[0];
            Assert.AreEqual("Inst1", row.Testinstance);
            Assert.AreEqual("1", row.Testnum);
            Assert.AreEqual("Test1", row.Testname);
            Assert.AreEqual("Cont", row.Type);
            Assert.AreEqual("Pin1", row.PatternPin);
            Assert.AreEqual("1.0", row.Lolimit);
            Assert.AreEqual("2.0", row.Hilimit);
            Assert.AreEqual("V", row.Unit);
            Assert.AreEqual("Pass", row.TestResult);
            Assert.AreEqual("AutoRange", row.Meterirangesetting);
            Assert.AreEqual("Yes", row.TPuselimit);
            Assert.AreEqual("Cmt", row.Comment);
            Assert.AreEqual("Note1", row.Notes);
            Assert.AreEqual("5", row.SoftBin);
            Assert.AreEqual("0.5", row.TpLoLimit);
            Assert.AreEqual("2.5", row.TpHiLimit);
            Assert.AreEqual("10", row.ExecutionProfileTestTime);
            Assert.AreEqual("CP1", row.ByStage);

            Assert.AreEqual("1", row.StatusDictionary["CP1_BinOut"]);
            Assert.AreEqual("Y", row.EnableDictionary["CP1"]);
            Assert.AreEqual("U1", row.BinOutEnableDictionary["CP1"]);
            Assert.AreEqual("L1", row.UpdatedLoLimitDic["CP1"]);
            Assert.AreEqual("H1", row.UpdatedHiLimitDic["CP1"]);

            CollectionAssert.Contains(jobs.ToArray(), "CP1");
        }

        [TestMethod]
        public void ReadSheet_NoDataRows_ReturnsEmptyRowsList()
        {
            // Arrange
            using var package = new ExcelPackage();
            ExcelWorksheet worksheet = BuildWorksheet(package, _headers, null);
            var reader = new BinOutStatusHipListReader(["CP1"]);

            // Act
            BinOutStatusHipListSheet? result = reader.ReadSheet(worksheet);

            // Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(0, result.Rows.Count);
        }
    }
}
