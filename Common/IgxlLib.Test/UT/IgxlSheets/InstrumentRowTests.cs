using System;
using System.Collections.Generic;
using System.IO;

using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace IgxlLib.Test.UT.IgxlSheets
{
    [TestClass]
    public class InstrumentRowTests
    {
        private string _tempExcelPath;

        [TestCleanup]
        public void Cleanup()
        {
            if (File.Exists(_tempExcelPath))
            {
                File.Delete(_tempExcelPath);
            }
        }

        [TestMethod]
        public void Reader_FileDoesNotExist_ReturnsEmptyList()
        {
            // Arrange
            string missingFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");

            // Act
            List<InstrumentRow> result = InstrumentRow.Reader(missingFile);

            // Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void Reader_MissingIFoldWorksheet_ReturnsEmptyList()
        {
            // Arrange
            _tempExcelPath = Path.Combine(Path.GetTempPath(), "MissingSheet.xlsx");
            using (var package = new ExcelPackage(new FileInfo(_tempExcelPath)))
            {
                package.Workbook.Worksheets.Add("WrongSheetName");
                package.Save();
            }

            // Act
            List<InstrumentRow> result = InstrumentRow.Reader(_tempExcelPath);

            // Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void Reader_MissingRequiredHeaders_ReturnsEmptyList()
        {
            // Arrange
            _tempExcelPath = Path.Combine(Path.GetTempPath(), "IncompleteHeaders.xlsx");
            using (var package = new ExcelPackage(new FileInfo(_tempExcelPath)))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets.Add("IFold");
                sheet.Cells[1, 1].Value = "Min";
                sheet.Cells[1, 2].Value = "Max";
                sheet.Cells[1, 3].Value = "Instrument";
                package.Save();
            }

            // Act
            List<InstrumentRow> result = InstrumentRow.Reader(_tempExcelPath);

            // Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void Reader_ValidExcelFile_ParsesDataRowsCorrectly()
        {
            // Arrange
            _tempExcelPath = Path.Combine(Path.GetTempPath(), "ValidInstrumentFile.xlsx");
            using (var package = new ExcelPackage(new FileInfo(_tempExcelPath)))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets.Add("IFold");

                // Write Header Row (Headers mapped dynamically via hardcoded values 1-6 in original logic)
                sheet.Cells[1, 1].Value = "Min";
                sheet.Cells[1, 2].Value = "Max";
                sheet.Cells[1, 3].Value = "CurrentType";
                sheet.Cells[1, 4].Value = "IsMerge";
                sheet.Cells[1, 5].Value = "Instrument";
                sheet.Cells[1, 6].Value = "IFold";

                // Write Data Row 1
                sheet.Cells[2, 1].Value = "10";
                sheet.Cells[2, 2].Value = "100";
                sheet.Cells[2, 3].Value = "DC";
                sheet.Cells[2, 4].Value = "True";
                sheet.Cells[2, 5].Value = "HVI";
                sheet.Cells[2, 6].Value = "FoldA";

                // Write Data Row 2 (Testing Null CurrentType)
                sheet.Cells[3, 1].Value = "20";
                sheet.Cells[3, 2].Value = "200";
                sheet.Cells[3, 3].Value = null;
                sheet.Cells[3, 4].Value = "False";
                sheet.Cells[3, 5].Value = "DMM";
                sheet.Cells[3, 6].Value = "FoldB";

                package.Save();
            }

            // Act
            List<InstrumentRow> result = InstrumentRow.Reader(_tempExcelPath);

            // Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(2, result.Count);

            // Verify Row 1
            Assert.AreEqual("10", result[0].Min);
            Assert.AreEqual("100", result[0].Max);
            Assert.AreEqual("DC", result[0].CurrentType);
            Assert.AreEqual("True", result[0].IsMerge);
            Assert.AreEqual("HVI", result[0].Instrument);
            Assert.AreEqual("FoldA", result[0].IFold);

            // Verify Row 2
            Assert.AreEqual("", result[1].CurrentType);
            Assert.AreEqual("DMM", result[1].Instrument);
        }
    }
}
