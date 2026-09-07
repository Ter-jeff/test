using System.IO;

using CommonLib.Extension;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace CommonLib.Test.UT.Extension
{
    [TestClass]
    public class EpplusExtensionsExcelPackageTests
    {
        private string _tempDir;

        [TestInitialize]
        public void Setup()
        {
            _tempDir = Path.Combine(Path.GetTempPath(), "EpplusExtTests_" + System.Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_tempDir);
        }

        [TestCleanup]
        public void Teardown()
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, recursive: true);
            }
        }

        // ─── ExcelPackage: ExportWorkBook2Txt ──────────────────────────────────
        [TestMethod]
        public void ExportWorkBook2Txt_WithWorksheets_CreatesOutputDirectory()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "Test";

            string outputPath = Path.Combine(_tempDir, "output");
            pkg.ExportWorkBook2Txt(outputPath);

            Assert.IsTrue(Directory.Exists(outputPath));
        }

        [TestMethod]
        public void ExportWorkBook2Txt_WithWorksheets_CreatesTxtFiles()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws1 = pkg.Workbook.Worksheets.Add("Sheet1");
            ws1.Cells[1, 1].Value = "Data1";
            ExcelWorksheet ws2 = pkg.Workbook.Worksheets.Add("Sheet2");
            ws2.Cells[1, 1].Value = "Data2";

            string outputPath = Path.Combine(_tempDir, "output");
            pkg.ExportWorkBook2Txt(outputPath);

            Assert.IsTrue(File.Exists(Path.Combine(outputPath, "Sheet1.txt")));
            Assert.IsTrue(File.Exists(Path.Combine(outputPath, "Sheet2.txt")));
        }

        [TestMethod]
        public void ExportWorkBook2Txt_ExistingDirectory_DoesNotThrow()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "Test";

            string outputPath = Path.Combine(_tempDir, "output");
            Directory.CreateDirectory(outputPath);

            // Should not throw when directory already exists
            pkg.ExportWorkBook2Txt(outputPath);

            Assert.IsTrue(File.Exists(Path.Combine(outputPath, "Sheet1.txt")));
        }

        [TestMethod]
        public void ExportWorkBook2Txt_EmptyWorkbook_CreatesDirectoryButNoFiles()
        {
            using var pkg = new ExcelPackage();
            string outputPath = Path.Combine(_tempDir, "output");
            pkg.ExportWorkBook2Txt(outputPath);

            Assert.IsTrue(Directory.Exists(outputPath));
        }
    }
}
