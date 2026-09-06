using System;
using System.Collections.Generic;
using System.IO;

using CommonLib.Utility;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace CommonLib.Test.UT.Utility
{
    [TestClass]
    public class FileManagerTests
    {
        private string _tempDir;
        private string _outputExcelPath;

        [TestInitialize]
        public void Setup()
        {
            _tempDir = Path.Combine(Path.GetTempPath(), "FileManagerTests_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_tempDir);
            _outputExcelPath = Path.Combine(_tempDir, "MergedOutput.xlsx");
        }

        [TestCleanup]
        public void Teardown()
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, recursive: true);
            }
        }

        [TestMethod]
        public void CopyFile_SourceExists_CreatesTargetFile()
        {
            string sourceFile = Path.Combine(_tempDir, "source.txt");
            string targetFile = Path.Combine(_tempDir, "target.txt");
            File.WriteAllText(sourceFile, "hello");

            FileManager.CopyFile(sourceFile, targetFile);

            Assert.IsTrue(File.Exists(targetFile));
            Assert.AreEqual("hello", File.ReadAllText(targetFile));
        }

        [TestMethod]
        public void CopyFile_TargetAlreadyExists_Overwrites()
        {
            string sourceFile = Path.Combine(_tempDir, "source.txt");
            string targetFile = Path.Combine(_tempDir, "target.txt");
            File.WriteAllText(sourceFile, "new content");
            File.WriteAllText(targetFile, "old content");

            FileManager.CopyFile(sourceFile, targetFile);

            Assert.AreEqual("new content", File.ReadAllText(targetFile));
        }

        [TestMethod]
        public void CopyFile_ReturnsTargetPath()
        {
            string sourceFile = Path.Combine(_tempDir, "source.txt");
            string targetFile = Path.Combine(_tempDir, "target.txt");
            File.WriteAllText(sourceFile, "data");

            string result = FileManager.CopyFile(sourceFile, targetFile);

            Assert.AreEqual(targetFile, result);
        }

        [TestMethod]
        public void CopyFolder_CopiesAllFiles()
        {
            string sourceDir = Path.Combine(_tempDir, "src");
            string destDir = Path.Combine(_tempDir, "dst");
            Directory.CreateDirectory(sourceDir);
            File.WriteAllText(Path.Combine(sourceDir, "a.txt"), "A");
            File.WriteAllText(Path.Combine(sourceDir, "b.txt"), "B");

            FileManager.CopyFolder(sourceDir, destDir);

            Assert.IsTrue(File.Exists(Path.Combine(destDir, "a.txt")));
            Assert.IsTrue(File.Exists(Path.Combine(destDir, "b.txt")));
        }

        [TestMethod]
        public void CopyFolder_CopiesSubDirectories()
        {
            string sourceDir = Path.Combine(_tempDir, "src");
            string subDir = Path.Combine(sourceDir, "sub");
            string destDir = Path.Combine(_tempDir, "dst");
            Directory.CreateDirectory(subDir);
            File.WriteAllText(Path.Combine(subDir, "c.txt"), "C");

            FileManager.CopyFolder(sourceDir, destDir);

            Assert.IsTrue(File.Exists(Path.Combine(destDir, "sub", "c.txt")));
        }

        [TestMethod]
        public void CopyFolder_CreatesDestinationIfNotExists()
        {
            string sourceDir = Path.Combine(_tempDir, "src");
            string destDir = Path.Combine(_tempDir, "newdest");
            Directory.CreateDirectory(sourceDir);
            File.WriteAllText(Path.Combine(sourceDir, "file.txt"), "data");

            FileManager.CopyFolder(sourceDir, destDir);

            Assert.IsTrue(Directory.Exists(destDir));
        }

        [TestMethod]
        public void CopyDirectory_CopiesFilesAndSubDirs()
        {
            string sourceDir = Path.Combine(_tempDir, "src");
            string subDir = Path.Combine(sourceDir, "sub");
            string destDir = Path.Combine(_tempDir, "dst");
            Directory.CreateDirectory(subDir);
            File.WriteAllText(Path.Combine(sourceDir, "root.txt"), "root");
            File.WriteAllText(Path.Combine(subDir, "child.txt"), "child");

            FileManager.CopyDirectory(sourceDir, destDir);

            Assert.IsTrue(File.Exists(Path.Combine(destDir, "root.txt")));
            Assert.IsTrue(File.Exists(Path.Combine(destDir, "sub", "child.txt")));
        }

        [TestMethod]
        public void CopyDirectory_CopySubDirsFalse_SkipsSubDirs()
        {
            string sourceDir = Path.Combine(_tempDir, "src");
            string subDir = Path.Combine(sourceDir, "sub");
            string destDir = Path.Combine(_tempDir, "dst");
            Directory.CreateDirectory(subDir);
            File.WriteAllText(Path.Combine(sourceDir, "root.txt"), "root");
            File.WriteAllText(Path.Combine(subDir, "child.txt"), "child");

            FileManager.CopyDirectory(sourceDir, destDir, copySubDirs: false);

            Assert.IsTrue(File.Exists(Path.Combine(destDir, "root.txt")));
            Assert.IsFalse(Directory.Exists(Path.Combine(destDir, "sub")));
        }

        [TestMethod]
        public void CopyDirectory_SourceNotExists_ThrowsException()
        {
            Assert.ThrowsException<DirectoryNotFoundException>(() => FileManager.CopyDirectory(Path.Combine(_tempDir, "nonexistent"), Path.Combine(_tempDir, "dst")));
        }

        [TestMethod]
        public void MergeTestPlanCsv_ValidCsvInput_LoadsDataAndCreatesSheet()
        {
            string csvFile = Path.Combine(_tempDir, "TestPlanA.csv");
            File.WriteAllText(csvFile, "Header1,Header2\r\nValue1,Value2");
            var files = new List<string> { csvFile };

            FileManager.MergeTestPlanCsv(files, _outputExcelPath);

            using var pkg = new ExcelPackage(new FileInfo(_outputExcelPath));
            Assert.AreEqual(1, pkg.Workbook.Worksheets.Count);
            ExcelWorksheet sheet = pkg.Workbook.Worksheets["TestPlanA"];
            Assert.IsNotNull(sheet);
            Assert.AreEqual("Header1", sheet.Cells[1, 1].Value.ToString());
            Assert.AreEqual("Value2", sheet.Cells[2, 2].Value.ToString());
        }

        [TestMethod]
        public void MergeTestPlanCsv_ValidCsvInput_LoadsDataAndCreatesSheet_1()
        {
            string csvFile = Path.Combine(_tempDir, "TestPlanA.csv");
            File.WriteAllText(csvFile, "Header1,Header2\nValue1,Value2");
            var files = new List<string> { csvFile };

            FileManager.MergeTestPlanCsv(files, _outputExcelPath);

            using var pkg = new ExcelPackage(new FileInfo(_outputExcelPath));
            Assert.AreEqual(1, pkg.Workbook.Worksheets.Count);
            ExcelWorksheet sheet = pkg.Workbook.Worksheets["TestPlanA"];
            Assert.IsNotNull(sheet);
            Assert.AreEqual("Header1", sheet.Cells[1, 1].Value.ToString());
            Assert.AreEqual("Value2", sheet.Cells[2, 2].Value.ToString());
        }

        [TestMethod]
        public void MergeTestPlanCsv_ValidExcelInput_CopiesCellsRowByRow()
        {
            string sourceExcel = Path.Combine(_tempDir, "SourcePlan.xlsx");
            using (var srcPkg = new ExcelPackage(new FileInfo(sourceExcel)))
            {
                ExcelWorksheet ws = srcPkg.Workbook.Worksheets.Add("SourceSheet");
                ws.Cells[1, 1].Value = "CellA1";
                ws.Cells[2, 3].Value = "CellC2";
                srcPkg.Save();
            }
            var files = new List<string> { sourceExcel };

            FileManager.MergeTestPlanCsv(files, _outputExcelPath);

            using var pkg = new ExcelPackage(new FileInfo(_outputExcelPath));
            ExcelWorksheet destSheet = pkg.Workbook.Worksheets["SourceSheet"];
            Assert.IsNotNull(destSheet);
            Assert.AreEqual("CellA1", destSheet.Cells[1, 1].Value.ToString());
            Assert.AreEqual("CellC2", destSheet.Cells[2, 3].Value.ToString());
        }

        [TestMethod]
        public void MergeTestPlanCsv_LongSheetName_TruncatesTo31Characters()
        {
            // Creating a filename context string clearly exceeding the 31 char max cap rule bounds
            string longName = "ThisIsAVeryLongFileNameThatExceedsTheThirtyOneCharacterLimitSpecification";
            string csvFile = Path.Combine(_tempDir, longName + ".csv");
            File.WriteAllText(csvFile, "Data");
            var files = new List<string> { csvFile };

            FileManager.MergeTestPlanCsv(files, _outputExcelPath);

            using var pkg = new ExcelPackage(new FileInfo(_outputExcelPath));
            Assert.AreEqual(1, pkg.Workbook.Worksheets.Count);
            string expectedName = longName[..31];
            Assert.IsNotNull(pkg.Workbook.Worksheets[expectedName]);
        }

        [TestMethod]
        public void MergeTestPlanCsv_DuplicateSheetNames_AppendsCounterAndMaintainsUniqueNaming()
        {
            string csvFile1 = Path.Combine(_tempDir, "SameName.csv");
            string csvFile2 = Path.Combine(_tempDir, "SameName.xlsx");
            File.WriteAllText(csvFile1, "CsvData");

            using (var srcPkg = new ExcelPackage(new FileInfo(csvFile2)))
            {
                // Forces sheet naming initialization collision target profiles
                srcPkg.Workbook.Worksheets.Add("SameName");
                srcPkg.Save();
            }
            var files = new List<string> { csvFile1, csvFile2 };

            FileManager.MergeTestPlanCsv(files, _outputExcelPath);

            using var pkg = new ExcelPackage(new FileInfo(_outputExcelPath));
            Assert.AreEqual(2, pkg.Workbook.Worksheets.Count);
            Assert.IsNotNull(pkg.Workbook.Worksheets["SameName"]);
            Assert.IsNotNull(pkg.Workbook.Worksheets["SameName_1"]);
        }

        [TestMethod]
        public void MergeTestPlanCsv_UnsortedInputFiles_SortsWorkbookSheetsAlphabetically()
        {
            string csvC = Path.Combine(_tempDir, "Charlie.csv");
            string csvA = Path.Combine(_tempDir, "Alpha.csv");
            string csvB = Path.Combine(_tempDir, "Bravo.csv");

            File.WriteAllText(csvC, "Data");
            File.WriteAllText(csvA, "Data");
            File.WriteAllText(csvB, "Data");

            // Inserted deliberately unsorted out of alphabetical alignment sequence mapping
            var files = new List<string> { csvC, csvA, csvB };

            FileManager.MergeTestPlanCsv(files, _outputExcelPath);

            using var pkg = new ExcelPackage(new FileInfo(_outputExcelPath));
            Assert.AreEqual(3, pkg.Workbook.Worksheets.Count);
            Assert.AreEqual("Alpha", pkg.Workbook.Worksheets[0].Name);
            Assert.AreEqual("Bravo", pkg.Workbook.Worksheets[1].Name);
            Assert.AreEqual("Charlie", pkg.Workbook.Worksheets[2].Name);
        }

        [TestMethod]
        public void MergeTestPlanCsv_ExcelWithEmptyWorksheetDimensions_HandlesGracefullyWithoutCrashing()
        {
            string sourceExcel = Path.Combine(_tempDir, "EmptySheet.xlsx");
            using (var srcPkg = new ExcelPackage(new FileInfo(sourceExcel)))
            {
                srcPkg.Workbook.Worksheets.Add("NoDimensionSheet");
                // Leaves worksheet entirely empty, causing srcSheet.Dimension to be null
                srcPkg.Save();
            }
            var files = new List<string> { sourceExcel };

            FileManager.MergeTestPlanCsv(files, _outputExcelPath);

            using var pkg = new ExcelPackage(new FileInfo(_outputExcelPath));
            Assert.AreEqual(1, pkg.Workbook.Worksheets.Count);
            Assert.IsNotNull(pkg.Workbook.Worksheets["NoDimensionSheet"]);
        }
    }
}
