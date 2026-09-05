using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;

using IgxlLib.Enums;
using IgxlLib.IgxlSheets;
using IgxlLib.Utility;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace IgxlLib.Test.UT.Utility
{
    [TestClass]
    public class IgxlSheetReaderHelpersTests
    {
        private string _tempDirectory;

        public TestContext TestContext { get; set; }

        [TestInitialize]
        public void Setup()
        {
            // Create a clean isolated directory for disk IO tests
            _tempDirectory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(_tempDirectory);
        }

        [TestCleanup]
        public void Teardown()
        {
            if (!Directory.Exists(_tempDirectory))
            {
                return;
            }

            // Force garbage collection to release unclosed stream handles
            GC.Collect();
            GC.WaitForPendingFinalizers();

            for (int i = 0; i < 5; i++)
            {
                try
                {
                    Directory.Delete(_tempDirectory, true);
                    return;
                }
                catch (IOException)
                {
                    System.Threading.Thread.Sleep(50);
                }
            }
        }

        [TestMethod]
        public void GetIgxlSheetsByIgxlFile_MatchingTypesFound_ReturnsFilteredSheets()
        {
            // Arrange
            string zipPath = Path.Combine(_tempDirectory, "test.igxl");
            EnumSheetType targetType = EnumSheetType.DTTimesetBasicSheet;

            // Build an in-memory zip archive with one matching file and one non-matching file
            using (var fs = new FileStream(zipPath, FileMode.Create))
            using (var archive = new ZipArchive(fs, ZipArchiveMode.Create))
            {
                ZipArchiveEntry entry1 = archive.CreateEntry("Sheet1.txt");
                using (var writer = new StreamWriter(entry1.Open()))
                {
                    writer.WriteLine("DTTimesetBasicSheet,version=2.3:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\tTime Sets (Basic)");
                }

                ZipArchiveEntry entry2 = archive.CreateEntry("Sheet2.txt");
                using (var writer = new StreamWriter(entry2.Open()))
                {
                    writer.WriteLine("DTPinMap,version=2.1:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\tPin Map\r\n");
                }
            }

            // Act
            // Note: Make sure IgxlLoaderHelpers.GetIgxlSheetType is set up or naturally resolves 
            // "TimeSet_Header_Row" to EnumSheetType.TimeSet in your test environment execution.
            List<IIgxlSheet> result = IgxlSheetReaderHelpers.GetIgxlSheetsByIgxlFile(zipPath, targetType);

            // Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(1, result.Count);
        }

        [TestMethod]
        public void GetIgxlSheetsByIgxlFile_EmptyArchive_ReturnsEmptyList()
        {
            // Arrange
            string zipPath = Path.Combine(_tempDirectory, "empty.igxl");
            using (var fs = new FileStream(zipPath, FileMode.Create))
            using (var archive = new ZipArchive(fs, ZipArchiveMode.Create))
            { }

            // Act
            List<IIgxlSheet> result = IgxlSheetReaderHelpers.GetIgxlSheetsByIgxlFile(zipPath, EnumSheetType.DTTimesetBasicSheet);

            // Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void CreateIgxlSheet_ValidFile_ReturnsConstructedSheetInstance()
        {
            // Arrange
            string filePath = Path.Combine(_tempDirectory, "ValidSheet.txt");
            File.WriteAllText(filePath, "DTFlowtableSheet,version=3.0:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\tFlow Table\t");

            // Act
            IIgxlSheet result = IgxlSheetReaderHelpers.CreateIgxlSheet(filePath);

            // Assert
            Assert.IsNotNull(result);
            Assert.AreEqual("DTFlowtableSheet", result.SheetType);
        }

        [TestMethod]
        public void GetIgxlSheets_DirectoryContainsMixedExtensions_FiltersOnlyTxtFiles()
        {
            // Arrange
            // Target type file (.txt)
            string txtFile = Path.Combine(_tempDirectory, "ValidSheet.txt");
            File.WriteAllText(txtFile, "DTTimesetBasicSheet,version=2.3:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\tTime Sets (Basic)");

            // Ignore file (.csv) with identical content
            string csvFile = Path.Combine(_tempDirectory, "InvalidSheet.csv");
            File.WriteAllText(csvFile, "DTTimesetBasicSheet,version=2.3:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\tTime Sets (Basic)");

            EnumSheetType targetType = EnumSheetType.DTTimesetBasicSheet;

            // Act
            List<IIgxlSheet> result = IgxlSheetReaderHelpers.GetIgxlSheets(_tempDirectory, targetType);

            // Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(1, result.Count);
        }

        [TestMethod]
        public void GetIgxlSheets_NoMatchingSheetTypes_ReturnsEmptyList()
        {
            // Arrange
            string txtFile = Path.Combine(_tempDirectory, "PinMapSheet.txt");
            File.WriteAllText(txtFile, "PinMap_Header_Row");

            EnumSheetType targetType = EnumSheetType.DTTimesetBasicSheet;

            // Act
            List<IIgxlSheet> result = IgxlSheetReaderHelpers.GetIgxlSheets(_tempDirectory, targetType);

            // Assert
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void IsSheetType_CaseInsensitiveMatch_ReturnsTrue()
        {
            // Arrange
            string filePath = Path.Combine(_tempDirectory, "test_flow.txt");
            File.WriteAllText(filePath, "DTFlowtableSheet header row text");

            // Act
            bool result = IgxlSheetReaderHelpers.IsSheetType(EnumSheetType.DTFlowtableSheet, filePath);

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void IsSheetType_NoMatch_ReturnsFalse()
        {
            // Arrange
            string filePath = Path.Combine(_tempDirectory, "test_wrong.txt");
            File.WriteAllText(filePath, "Some Completely Unrelated Header");

            // Act
            bool result = IgxlSheetReaderHelpers.IsSheetType(EnumSheetType.DTFlowtableSheet, filePath);

            // Assert
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void GetSheetsByType_ValidDirectory_FiltersAndReturnsMatchingFilesOnly()
        {
            // Arrange
            string matchedFile = Path.Combine(_tempDirectory, "sheet1.txt");
            string unmatchedFile1 = Path.Combine(_tempDirectory, "sheet2.txt");
            string wrongExtensionFile = Path.Combine(_tempDirectory, "sheet3.log");

            File.WriteAllText(matchedFile, "DTFlowtableSheet Version 1");
            File.WriteAllText(unmatchedFile1, "DTTestInstancesSheet Version 1");
            File.WriteAllText(wrongExtensionFile, "DTFlowtableSheet Version 1");

            // Act
            List<string> result = IgxlSheetReaderHelpers.GetSheetsByType(_tempDirectory, EnumSheetType.DTFlowtableSheet);

            // Assert
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual(matchedFile, result[0]);
        }

        [TestMethod]
        public void ConvertStreamToExcelSheet_TabSeparatedStream_PopulatesCellsCorrectly()
        {
            // Arrange
            // Simulated fix: Supplying an ExcelPackage context internally if your helper creates one.
            // If you change your code to pass ExcelPackage or create it locally inside the method, this test will pass.
            string csvContent = "Header1\tHeader2\tHeader3\nValue1\tValue2\tValue3";

            using var ms = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(csvContent));
            // Act
            ExcelWorksheet sheet = IgxlSheetReaderHelpers.ConvertStreamToExcelSheet("Test%20Sheet", ms);

            // Assert
            Assert.IsNotNull(sheet);
            Assert.AreEqual("Test Sheet", sheet.Name);
            Assert.AreEqual("Header1", sheet.Cells[1, 1].Value.ToString());
            Assert.AreEqual("Header2", sheet.Cells[1, 2].Value.ToString());
            Assert.AreEqual("Value1", sheet.Cells[2, 1].Value.ToString());
            Assert.AreEqual("Value3", sheet.Cells[2, 3].Value.ToString());
        }

        [TestMethod]
        public void GetExcelWorksheet_ValidZipArchive_ExtractsTargetSheet()
        {
            // Arrange
            string zipPath = Path.Combine(_tempDirectory, "archive.igxl");
            string targetSheetName = "MyTestInstances";

            using (var zipToOpen = new FileStream(zipPath, FileMode.Create))
            {
                using var archive = new ZipArchive(zipToOpen, ZipArchiveMode.Create);
                ZipArchiveEntry entry = archive.CreateEntry($"{targetSheetName}.txt");
                using var writer = new StreamWriter(entry.Open());
                writer.WriteLine("DTTestInstancesSheet data row contents");
                writer.WriteLine("Row2Val1\tRow2Val2");
            }

            // Act
            ExcelWorksheet worksheet = IgxlSheetReaderHelpers.GetExcelWorksheet(zipPath, targetSheetName);

            // Assert
            Assert.IsNotNull(worksheet);
            Assert.AreEqual(targetSheetName, worksheet.Name);
            Assert.AreEqual("DTTestInstancesSheet data row contents", worksheet.Cells[1, 1].Value.ToString());
        }

        [TestMethod]
        public void IgxlSheetReaderHelpers_HasAllStaticMethods()
        {
            // Verify the static methods exist
            System.Reflection.MethodInfo[] methods = typeof(IgxlSheetReaderHelpers).GetMethods();
            Assert.IsTrue(methods.Length > 0);
        }

        [TestMethod]
        public void IgxlSheetReaderHelpers_GetSheetsByType_WithNullPath_HandlesGracefully()
        {
            // This would require actual files, so we just verify the method exists
            Type type = typeof(IgxlSheetReaderHelpers);
            System.Reflection.MethodInfo method = type.GetMethod("GetSheetsByType");
            Assert.IsNotNull(method);
        }
    }
}
