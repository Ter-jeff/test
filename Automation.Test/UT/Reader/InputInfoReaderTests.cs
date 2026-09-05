using System;
using System.IO;

using CommonReaderLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace Automation.Test.UT.Reader
{
    [TestClass]
    public class InputInfoReaderTests
    {
        private ExcelPackage _package = null!;

        [TestInitialize]
        public void Setup()
        {
            _package = new ExcelPackage();
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package?.Dispose();
        }

        [TestMethod]
        public void ReadSheet_NullWorksheet_ReturnsNull()
        {
            // Arrange
            var reader = new InputInfoReader(Path.Combine("C:", "input", "input.xlsx"));

            // Act
            InputInfoSheet? result = reader.ReadSheet(null!);

            // Assert
            Assert.AreEqual(null, result);
        }

        [TestMethod]
        public void ReadSheet_EmptyWorksheet_ReturnsNull()
        {
            // Arrange
            var reader = new InputInfoReader(Path.Combine("C:", "input", "input.xlsx"));
            ExcelWorksheet ws = _package.Workbook.Worksheets.Add("Sheet1");

            // Act
            InputInfoSheet? result = reader.ReadSheet(ws);

            // Assert
            Assert.AreEqual(null, result);
        }

        [TestMethod]
        public void ReadSheet_WorksheetWithoutItemHeader_ReturnsNull()
        {
            // Arrange
            var reader = new InputInfoReader(Path.Combine("C:", "input", "input.xlsx"));
            ExcelWorksheet ws = _package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "NotItem";
            ws.Cells[1, 2].Value = "Other";

            // Act
            InputInfoSheet? result = reader.ReadSheet(ws);

            // Assert
            Assert.AreEqual(null, result);
        }

        [TestMethod]
        public void ReadSheet_ValidWorksheetAllHeaders_SetsCorrectIndices()
        {
            // Arrange
            var reader = new InputInfoReader(Path.Combine("C:", "input", "input.xlsx"));
            ExcelWorksheet ws = _package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "Item";
            ws.Cells[1, 2].Value = "Value";
            ws.Cells[1, 3].Value = "Description";
            ws.Cells[2, 1].Value = "someitem";
            ws.Cells[2, 2].Value = "somevalue";
            ws.Cells[2, 3].Value = "somedesc";

            // Act
            InputInfoSheet? result = reader.ReadSheet(ws);

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.AreEqual(1, result!.IndexItem);
            Assert.AreEqual(2, result.IndexValue);
            Assert.AreEqual(3, result.IndexDescription);
        }

        [TestMethod]
        public void ReadSheet_ValidWorksheetOnlyItemHeader_ValueAndDescriptionIndexAreMinus1()
        {
            // Arrange
            var reader = new InputInfoReader(Path.Combine("C:", "input", "input.xlsx"));
            ExcelWorksheet ws = _package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "Item";
            ws.Cells[2, 1].Value = "someitem";

            // Act
            InputInfoSheet? result = reader.ReadSheet(ws);

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.AreEqual(1, result!.IndexItem);
            Assert.AreEqual(-1, result.IndexValue);
            Assert.AreEqual(-1, result.IndexDescription);
        }

        [TestMethod]
        public void ReadSheet_ValidWorksheet_SheetNameMatchesWorksheetName()
        {
            // Arrange
            var reader = new InputInfoReader(Path.Combine("C:", "input", "input.xlsx"));
            ExcelWorksheet ws = _package.Workbook.Worksheets.Add("MyInputSheet");
            ws.Cells[1, 1].Value = "Item";
            ws.Cells[2, 1].Value = "someitem";

            // Act
            InputInfoSheet? result = reader.ReadSheet(ws);

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.AreEqual("MyInputSheet", result!.SheetName);
        }

        [TestMethod]
        public void ReadSheet_ValidWorksheetWithDataRows_AllRowsRead()
        {
            // Arrange
            var reader = new InputInfoReader(Path.Combine("C:", "input", "input.xlsx"));
            ExcelWorksheet ws = _package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "Item";
            ws.Cells[1, 2].Value = "Value";
            ws.Cells[2, 1].Value = "item1";
            ws.Cells[2, 2].Value = "value1";
            ws.Cells[3, 1].Value = "item2";
            ws.Cells[3, 2].Value = "value2";

            // Act
            InputInfoSheet? result = reader.ReadSheet(ws);

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.AreEqual(2, result!.Rows.Count);
            Assert.AreEqual("item1", result.Rows[0].Item);
            Assert.AreEqual("value1", result.Rows[0].Value);
            Assert.AreEqual("item2", result.Rows[1].Item);
            Assert.AreEqual("value2", result.Rows[1].Value);
        }

        [TestMethod]
        public void ReadSheet_HeaderIsCaseInsensitive_DetectsAllHeaders()
        {
            // Arrange
            var reader = new InputInfoReader(Path.Combine("C:", "input", "input.xlsx"));
            ExcelWorksheet ws = _package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "ITEM";
            ws.Cells[1, 2].Value = "VALUE";
            ws.Cells[1, 3].Value = "DESCRIPTION";
            ws.Cells[2, 1].Value = "a";

            // Act
            InputInfoSheet? result = reader.ReadSheet(ws);

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.AreEqual(1, result!.IndexItem);
            Assert.AreEqual(2, result.IndexValue);
            Assert.AreEqual(3, result.IndexDescription);
        }

        [TestMethod]
        public void ReadSheet_SingleDataRow_RowNumIsSetCorrectly()
        {
            // Arrange
            var reader = new InputInfoReader(Path.Combine("C:", "input", "input.xlsx"));
            ExcelWorksheet ws = _package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "Item";
            ws.Cells[2, 1].Value = "myitem";

            // Act
            InputInfoSheet? result = reader.ReadSheet(ws);

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.AreEqual(1, result!.Rows.Count);
            Assert.AreEqual(2, result.Rows[0].RowNum);
        }
    }

    [TestClass]
    public class InputInfoSheetModifyPathTests
    {
        [TestMethod]
        public void ModifyPath_AbsolutePath_PathRemainsUnchanged()
        {
            // Arrange
            var sheet = new InputInfoSheet("Sheet1");
            sheet.Rows.Add(new InputInfoRow("Sheet1") { Item = "testplanfile", Value = Path.Combine("C:", "absolute", "path", "plan.xlsx") });

            // Act
            sheet.ModifyPath(@"C:\some\dir\input.xlsx");

            // Assert
            Assert.AreEqual(Path.Combine("C:", "absolute", "path", "plan.xlsx"), sheet.Rows[0].Value);
        }

        [TestMethod]
        public void ModifyPath_RelativePath_CombinesWithInputInfoDirectory()
        {
            // Arrange
            var sheet = new InputInfoSheet("Sheet1");
            sheet.Rows.Add(new InputInfoRow("Sheet1") { Item = "testplanfile", Value = "subdir/plan.xlsx" });

            // Act
            sheet.ModifyPath(@"C:\some\dir\input.xlsx");

            // Assert
            string expected = Path.GetFullPath(Path.Combine(@"C:\some\dir", "subdir/plan.xlsx"));
            Assert.AreEqual(expected, sheet.Rows[0].Value);
        }

        [TestMethod]
        public void ModifyPath_DotPath_ReturnsInputInfoDirectory()
        {
            // Arrange
            var sheet = new InputInfoSheet("Sheet1");
            sheet.Rows.Add(new InputInfoRow("Sheet1") { Item = "repo", Value = "." });

            // Act
            sheet.ModifyPath(@"C:\some\dir\input.xlsx");

            // Assert
            string expected = Path.GetFullPath(@"C:\some\dir");
            Assert.AreEqual(expected, sheet.Rows[0].Value);
        }

        [TestMethod]
        public void ModifyPath_DotSlashPath_PathCombinedWithCurrentDir()
        {
            // Arrange
            var sheet = new InputInfoSheet("Sheet1");
            sheet.Rows.Add(new InputInfoRow("Sheet1") { Item = "repo", Value = "./" });

            // Act
            sheet.ModifyPath(@"C:\some\dir\input.xlsx");

            // Assert
            string expected = Path.GetFullPath(Path.Combine(@"C:\some\dir", "./"));
            Assert.AreEqual(expected, sheet.Rows[0].Value);
        }

        [TestMethod]
        public void ModifyPath_ParentPath_ReturnsParentDirectory()
        {
            // Arrange
            var sheet = new InputInfoSheet("Sheet1");
            sheet.Rows.Add(new InputInfoRow("Sheet1") { Item = "repo", Value = ".." });

            // Act
            sheet.ModifyPath(@"C:\some\dir\sub\input.xlsx");

            // Assert
            string expected = Path.GetFullPath(@"C:\some\dir");
            Assert.AreEqual(expected, sheet.Rows[0].Value);
        }

        [TestMethod]
        public void ModifyPath_ParentPathWithSubfolder_CombinesCorrectly()
        {
            // Arrange
            var sheet = new InputInfoSheet("Sheet1");
            sheet.Rows.Add(new InputInfoRow("Sheet1") { Item = "repo", Value = "../sibling" });

            // Act
            sheet.ModifyPath(@"C:\some\dir\sub\input.xlsx");

            // Assert
            string expected = Path.GetFullPath(@"C:\some\dir\sibling");
            Assert.AreEqual(expected, sheet.Rows[0].Value);
        }

        [TestMethod]
        public void ModifyPath_MultipleCommaSeparatedAbsolutePaths_EachPathNormalized()
        {
            // Arrange
            var sheet = new InputInfoSheet("Sheet1");
            sheet.Rows.Add(new InputInfoRow("Sheet1") { Item = "testplanfile", Value = $"{Path.Combine("C:", "abs", "a.xlsx")},{Path.Combine("C:", "abs", "b.xlsx")}" });

            // Act
            sheet.ModifyPath(@"C:\some\dir\input.xlsx");

            // Assert
            string expected = Path.GetFullPath(Path.Combine("C:", "abs", "a.xlsx")) + "," + Path.GetFullPath(Path.Combine("C:", "abs", "b.xlsx"));
            Assert.AreEqual(expected, sheet.Rows[0].Value);
        }

        [TestMethod]
        public void ModifyPath_RowWithEmptyItem_ValueUnchanged()
        {
            // Arrange
            var sheet = new InputInfoSheet("Sheet1");
            sheet.Rows.Add(new InputInfoRow("Sheet1") { Item = "", Value = "relative/path.xlsx" });

            // Act
            sheet.ModifyPath(@"C:\some\dir\input.xlsx");

            // Assert
            Assert.AreEqual("relative/path.xlsx", sheet.Rows[0].Value);
        }

        [TestMethod]
        public void ModifyPath_RowWithEmptyValue_ValueRemainsEmpty()
        {
            // Arrange
            var sheet = new InputInfoSheet("Sheet1");
            sheet.Rows.Add(new InputInfoRow("Sheet1") { Item = "testplanfile", Value = "" });

            // Act
            sheet.ModifyPath(@"C:\some\dir\input.xlsx");

            // Assert
            Assert.AreEqual("", sheet.Rows[0].Value);
        }

        [TestMethod]
        public void ModifyPath_ItemNotInModifySet_ValueUnchanged()
        {
            // Arrange
            var sheet = new InputInfoSheet("Sheet1");
            sheet.Rows.Add(new InputInfoRow("Sheet1") { Item = "unknownitem", Value = "relative/path.xlsx" });

            // Act
            sheet.ModifyPath(@"C:\some\dir\input.xlsx");

            // Assert
            Assert.AreEqual("relative/path.xlsx", sheet.Rows[0].Value);
        }

        [TestMethod]
        public void ModifyPath_ItemMatchIsCaseInsensitive_PathIsProcessed()
        {
            // Arrange
            var sheet = new InputInfoSheet("Sheet1");
            sheet.Rows.Add(new InputInfoRow("Sheet1") { Item = "TESTPLANFILE", Value = "sub/plan.xlsx" });

            // Act
            sheet.ModifyPath(@"C:\some\dir\input.xlsx");

            // Assert
            string expected = Path.GetFullPath(Path.Combine(@"C:\some\dir", "sub/plan.xlsx"));
            Assert.AreEqual(expected, sheet.Rows[0].Value);
        }
    }

    [TestClass]
    public class InputInfoSheetCheckTests
    {
        [TestMethod]
        public void Check_ExistingFile_AddsNoErrors()
        {
            // Arrange
            string tempFile = Path.GetTempFileName();
            try
            {
                var sheet = new InputInfoSheet("Sheet1");
                sheet.Rows.Add(new InputInfoRow("Sheet1") { Item = "testplanfile", Value = tempFile, RowNum = 2 });

                // Act
                sheet.Check();

                // Assert
                Assert.AreEqual(0, sheet.GetErrors().Count);
            }
            finally
            {
                File.Delete(tempFile);
            }
        }

        [TestMethod]
        public void Check_MissingFile_AddsOneError()
        {
            // Arrange
            string missingPath = Path.Combine(Path.GetTempPath(), $"missing_{Guid.NewGuid()}.xlsx");
            var sheet = new InputInfoSheet("Sheet1");
            sheet.Rows.Add(new InputInfoRow("Sheet1") { Item = "voltagetablefile", Value = missingPath, RowNum = 2 });

            // Act
            sheet.Check();

            // Assert
            Assert.AreEqual(1, sheet.GetErrors().Count);
        }

        [TestMethod]
        public void Check_ExistingDirectory_AddsNoErrors()
        {
            // Arrange
            string tempDir = Path.Combine(Path.GetTempPath(), $"InputInfoTest_{Guid.NewGuid()}");
            Directory.CreateDirectory(tempDir);
            try
            {
                var sheet = new InputInfoSheet("Sheet1");
                sheet.Rows.Add(new InputInfoRow("Sheet1") { Item = "repo", Value = tempDir, RowNum = 2 });

                // Act
                sheet.Check();

                // Assert
                Assert.AreEqual(0, sheet.GetErrors().Count);
            }
            finally
            {
                Directory.Delete(tempDir);
            }
        }

        [TestMethod]
        public void Check_MissingDirectory_AddsOneError()
        {
            // Arrange
            string missingDir = Path.Combine(Path.GetTempPath(), $"missing_{Guid.NewGuid()}");
            var sheet = new InputInfoSheet("Sheet1");
            sheet.Rows.Add(new InputInfoRow("Sheet1") { Item = "repo", Value = missingDir, RowNum = 2 });

            // Act
            sheet.Check();

            // Assert
            Assert.AreEqual(1, sheet.GetErrors().Count);
        }

        [TestMethod]
        public void Check_EmptyRows_AddsNoErrors()
        {
            // Arrange
            var sheet = new InputInfoSheet("Sheet1");

            // Act
            sheet.Check();

            // Assert
            Assert.AreEqual(0, sheet.GetErrors().Count);
        }

        [TestMethod]
        public void Check_RowWithEmptyItem_IsSkipped()
        {
            // Arrange
            string missingPath = Path.Combine(Path.GetTempPath(), $"missing_{Guid.NewGuid()}.xlsx");
            var sheet = new InputInfoSheet("Sheet1");
            sheet.Rows.Add(new InputInfoRow("Sheet1") { Item = "", Value = missingPath, RowNum = 2 });

            // Act
            sheet.Check();

            // Assert
            Assert.AreEqual(0, sheet.GetErrors().Count);
        }

        [TestMethod]
        public void Check_RowWithEmptyValue_IsSkipped()
        {
            // Arrange
            var sheet = new InputInfoSheet("Sheet1");
            sheet.Rows.Add(new InputInfoRow("Sheet1") { Item = "testplanfile", Value = "", RowNum = 2 });

            // Act
            sheet.Check();

            // Assert
            Assert.AreEqual(0, sheet.GetErrors().Count);
        }

        [TestMethod]
        public void Check_ItemNotInCheckSet_AddsNoErrors()
        {
            // Arrange
            string missingPath = Path.Combine(Path.GetTempPath(), $"missing_{Guid.NewGuid()}.xlsx");
            var sheet = new InputInfoSheet("Sheet1");
            sheet.Rows.Add(new InputInfoRow("Sheet1") { Item = "unknownitem", Value = missingPath, RowNum = 2 });

            // Act
            sheet.Check();

            // Assert
            Assert.AreEqual(0, sheet.GetErrors().Count);
        }

        [TestMethod]
        public void Check_MultiplePathsCommaSeparated_EachPathChecked()
        {
            // Arrange
            string missingA = Path.Combine(Path.GetTempPath(), $"missing_{Guid.NewGuid()}.xlsx");
            string missingB = Path.Combine(Path.GetTempPath(), $"missing_{Guid.NewGuid()}.xlsx");
            var sheet = new InputInfoSheet("Sheet1");
            sheet.Rows.Add(new InputInfoRow("Sheet1") { Item = "voltagetablefile", Value = $"{missingA},{missingB}", RowNum = 2 });

            // Act
            sheet.Check();

            // Assert
            Assert.AreEqual(2, sheet.GetErrors().Count);
        }
    }

    [TestClass]
    public class InputInfoRowTests
    {
        [TestMethod]
        public void InputInfoRow_Constructor_SetsSourceSheetName()
        {
            // Arrange & Act
            var row = new InputInfoRow("TestSheet");

            // Assert
            Assert.AreEqual("TestSheet", row.SourceSheetName);
        }

        [TestMethod]
        public void InputInfoRow_DefaultProperties_AreEmptyStrings()
        {
            // Arrange & Act
            var row = new InputInfoRow("Sheet1");

            // Assert
            Assert.AreEqual(string.Empty, row.Item);
            Assert.AreEqual(string.Empty, row.Value);
            Assert.AreEqual(string.Empty, row.Description);
        }

        [TestMethod]
        public void InputInfoRow_SetProperties_ReturnsSetValues()
        {
            // Arrange
            var row = new InputInfoRow("Sheet1")
            {
                // Act
                Item = "myItem",
                Value = "myValue",
                Description = "myDesc"
            };

            // Assert
            Assert.AreEqual("myItem", row.Item);
            Assert.AreEqual("myValue", row.Value);
            Assert.AreEqual("myDesc", row.Description);
        }
    }
}
