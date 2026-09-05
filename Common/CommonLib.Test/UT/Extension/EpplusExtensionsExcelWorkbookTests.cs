using System;
using System.Collections.Generic;
using System.IO;

using CommonLib.Extension;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace CommonLib.Test.UT.Extension
{
    [TestClass]
    public class EpplusExtensionsExcelWorkbookTests
    {
        private string _tempDir;
        private string _txtFilePath;

        [TestInitialize]
        public void Setup()
        {
            _tempDir = Path.Combine(Path.GetTempPath(), "ExcelExtensionTests_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_tempDir);
            _txtFilePath = Path.Combine(_tempDir, "output.txt");
        }

        [TestCleanup]
        public void Teardown()
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, recursive: true);
            }
        }

        // ─── ExcelWorkbook: IsExist ───────────────────────────────────────────
        [TestMethod]
        public void IsExist_SheetPresent_ReturnsTrue()
        {
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("Alpha");
            Assert.IsTrue(pkg.Workbook.IsExist("Alpha"));
        }

        [TestMethod]
        public void IsExist_SheetAbsent_ReturnsFalse()
        {
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("Alpha");
            Assert.IsFalse(pkg.Workbook.IsExist("Beta"));
        }

        [TestMethod]
        public void IsExist_CaseInsensitive_FindsSheet()
        {
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("TestSheet");
            ExcelWorksheet result = pkg.Workbook.Worksheets["testsheet"];
            Assert.IsTrue(pkg.Workbook.IsExist("testsheet"));
            Assert.IsNotNull(result);
        }

        [TestMethod]
        public void IsExist_EmptyWorkbook_ReturnsFalse()
        {
            using var pkg = new ExcelPackage();
            Assert.IsFalse(pkg.Workbook.IsExist("AnySheet"));
        }

        [TestMethod]
        public void IsExist_NullName_ReturnsFalse()
        {
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("Alpha");
            Assert.IsFalse(pkg.Workbook.IsExist(null));
        }

        // ─── ExcelWorkbook: GetMatchPlanSheets ────────────────────────────────
        [TestMethod]
        public void GetMatchPlanSheets_PatternMatch_ReturnsMatchingSheets()
        {
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("PlanA");
            pkg.Workbook.Worksheets.Add("PlanB");
            pkg.Workbook.Worksheets.Add("Data");

            List<string> result = pkg.Workbook.GetMatchPlanSheets("^Plan");

            Assert.AreEqual(2, result.Count);
            CollectionAssert.Contains(result, "PlanA");
            CollectionAssert.Contains(result, "PlanB");
        }

        [TestMethod]
        public void GetMatchPlanSheets_NoMatch_ReturnsEmptyList()
        {
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("Alpha");
            List<string> result = pkg.Workbook.GetMatchPlanSheets("^XYZ");
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void GetMatchPlanSheets_CaseInsensitive_Matches()
        {
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("PLANX");
            List<string> result = pkg.Workbook.GetMatchPlanSheets("^plan");
            Assert.AreEqual(1, result.Count);
        }

        [TestMethod]
        public void GetMatchPlanSheets_EmptyWorkbook_ReturnsEmpty()
        {
            using var pkg = new ExcelPackage();
            List<string> result = pkg.Workbook.GetMatchPlanSheets("^Plan");
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void GetMatchPlanSheets_ComplexPattern_ReturnsMatching()
        {
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("Plan_01");
            pkg.Workbook.Worksheets.Add("Plan_02");
            pkg.Workbook.Worksheets.Add("Plan_A");
            pkg.Workbook.Worksheets.Add("Other");

            List<string> result = pkg.Workbook.GetMatchPlanSheets(@"^Plan_\d+");

            Assert.AreEqual(2, result.Count);
            CollectionAssert.Contains(result, "Plan_01");
            CollectionAssert.Contains(result, "Plan_02");
        }

        // ─── ExcelWorkbook: GetPlanSheets ─────────────────────────────────────
        [TestMethod]
        public void GetPlanSheets_PatternMatch_ReturnsMatchingSheets()
        {
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("TestPlan1");
            pkg.Workbook.Worksheets.Add("TestPlan2");
            pkg.Workbook.Worksheets.Add("Summary");

            List<string> result = pkg.Workbook.GetPlanSheets("TestPlan");

            Assert.AreEqual(2, result.Count);
        }

        [TestMethod]
        public void GetPlanSheets_NoMatch_ReturnsEmpty()
        {
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("Summary");
            List<string> result = pkg.Workbook.GetPlanSheets("TestPlan");
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void GetPlanSheets_EmptyWorkbook_ReturnsEmpty()
        {
            using var pkg = new ExcelPackage();
            List<string> result = pkg.Workbook.GetPlanSheets("TestPlan");
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void GetPlanSheets_CaseInsensitive_FindsSheets()
        {
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("TESTPLAN_A");
            pkg.Workbook.Worksheets.Add("TESTPLAN_B");

            List<string> result = pkg.Workbook.GetPlanSheets("testplan");

            Assert.AreEqual(2, result.Count);
        }

        // ─── ExcelWorkbook: DeleteSheet ───────────────────────────────────────
        [TestMethod]
        public void DeleteSheet_ExistingSheet_RemovesIt()
        {
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("Alpha");
            pkg.Workbook.Worksheets.Add("Beta");
            pkg.Workbook.DeleteSheet("Alpha");
            Assert.IsFalse(pkg.Workbook.IsExist("Alpha"));
            Assert.IsTrue(pkg.Workbook.IsExist("Beta"));
        }

        [TestMethod]
        public void DeleteSheet_NonExistingSheet_DoesNotThrow()
        {
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("Alpha");
            pkg.Workbook.DeleteSheet("NotHere");
            Assert.AreEqual(1, pkg.Workbook.Worksheets.Count);
        }

        [TestMethod]
        public void DeleteSheet_CaseInsensitive_RemovesSheet()
        {
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("Alpha");
            pkg.Workbook.Worksheets.Add("Beta");
            pkg.Workbook.DeleteSheet("ALPHA");
            Assert.IsFalse(pkg.Workbook.IsExist("Alpha"));
        }

        [TestMethod]
        public void DeleteSheet_NullName_DoesNotThrow()
        {
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("Alpha");
            pkg.Workbook.DeleteSheet(null);
            Assert.AreEqual(1, pkg.Workbook.Worksheets.Count);
        }

        // ─── ExcelWorkbook: AddSheet(string) ─────────────────────────────────
        [TestMethod]
        public void WorkbookAddSheet_NewName_AddsSheet()
        {
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("Existing");
            pkg.Workbook.AddSheet("NewSheet");
            Assert.IsTrue(pkg.Workbook.IsExist("NewSheet"));
        }

        [TestMethod]
        public void WorkbookAddSheet_ExistingName_ClearsContent()
        {
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("Anchor");
            pkg.Workbook.Worksheets.Add("Target");
            pkg.Workbook.Worksheets["Target"].Cells[1, 1].Value = "old data";

            pkg.Workbook.AddSheet("Target");

            Assert.IsNull(pkg.Workbook.Worksheets["Target"].Cells[1, 1].Value);
        }

        [TestMethod]
        public void WorkbookAddSheet_NewSheet_IsMovedToFront()
        {
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("First");
            pkg.Workbook.AddSheet("Second");
            Assert.AreEqual("Second", pkg.Workbook.Worksheets[0].Name);
        }

        [TestMethod]
        public void WorkbookAddSheet_MultipleNewSheets_LastOneAtFront()
        {
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("Original");
            pkg.Workbook.AddSheet("First");
            pkg.Workbook.AddSheet("Second");
            Assert.AreEqual("Second", pkg.Workbook.Worksheets[0].Name);
        }

        [TestMethod]
        public void ExportToTxt_NullOrEmptyDimension_ReturnsImmediately()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet sheet = pkg.Workbook.Worksheets.Add("Empty");

            sheet.ExportToTxt(_txtFilePath);

            Assert.IsFalse(File.Exists(_txtFilePath));
        }

        [TestMethod]
        public void ExportToTxt_StandardSheet_ExportsCorrectDelimiterSeparatedValues()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet sheet = pkg.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "A1";
            sheet.Cells[1, 2].Value = "B1";
            sheet.Cells[2, 1].Value = "A2";
            sheet.Cells[2, 2].Value = "B2";

            sheet.ExportToTxt(_txtFilePath, delimiter: ",");

            Assert.IsTrue(File.Exists(_txtFilePath));
            string[] lines = File.ReadAllLines(_txtFilePath);
            Assert.AreEqual(2, lines.Length);
            Assert.AreEqual("A1,B1", lines[0]);
            Assert.AreEqual("A2,B2", lines[1]);
        }

        [TestMethod]
        public void ExportToTxt_WithSkipColumnList_OmitsSpecifiedColumns()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet sheet = pkg.Workbook.Worksheets.Add("SkipData");
            sheet.Cells[1, 1].Value = "Keep1";
            sheet.Cells[1, 2].Value = "SkipMe";
            sheet.Cells[1, 3].Value = "Keep2";

            var skipColumns = new List<int> { 2 };
            sheet.ExportToTxt(_txtFilePath, delimiter: "\t", skipColumnList: skipColumns);

            string[] lines = File.ReadAllLines(_txtFilePath);
            Assert.AreEqual("Keep1\tKeep2", lines[0]);
        }

        [TestMethod]
        public void ExportToTxt_WithFixedValueMatrix_InsertsInjectedColumnsCorrectly()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet sheet = pkg.Workbook.Worksheets.Add("FixedValueData");
            sheet.Cells[1, 1].Value = "H1";
            sheet.Cells[2, 1].Value = "R2";

            // Matrix Definition layout: [Row index, 0] = HeaderValue, [Row index, 1] = RowValue, [Row index, 2] = InsertIndex
            string[,] fixedValues = new string[,] { { "NewHeader", "NewValue", "0" } };

            sheet.ExportToTxt(_txtFilePath, delimiter: ",", pColumnNameFixedValue: fixedValues);

            string[] lines = File.ReadAllLines(_txtFilePath);
            // First row matches pColumnNameFixedValue[x, 0]
            Assert.AreEqual("NewHeader,H1", lines[0]);
            // Sub-rows match pColumnNameFixedValue[x, 1]
            Assert.AreEqual("NewValue,R2", lines[1]);
        }

        [TestMethod]
        public void ExportToTxt_WithExtraRowsDictionary_AppendsRowsAtExplicitIndices()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet sheet = pkg.Workbook.Worksheets.Add("ExtraRowsData");
            sheet.Cells[1, 1].Value = "BaseCell";

            var extraRows = new Dictionary<int, List<string>>
                {
                    { 0, new List<string> { "ExtraA" } },
                    { 2, new List<string> { "ExtraC" } }
                };

            sheet.ExportToTxt(_txtFilePath, delimiter: ",", extraRows: extraRows);

            string[] lines = File.ReadAllLines(_txtFilePath);
            Assert.AreEqual(2, lines.Length);
            Assert.AreEqual("BaseCell", lines[0]);
            // Checks matching index placement validation array layout structure [ExtraA, "", ExtraC]
            Assert.AreEqual("ExtraA,,ExtraC", lines[1]);
        }

        [TestMethod]
        public void PrintInputFileInformation_GeneratesStyledSheetStructure()
        {
            using var pkg = new ExcelPackage();
            var infoDic = new Dictionary<string, List<string>>
                {
                    { "File_Alpha.csv", new List<string> { "Info1", "Info2" } }
                };

            pkg.Workbook.PrintInputFileInformation(infoDic);

            ExcelWorksheet sheet = pkg.Workbook.Worksheets["InputFile_Information"];
            Assert.IsNotNull(sheet);

            // 1. Verify exact Header definitions and dimensions setup profiles
            Assert.AreEqual("Input File", sheet.Cells[1, 1].Value);
            Assert.AreEqual("Information", sheet.Cells[1, 2].Value);
            Assert.AreEqual(26.29, sheet.Column(1).Width);
            Assert.AreEqual(66.86, sheet.Column(2).Width);

            // 2. Verify cells layout style definitions configurations (Color matching blue: #305496)
            ExcelColor headerColor = sheet.Cells[1, 1].Style.Fill.BackgroundColor;
            Assert.AreEqual("FF305496", headerColor.Rgb);
            Assert.AreEqual(ExcelHorizontalAlignment.CenterContinuous, sheet.Cells[1, 1].Style.HorizontalAlignment);

            // 3. Verify Dictionary content unpacking records layout values matching indices loop mapping
            Assert.AreEqual("File_Alpha.csv", sheet.Cells[2, 1].Value);
            Assert.AreEqual("Info1", sheet.Cells[2, 2].Value);
            Assert.AreEqual("Info2", sheet.Cells[3, 2].Value);

            // Data key value fill color verification matching green: #00B050
            ExcelColor dataKeyColor = sheet.Cells[2, 1].Style.Fill.BackgroundColor;
            Assert.AreEqual("FF00B050", dataKeyColor.Rgb);
        }

        [TestMethod]
        public void PrintInputFileInformation_EmptyValueList_IncrementsIndexSafely()
        {
            using var pkg = new ExcelPackage();
            var infoDic = new Dictionary<string, List<string>>
                {
                    { "EmptyFile.csv", new List<string>() },
                    { "NextFile.csv", new List<string> { "ValidInfo" } }
                };

            pkg.Workbook.PrintInputFileInformation(infoDic);

            ExcelWorksheet sheet = pkg.Workbook.Worksheets["InputFile_Information"];
            Assert.AreEqual("EmptyFile.csv", sheet.Cells[2, 1].Value);
            // Empty item safely moved tracking target index to row 3
            Assert.AreEqual("NextFile.csv", sheet.Cells[3, 1].Value);
        }

        [TestMethod]
        public void CopyWorkSheet_EmptyOrNullFileName_ReturnsImmediately()
        {
            using var pkg = new ExcelPackage();
            pkg.Workbook.CopyWorkSheet("");
            pkg.Workbook.CopyWorkSheet(null);
            Assert.AreEqual(0, pkg.Workbook.Worksheets.Count);
        }

        [TestMethod]
        public void CopyWorkSheet_CsvFile_ParsesCommasAndMovesSheetToFront()
        {
            string csvFile = Path.Combine(_tempDir, "DataLog.csv");
            File.WriteAllText(csvFile, "Header1,Header2\nVal1,Val2");

            using var pkg = new ExcelPackage();
            // Pre-seed an existing sheet to verify the MoveBefore sorting routine works
            pkg.Workbook.Worksheets.Add("ExistingSheet");

            pkg.Workbook.CopyWorkSheet(csvFile);

            Assert.AreEqual(2, pkg.Workbook.Worksheets.Count);
            // The newly added sheet from the CSV filename should be moved to position 1
            Assert.AreEqual("DataLog", pkg.Workbook.Worksheets[0].Name);
            Assert.AreEqual("Val1", pkg.Workbook.Worksheets[0].Cells["A2"].Value.ToString());
        }

        [TestMethod]
        public void CopyWorkSheet_TxtFile_ParsesTabsAndAppendsSheet()
        {
            string txtFile = Path.Combine(_tempDir, "DumpFile.txt");
            File.WriteAllText(txtFile, "ColA\tColB\r\nTextA\tTextB");

            using var pkg = new ExcelPackage();
            pkg.Workbook.CopyWorkSheet(txtFile);

            Assert.AreEqual(1, pkg.Workbook.Worksheets.Count);
            ExcelWorksheet sheet = pkg.Workbook.Worksheets["DumpFile"];
            Assert.IsNotNull(sheet);
            Assert.AreEqual("TextB", sheet.Cells["B2"].Value.ToString());
        }

        [TestMethod]
        public void CopyWorkSheet_ExcelFile_IteratesSheetsAndCallsAddSheet()
        {
            string sourceExcel = Path.Combine(_tempDir, "SourceBook.xlsx");
            using (var srcPkg = new ExcelPackage(new FileInfo(sourceExcel)))
            {
                srcPkg.Workbook.Worksheets.Add("ExtractedSheet1");
                srcPkg.Workbook.Worksheets.Add("ExtractedSheet2");
                srcPkg.Save();
            }

            using var pkg = new ExcelPackage();
            pkg.Workbook.CopyWorkSheet(sourceExcel);

            // Verifies both sheets were read and attached by the fallback routine
            Assert.AreEqual(2, pkg.Workbook.Worksheets.Count);
            Assert.IsNotNull(pkg.Workbook.Worksheets["ExtractedSheet1"]);
            Assert.IsNotNull(pkg.Workbook.Worksheets["ExtractedSheet2"]);
        }

        [TestMethod]
        public void CopyWorkSheets_ListWithMultipleFormats_ProcessesAllSequentially()
        {
            string csvFile = Path.Combine(_tempDir, "Batch1.csv");
            string txtFile = Path.Combine(_tempDir, "Batch2.txt");
            File.WriteAllText(csvFile, "A,B");
            File.WriteAllText(txtFile, "C\tD");

            var fileList = new List<string> { csvFile, txtFile };

            using var pkg = new ExcelPackage();
            pkg.Workbook.CopyWorkSheets(fileList);

            Assert.AreEqual(2, pkg.Workbook.Worksheets.Count);
            Assert.IsNotNull(pkg.Workbook.Worksheets["Batch1"]);
            Assert.IsNotNull(pkg.Workbook.Worksheets["Batch2"]);
        }

        [TestMethod]
        public void OutLineColumnByCell_ConsecutiveMatchingCells_SetsOutlineLevelsAndCollapses()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet sheet = pkg.Workbook.Worksheets.Add("OutlineData");
            sheet.Cells[1, 1].Value = "GroupA";
            sheet.Cells[2, 1].Value = "GroupA";
            sheet.Cells[3, 1].Value = "GroupA";
            sheet.Cells[4, 1].Value = "Different";

            sheet.OutLineColumnByCell(colNum: 1, rowNum: 1);

            // Verification: Rows 1 and 2 inside the matching group range should be outlined/collapsed
            Assert.AreEqual(1, sheet.Row(1).OutlineLevel);
            Assert.IsTrue(sheet.Row(1).Collapsed);
            Assert.AreEqual(1, sheet.Row(2).OutlineLevel);
            Assert.IsTrue(sheet.Row(2).Collapsed);

            // Row 3 is the boundary termination index cell, so it remains unchanged by logic
            Assert.AreEqual(0, sheet.Row(3).OutlineLevel);
            Assert.IsFalse(sheet.Row(3).Collapsed);
        }

        [TestMethod]
        public void MergeColumn_ConsecutiveMatchingCells_MergesCellsAndCentersText()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet sheet = pkg.Workbook.Worksheets.Add("MergeData");
            sheet.Cells[1, 1].Value = "MergeMe";
            sheet.Cells[2, 1].Value = "MergeMe";
            sheet.Cells[3, 1].Value = "MergeMe";

            sheet.MergeColumn(colNum: 1);

            ExcelRange cellRange = sheet.Cells[1, 1, 3, 1];
            Assert.IsTrue(cellRange.Merge);
            Assert.AreEqual(ExcelVerticalAlignment.Center, cellRange.Style.VerticalAlignment);
        }

        [TestMethod]
        public void SetPercentFormat_WithFillTrue_AppliesPercentageAndDataBarAttributes()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet sheet = pkg.Workbook.Worksheets.Add("FormatData");
            ExcelRange targetRange = sheet.Cells["A1:A5"];

            targetRange.SetPercentFormat("0.75", fill: true);

            Assert.AreEqual("0.75", targetRange.Formula);
            Assert.AreEqual("0.00%", targetRange.Style.Numberformat.Format);
        }

        [TestMethod]
        public void AddTxt_ValidTabSeparatedFile_ImportsDataGridIntoWorksheet()
        {
            string txtFile = Path.Combine(_tempDir, "MetricsReport.txt");
            File.WriteAllLines(txtFile, ["HeaderA\tHeaderB", "Val1\tVal2"]);

            using var pkg = new ExcelPackage();
            pkg.Workbook.AddTxt(txtFile);

            ExcelWorksheet sheet = pkg.Workbook.Worksheets["MetricsReport"];
            Assert.IsNotNull(sheet);

            // Extract values array block dropped into cells
            object[,] cellValues = sheet.Cells[1, 1, 3, 2].Value as object[,];
            Assert.IsNotNull(cellValues);
            Assert.AreEqual("HeaderA", cellValues[0, 0]);
            Assert.AreEqual("HeaderB", cellValues[0, 1]);
            Assert.AreEqual("Val1", cellValues[1, 0]);
            Assert.AreEqual("Val2", cellValues[1, 1]);
        }

        [TestMethod]
        public void ExportSheet_StandardWorksheetData_WritesDelimiterSeparatedTextFile()
        {
            string exportPath = Path.Combine(_tempDir, "ExportedData.txt");

            using var pkg = new ExcelPackage();
            ExcelWorksheet sheet = pkg.Workbook.Worksheets.Add("ExportData");
            sheet.Cells[1, 1].Value = "Cell1";
            sheet.Cells[1, 2].Value = "Cell2";
            sheet.Cells[2, 1].Value = "Cell3";
            sheet.Cells[2, 2].Value = "Cell4";

            sheet.ExportSheet(exportPath, delimiter: ",");

            Assert.IsTrue(File.Exists(exportPath));
            string[] fileLines = File.ReadAllLines(exportPath);
            Assert.AreEqual(2, fileLines.Length);
            Assert.AreEqual("Cell1,Cell2", fileLines[0]);
            Assert.AreEqual("Cell3,Cell4", fileLines[1]);
        }

        [TestMethod]
        public void SetHeaderStyle_ValidWorksheet_AppliesHeaderStyleToFirstRow()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet sheet = pkg.Workbook.Worksheets.Add("HeaderStyleSheet");
            sheet.Cells["A1"].Value = "ID";
            sheet.Cells["B1"].Value = "Name";
            // Data Row 2
            sheet.Cells["A2"].Value = "101";

            sheet.SetHeaderStyle();

            // Asserts that row 1 targets the named style property boundary
            Assert.AreEqual("Header", sheet.Cells["A1"].StyleName);
            Assert.AreEqual("Header", sheet.Cells["B1"].StyleName);

            // Row 2 data row should not have its StyleName altered
            Assert.AreNotEqual("Header", sheet.Cells["A2"].StyleName);
        }

        [TestMethod]
        public void PrintExcelCol_ValidArrayInput_LoadsCollectionIntoTargetCells()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet sheet = pkg.Workbook.Worksheets.Add("ColumnPrintSheet");
            string[] data = ["Alpha", "Bravo", "Charlie"];
            ExcelRange range = sheet.Cells["B2"];

            range.PrintExcelCol(data);

            // LoadFromCollection prints sequentially downward starting at target range element position
            Assert.AreEqual("Alpha", sheet.Cells["B2"].Value.ToString());
            Assert.AreEqual("Bravo", sheet.Cells["B3"].Value.ToString());
            Assert.AreEqual("Charlie", sheet.Cells["B4"].Value.ToString());
        }
    }
}
