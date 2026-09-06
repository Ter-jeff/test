using CommonLib.Extension;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace CommonLib.Test.UT.Extension
{
    [TestClass]
    public class EpplusExtensionsExcelWorksheetsTests
    {
        // ─── ExcelWorksheets: AddSheet(name) ────────────────────────────────────

        [TestMethod]
        public void WorksheetsAddSheet_NewSheet_MovesToFront()
        {
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("First");
            pkg.Workbook.Worksheets.AddSheet("Second");
            Assert.AreEqual("Second", pkg.Workbook.Worksheets[0].Name);
        }

        [TestMethod]
        public void WorksheetsAddSheet_ExistingSheet_ClearsContent()
        {
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("Anchor");
            pkg.Workbook.Worksheets.Add("Sheet1");
            pkg.Workbook.Worksheets["Sheet1"].Cells[1, 1].Value = "old";

            pkg.Workbook.Worksheets.AddSheet("Sheet1");

            Assert.IsNull(pkg.Workbook.Worksheets["Sheet1"].Cells[1, 1].Value);
        }

        [TestMethod]
        public void WorksheetsAddSheet_EmptyWorksheets_AddsSheet()
        {
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.AddSheet("NewSheet");
            Assert.IsTrue(pkg.Workbook.IsExist("NewSheet"));
        }

        [TestMethod]
        public void WorksheetsAddSheet_PreservesOtherSheets()
        {
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("First");
            pkg.Workbook.Worksheets.Add("Second");
            pkg.Workbook.Worksheets.AddSheet("Third");

            Assert.IsTrue(pkg.Workbook.IsExist("First"));
            Assert.IsTrue(pkg.Workbook.IsExist("Second"));
            Assert.IsTrue(pkg.Workbook.IsExist("Third"));
        }

        // ─── ExcelWorksheets: AddSheet(name, insertFlag) ─────────────────────
        [TestMethod]
        public void WorksheetsAddSheet_WithInsertFlagTrue_MovesToFront()
        {
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("First");
            pkg.Workbook.Worksheets.AddSheet("New", insertFlag: true);
            Assert.AreEqual("New", pkg.Workbook.Worksheets[0].Name);
        }

        [TestMethod]
        public void WorksheetsAddSheet_WithInsertFlagFalse_DoesNotMoveToFront()
        {
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("First");
            pkg.Workbook.Worksheets.Add("Second");
            pkg.Workbook.Worksheets.AddSheet("New", insertFlag: false);
            Assert.AreNotEqual("New", pkg.Workbook.Worksheets[1].Name);
        }

        [TestMethod]
        public void WorksheetsAddSheet_InsertFlagFalse_AppendsToEnd()
        {
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("First");
            pkg.Workbook.Worksheets.Add("Second");
            int countBefore = pkg.Workbook.Worksheets.Count;
            pkg.Workbook.Worksheets.AddSheet("New", insertFlag: false);

            Assert.AreEqual(countBefore + 1, pkg.Workbook.Worksheets.Count);
            Assert.AreEqual("New", pkg.Workbook.Worksheets[countBefore].Name);
        }

        [TestMethod]
        public void WorksheetsAddSheet_ExistingWithInsertFlagTrue_ClearsAndMoves()
        {
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("First");
            pkg.Workbook.Worksheets.Add("Anchor");
            pkg.Workbook.Worksheets.Add("Target");
            pkg.Workbook.Worksheets["Target"].Cells[1, 1].Value = "old";

            pkg.Workbook.Worksheets.AddSheet("Target", insertFlag: true);

            Assert.AreEqual("Target", pkg.Workbook.Worksheets[2].Name);
            Assert.IsNull(pkg.Workbook.Worksheets["Target"].Cells[1, 1].Value);
        }

        // ─── ExcelWorksheets: InsertSheet ─────────────────────────────────────
        [TestMethod]
        public void InsertSheet_NewSheet_MovesToFront()
        {
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("First");
            pkg.Workbook.Worksheets.InsertSheet("Inserted");
            Assert.AreEqual("Inserted", pkg.Workbook.Worksheets[0].Name);
        }

        [TestMethod]
        public void InsertSheet_EmptyWorksheets_AddsNewSheet()
        {
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.InsertSheet("NewSheet");
            Assert.IsTrue(pkg.Workbook.IsExist("NewSheet"));
        }

        [TestMethod]
        public void InsertSheet_MultipleInsertions_OrderCorrect()
        {
            using var pkg = new ExcelPackage();
            pkg.Workbook.Worksheets.Add("Original");
            pkg.Workbook.Worksheets.InsertSheet("First");
            pkg.Workbook.Worksheets.InsertSheet("Second");

            Assert.AreEqual("Second", pkg.Workbook.Worksheets[0].Name);
            Assert.AreEqual("First", pkg.Workbook.Worksheets[1].Name);
            Assert.AreEqual("Original", pkg.Workbook.Worksheets[2].Name);
        }

        // ─── ExcelWorksheets: AddSheet(ExcelWorksheet) ───────────────────────
        [TestMethod]
        public void WorksheetsAddSheet_ExcelWorksheet_AddsSheet()
        {
            using var pkg1 = new ExcelPackage();
            using var pkg2 = new ExcelPackage();
            ExcelWorksheet sourceSheet = pkg1.Workbook.Worksheets.Add("Source");
            sourceSheet.Cells[1, 1].Value = "TestData";

            ExcelWorksheet targetSheet = pkg2.Workbook.Worksheets.Add("Dummy");
            pkg2.Workbook.Worksheets.AddSheet(sourceSheet);

            Assert.IsTrue(pkg2.Workbook.IsExist("Source"));
        }

        [TestMethod]
        public void WorksheetsAddSheet_ExcelWorksheet_CopiesContent()
        {
            using var pkg1 = new ExcelPackage();
            using var pkg2 = new ExcelPackage();
            ExcelWorksheet sourceSheet = pkg1.Workbook.Worksheets.Add("Source");
            sourceSheet.Cells[1, 1].Value = "Data";

            ExcelWorksheet targetSheet = pkg2.Workbook.Worksheets.Add("Dummy");
            pkg2.Workbook.Worksheets.AddSheet(sourceSheet);

            // Content should be copied to the new sheet
            ExcelWorksheet added = pkg2.Workbook.Worksheets["Source"];
            Assert.IsNotNull(added.Cells[1, 1].Value);
        }

        [TestMethod]
        public void WorksheetsAddSheet_ExcelWorksheet_ClearsExistingContent()
        {
            using var pkg1 = new ExcelPackage();
            using var pkg2 = new ExcelPackage();
            ExcelWorksheet sourceSheet = pkg1.Workbook.Worksheets.Add("Sheet");
            sourceSheet.Cells[1, 1].Value = "New";

            ExcelWorksheet existing = pkg2.Workbook.Worksheets.Add("Sheet");
            existing.Cells[1, 1].Value = "Old";
            existing.Cells[2, 1].Value = "ToDelete";

            pkg2.Workbook.Worksheets.AddSheet(sourceSheet);

            ExcelWorksheet result = pkg2.Workbook.Worksheets["Sheet"];
            Assert.AreEqual("New", result.Cells[1, 1].Value);
        }
    }
}
