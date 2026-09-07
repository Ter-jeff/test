using System.Collections.Generic;
using System.Data;

using CommonLib.Extension;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace CommonLib.Test.UT.Extension
{
    [TestClass]
    public class EpplusExtensionsExcelWorksheetTests
    {
        // ─── ExcelWorksheet: GetCellValue ─────────────────────────────────────

        [TestMethod]
        public void GetCellValue_StringCell_ReturnsString()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].Value = "Hello";
            Assert.AreEqual("Hello", ws.GetCellValue(1, 1));
        }

        [TestMethod]
        public void GetCellValue_DoubleCell_ReturnsString()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].Value = 3.14;
            Assert.AreEqual("3.14", ws.GetCellValue(1, 1));
        }

        [TestMethod]
        public void GetCellValue_BoolCell_ReturnsString()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].Value = true;
            Assert.AreEqual("True", ws.GetCellValue(1, 1));
        }

        [TestMethod]
        public void GetCellValue_EmptyCell_ReturnsEmpty()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].Value = "seed";
            Assert.AreEqual("", ws.GetCellValue(1, 5));
        }

        [TestMethod]
        public void GetCellValue_InvalidRowOrCol_ReturnsEmpty()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            Assert.AreEqual("", ws.GetCellValue(0, 1));
            Assert.AreEqual("", ws.GetCellValue(1, 0));
            Assert.AreEqual("", ws.GetCellValue(-1, 1));
        }

        [TestMethod]
        public void GetCellValue_IntegerCell_ReturnsString()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].Value = 42;
            Assert.AreEqual("42", ws.GetCellValue(1, 1));
        }

        // ─── ExcelWorksheet: GetCellLine ──────────────────────────────────────
        [TestMethod]
        public void GetCellLine_SingleRow_ReturnsTabSeparated()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].Value = "A";
            ws.Cells[1, 2].Value = "B";
            ws.Cells[1, 3].Value = "C";
            Assert.AreEqual("A\tB\tC", ws.GetCellLine(1));
        }

        [TestMethod]
        public void GetCellLine_EmptySheet_ReturnsEmpty()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            Assert.AreEqual("", ws.GetCellLine(1));
        }

        [TestMethod]
        public void GetCellLine_SecondRow_ReturnsSecondRowValues()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].Value = "Header";
            ws.Cells[2, 1].Value = "Data1";
            ws.Cells[2, 2].Value = "Data2";
            Assert.AreEqual("Data1\tData2", ws.GetCellLine(2));
        }

        [TestMethod]
        public void GetCellLine_WithNumbers_FormatsAsStrings()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].Value = 100;
            ws.Cells[1, 2].Value = 200;
            Assert.AreEqual("100\t200", ws.GetCellLine(1));
        }

        // ─── ExcelWorksheet: GetCellValueAndAddress ───────────────────────────
        [TestMethod]
        public void GetCellValueAndAddress_ValidCell_ReturnsValueAndAddress()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[2, 3].Value = "TestValue";
            string value = ws.GetCellValueAndAddress(2, 3, out string address);
            Assert.AreEqual("TestValue", value);
            Assert.AreEqual("C2", address);
        }

        [TestMethod]
        public void GetCellValueAndAddress_InvalidRow_ReturnsEmpty()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            Assert.AreEqual("", ws.GetCellValueAndAddress(0, 1, out _));
        }

        [TestMethod]
        public void GetCellValueAndAddress_InvalidCol_ReturnsEmpty()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            Assert.AreEqual("", ws.GetCellValueAndAddress(1, 0, out _));
        }

        [TestMethod]
        public void GetCellValueAndAddress_EmptyCell_ReturnsEmpty()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].Value = "data";
            string value = ws.GetCellValueAndAddress(2, 2, out string address);
            Assert.AreEqual("", value);
        }

        [TestMethod]
        public void GetCellValueAndAddress_HighRowCol_ReturnsCorrectAddress()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[10, 27].Value = "Data";
            string value = ws.GetCellValueAndAddress(10, 27, out string address);
            Assert.AreEqual("Data", value);
            Assert.AreEqual("AA10", address);
        }

        // ─── ExcelWorksheet: FindCellByValue ──────────────────────────────────
        [TestMethod]
        public void FindCellByValue_ValuePresent_SetsRowAndCol()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].Value = "Header";
            ws.Cells[2, 3].Value = "Target";
            int row = 0, col = 0;
            ws.FindCellByValue(ref row, ref col, "Target");
            Assert.AreEqual(2, row);
            Assert.AreEqual(3, col);
        }

        [TestMethod]
        public void FindCellByValue_CaseInsensitive_FindsCell()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].Value = "HELLO";
            int row = 0, col = 0;
            ws.FindCellByValue(ref row, ref col, "hello");
            Assert.AreEqual(1, row);
            Assert.AreEqual(1, col);
        }

        [TestMethod]
        public void FindCellByValue_ValueAbsent_KeepsOriginalValues()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].Value = "something";
            int row = 99, col = 99;
            ws.FindCellByValue(ref row, ref col, "nothere");
            Assert.AreEqual(99, row);
            Assert.AreEqual(99, col);
        }

        [TestMethod]
        public void FindCellByValue_EmptySheet_KeepsOriginalValues()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            int row = 5, col = 5;
            ws.FindCellByValue(ref row, ref col, "anything");
            Assert.AreEqual(5, row);
            Assert.AreEqual(5, col);
        }

        // ─── ExcelWorksheet: ReadSheetToDataSet ───────────────────────────────
        [TestMethod]
        public void ReadSheetToDataSet_WithData_ReturnsCorrectDataTable()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].Value = "ColA";
            ws.Cells[1, 2].Value = "ColB";
            ws.Cells[2, 1].Value = "Val1";
            ws.Cells[2, 2].Value = "Val2";
            ws.Cells[3, 1].Value = "Val3";
            ws.Cells[3, 2].Value = "Val4";

            DataTable dt = ws.ReadSheetToDataSet();

            Assert.IsNotNull(dt);
            Assert.AreEqual(2, dt.Columns.Count);
            Assert.AreEqual(2, dt.Rows.Count);
            Assert.AreEqual("Val1", dt.Rows[0][0].ToString());
            Assert.AreEqual("Val4", dt.Rows[1][1].ToString());
        }

        [TestMethod]
        public void ReadSheetToDataSet_HeadersAsColumnNames()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].Value = "Name";
            ws.Cells[1, 2].Value = "Age";
            ws.Cells[2, 1].Value = "Alice";
            ws.Cells[2, 2].Value = 30;

            DataTable dt = ws.ReadSheetToDataSet();

            Assert.AreEqual("Name", dt.Columns[0].ColumnName);
            Assert.AreEqual("Age", dt.Columns[1].ColumnName);
        }

        [TestMethod]
        public void ReadSheetToDataSet_EmptySheet_ReturnsNull()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            Assert.IsNull(ws.ReadSheetToDataSet());
        }

        [TestMethod]
        public void ReadSheetToDataSet_OnlyHeaders_ReturnsEmptyDataTable()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].Value = "Col1";
            ws.Cells[1, 2].Value = "Col2";

            DataTable dt = ws.ReadSheetToDataSet();

            Assert.IsNotNull(dt);
            Assert.AreEqual(2, dt.Columns.Count);
            Assert.AreEqual(0, dt.Rows.Count);
        }

        // ─── ExcelWorksheet: ConvertToLines ───────────────────────────────────
        [TestMethod]
        public void ConvertToLines_WithData_ReturnsTabSeparatedLines()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].Value = "A";
            ws.Cells[1, 2].Value = "B";
            ws.Cells[2, 1].Value = "C";
            ws.Cells[2, 2].Value = "D";

            List<string> lines = ws.ConvertToLines();

            Assert.AreEqual(2, lines.Count);
            Assert.AreEqual("A\tB", lines[0]);
            Assert.AreEqual("C\tD", lines[1]);
        }

        [TestMethod]
        public void ConvertToLines_SingleRow_ReturnsOneEntry()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].Value = "Only";

            List<string> lines = ws.ConvertToLines();

            Assert.AreEqual(1, lines.Count);
            Assert.AreEqual("Only", lines[0]);
        }

        [TestMethod]
        public void ConvertToLines_EmptySheet_ReturnsEmptyList()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            List<string> lines = ws.ConvertToLines();
            Assert.AreEqual(0, lines.Count);
        }

        // ─── ExcelWorksheet: SplitByEmptyRow ──────────────────────────────────
        [TestMethod]
        public void SplitByEmptyRow_WithEmptyRow_ContainsEmptyRowIndex()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].Value = "Data";
            ws.Cells[3, 1].Value = "More";

            List<int> result = ws.SplitByEmptyRow();

            CollectionAssert.Contains(result, 2);
        }

        [TestMethod]
        public void SplitByEmptyRow_NoEmptyRows_ReturnsBoundaryRow()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].Value = "A";
            ws.Cells[2, 1].Value = "B";
            ws.Cells[3, 1].Value = "C";

            List<int> result = ws.SplitByEmptyRow();

            Assert.AreEqual(1, result.Count);
        }

        [TestMethod]
        public void SplitByEmptyRow_EmptySheet_ReturnsEmpty()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            List<int> result = ws.SplitByEmptyRow();
            Assert.AreEqual(0, result.Count);
        }

        // ─── Static utilities: GetHeaderOrder ────────────────────────────────
        [TestMethod]
        public void GetHeaderOrder_WithHeaders_ReturnsDictionary()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].Value = "Name";
            ws.Cells[1, 2].Value = "Age";
            ws.Cells[1, 3].Value = "City";

            Dictionary<string, int> headers = EpplusExtensions.GetHeaderOrder(ws);

            Assert.AreEqual(1, headers["Name"]);
            Assert.AreEqual(2, headers["Age"]);
            Assert.AreEqual(3, headers["City"]);
        }

        [TestMethod]
        public void GetHeaderOrder_CaseInsensitive_FindsHeaders()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].Value = "Name";

            Dictionary<string, int> headers = EpplusExtensions.GetHeaderOrder(ws);

            Assert.IsTrue(headers.ContainsKey("name"));
            Assert.IsTrue(headers.ContainsKey("NAME"));
        }

        [TestMethod]
        public void GetHeaderOrder_EmptySheet_ReturnsEmptyDictionary()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            Dictionary<string, int> headers = EpplusExtensions.GetHeaderOrder(ws);
            Assert.AreEqual(0, headers.Count);
        }

        [TestMethod]
        public void GetHeaderOrder_StartRowParam_ReadsFromSpecifiedRow()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].Value = "Ignored";
            ws.Cells[2, 1].Value = "RealHeader";
            ws.Cells[2, 2].Value = "Second";

            Dictionary<string, int> headers = EpplusExtensions.GetHeaderOrder(ws, startRow: 2);

            Assert.IsTrue(headers.ContainsKey("RealHeader"));
            Assert.IsTrue(headers.ContainsKey("Second"));
            Assert.IsFalse(headers.ContainsKey("Ignored"));
        }

        // ─── Static utilities: GetMergedCellValue ────────────────────────────
        [TestMethod]
        public void GetStaticMergedCellValue_NonMergedStringCell_ReturnsValue()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].Value = "Hello";
            string result = EpplusExtensions.GetMergedCellValue(ws, 1, 1);
            Assert.AreEqual("Hello", result);
        }

        [TestMethod]
        public void GetStaticMergedCellValue_EmptyCell_ReturnsEmpty()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].Value = "seed";
            string result = EpplusExtensions.GetMergedCellValue(ws, 1, 5);
            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void GetStaticMergedCellValue_InvalidRowOrCol_ReturnsEmpty()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            Assert.AreEqual("", EpplusExtensions.GetMergedCellValue(ws, 0, 1));
            Assert.AreEqual("", EpplusExtensions.GetMergedCellValue(ws, 1, 0));
        }

        // ─── Static utilities: GetCellFormula ────────────────────────────────
        [TestMethod]
        public void GetCellFormula_NullSheet_ReturnsEmpty()
        {
            Assert.AreEqual("", EpplusExtensions.GetCellFormula(null, 1, 1));
        }

        [TestMethod]
        public void GetCellFormula_InvalidRowOrCol_ReturnsEmpty()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            Assert.AreEqual("", EpplusExtensions.GetCellFormula(ws, 0, 1));
            Assert.AreEqual("", EpplusExtensions.GetCellFormula(ws, 1, 0));
        }

        [TestMethod]
        public void GetCellFormula_CellWithValue_ReturnsValue()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].Value = "TestVal";
            Assert.AreEqual("TestVal", EpplusExtensions.GetCellFormula(ws, 1, 1));
        }

        // ─── Static utilities: GetCellText ──────────────────────────────────
        [TestMethod]
        public void GetCellText_NullSheet_ReturnsEmpty()
        {
            Assert.AreEqual("", EpplusExtensions.GetCellText(null, 1, 1));
        }

        [TestMethod]
        public void GetCellText_InvalidRowOrCol_ReturnsEmpty()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            Assert.AreEqual("", EpplusExtensions.GetCellText(ws, 0, 1));
            Assert.AreEqual("", EpplusExtensions.GetCellText(ws, 1, 0));
        }

        [TestMethod]
        public void GetCellText_CellWithStringValue_ReturnsValue()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].Value = "Hello";
            Assert.AreEqual("Hello", EpplusExtensions.GetCellText(ws, 1, 1));
        }

        [TestMethod]
        public void GetCellText_EmptyCell_ReturnsEmpty()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            string result = EpplusExtensions.GetCellText(ws, 1, 1);
            Assert.AreEqual("", result);
        }
    }
}
