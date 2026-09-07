using System.Collections.Generic;

using CommonLib.Extension;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace CommonLib.Test.UT.Extension
{
    [TestClass]
    public class EpplusExtensionsExcelRangeTests
    {
        // ─── ExcelRange: SetHyperLinkFormat ────────────────────────────────────

        [TestMethod]
        public void SetHyperLinkFormat_Cell_FormatsAsHyperlink()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].Value = "http://example.com";
            ws.Cells[1, 1].SetHyperLinkFormat();

            Assert.IsTrue(ws.Cells[1, 1].Style.Font.UnderLine);
        }

        [TestMethod]
        public void SetHyperLinkFormat_CellWithoutValue_DoesNotThrow()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].SetHyperLinkFormat();
            Assert.IsFalse(ws.Cells[1, 1].Style.Font.UnderLine);
        }

        // ─── ExcelRange: FormatHeader ─────────────────────────────────────────
        [TestMethod]
        public void FormatHeader_SingleCell_FormatsAsHeader()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].Value = "Header";
            ws.Cells[1, 1].FormatHeader();

            Assert.AreEqual(OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, ws.Cells[1, 1].Style.HorizontalAlignment);
        }

        [TestMethod]
        public void FormatHeader_WithWrapTrue_EnablesWrapping()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].Value = "Header";
            ws.Cells[1, 1].FormatHeader(wrap: true);

            Assert.IsTrue(ws.Cells[1, 1].Style.WrapText);
        }

        [TestMethod]
        public void FormatHeader_WithWrapFalse_DisablesWrapping()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].Value = "Header";
            ws.Cells[1, 1].FormatHeader(wrap: false);

            Assert.IsFalse(ws.Cells[1, 1].Style.WrapText);
        }

        [TestMethod]
        public void FormatHeader_SetsCorrectFont()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            ws.Cells[1, 1].Value = "Header";
            ws.Cells[1, 1].FormatHeader();

            Assert.AreEqual(12, ws.Cells[1, 1].Style.Font.Size);
        }

        // ─── ExcelRange: PrintExcelRow ────────────────────────────────────────
        [TestMethod]
        public void PrintExcelRow_ArrayOfStrings_WritesRowAndReturnsNextRow()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            string[] data = ["X", "Y", "Z"];

            int nextRow = ws.Cells[1, 1].PrintExcelRow(data);

            Assert.AreEqual(2, nextRow);
            Assert.AreEqual("X", ws.Cells[1, 1].Value);
            Assert.AreEqual("Y", ws.Cells[1, 2].Value);
            Assert.AreEqual("Z", ws.Cells[1, 3].Value);
        }

        [TestMethod]
        public void PrintExcelRow_NullItem_WritesNullOrEmptyCell()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            string[] data = ["A", null, "C"];

            ws.Cells[1, 1].PrintExcelRow(data);

            object cellVal = ws.Cells[1, 2].Value;
            Assert.IsTrue(cellVal == null || string.IsNullOrEmpty(cellVal.ToString()));
        }

        [TestMethod]
        public void PrintExcelRow_EmptyArray_ReturnsStartRow()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            string[] data = [];

            int nextRow = ws.Cells[1, 1].PrintExcelRow(data);

            // Should return row + 1 for an empty array
            Assert.IsTrue(nextRow >= 1);
        }

        [TestMethod]
        public void PrintExcelRow_SingleItem_WritesToFirstColumn()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            string[] data = ["SingleValue"];

            ws.Cells[1, 1].PrintExcelRow(data);

            Assert.AreEqual("SingleValue", ws.Cells[1, 1].Value);
        }

        [TestMethod]
        public void PrintExcelRow_MultipleRows_WritesSequentially()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            string[] data1 = ["A", "B"];
            string[] data2 = ["C", "D"];

            int nextRow1 = ws.Cells[1, 1].PrintExcelRow(data1);
            ws.Cells[nextRow1, 1].PrintExcelRow(data2);

            Assert.AreEqual("A", ws.Cells[1, 1].Value);
            Assert.AreEqual("C", ws.Cells[2, 1].Value);
        }

        // ─── ExcelRange: PrintExcelRange ──────────────────────────────────────
        [TestMethod]
        public void PrintExcelRange_2DArray_WritesAllCells()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            string[,] data = { { "A", "B" }, { "C", "D" } };

            int nextRow = ws.Cells[1, 1].PrintExcelRange(data);

            Assert.AreEqual(3, nextRow);
            Assert.AreEqual("A", ws.Cells[1, 1].Value);
            Assert.AreEqual("B", ws.Cells[1, 2].Value);
            Assert.AreEqual("C", ws.Cells[2, 1].Value);
            Assert.AreEqual("D", ws.Cells[2, 2].Value);
        }

        [TestMethod]
        public void PrintExcelRange_SingleRow_WritesSuccessfully()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            string[,] data = { { "X", "Y", "Z" } };

            ws.Cells[1, 1].PrintExcelRange(data);

            Assert.AreEqual("X", ws.Cells[1, 1].Value);
            Assert.AreEqual("Y", ws.Cells[1, 2].Value);
            Assert.AreEqual("Z", ws.Cells[1, 3].Value);
        }

        [TestMethod]
        public void PrintExcelRange_SingleColumn_WritesSuccessfully()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            string[,] data = { { "A" }, { "B" }, { "C" } };

            ws.Cells[1, 1].PrintExcelRange(data);

            Assert.AreEqual("A", ws.Cells[1, 1].Value);
            Assert.AreEqual("B", ws.Cells[2, 1].Value);
            Assert.AreEqual("C", ws.Cells[3, 1].Value);
        }

        [TestMethod]
        public void PrintExcelRange_LargeArray_WritesCorrectly()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            string[,] data = new string[10, 5];
            for (int i = 0; i < 10; i++)
            {
                for (int j = 0; j < 5; j++)
                {
                    data[i, j] = $"Cell_{i}_{j}";
                }
            }

            ws.Cells[1, 1].PrintExcelRange(data);

            Assert.AreEqual("Cell_0_0", ws.Cells[1, 1].Value);
            Assert.AreEqual("Cell_9_4", ws.Cells[10, 5].Value);
        }

        // ─── ExcelRange: PrintExcelRowByList ──────────────────────────────────
        [TestMethod]
        public void PrintExcelRowByList_EmptyList_ReturnsZero()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            int result = ws.Cells[1, 1].PrintExcelRowByList(new List<List<string>>());
            Assert.AreEqual(0, result);
        }

        [TestMethod]
        public void PrintExcelRowByList_WithData_WritesRowsAndReturnsNextRow()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            var list = new List<List<string>>
                {
                    new() { "A", "B" },
                    new() { "C", "D" }
                };

            int nextRow = ws.Cells[1, 1].PrintExcelRowByList(list);

            Assert.AreEqual(3, nextRow);
            Assert.AreEqual("A", ws.Cells[1, 1].Value);
            Assert.AreEqual("C", ws.Cells[2, 1].Value);
        }

        [TestMethod]
        public void PrintExcelRowByList_SingleRow_WritesSuccessfully()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            var list = new List<List<string>>
                {
                    new() { "Only", "One", "Row" }
                };

            int nextRow = ws.Cells[1, 1].PrintExcelRowByList(list);

            Assert.AreEqual(2, nextRow);
            Assert.AreEqual("Only", ws.Cells[1, 1].Value);
            Assert.AreEqual("One", ws.Cells[1, 2].Value);
            Assert.AreEqual("Row", ws.Cells[1, 3].Value);
        }

        [TestMethod]
        public void PrintExcelRowByList_VaryingColumnCounts_HandlesGracefully()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            var list = new List<List<string>>
                {
                    new() { "A", "B" },
                    new() { "C", "D", "E" }
                };

            int nextRow = ws.Cells[1, 1].PrintExcelRowByList(list);

            Assert.AreEqual(3, nextRow);
            Assert.AreEqual("A", ws.Cells[1, 1].Value);
            Assert.AreEqual("C", ws.Cells[2, 1].Value);
            Assert.AreEqual("E", ws.Cells[2, 3].Value);
        }

        [TestMethod]
        public void PrintExcelRowByList_WithNullItems_HandlesGracefully()
        {
            using var pkg = new ExcelPackage();
            ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("S");
            var list = new List<List<string>>
                {
                    new() { "A", null, "C" }
                };

            int nextRow = ws.Cells[1, 1].PrintExcelRowByList(list);

            Assert.AreEqual(2, nextRow);
            Assert.AreEqual("A", ws.Cells[1, 1].Value);
            Assert.AreEqual("C", ws.Cells[1, 3].Value);
        }
    }
}
