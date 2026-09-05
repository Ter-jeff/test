using System.Collections.Generic;

using IgxlLib.Const;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

namespace IgxlLib.Test.UT.IgxlSheets
{
    [TestClass]
    public class DcSpecSheetTests
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void DcSpecSheet_Constructor_WithSheetName()
        {
            // Arrange
            string sheetName = "DcSpec";

            // Act
            var dcSpecSheet = new DcSpecSheet(sheetName);

            // Assert
            Assert.IsNotNull(dcSpecSheet);
            Assert.AreEqual(sheetName, dcSpecSheet.Name);
            Assert.AreEqual("DTDCSpecSheet", dcSpecSheet.SheetType);
            Assert.AreEqual(IgxlSheetNames.DcSpec, dcSpecSheet.IgxlSheetName);
            Assert.AreEqual(0, dcSpecSheet.Rows.Count);
        }

        [TestMethod]
        public void DcSpecSheet_Constructor_WithCategoryAndSelectorLists()
        {
            // Arrange
            string sheetName = "DcSpec";
            var categoryList = new List<string> { "Category1", "Category2" };
            var selectorList = new List<string> { "Selector1", "Selector2" };

            // Act
            var dcSpecSheet = new DcSpecSheet(sheetName, categoryList, selectorList);

            // Assert
            Assert.IsNotNull(dcSpecSheet);
            Assert.AreEqual(sheetName, dcSpecSheet.Name);
            Assert.AreEqual("DTDCSpecSheet", dcSpecSheet.SheetType);
            Assert.AreEqual(IgxlSheetNames.DcSpec, dcSpecSheet.IgxlSheetName);
            Assert.AreEqual(2, dcSpecSheet.CategoryList.Count);
            Assert.AreEqual(2, dcSpecSheet.SelectorNameList.Count);
        }

        [TestMethod]
        public void DcSpecSheet_AddRow()
        {
            // Arrange
            var dcSpecSheet = new DcSpecSheet("DcSpec");
            var dcSpec = new DcSpec("TestSpec", "1.5", "Test comment");

            // Act
            dcSpecSheet.AddRow(dcSpec);

            // Assert
            Assert.AreEqual(1, dcSpecSheet.Rows.Count);
            Assert.AreEqual("TestSpec", dcSpecSheet.Rows[0].Symbol);
            Assert.AreEqual("1.5", dcSpecSheet.Rows[0].Value);
        }

        [TestMethod]
        public void DcSpecSheet_AddRows()
        {
            // Arrange
            var dcSpecSheet = new DcSpecSheet("DcSpec");
            var rows = new List<DcSpec>
            {
                new("Spec1", "1.0", "Comment 1"),
                new("Spec2", "2.0", "Comment 2"),
                new("Spec3", "3.0", "Comment 3")
            };

            // Act
            dcSpecSheet.AddRows(rows);

            // Assert
            Assert.AreEqual(3, dcSpecSheet.Rows.Count);
        }

        [TestMethod]
        public void DcSpecSheet_RemoveRow()
        {
            // Arrange
            var dcSpecSheet = new DcSpecSheet("DcSpec");
            var row1 = new DcSpec("Spec1", "1.0", "Comment 1");
            var row2 = new DcSpec("Spec2", "2.0", "Comment 2");
            dcSpecSheet.AddRow(row1);
            dcSpecSheet.AddRow(row2);

            // Act
            dcSpecSheet.RemoveRow(row1);

            // Assert
            Assert.AreEqual(1, dcSpecSheet.Rows.Count);
            Assert.AreEqual("Spec2", dcSpecSheet.Rows[0].Symbol);
        }

        [TestMethod]
        public void DcSpecSheet_InsertRow()
        {
            // Arrange
            var dcSpecSheet = new DcSpecSheet("DcSpec");
            var row1 = new DcSpec("Spec1", "1.0", "Comment 1");
            var row3 = new DcSpec("Spec3", "3.0", "Comment 3");
            var rowToInsert = new DcSpec("Spec2", "2.0", "Comment 2");
            dcSpecSheet.AddRow(row1);
            dcSpecSheet.AddRow(row3);

            // Act
            int index = dcSpecSheet.InsertRow(1, rowToInsert);

            // Assert
            Assert.AreEqual(2, index);
            Assert.AreEqual(3, dcSpecSheet.Rows.Count);
            Assert.AreEqual("Spec2", dcSpecSheet.Rows[1].Symbol);
            Assert.AreEqual("Spec3", dcSpecSheet.Rows[2].Symbol);
        }

        [TestMethod]
        public void DcSpecSheet_SheetType_IsCorrect()
        {
            // Arrange & Act
            var dcSpecSheet = new DcSpecSheet("DcSpec");

            // Assert
            Assert.AreEqual("DTDCSpecSheet", dcSpecSheet.SheetType);
        }

        [TestMethod]
        public void DcSpecSheet_Errors_InitializedEmpty()
        {
            // Arrange & Act
            var dcSpecSheet = new DcSpecSheet("DcSpec");

            // Assert
            Assert.IsNotNull(dcSpecSheet.GetErrors());
            Assert.AreEqual(0, dcSpecSheet.GetErrors().Count);
        }

        [TestMethod]
        public void DcSpecSheet_Name_CanBeSet()
        {
            // Arrange
            var dcSpecSheet = new DcSpecSheet("DcSpec")
            {
                // Act
                Name = "NewDcSpecName"
            };

            // Assert
            Assert.AreEqual("NewDcSpecName", dcSpecSheet.Name);
        }

        [TestMethod]
        public void DcSpecSheet_DcSpec_WithSelectors()
        {
            // Arrange
            var dcSpecSheet = new DcSpecSheet("DcSpec");
            var selectors = new List<Selector>
            {
                new("Selector1", "Value1"),
                new("Selector2", "Value2")
            };
            var dcSpec = new DcSpec("TestSpec", selectors, "1.5", "Test comment");

            // Act
            dcSpecSheet.AddRow(dcSpec);

            // Assert
            Assert.AreEqual(1, dcSpecSheet.Rows.Count);
            Assert.AreEqual(2, dcSpecSheet.Rows[0].SelectorList.Count);
        }

        [TestMethod]
        public void DcSpecSheet_DcSpec_CategoryCount()
        {
            // Arrange
            var dcSpecSheet = new DcSpecSheet("DcSpec");
            var dcSpec = new DcSpec("TestSpec", "1.5", "Test comment");
            dcSpec.AddCategory(new CategoryInSpec("Category1", "DC", "0", "10"));
            dcSpec.AddCategory(new CategoryInSpec("Category2", "AC", "5", "15"));

            // Act
            dcSpecSheet.AddRow(dcSpec);

            // Assert
            Assert.AreEqual(1, dcSpecSheet.Rows.Count);
            Assert.AreEqual(2, dcSpecSheet.Rows[0].CategoryCount);
        }

        [TestMethod]
        public void DcSpecSheet_ClearRows()
        {
            // Arrange
            var dcSpecSheet = new DcSpecSheet("DcSpec");
            dcSpecSheet.AddRow(new DcSpec("Spec1", "1.0", "Comment 1"));
            dcSpecSheet.AddRow(new DcSpec("Spec2", "2.0", "Comment 2"));
            Assert.AreEqual(2, dcSpecSheet.Rows.Count);

            // Act
            dcSpecSheet.Rows.Clear();

            // Assert
            Assert.AreEqual(0, dcSpecSheet.Rows.Count);
        }
    }
}
