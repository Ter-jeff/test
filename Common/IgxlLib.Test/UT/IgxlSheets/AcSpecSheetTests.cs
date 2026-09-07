using System.Collections.Generic;

using IgxlLib.Const;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

namespace IgxlLib.Test.UT.IgxlSheets
{
    [TestClass]
    public class AcSpecSheetTests
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void AcSpecSheet_Constructor_WithSheetName()
        {
            // Arrange
            string sheetName = "AcSpec";

            // Act
            var acSpecSheet = new AcSpecSheet(sheetName);

            // Assert
            Assert.IsNotNull(acSpecSheet);
            Assert.AreEqual(sheetName, acSpecSheet.Name);
            Assert.AreEqual("DTACSpecSheet", acSpecSheet.SheetType);
            Assert.AreEqual(IgxlSheetNames.AcSpec, acSpecSheet.IgxlSheetName);
            Assert.AreEqual(0, acSpecSheet.Rows.Count);
        }

        [TestMethod]
        public void AcSpecSheet_Constructor_WithCategoryAndSelectorLists()
        {
            // Arrange
            string sheetName = "AcSpec";
            var categoryList = new List<string> { "Category1", "Category2" };
            var selectorList = new List<string> { "Selector1", "Selector2" };

            // Act
            var acSpecSheet = new AcSpecSheet(sheetName, categoryList, selectorList);

            // Assert
            Assert.IsNotNull(acSpecSheet);
            Assert.AreEqual(sheetName, acSpecSheet.Name);
            Assert.AreEqual("DTACSpecSheet", acSpecSheet.SheetType);
            Assert.AreEqual(IgxlSheetNames.AcSpec, acSpecSheet.IgxlSheetName);
            Assert.AreEqual(2, acSpecSheet.CategoryList.Count);
            Assert.AreEqual(2, acSpecSheet.SelectorNameList.Count);
        }

        [TestMethod]
        public void AcSpecSheet_AddRow()
        {
            // Arrange
            var acSpecSheet = new AcSpecSheet("AcSpec");
            var acSpec = new AcSpec("TestSpec", null, "1.5", "Test comment");

            // Act
            acSpecSheet.AddRow(acSpec);

            // Assert
            Assert.AreEqual(1, acSpecSheet.Rows.Count);
            Assert.AreEqual("TestSpec", acSpecSheet.Rows[0].Symbol);
            Assert.AreEqual("1.5", acSpecSheet.Rows[0].Value);
        }

        [TestMethod]
        public void AcSpecSheet_AddRows()
        {
            // Arrange
            var acSpecSheet = new AcSpecSheet("AcSpec");
            var rows = new List<AcSpec>
            {
                new("Spec1", null, "1.0", "Comment 1"),
                new("Spec2", null, "2.0", "Comment 2"),
                new("Spec3", null, "3.0", "Comment 3")
            };

            // Act
            acSpecSheet.AddRows(rows);

            // Assert
            Assert.AreEqual(3, acSpecSheet.Rows.Count);
        }

        [TestMethod]
        public void AcSpecSheet_RemoveRow()
        {
            // Arrange
            var acSpecSheet = new AcSpecSheet("AcSpec");
            var row1 = new AcSpec("Spec1", null, "1.0", "Comment 1");
            var row2 = new AcSpec("Spec2", null, "2.0", "Comment 2");
            acSpecSheet.AddRow(row1);
            acSpecSheet.AddRow(row2);

            // Act
            acSpecSheet.RemoveRow(row1);

            // Assert
            Assert.AreEqual(1, acSpecSheet.Rows.Count);
            Assert.AreEqual("Spec2", acSpecSheet.Rows[0].Symbol);
        }

        [TestMethod]
        public void AcSpecSheet_InsertRow()
        {
            // Arrange
            var acSpecSheet = new AcSpecSheet("AcSpec");
            var row1 = new AcSpec("Spec1", null, "1.0", "Comment 1");
            var row3 = new AcSpec("Spec3", null, "3.0", "Comment 3");
            var rowToInsert = new AcSpec("Spec2", null, "2.0", "Comment 2");
            acSpecSheet.AddRow(row1);
            acSpecSheet.AddRow(row3);

            // Act
            int index = acSpecSheet.InsertRow(1, rowToInsert);

            // Assert
            Assert.AreEqual(2, index);
            Assert.AreEqual(3, acSpecSheet.Rows.Count);
            Assert.AreEqual("Spec2", acSpecSheet.Rows[1].Symbol);
            Assert.AreEqual("Spec3", acSpecSheet.Rows[2].Symbol);
        }

        [TestMethod]
        public void AcSpecSheet_SheetType_IsCorrect()
        {
            // Arrange & Act
            var acSpecSheet = new AcSpecSheet("AcSpec");

            // Assert
            Assert.AreEqual("DTACSpecSheet", acSpecSheet.SheetType);
        }

        [TestMethod]
        public void AcSpecSheet_IgxlSheetName_IsCorrect()
        {
            // Arrange & Act
            var acSpecSheet = new AcSpecSheet("AcSpec");

            // Assert
            Assert.AreEqual(IgxlSheetNames.AcSpec, acSpecSheet.IgxlSheetName);
        }

        [TestMethod]
        public void AcSpecSheet_Errors_InitializedEmpty()
        {
            // Arrange & Act
            var acSpecSheet = new AcSpecSheet("AcSpec");

            // Assert
            Assert.IsNotNull(acSpecSheet.GetErrors());
            Assert.AreEqual(0, acSpecSheet.GetErrors().Count);
        }

        [TestMethod]
        public void AcSpecSheet_Name_CanBeSet()
        {
            // Arrange
            var acSpecSheet = new AcSpecSheet("AcSpec")
            {
                // Act
                Name = "NewAcSpecName"
            };

            // Assert
            Assert.AreEqual("NewAcSpecName", acSpecSheet.Name);
        }

        [TestMethod]
        public void AcSpecSheet_AcSpec_WithSelectors()
        {
            // Arrange
            var acSpecSheet = new AcSpecSheet("AcSpec");
            var selectors = new List<Selector>
            {
                new("Selector1", "Value1"),
                new("Selector2", "Value2")
            };
            var acSpec = new AcSpec("TestSpec", selectors, "1.5", "Test comment");

            // Act
            acSpecSheet.AddRow(acSpec);

            // Assert
            Assert.AreEqual(1, acSpecSheet.Rows.Count);
            Assert.AreEqual(2, acSpecSheet.Rows[0].SelectorList.Count);
        }

        [TestMethod]
        public void AcSpecSheet_AcSpec_AddCategory()
        {
            // Arrange
            var acSpec = new AcSpec("TestSpec", null, "1.5", "Test comment");
            var category1 = new CategoryInSpec("Category1", "DC", "0", "10");
            var category2 = new CategoryInSpec("Category2", "AC", "5", "15");

            // Act
            acSpec.AddCategory(category1);
            acSpec.AddCategory(category2);

            // Assert
            Assert.AreEqual(2, acSpec.CategoryList.Count);
            Assert.IsTrue(acSpec.ContainsCategory("Category1"));
            Assert.IsTrue(acSpec.ContainsCategory("Category2"));
        }

        [TestMethod]
        public void AcSpecSheet_AcSpec_GetCategory()
        {
            // Arrange
            var acSpec = new AcSpec("TestSpec", null, "1.5", "Test comment");
            var category = new CategoryInSpec("TestCategory", "DC", "0", "10");
            acSpec.AddCategory(category);

            // Act
            CategoryInSpec retrievedCategory = acSpec.GetCategoryItem("TestCategory");

            // Assert
            Assert.IsNotNull(retrievedCategory);
            Assert.AreEqual("TestCategory", retrievedCategory.Name);
            Assert.AreEqual("DC", retrievedCategory.Typ);
        }

        [TestMethod]
        public void AcSpecSheet_ClearRows()
        {
            // Arrange
            var acSpecSheet = new AcSpecSheet("AcSpec");
            acSpecSheet.AddRow(new AcSpec("Spec1", null, "1.0", "Comment 1"));
            acSpecSheet.AddRow(new AcSpec("Spec2", null, "2.0", "Comment 2"));
            Assert.AreEqual(2, acSpecSheet.Rows.Count);

            // Act
            acSpecSheet.Rows.Clear();

            // Assert
            Assert.AreEqual(0, acSpecSheet.Rows.Count);
        }
    }
}
