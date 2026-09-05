using System.Collections.Generic;

using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class SpecTests
    {
        [TestMethod]
        public void AcSpec_Constructor_InitializesPropertiesCorrectly()
        {
            // Arrange
            var selectors = new List<Selector> { new("Sel1", "Val1") };
            string symbol = "TestSpec";
            string value = "100";
            string comment = "Test comment";

            // Act
            var acSpec = new AcSpec(symbol, selectors, value, comment);

            // Assert
            Assert.AreEqual(symbol, acSpec.Symbol);
            Assert.AreEqual(value, acSpec.Value);
            Assert.AreEqual(comment, acSpec.Comment);
            Assert.AreEqual(selectors, acSpec.SelectorList);
            Assert.AreEqual(0, acSpec.CategoryList.Count);
        }

        [TestMethod]
        public void AcSpec_Constructor_WithDefaultValue_SetsValueToZero()
        {
            // Arrange
            var selectors = new List<Selector>();
            string symbol = "TestSpec";

            // Act
            var acSpec = new AcSpec(symbol, selectors);

            // Assert
            Assert.AreEqual("0", acSpec.Value);
            Assert.AreEqual("", acSpec.Comment);
        }

        [TestMethod]
        public void AcSpec_AddCategory_AddsItemToList()
        {
            // Arrange
            var selectors = new List<Selector>();
            var acSpec = new AcSpec("TestSpec", selectors);
            var category = new CategoryInSpec("Category1", "TypeA", "10", "20");

            // Act
            acSpec.AddCategory(category);

            // Assert
            Assert.AreEqual(1, acSpec.CategoryList.Count);
            Assert.AreEqual(category, acSpec.CategoryList[0]);
        }

        [TestMethod]
        public void AcSpec_AddCategory_MultipleTimes_AddsAllItems()
        {
            // Arrange
            var selectors = new List<Selector>();
            var acSpec = new AcSpec("TestSpec", selectors);
            var category1 = new CategoryInSpec("Category1", "TypeA", "10", "20");
            var category2 = new CategoryInSpec("Category2", "TypeB", "30", "40");

            // Act
            acSpec.AddCategory(category1);
            acSpec.AddCategory(category2);

            // Assert
            Assert.AreEqual(2, acSpec.CategoryList.Count);
        }

        [TestMethod]
        public void AcSpec_ContainsCategory_WithExistingCategory_ReturnsTrue()
        {
            // Arrange
            var selectors = new List<Selector>();
            var acSpec = new AcSpec("TestSpec", selectors);
            var category = new CategoryInSpec("Category1", "TypeA", "10", "20");
            acSpec.AddCategory(category);

            // Act
            bool result = acSpec.ContainsCategory("Category1");

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void AcSpec_ContainsCategory_WithNonexistentCategory_ReturnsFalse()
        {
            // Arrange
            var selectors = new List<Selector>();
            var acSpec = new AcSpec("TestSpec", selectors);
            var category = new CategoryInSpec("Category1", "TypeA", "10", "20");
            acSpec.AddCategory(category);

            // Act
            bool result = acSpec.ContainsCategory("NonexistentCategory");

            // Assert
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void AcSpec_GetCategoryItem_WithExistingCategory_ReturnsCategory()
        {
            // Arrange
            var selectors = new List<Selector>();
            var acSpec = new AcSpec("TestSpec", selectors);
            var category = new CategoryInSpec("Category1", "TypeA", "10", "20");
            acSpec.AddCategory(category);

            // Act
            CategoryInSpec result = acSpec.GetCategoryItem("Category1");

            // Assert
            Assert.AreEqual(category, result);
            Assert.AreEqual("Category1", result.Name);
            Assert.AreEqual("TypeA", result.Typ);
        }

        [TestMethod]
        public void AcSpec_GetCategoryItem_WithNonexistentCategory_ReturnsNull()
        {
            // Arrange
            var selectors = new List<Selector>();
            var acSpec = new AcSpec("TestSpec", selectors);
            var category = new CategoryInSpec("Category1", "TypeA", "10", "20");
            acSpec.AddCategory(category);

            // Act
            CategoryInSpec result = acSpec.GetCategoryItem("NonexistentCategory");

            // Assert
            Assert.IsNull(result);
        }

        [TestMethod]
        public void GlobalSpec_Constructor_WithSymbolOnly_InitializesProperties()
        {
            // Arrange & Act
            var globalSpec = new GlobalSpec("GlobalSpec1");

            // Assert
            Assert.AreEqual("GlobalSpec1", globalSpec.Symbol);
            Assert.AreEqual("", globalSpec.Job);
            Assert.AreEqual("", globalSpec.Value);
            Assert.AreEqual("", globalSpec.Comment);
        }

        [TestMethod]
        public void GlobalSpec_Constructor_WithSymbolAndValue_InitializesProperties()
        {
            // Arrange & Act
            var globalSpec = new GlobalSpec("GlobalSpec1", "100");

            // Assert
            Assert.AreEqual("GlobalSpec1", globalSpec.Symbol);
            Assert.AreEqual("100", globalSpec.Value);
            Assert.AreEqual("", globalSpec.Job);
        }

        [TestMethod]
        public void GlobalSpec_Constructor_WithAllParameters_InitializesProperties()
        {
            // Arrange & Act
            var globalSpec = new GlobalSpec("GlobalSpec1", "Job1", "100", "Test comment");

            // Assert
            Assert.AreEqual("GlobalSpec1", globalSpec.Symbol);
            Assert.AreEqual("Job1", globalSpec.Job);
            Assert.AreEqual("100", globalSpec.Value);
            Assert.AreEqual("Test comment", globalSpec.Comment);
        }

        [TestMethod]
        public void GlobalSpec_SetJob_UpdatesJobProperty()
        {
            // Arrange
            var globalSpec = new GlobalSpec("GlobalSpec1")
            {
                // Act
                Job = "NewJob"
            };

            // Assert
            Assert.AreEqual("NewJob", globalSpec.Job);
        }

        [TestMethod]
        public void DcSpec_Constructor_InitializesWithAcSpecProperties()
        {
            // Arrange
            var selectors = new List<Selector> { new("Sel1", "Val1") };
            string symbol = "DcSpec1";
            string value = "50";
            string comment = "DC spec comment";

            // Act
            var dcSpec = new DcSpec(symbol, selectors, value, comment);

            // Assert
            Assert.AreEqual(symbol, dcSpec.Symbol);
            Assert.AreEqual(value, dcSpec.Value);
            Assert.AreEqual(comment, dcSpec.Comment);
        }

        [TestMethod]
        public void AcSpec_Inherits_FromSpec_And_IgxlRow()
        {
            // Arrange
            var selectors = new List<Selector>();
            var acSpec = new AcSpec("TestSpec", selectors);

            // Act & Assert
            Assert.IsInstanceOfType(acSpec, typeof(Spec));
            Assert.IsInstanceOfType(acSpec, typeof(IgxlRow));
        }

        [TestMethod]
        public void GlobalSpec_Inherits_FromSpec()
        {
            // Arrange & Act
            var globalSpec = new GlobalSpec("TestSpec");

            // Assert
            Assert.IsInstanceOfType(globalSpec, typeof(Spec));
            Assert.IsInstanceOfType(globalSpec, typeof(IgxlRow));
        }
    }
}
