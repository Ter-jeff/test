using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class CategoryInSpecTests
    {
        [TestMethod]
        public void CategoryInSpec_Constructor_WithFullParameters_InitializesAllProperties()
        {
            // Arrange & Act
            var category = new CategoryInSpec("TestCategory", "TypeA", "10", "20");

            // Assert
            Assert.AreEqual("TestCategory", category.Name);
            Assert.AreEqual("TypeA", category.Typ);
            Assert.AreEqual("10", category.Min);
            Assert.AreEqual("20", category.Max);
        }

        [TestMethod]
        public void CategoryInSpec_Constructor_WithNameOnly_InitializesNameProperty()
        {
            // Arrange & Act
            var category = new CategoryInSpec("TestCategory");

            // Assert
            Assert.AreEqual("TestCategory", category.Name);
            Assert.IsNull(category.Typ);
            Assert.IsNull(category.Min);
            Assert.IsNull(category.Max);
        }

        [TestMethod]
        public void CategoryInSpec_SetProperties_UpdatesValuesCorrectly()
        {
            // Arrange
            var category = new CategoryInSpec("TestCategory")
            {
                // Act
                Typ = "TypeB",
                Min = "5",
                Max = "15"
            };

            // Assert
            Assert.AreEqual("TypeB", category.Typ);
            Assert.AreEqual("5", category.Min);
            Assert.AreEqual("15", category.Max);
        }

        [TestMethod]
        public void CategoryInSpec_MultipleInstances_AreIndependent()
        {
            // Arrange & Act
            var category1 = new CategoryInSpec("Cat1", "Type1", "0", "10");
            var category2 = new CategoryInSpec("Cat2", "Type2", "20", "30");

            // Assert
            Assert.AreEqual("Cat1", category1.Name);
            Assert.AreEqual("Cat2", category2.Name);
            Assert.AreNotEqual(category1.Name, category2.Name);
        }

        [TestMethod]
        public void CategoryInSpec_Properties_CanBeEmptyStrings()
        {
            // Arrange & Act
            var category = new CategoryInSpec("", "", "", "");

            // Assert
            Assert.AreEqual("", category.Name);
            Assert.AreEqual("", category.Typ);
            Assert.AreEqual("", category.Min);
            Assert.AreEqual("", category.Max);
        }
    }
}
