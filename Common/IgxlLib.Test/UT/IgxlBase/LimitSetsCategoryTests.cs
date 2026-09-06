using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class LimitSetsCategoryTests
    {
        [TestMethod]
        public void LimitSetsCategory_DefaultConstructor_InitializesWithDefaults()
        {
            // Arrange & Act
            var category = new LimitSetsItem();

            // Assert
            Assert.AreEqual("", category.CategoryName);
            Assert.AreEqual("NA", category.LoLim);
            Assert.AreEqual("NA", category.HiLim);
        }

        [TestMethod]
        public void LimitSetsCategory_Constructor_WithParameters_InitializesProperties()
        {
            // Arrange & Act
            var category = new LimitSetsItem("Category1", "1.0", "10.0");

            // Assert
            Assert.AreEqual("Category1", category.CategoryName);
            Assert.AreEqual("1.0", category.LoLim);
            Assert.AreEqual("10.0", category.HiLim);
        }

        [TestMethod]
        public void LimitSetsCategory_Constructor_WithEmptyLimits_KeepsNA()
        {
            // Arrange & Act
            var category = new LimitSetsItem("Category2", "", "");

            // Assert
            Assert.AreEqual("Category2", category.CategoryName);
            Assert.AreEqual("NA", category.LoLim);
            Assert.AreEqual("NA", category.HiLim);
        }

        [TestMethod]
        public void LimitSetsCategory_Constructor_WithNullLimits_KeepsNA()
        {
            // Arrange & Act
            var category = new LimitSetsItem("Category3", null, null);

            // Assert
            Assert.AreEqual("Category3", category.CategoryName);
            Assert.AreEqual("NA", category.LoLim);
            Assert.AreEqual("NA", category.HiLim);
        }

        [TestMethod]
        public void LimitSetsCategory_Constructor_WithPartialLimits_UpdatesProvided()
        {
            // Arrange & Act
            var category = new LimitSetsItem("Category4", "0.5", "");

            // Assert
            Assert.AreEqual("Category4", category.CategoryName);
            Assert.AreEqual("0.5", category.LoLim);
            Assert.AreEqual("NA", category.HiLim);
        }

        [TestMethod]
        public void LimitSetsCategory_SetProperties_UpdatesValues()
        {
            // Arrange
            var category = new LimitSetsItem
            {
                // Act
                CategoryName = "UpdatedCat",
                LoLim = "5.0",
                HiLim = "15.0"
            };

            // Assert
            Assert.AreEqual("UpdatedCat", category.CategoryName);
            Assert.AreEqual("5.0", category.LoLim);
            Assert.AreEqual("15.0", category.HiLim);
        }
    }
}
