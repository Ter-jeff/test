using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class SelectorTests
    {
        [TestMethod]
        public void Selector_DefaultConstructor_CreatesEmptyInstance()
        {
            // Arrange & Act
            var selector = new Selector();

            // Assert
            Assert.IsNull(selector.SelectorName);
            Assert.IsNull(selector.SelectorValue);
        }

        [TestMethod]
        public void Selector_Constructor_WithParameters_InitializesProperties()
        {
            // Arrange & Act
            var selector = new Selector("SelectorName", "SelectorValue");

            // Assert
            Assert.AreEqual("SelectorName", selector.SelectorName);
            Assert.AreEqual("SelectorValue", selector.SelectorValue);
        }

        [TestMethod]
        public void Selector_SetProperties_UpdatesValuesCorrectly()
        {
            // Arrange
            var selector = new Selector
            {
                // Act
                SelectorName = "TestName",
                SelectorValue = "TestValue"
            };

            // Assert
            Assert.AreEqual("TestName", selector.SelectorName);
            Assert.AreEqual("TestValue", selector.SelectorValue);
        }

        [TestMethod]
        public void Selector_MultipleInstances_AreIndependent()
        {
            // Arrange & Act
            var selector1 = new Selector("Name1", "Value1");
            var selector2 = new Selector("Name2", "Value2");

            // Assert
            Assert.AreEqual("Name1", selector1.SelectorName);
            Assert.AreEqual("Value1", selector1.SelectorValue);
            Assert.AreEqual("Name2", selector2.SelectorName);
            Assert.AreEqual("Value2", selector2.SelectorValue);
        }
    }
}
