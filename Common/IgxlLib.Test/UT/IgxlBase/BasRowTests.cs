using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class BasRowTests
    {
        [TestMethod]
        public void BasRow_DefaultConstructor_CreatesEmptyInstance()
        {
            // Arrange & Act
            var basRow = new BasRow();

            // Assert
            Assert.IsNull(basRow.Text);
        }

        [TestMethod]
        public void BasRow_Constructor_WithText_InitializesTextProperty()
        {
            // Arrange & Act
            var basRow = new BasRow("Test text content");

            // Assert
            Assert.AreEqual("Test text content", basRow.Text);
        }

        [TestMethod]
        public void BasRow_SetText_UpdatesTextProperty()
        {
            // Arrange
            var basRow = new BasRow
            {
                Text = "Updated text"
            };

            // Assert
            Assert.AreEqual("Updated text", basRow.Text);
        }

        [TestMethod]
        public void BasRow_MultipleInstances_AreIndependent()
        {
            // Arrange & Act
            var row1 = new BasRow("Text1");
            var row2 = new BasRow("Text2");

            // Assert
            Assert.AreEqual("Text1", row1.Text);
            Assert.AreEqual("Text2", row2.Text);
            Assert.AreNotEqual(row1.Text, row2.Text);
        }

        [TestMethod]
        public void BasRow_TextProperty_CanBeNull()
        {
            // Arrange
            var basRow = new BasRow("Initial")
            {
                // Act
                Text = null
            };

            // Assert
            Assert.IsNull(basRow.Text);
        }

        [TestMethod]
        public void BasRow_TextProperty_CanBeEmpty()
        {
            // Arrange & Act
            var basRow = new BasRow("");

            // Assert
            Assert.AreEqual("", basRow.Text);
        }

        [TestMethod]
        public void BasRow_Inherits_FromIgxlRow()
        {
            // Arrange & Act
            var basRow = new BasRow("Text");

            // Assert
            Assert.IsInstanceOfType(basRow, typeof(IgxlRow));
        }
    }
}
