using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class LevelRowTests
    {
        [TestMethod]
        public void LevelRow_DefaultConstructor_CreatesEmptyInstance()
        {
            // Arrange & Act
            var levelRow = new LevelRow();

            // Assert
            Assert.IsNull(levelRow.PinName);
            Assert.IsNull(levelRow.Parameter);
            Assert.IsNull(levelRow.Value);
            Assert.IsNull(levelRow.Comment);
        }

        [TestMethod]
        public void LevelRow_Constructor_WithParameters_InitializesProperties()
        {
            // Arrange & Act
            var levelRow = new LevelRow("Pin1", "Voltage", "1.8V", "Level comment", 5);

            // Assert
            Assert.AreEqual("Pin1", levelRow.PinName);
            Assert.AreEqual("Voltage", levelRow.Parameter);
            Assert.AreEqual("1.8V", levelRow.Value);
            Assert.AreEqual("Level comment", levelRow.Comment);
            Assert.AreEqual(5, levelRow.RowNum);
        }

        [TestMethod]
        public void LevelRow_Constructor_WithoutRowNum_DefaultsToZero()
        {
            // Arrange & Act
            var levelRow = new LevelRow("Pin2", "Current", "100mA", "Current comment");

            // Assert
            Assert.AreEqual("Pin2", levelRow.PinName);
            Assert.AreEqual(0, levelRow.RowNum);
        }

        [TestMethod]
        public void LevelRow_SetProperties_UpdatesValuesCorrectly()
        {
            // Arrange
            var levelRow = new LevelRow
            {
                // Act
                PinName = "Pin3",
                Seq = "Seq1",
                Parameter = "Power",
                Value = "5.0V",
                Comment = "Updated"
            };

            // Assert
            Assert.AreEqual("Pin3", levelRow.PinName);
            Assert.AreEqual("Seq1", levelRow.Seq);
            Assert.AreEqual("Power", levelRow.Parameter);
            Assert.AreEqual("5.0V", levelRow.Value);
        }

        [TestMethod]
        public void LevelRow_CopyConstructor_CopiesAllProperties()
        {
            // Arrange
            var originalRow = new LevelRow("Pin4", "Timing", "10ns", "Timing comment", 10)
            {
                Seq = "Seq2",
                SheetName = "LevelSheet",
                IsBackup = true,
                ColumnA = "ColumnA"
            };

            // Act
            var copiedRow = new LevelRow(originalRow);

            // Assert
            Assert.AreEqual(originalRow.PinName, copiedRow.PinName);
            Assert.AreEqual(originalRow.Parameter, copiedRow.Parameter);
            Assert.AreEqual(originalRow.Value, copiedRow.Value);
            Assert.AreEqual(originalRow.Seq, copiedRow.Seq);
            Assert.AreEqual(originalRow.RowNum, copiedRow.RowNum);
        }

        [TestMethod]
        public void LevelRow_CopyConstructor_WithNull_CreatesEmptyInstance()
        {
            // Arrange & Act
            var levelRow = new LevelRow(null);

            // Assert
            Assert.IsNull(levelRow.PinName);
            Assert.IsNull(levelRow.Parameter);
        }

        [TestMethod]
        public void LevelRow_Copy_CreatesIndependentCopy()
        {
            // Arrange
            var originalRow = new LevelRow("Pin5", "Voltage", "3.3V", "Test");

            // Act
            LevelRow copiedRow = originalRow.Copy();

            // Assert
            Assert.AreEqual(originalRow.PinName, copiedRow.PinName);
            Assert.AreNotSame(originalRow, copiedRow);

            // Verify independence
            originalRow.PinName = "Pin5_Modified";
            Assert.AreEqual("Pin5", copiedRow.PinName);
        }

        [TestMethod]
        public void LevelRow_Inherits_FromIgxlRow()
        {
            // Arrange & Act
            var levelRow = new LevelRow("Pin", "Param", "Value", "Comment");

            // Assert
            Assert.IsInstanceOfType(levelRow, typeof(IgxlRow));
        }

        [TestMethod]
        public void LevelRow_MultipleInstances_AreIndependent()
        {
            // Arrange & Act
            var row1 = new LevelRow("Pin1", "Param1", "Value1", "Comment1");
            var row2 = new LevelRow("Pin2", "Param2", "Value2", "Comment2");

            // Assert
            Assert.AreEqual("Pin1", row1.PinName);
            Assert.AreEqual("Pin2", row2.PinName);
            Assert.AreNotEqual(row1.Parameter, row2.Parameter);
        }
    }
}
