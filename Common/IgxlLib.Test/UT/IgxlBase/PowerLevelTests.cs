using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class PowerLevelTests
    {
        [TestMethod]
        public void PowerLevel_Constructor_WithoutVSlewRate_InitializesProperties()
        {
            // Arrange & Act
            var powerLevel = new PowerLevel("VDD", "1.8", "1.5", "FoldA", "100ns", "Test power");

            // Assert
            Assert.AreEqual("VDD", powerLevel.PinName);
            Assert.AreEqual("1.8", powerLevel.Vmain);
            Assert.AreEqual("1.5", powerLevel.Valt);
            Assert.AreEqual("FoldA", powerLevel.IFoldLevel);
            Assert.AreEqual("100ns", powerLevel.Tdelay);
            Assert.AreEqual("Test power", powerLevel.Comment);
        }

        [TestMethod]
        public void PowerLevel_Constructor_WithVSlewRate_InitializesAllProperties()
        {
            // Arrange & Act
            var powerLevel = new PowerLevel("VCC", "3.3", "3.0", "FoldB", "50ns", "0.5V/ns", "Extended power");

            // Assert
            Assert.AreEqual("VCC", powerLevel.PinName);
            Assert.AreEqual("3.3", powerLevel.Vmain);
            Assert.AreEqual("3.0", powerLevel.Valt);
            Assert.AreEqual("FoldB", powerLevel.IFoldLevel);
            Assert.AreEqual("50ns", powerLevel.Tdelay);
            Assert.AreEqual("0.5V/ns", powerLevel.VSlewRate);
            Assert.AreEqual("Extended power", powerLevel.Comment);
        }

        [TestMethod]
        public void PowerLevel_Constructor_WithDefaultComment_InitializesWithEmptyComment()
        {
            // Arrange & Act
            var powerLevel = new PowerLevel("GND", "0", "0", "FoldC", "0ns");

            // Assert
            Assert.AreEqual("GND", powerLevel.PinName);
            Assert.AreEqual("", powerLevel.Comment);
        }

        [TestMethod]
        public void PowerLevel_SetProperties_UpdatesValuesCorrectly()
        {
            // Arrange
            var powerLevel = new PowerLevel("VDD", "1.8", "1.5", "FoldA", "100ns")
            {
                // Act
                PinName = "VDD_IO",
                Vmain = "2.0",
                VSlewRate = "1.0V/ns",
                Comment = "Updated"
            };

            // Assert
            Assert.AreEqual("VDD_IO", powerLevel.PinName);
            Assert.AreEqual("2.0", powerLevel.Vmain);
            Assert.AreEqual("1.0V/ns", powerLevel.VSlewRate);
            Assert.AreEqual("Updated", powerLevel.Comment);
        }

        [TestMethod]
        public void PowerLevel_MultipleInstances_AreIndependent()
        {
            // Arrange & Act
            var powerLevel1 = new PowerLevel("VDD", "1.8", "1.5", "FoldA", "100ns");
            var powerLevel2 = new PowerLevel("VCC", "3.3", "3.0", "FoldB", "50ns");

            // Assert
            Assert.AreEqual("VDD", powerLevel1.PinName);
            Assert.AreEqual("VCC", powerLevel2.PinName);
            Assert.AreNotEqual(powerLevel1.PinName, powerLevel2.PinName);
        }
    }
}
