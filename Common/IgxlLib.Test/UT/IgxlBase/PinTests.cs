using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class PinTests
    {
        [TestMethod]
        public void Pin_DefaultConstructor_CreatesEmptyInstance()
        {
            // Arrange & Act
            var pin = new Pin();

            // Assert
            Assert.AreEqual(string.Empty, pin.PinName);
            Assert.AreEqual(string.Empty, pin.PinType);
            Assert.AreEqual(string.Empty, pin.Comment);
            Assert.AreEqual(string.Empty, pin.ChannelType);
        }

        [TestMethod]
        public void Pin_Constructor_WithPinNameAndType_InitializesProperties()
        {
            // Arrange & Act
            var pin = new Pin("PinA", "DigitalPin");

            // Assert
            Assert.AreEqual("PinA", pin.PinName);
            Assert.AreEqual("DigitalPin", pin.PinType);
            Assert.AreEqual("", pin.Comment);
            Assert.AreEqual("", pin.ChannelType);
        }

        [TestMethod]
        public void Pin_Constructor_WithAllParameters_InitializesAllProperties()
        {
            // Arrange & Act
            var pin = new Pin("PinB", "AnalogPin", "Test comment");

            // Assert
            Assert.AreEqual("PinB", pin.PinName);
            Assert.AreEqual("AnalogPin", pin.PinType);
            Assert.AreEqual("Test comment", pin.Comment);
        }

        [TestMethod]
        public void Pin_CopyConstructor_CopiesAllProperties()
        {
            // Arrange
            var originalPin = new Pin("PinC", "DigitalPin", "Original comment")
            {
                ChannelType = "ChannelA",
                InstrumentType = "InstrumentX",
                SheetName = "TestSheet",
                RowNum = 5,
                IsBackup = true,
                ColumnA = "TestColumn"
            };

            // Act
            var copiedPin = new Pin(originalPin);

            // Assert
            Assert.AreEqual(originalPin.PinName, copiedPin.PinName);
            Assert.AreEqual(originalPin.PinType, copiedPin.PinType);
            Assert.AreEqual(originalPin.Comment, copiedPin.Comment);
            Assert.AreEqual(originalPin.ChannelType, copiedPin.ChannelType);
            Assert.AreEqual(originalPin.InstrumentType, copiedPin.InstrumentType);
            Assert.AreEqual(originalPin.SheetName, copiedPin.SheetName);
            Assert.AreEqual(originalPin.RowNum, copiedPin.RowNum);
            Assert.AreEqual(originalPin.IsBackup, copiedPin.IsBackup);
        }

        [TestMethod]
        public void Pin_CopyConstructor_WithNull_CreatesEmptyInstance()
        {
            // Arrange & Act
            var pin = new Pin(null);

            // Assert
            Assert.AreEqual(string.Empty, pin.PinName);
            Assert.AreEqual(string.Empty, pin.PinType);
            Assert.AreEqual(string.Empty, pin.Comment);
        }

        [TestMethod]
        public void Pin_Copy_CreatesIndependentCopy()
        {
            // Arrange
            var originalPin = new Pin("PinD", "DigitalPin", "Original")
            {
                ChannelType = "Channel1"
            };

            // Act
            Pin copiedPin = originalPin.Copy();

            // Assert
            Assert.AreEqual(originalPin.PinName, copiedPin.PinName);
            Assert.AreNotSame(originalPin, copiedPin);

            // Verify independence
            originalPin.PinName = "ChangedPin";
            Assert.AreEqual("PinD", copiedPin.PinName);
            Assert.AreEqual("ChangedPin", originalPin.PinName);
        }

        [TestMethod]
        public void Pin_Inherits_FromPinBase_And_IgxlRow()
        {
            // Arrange & Act
            var pin = new Pin("Pin", "Type");

            // Assert
            Assert.IsInstanceOfType(pin, typeof(PinBase));
            Assert.IsInstanceOfType(pin, typeof(IgxlRow));
        }
    }
}
