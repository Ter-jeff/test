using System.Collections.Generic;

using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class PinGroupTests
    {
        [TestMethod]
        public void PinGroup_Constructor_WithNameAndType_InitializesProperties()
        {
            // Arrange & Act
            var pinGroup = new PinGroup("GroupA", "DigitalGroup");

            // Assert
            Assert.AreEqual("GroupA", pinGroup.PinName);
            Assert.AreEqual("DigitalGroup", pinGroup.PinType);
            Assert.AreEqual(0, pinGroup.PinList.Count);
        }

        [TestMethod]
        public void PinGroup_Constructor_WithNameOnly_InitializesProperties()
        {
            // Arrange & Act
            var pinGroup = new PinGroup("GroupB");

            // Assert
            Assert.AreEqual("GroupB", pinGroup.PinName);
            Assert.AreEqual(0, pinGroup.PinList.Count);
        }

        [TestMethod]
        public void PinGroup_AddPin_WithPinObject_AddsToList()
        {
            // Arrange
            var pinGroup = new PinGroup("GroupC", "DigitalGroup");
            var pin = new Pin("Pin1", "DigitalPin");

            // Act
            pinGroup.AddPin(pin);

            // Assert
            Assert.AreEqual(1, pinGroup.PinList.Count);
            Assert.AreEqual("Pin1", pinGroup.PinList[0].PinName);
        }

        [TestMethod]
        public void PinGroup_AddPin_WithPinNameAndType_AddsToList()
        {
            // Arrange
            var pinGroup = new PinGroup("GroupD", "DigitalGroup");

            // Act
            pinGroup.AddPin("Pin2", "DigitalPin");

            // Assert
            Assert.AreEqual(1, pinGroup.PinList.Count);
            Assert.AreEqual("Pin2", pinGroup.PinList[0].PinName);
            Assert.AreEqual("DigitalPin", pinGroup.PinList[0].PinType);
        }

        [TestMethod]
        public void PinGroup_AddPin_WithDuplicateName_DoesNotAddDuplicate()
        {
            // Arrange
            var pinGroup = new PinGroup("GroupE", "DigitalGroup");
            var pin1 = new Pin("Pin3", "DigitalPin");
            var pin2 = new Pin("Pin3", "DigitalPin");

            // Act
            pinGroup.AddPin(pin1);
            pinGroup.AddPin(pin2);

            // Assert
            Assert.AreEqual(1, pinGroup.PinList.Count);
        }

        [TestMethod]
        public void PinGroup_AddPin_CaseInsensitive_DoesNotAddDuplicate()
        {
            // Arrange
            var pinGroup = new PinGroup("GroupF", "DigitalGroup");

            // Act
            pinGroup.AddPin("Pin4", "DigitalPin");
            pinGroup.AddPin("pin4", "DigitalPin");

            // Assert
            Assert.AreEqual(1, pinGroup.PinList.Count);
        }

        [TestMethod]
        public void PinGroup_AddPin_WithNameOnly_AddsWithEmptyType()
        {
            // Arrange
            var pinGroup = new PinGroup("GroupG", "DigitalGroup");

            // Act
            pinGroup.AddPin("Pin5");

            // Assert
            Assert.AreEqual(1, pinGroup.PinList.Count);
            Assert.AreEqual("Pin5", pinGroup.PinList[0].PinName);
            Assert.AreEqual("", pinGroup.PinList[0].PinType);
        }

        [TestMethod]
        public void PinGroup_AddPins_WithMultiplePins_AddsAll()
        {
            // Arrange
            var pinGroup = new PinGroup("GroupH", "DigitalGroup");
            var pins = new List<Pin>
            {
                new("Pin6", "DigitalPin", "Comment1"),
                new("Pin7", "DigitalPin", "Comment2"),
                new("Pin8", "DigitalPin", "Comment3")
            };

            // Act
            pinGroup.AddPins(pins);

            // Assert
            Assert.AreEqual(3, pinGroup.PinList.Count);
        }

        [TestMethod]
        public void PinGroup_AddPins_WithDuplicates_OnlyAddsUnique()
        {
            // Arrange
            var pinGroup = new PinGroup("GroupI", "DigitalGroup");
            pinGroup.AddPin("Pin9", "DigitalPin");

            var pins = new List<Pin>
            {
                new("Pin9", "DigitalPin"),  // Duplicate
                new("Pin10", "DigitalPin"),
                new("Pin11", "DigitalPin")
            };

            // Act
            pinGroup.AddPins(pins);

            // Assert
            Assert.AreEqual(3, pinGroup.PinList.Count);
        }

        [TestMethod]
        public void PinGroup_Inherits_FromPinBase()
        {
            // Arrange & Act
            var pinGroup = new PinGroup("Group", "Type");

            // Assert
            Assert.IsInstanceOfType(pinGroup, typeof(PinBase));
            Assert.IsInstanceOfType(pinGroup, typeof(IgxlRow));
        }

        [TestMethod]
        public void PinGroup_MultipleInstances_AreIndependent()
        {
            // Arrange & Act
            var group1 = new PinGroup("Group1", "Type1");
            var group2 = new PinGroup("Group2", "Type2");

            group1.AddPin("Pin1", "DigitalPin");
            group2.AddPin("Pin2", "AnalogPin");

            // Assert
            Assert.AreEqual(1, group1.PinList.Count);
            Assert.AreEqual(1, group2.PinList.Count);
            Assert.AreEqual("Pin1", group1.PinList[0].PinName);
            Assert.AreEqual("Pin2", group2.PinList[0].PinName);
        }
    }
}
