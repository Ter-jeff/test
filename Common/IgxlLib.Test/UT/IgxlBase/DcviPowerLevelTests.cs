using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class DcviPowerLevelTests
    {
        [TestMethod]
        public void DcviPowerLevel_Constructor_InitializesAllProperties()
        {
            // Arrange & Act
            var dcviPowerLevel = new DcviPowerLevel("VCAP", "5.0V", "100mA", "10ns", "Power supply pin");

            // Assert
            Assert.AreEqual("VCAP", dcviPowerLevel.PinName);
            Assert.AreEqual("5.0V", dcviPowerLevel.Vps);
            Assert.AreEqual("100mA", dcviPowerLevel.Isc);
            Assert.AreEqual("10ns", dcviPowerLevel.TDelay);
            Assert.AreEqual("Power supply pin", dcviPowerLevel.Comment);
        }

        [TestMethod]
        public void DcviPowerLevel_SetProperties_UpdatesValuesCorrectly()
        {
            // Arrange
            var dcviPowerLevel = new DcviPowerLevel("VCAP", "5.0V", "100mA", "10ns", "Original")
            {
                // Act
                PinName = "VCAP_IO",
                Vps = "3.3V",
                Isc = "50mA",
                Comment = "Updated"
            };

            // Assert
            Assert.AreEqual("VCAP_IO", dcviPowerLevel.PinName);
            Assert.AreEqual("3.3V", dcviPowerLevel.Vps);
            Assert.AreEqual("50mA", dcviPowerLevel.Isc);
            Assert.AreEqual("Updated", dcviPowerLevel.Comment);
        }

        [TestMethod]
        public void DcviPowerLevel_MultipleInstances_AreIndependent()
        {
            // Arrange & Act
            var dcviPower1 = new DcviPowerLevel("VDD", "3.3V", "200mA", "5ns", "Main power");
            var dcviPower2 = new DcviPowerLevel("GND", "0V", "0mA", "0ns", "Ground");

            // Assert
            Assert.AreEqual("VDD", dcviPower1.PinName);
            Assert.AreEqual("GND", dcviPower2.PinName);
            Assert.AreEqual("3.3V", dcviPower1.Vps);
            Assert.AreEqual("0V", dcviPower2.Vps);
        }
    }
}
