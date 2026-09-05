using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class IoLevelTests
    {
        [TestMethod]
        public void IoLevel_Constructor_InitializesAllProperties()
        {
            // Arrange & Act
            var ioLevel = new IoLevel("IO_PIN", "0.2", "0.8", "0.1", "2.9", "2.95", "2.97", "10mA", "20mA", "1.4", "0.5", "2.0", "0.05", "2.9", "CMOS");

            // Assert
            Assert.AreEqual("IO_PIN", ioLevel.PinName);
            Assert.AreEqual("0.2", ioLevel.Vil);
            Assert.AreEqual("0.8", ioLevel.Vih);
            Assert.AreEqual("0.1", ioLevel.Vol);
            Assert.AreEqual("2.9", ioLevel.Voh);
            Assert.AreEqual("2.95", ioLevel.VohAlt1);
            Assert.AreEqual("2.97", ioLevel.VohAlt2);
            Assert.AreEqual("10mA", ioLevel.Iol);
            Assert.AreEqual("20mA", ioLevel.Ioh);
            Assert.AreEqual("1.4", ioLevel.Vt);
            Assert.AreEqual("0.5", ioLevel.Vcl);
            Assert.AreEqual("2.0", ioLevel.Vch);
            Assert.AreEqual("0.05", ioLevel.VOutLoTyp);
            Assert.AreEqual("2.9", ioLevel.VOutHiTyp);
            Assert.AreEqual("CMOS", ioLevel.DriverMode);
        }

        [TestMethod]
        public void IoLevel_SetProperties_UpdatesValuesCorrectly()
        {
            // Arrange
            var ioLevel = new IoLevel("IO_PAD", "0.2", "0.8", "0.1", "2.9", "2.95", "2.97", "10mA", "20mA", "1.4", "0.5", "2.0", "0.05", "2.9", "CMOS")
            {
                // Act
                PinName = "IO_PAD_NEW",
                Vil = "0.3",
                Vih = "0.9",
                DriverMode = "TTL"
            };

            // Assert
            Assert.AreEqual("IO_PAD_NEW", ioLevel.PinName);
            Assert.AreEqual("0.3", ioLevel.Vil);
            Assert.AreEqual("0.9", ioLevel.Vih);
            Assert.AreEqual("TTL", ioLevel.DriverMode);
        }

        [TestMethod]
        public void IoLevel_MultipleInstances_AreIndependent()
        {
            // Arrange & Act
            var ioLevel1 = new IoLevel("PIN1", "0.2", "0.8", "0.1", "2.9", "2.95", "2.97", "10mA", "20mA", "1.4", "0.5", "2.0", "0.05", "2.9", "CMOS");

            var ioLevel2 = new IoLevel("PIN2", "0.3", "0.9", "0.2", "3.0", "3.05", "3.07", "15mA", "25mA", "1.5", "0.6", "2.1", "0.1", "3.0", "TTL");

            // Assert
            Assert.AreEqual("PIN1", ioLevel1.PinName);
            Assert.AreEqual("PIN2", ioLevel2.PinName);
            Assert.AreEqual("CMOS", ioLevel1.DriverMode);
            Assert.AreEqual("TTL", ioLevel2.DriverMode);
        }

        [TestMethod]
        public void IoLevel_VoltageProperties_CanBeDecimalValues()
        {
            // Arrange & Act
            var ioLevel = new IoLevel("ANALOG_PIN", "0.05", "3.15", "0.0", "3.3", "3.32", "3.35", "5mA", "15mA", "1.65", "0.3", "3.0", "0.02", "3.28", "Analog");

            // Assert
            Assert.AreEqual("0.05", ioLevel.Vil);
            Assert.AreEqual("3.15", ioLevel.Vih);
            Assert.AreEqual("3.32", ioLevel.VohAlt1);
            Assert.AreEqual("3.35", ioLevel.VohAlt2);
        }
    }
}
