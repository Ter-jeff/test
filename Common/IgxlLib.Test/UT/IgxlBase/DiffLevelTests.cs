using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class DiffLevelTests
    {
        [TestMethod]
        public void DiffLevel_Constructor_InitializesAllProperties()
        {
            // Arrange & Act
            var diffLevel = new DiffLevel(
                "DIFF_PIN",
                "0.9",          // vicm
                "0.2",          // vid
                "0.05",         // dVid0
                "0.06",         // dVid1
                "0.03",         // dVicm0
                "0.04",         // dVicm1
                "0.5",          // vod
                "0.52",         // vodAlt1
                "0.54",         // vodAlt2
                "0.02",         // dVod0
                "0.025",        // dVod1
                "5mA",          // iol
                "10mA",         // ioh
                "0.5",          // vodTyp
                "0.9",          // vocmTyp
                "0.45",         // vt
                "0.3",          // vcl
                "1.2",          // vch
                "LVDS"          // driverMode
            );

            // Assert
            Assert.AreEqual("DIFF_PIN", diffLevel.PinName);
            Assert.AreEqual("0.9", diffLevel.Vicm);
            Assert.AreEqual("0.2", diffLevel.Vid);
            Assert.AreEqual("0.05", diffLevel.DVid0);
            Assert.AreEqual("0.06", diffLevel.DVid1);
            Assert.AreEqual("0.03", diffLevel.DVicm0);
            Assert.AreEqual("0.04", diffLevel.DVicm1);
            Assert.AreEqual("0.5", diffLevel.Vod);
            Assert.AreEqual("0.52", diffLevel.VodAlt1);
            Assert.AreEqual("0.54", diffLevel.VodAlt2);
            Assert.AreEqual("0.02", diffLevel.DVod0);
            Assert.AreEqual("0.025", diffLevel.DVod1);
            Assert.AreEqual("5mA", diffLevel.Iol);
            Assert.AreEqual("10mA", diffLevel.Ioh);
            Assert.AreEqual("0.5", diffLevel.VodTyp);
            Assert.AreEqual("0.9", diffLevel.VocmTyp);
            Assert.AreEqual("0.45", diffLevel.Vt);
            Assert.AreEqual("0.3", diffLevel.Vcl);
            Assert.AreEqual("1.2", diffLevel.Vch);
            Assert.AreEqual("LVDS", diffLevel.DriverMode);
        }

        [TestMethod]
        public void DiffLevel_SetProperties_UpdatesValuesCorrectly()
        {
            // Arrange
            var diffLevel = new DiffLevel("PIN1", "0.9", "0.2", "0.05", "0.06", "0.03", "0.04", "0.5", "0.52", "0.54", "0.02", "0.025", "5mA", "10mA", "0.5", "0.9", "0.45", "0.3", "1.2", "LVDS")
            {
                // Act
                PinName = "UPDATED_PIN",
                Vicm = "1.0",
                DriverMode = "SLVS"
            };

            // Assert
            Assert.AreEqual("UPDATED_PIN", diffLevel.PinName);
            Assert.AreEqual("1.0", diffLevel.Vicm);
            Assert.AreEqual("SLVS", diffLevel.DriverMode);
        }

        [TestMethod]
        public void DiffLevel_MultipleInstances_AreIndependent()
        {
            // Arrange & Act
            var diffLevel1 = new DiffLevel("PIN1", "0.9", "0.2", "0.05", "0.06", "0.03", "0.04", "0.5", "0.52", "0.54", "0.02", "0.025", "5mA", "10mA", "0.5", "0.9", "0.45", "0.3", "1.2", "LVDS");

            var diffLevel2 = new DiffLevel("PIN2", "1.0", "0.3", "0.06", "0.07", "0.04", "0.05", "0.6", "0.62", "0.64", "0.03", "0.035", "8mA", "15mA", "0.6", "1.0", "0.5", "0.4", "1.3", "SLVS");

            // Assert
            Assert.AreEqual("PIN1", diffLevel1.PinName);
            Assert.AreEqual("PIN2", diffLevel2.PinName);
            Assert.AreEqual("LVDS", diffLevel1.DriverMode);
            Assert.AreEqual("SLVS", diffLevel2.DriverMode);
        }
    }
}
