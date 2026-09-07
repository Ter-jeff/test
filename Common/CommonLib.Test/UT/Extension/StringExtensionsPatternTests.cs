using CommonLib.Enums;
using CommonLib.Extension;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test.UT.Extension
{
    [TestClass]
    public class StringExtensionsPatternTests
    {
        [TestMethod]
        public void GetPatternType_RTOSPattern_ReturnsRTOS()
        {
            Assert.AreEqual(EnumPatternType.RTOS, "_RTS_".GetPatternType());
            Assert.AreEqual(EnumPatternType.RTOS, "_rts_".GetPatternType());
        }

        [TestMethod]
        public void GetPatternType_HARDIPPattern_ReturnsHARDIP()
        {
            Assert.AreEqual(EnumPatternType.HARDIP, "_AN_".GetPatternType());
        }

        [TestMethod]
        public void GetPatternType_InitPattern_ReturnsInit()
        {
            // Pattern: _IN followed by 2 word chars followed by _
            Assert.AreEqual(EnumPatternType.Init, "_INAB_".GetPatternType());
            Assert.AreEqual(EnumPatternType.Init, "_INXY_".GetPatternType());
        }

        [TestMethod]
        public void GetPatternType_PayloadPattern_ReturnsPayload()
        {
            // Pattern: _PL followed by 2 word chars followed by _
            Assert.AreEqual(EnumPatternType.Payload, "_PLAB_".GetPatternType());
            Assert.AreEqual(EnumPatternType.Payload, "_FULP_".GetPatternType());
        }

        [TestMethod]
        public void GetPatternType_RetentionPattern_ReturnsRetentionWait()
        {
            Assert.AreEqual(EnumPatternType.RetentionWait, "RETENTION_PAUSE_TEST".GetPatternType());
        }

        [TestMethod]
        public void GetPatternType_UnknownPattern_ReturnsUnknown()
        {
            Assert.AreEqual(EnumPatternType.Unknown, "UNKNOWN_PATTERN".GetPatternType());
        }

        [TestMethod]
        public void IsInit_InitPattern_ReturnsTrue()
        {
            // Pattern: _IN followed by 2 word chars followed by _
            Assert.IsTrue("_INAB_".IsInit());
            Assert.IsTrue("_INXY_".IsInit());
        }

        [TestMethod]
        public void IsInit_NonInitPattern_ReturnsFalse()
        {
            Assert.IsFalse("_RTS_".IsInit());
            Assert.IsFalse("UNKNOWN".IsInit());
        }
    }
}
