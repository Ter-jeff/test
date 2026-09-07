using Automation.Reader.ConfigFile.RtosCategory;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Rtos
{
    [TestClass]
    public class RtosCategoryConfigReaderTests
    {
        [TestMethod]
        public void RtosConfigBitRow_GetMaxMinBit_BitNum_Works()
        {
            var row = new RtosConfigBitRow("[3:7]", "ModeX", "VDD,VSS", "Y", "10");

            Assert.AreEqual(7, row.GetMaxBit());
            Assert.AreEqual(3, row.GetMinBit());
            Assert.AreEqual(5, row.BitNum());
        }

        [TestMethod]
        public void RtosConfigBitRow_UseProperty_Works()
        {
            var rowY = new RtosConfigBitRow("[0:0]", "ModeX", "VDD", "Y", "0");
            var rowN = new RtosConfigBitRow("[0:0]", "ModeX", "VDD", "N", "0");
            var rowOther = new RtosConfigBitRow("[0:0]", "ModeX", "VDD", "?", "0");

            Assert.IsTrue(rowY.Use);
            Assert.IsFalse(rowN.Use);
            Assert.IsFalse(rowOther.Use);
        }

        [TestMethod]
        public void RtosConfigBitRow_GetMode_PerformanceOffset_Works()
        {
            var row = new RtosConfigBitRow("[3:0]", "Mode1XXXX", "VDD", "Y", "0010");
            string mode = row.GetMode("1011");
            Assert.AreEqual("Mode10013", mode);
        }

        [TestMethod]
        public void RtosCategoryConfigReader_GetCategoryName_ValidHex_Works()
        {
            var reader = new RtosCategoryConfigReader
            {
                Rows =
                [
                    new("[0:3]", "ModeXXXX", "VDD", "Y", "0")
                ]
            };

            string categoryName = reader.GetCategoryName("F");
            Assert.AreEqual("f", categoryName);
        }

        [TestMethod]
        public void RtosCategoryConfigReader_GetCategoryName_InvalidHex_ReturnsEmpty()
        {
            var reader = new RtosCategoryConfigReader();
            string result = reader.GetCategoryName("GHI");
            Assert.AreEqual("", result);
        }
    }
}
