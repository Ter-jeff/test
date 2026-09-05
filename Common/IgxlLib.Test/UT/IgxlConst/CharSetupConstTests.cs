using IgxlLib.IgxlConst;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlConst
{
    [TestClass]
    public class CharSetupConstTests
    {
        [TestMethod]
        public void ParameterType_ShouldBeCaseInsensitive()
        {
            Assert.AreEqual("AC Spec", CharSetupConst.ParameterType["acspec"]);
            Assert.AreEqual("AC Spec", CharSetupConst.ParameterType["ACSpec"]);
        }

        [TestMethod]
        public void ParameterType_ShouldReturnCorrectMapping()
        {
            Assert.AreEqual("Global Spec", CharSetupConst.ParameterType["GlobalSpec"]);
        }

        [TestMethod]
        public void ParameterType_ShouldNotContainInvalidKey()
        {
            Assert.IsFalse(CharSetupConst.ParameterType.ContainsKey("Invalid"));
        }

        [TestMethod]
        public void ParameterName_ShouldReturnCorrectValue()
        {
            Assert.AreEqual("Master Period", CharSetupConst.ParameterName["MasterPeriod"]);
        }

        [TestMethod]
        public void ParameterType_ShouldHandleAllDefinedKeys()
        {
            Assert.AreEqual("DC Spec", CharSetupConst.ParameterType["DCSpec"]);
            Assert.AreEqual("Edge", CharSetupConst.ParameterType["Edge"]);
            Assert.AreEqual("Protocol Aware", CharSetupConst.ParameterType["ProtocolAware"]);
            Assert.AreEqual("Serial Timing", CharSetupConst.ParameterType["SerialTiming"]);
            Assert.AreEqual("VBT Parameter", CharSetupConst.ParameterType["VBTParameter"]);
        }

        [TestMethod]
        public void ParameterName_ShouldBeCaseInsensitive()
        {
            Assert.AreEqual("Master Period", CharSetupConst.ParameterName["masterperiod"]);
            Assert.AreEqual("Ref Offset", CharSetupConst.ParameterName["refoffset"]);
        }

        [TestMethod]
        public void ParameterName_ShouldHandleAllDefinedKeys()
        {
            Assert.AreEqual("DUT Period", CharSetupConst.ParameterName["DUTPeriod"]);
            Assert.AreEqual("Clock Offset", CharSetupConst.ParameterName["ClockOffset"]);
            Assert.AreEqual("Drive Delay", CharSetupConst.ParameterName["DriveDelay"]);
            Assert.AreEqual("Receive Delay", CharSetupConst.ParameterName["ReceiveDelay"]);
            Assert.AreEqual("HiZ Delay", CharSetupConst.ParameterName["HiZDelay"]);
            Assert.AreEqual("Reference Offset", CharSetupConst.ParameterName["ReferenceOffset"]);
        }

        [TestMethod]
        public void TestMethod_ShouldBeCaseInsensitive()
        {
            Assert.AreEqual("Retest", CharSetupConst.TestMethod["retest"]);
            Assert.AreEqual("Run Pattern", CharSetupConst.TestMethod["runpattern"]);
        }

        [TestMethod]
        public void TestMethod_ShouldReturnCorrectMapping()
        {
            Assert.AreEqual("Reburst", CharSetupConst.TestMethod["Reburst"]);
            Assert.AreEqual("Reburst Serial", CharSetupConst.TestMethod["ReburstSerial"]);
            Assert.AreEqual("Run Function", CharSetupConst.TestMethod["RunFunction"]);
        }

        [TestMethod]
        public void TestMethod_ShouldNotContainInvalidKey()
        {
            Assert.IsFalse(CharSetupConst.TestMethod.ContainsKey("Invalid"));
        }

    }
}
