using Automation.GenerateIgxl.HardIp.InputObject;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class HardIpInfoTests
    {
        private HardIpInfo _target = null!;

        [TestInitialize]
        public void Setup()
        {
            _target = new HardIpInfo();
        }

        [TestMethod]
        public void IsMsbFirst_WhenLsbIsFalse_ReturnsTrue()
        {
            _target.IsSendLsbFirstInfo = "true+false+true";
            Assert.AreEqual("TRUE", _target.IsMsbFirst());
        }

        [TestMethod]
        public void IsMsbFirst_WhenAllLsbTrue_ReturnsEmpty()
        {
            _target.IsSendLsbFirstInfo = "true+true";
            _target.IsCapLsbFirstInfo = "true";
            Assert.AreEqual("", _target.IsMsbFirst());
        }

        [TestMethod]
        public void IsRegisterReverse_WhenRegisterMatchesFalse_ReturnsTrue()
        {
            _target.SendBitName = "RegA+RegB";
            _target.IsSendLsbFirstInfo = "true+false";

            bool result = _target.IsRegisterReverse("RegB");

            Assert.IsTrue(result);
        }

        [TestMethod]
        public void IsRegisterReverse_WhenIndexOutOfRange_ReturnsFalse()
        {
            _target.SendBitName = "RegA+RegB";
            // Missing second index
            _target.IsSendLsbFirstInfo = "true";

            bool result = _target.IsRegisterReverse("RegB");

            Assert.IsFalse(result);
        }

        [TestMethod]
        public void ExtractForcePrePat_ACFormat_ParsesCorrectly()
        {
            // Format: AC:Pin1&Pin2:Value
            _target.ForcePrePatRaw = ["AC:CLK&DATA:50MHz"];

            ForceCondition result = _target.ExtractForcePrePat();

            Assert.AreEqual(1, result.ForcePins.Count);
            Assert.AreEqual("AC", result.ForcePins[0].ForceType);
            // & becomes ::
            Assert.AreEqual("CLK::DATA", result.ForcePins[0].PinName);
            Assert.AreEqual("50MHz", result.ForcePins[0].ForceValue);
        }

        [TestMethod]
        public void ExtractForcePrePat_StandardFormat_ParsesMultiplePins()
        {
            // Format: Pin1,Pin2:Type:Value
            _target.ForcePrePatRaw = ["PinA,PinB:V:1.8V"];

            ForceCondition result = _target.ExtractForcePrePat();

            Assert.AreEqual(2, result.ForcePins.Count);
            Assert.AreEqual("PinA", result.ForcePins[0].PinName);
            Assert.AreEqual("PinB", result.ForcePins[1].PinName);
            Assert.AreEqual("V", result.ForcePins[0].ForceType);
            Assert.AreEqual("1.8V", result.ForcePins[0].ForceValue);
        }

        [TestMethod]
        public void ExtractForcePrePat_EmptyList_ReturnsEmptyCondition()
        {
            _target.ForcePrePatRaw = [];

            ForceCondition result = _target.ExtractForcePrePat();

            Assert.AreNotEqual(null, result.ForcePins);
            Assert.AreEqual(0, result.ForcePins.Count);
        }
    }
}
