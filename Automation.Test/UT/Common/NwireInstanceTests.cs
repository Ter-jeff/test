using Automation.GenerateIgxl.Basic.Business.GenNwire.Business;
using Automation.Reader;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Common
{
    [TestClass]
    public class NwireInstanceTests
    {
        [TestMethod]
        public void ResolveConfigFreq_Enable_ValidUnit()
        {
            var pin = new ProtocolAwarePin
            {
                ControlAction = "Enable@10MHz",
                Freq = 1000000
            };

            string result = NwireInstance.ResolveConfigFreq(pin);

            Assert.AreEqual("10000000", result);
        }

        [TestMethod]
        public void ResolveConfigFreq_Enable_InvalidUnit_ShouldFallback()
        {
            var pin = new ProtocolAwarePin
            {
                ControlAction = "Enable@ABC",
                Freq = 2000000
            };

            string result = NwireInstance.ResolveConfigFreq(pin);

            Assert.AreEqual("2000000", result);
        }

        [TestMethod]
        public void ResolveConfigFreq_NoEnable_ShouldFallback()
        {
            var pin = new ProtocolAwarePin
            {
                ControlAction = "Disable@",
                Freq = 3000000
            };

            string result = NwireInstance.ResolveConfigFreq(pin);

            Assert.AreEqual("3000000", result);
        }

        [TestMethod]
        public void ResolveConfigFreq_Enable_NoValueOrUnit_ShouldFallback()
        {
            var pin = new ProtocolAwarePin
            {
                ControlAction = "Enable@",
                Freq = 4000000
            };

            string result = NwireInstance.ResolveConfigFreq(pin);

            Assert.AreEqual("4000000", result);
        }
    }
}
