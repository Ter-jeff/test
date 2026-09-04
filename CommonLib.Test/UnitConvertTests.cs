using CommonLib.Utility.StringExtension;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test
{
    [TestClass]
    public class UnitConvertTests
    {
        [TestMethod]
        public void Test1()
        {
            string unit;
            string scale;
            string number = "".ConvertUnit(out unit, out scale);
            Assert.AreEqual(number, "");
            Assert.AreEqual(scale, "");
            Assert.AreEqual(unit, "");
            number = "100.1".ConvertUnit(out unit, out scale);
            Assert.AreEqual(number, "100.1");
            Assert.AreEqual(scale, "");
            Assert.AreEqual(unit, "");
            number = "1.2E-5mV".ConvertUnit(out unit, out scale);
            Assert.AreEqual(number, "1.2E-08");
            Assert.AreEqual(scale, "m");
            Assert.AreEqual(unit, "V");
            number = "1.2e5V".ConvertUnit(out unit, out scale);
            Assert.AreEqual(number, "1.2E+05");
            Assert.AreEqual(scale, "");
            Assert.AreEqual(unit, "V");
            number = "0.000001V".ConvertUnit(out unit, out scale);
            Assert.AreEqual(number, "1E-06");
            Assert.AreEqual(scale, "");
            Assert.AreEqual(unit, "V");
            number = "100mV".ConvertUnit(out unit, out scale);
            Assert.AreEqual(number, "0.1");
            Assert.AreEqual(scale, "m");
            Assert.AreEqual(unit, "V");
            number = "100.1*mV".ConvertUnit(out unit, out scale);
            Assert.AreEqual(number, "0.1001");
            Assert.AreEqual(scale, "m");
            Assert.AreEqual(unit, "V");
            number = "100.1*%".ConvertUnit(out unit, out scale);
            Assert.AreEqual(number, "1.001");
            Assert.AreEqual(scale, "%");
            Assert.AreEqual(unit, "");
            number = "abc".ConvertUnit(out unit, out scale);
            Assert.AreEqual(number, "abc");
            Assert.AreEqual(scale, "");
            Assert.AreEqual(unit, "");
            number = "100.1abc".ConvertUnit(out unit, out scale);
            Assert.AreEqual(number, "100.1");
            Assert.AreEqual(scale, "");
            Assert.AreEqual(unit, "abc");

        }
    }
}
