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
            Assert.AreEqual("", number);
            Assert.AreEqual("", scale);
            Assert.AreEqual("", unit);
            number = "100.1".ConvertUnit(out unit, out scale);
            Assert.AreEqual("100.1", number);
            Assert.AreEqual("", scale);
            Assert.AreEqual("", unit);
            number = "1.2E-5mV".ConvertUnit(out unit, out scale);
            Assert.AreEqual("1.2E-08", number);
            Assert.AreEqual("m", scale);
            Assert.AreEqual("V", unit);
            number = "1.2e5V".ConvertUnit(out unit, out scale);
            Assert.AreEqual("1.2E+05", number);
            Assert.AreEqual("", scale);
            Assert.AreEqual("V", unit);
            number = "0.000001V".ConvertUnit(out unit, out scale);
            Assert.AreEqual("1E-06", number);
            Assert.AreEqual("", scale);
            Assert.AreEqual("V", unit);
            number = "100mV".ConvertUnit(out unit, out scale);
            Assert.AreEqual("0.1", number);
            Assert.AreEqual("m", scale);
            Assert.AreEqual("V", unit);
            number = "100.1*mV".ConvertUnit(out unit, out scale);
            Assert.AreEqual("0.1001", number);
            Assert.AreEqual("m", scale);
            Assert.AreEqual("V", unit);
            number = "100.1*%".ConvertUnit(out unit, out scale);
            Assert.AreEqual("1.001", number);
            Assert.AreEqual("%", scale);
            Assert.AreEqual("", unit);
            number = "abc".ConvertUnit(out unit, out scale);
            Assert.AreEqual("abc", number);
            Assert.AreEqual("", scale);
            Assert.AreEqual("", unit);
            number = "100.1abc".ConvertUnit(out unit, out scale);
            Assert.AreEqual("100.1", number);
            Assert.AreEqual("", scale);
            Assert.AreEqual("abc", unit);

        }
    }
}
