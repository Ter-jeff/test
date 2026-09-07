using CommonLib.Extension;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test.UT.Extension
{
    [TestClass]
    public class StringExtensionsUnitTests
    {
        [TestMethod]
        public void ConvertNumber_Empty_ReturnsEmptyComponents()
        {
            Assert.AreEqual("", "".ConvertNumber(out string unit, out string scale));
            Assert.AreEqual("", unit);
            Assert.AreEqual("", scale);
        }

        [TestMethod]
        public void ConvertNumber_UnitsAndScale_ReturnsNormalizedValue()
        {
            string number = "1.2E-5mV".ConvertNumber(out string unit, out string scale);
            Assert.AreEqual("1.2E-08", number);
            Assert.AreEqual("m", scale);
            Assert.AreEqual("V", unit);
        }

        [TestMethod]
        public void ConvertNumber_ScientificUpperCase_ReturnsNormalizedValue()
        {
            string number = "1.2e5V".ConvertNumber(out string unit, out string scale);
            Assert.AreEqual("1.2E+05", number);
            Assert.AreEqual("", scale);
            Assert.AreEqual("V", unit);
        }

        [TestMethod]
        public void ConvertNumber_MilliVoltToVolt_ReturnsCorrectValue()
        {
            string number = "100mV".ConvertNumber(out string unit, out string scale);
            Assert.AreEqual("0.1", number);
            Assert.AreEqual("m", scale);
            Assert.AreEqual("V", unit);
        }

        [TestMethod]
        public void ConvertNumber_PercentScale_ReturnsCorrectValue()
        {
            string number = "100.1*%".ConvertNumber(out string unit, out string scale);
            Assert.AreEqual("1.001", number);
            Assert.AreEqual("%", scale);
            Assert.AreEqual("", unit);
        }

        [TestMethod]
        public void ConvertNumber_InvalidText_ReturnsOriginal()
        {
            string number = "abc".ConvertNumber(out string unit, out string scale);
            Assert.AreEqual("abc", number);
            Assert.AreEqual("", unit);
            Assert.AreEqual("", scale);
        }

        [TestMethod]
        public void ConvertNumberAndUnit_ExtractsNumberAndUnit()
        {
            (string number, string unit) = "100.1mV".ConvertNumberAndUnit();
            Assert.AreEqual("100.1", number);
            Assert.AreEqual("mV", unit);
        }

        [TestMethod]
        public void TryConvertToVolt_ValidVoltage_ReturnsTrue()
        {
            Assert.IsTrue("100mV".TryConvertToVolt(out string result));
            Assert.AreEqual("0.1", result);
        }

        [TestMethod]
        public void TryConvertToVolt_InvalidVoltage_ReturnsFalse()
        {
            Assert.IsFalse("abc".TryConvertToVolt(out _));
        }

        [TestMethod]
        public void TryConvertToFreq_ValidFrequency_ReturnsTrue()
        {
            Assert.IsTrue("1KHz".TryConvertToFreq(out string result));
            Assert.AreEqual("1000", result);
        }

        [TestMethod]
        public void TryConvertToFreq_InvalidFrequency_ReturnsFalse()
        {
            Assert.IsFalse("abc".TryConvertToFreq(out _));
        }

        [TestMethod]
        public void TryConvertToAmpere_ValidCurrent_ReturnsTrue()
        {
            Assert.IsTrue("100mA".TryConvertToAmpere(out string result));
            Assert.AreEqual("0.1", result);
        }

        [TestMethod]
        public void TryConvertToAmpere_InvalidCurrent_ReturnsFalse()
        {
            Assert.IsFalse("abc".TryConvertToAmpere(out _));
        }

        [TestMethod]
        public void TryCombineVolt_ValidInput_ReturnsTrue()
        {
            Assert.IsTrue("0.1".TryCombineVolt("mV", out string result));
            Assert.AreEqual("0.0001", result);
        }

        [TestMethod]
        public void TryCombineVolt_InvalidInput_ReturnsFalse()
        {
            Assert.IsFalse("abc".TryCombineVolt("V", out _));
            Assert.IsFalse("0.1".TryCombineVolt("InvalidUnit", out _));
        }

        [TestMethod]
        public void TryCombineHz_ValidInput_ReturnsTrue()
        {
            Assert.IsTrue("1".TryCombineHz("KHz", out string result));
            Assert.AreEqual("1000", result);
        }

        [TestMethod]
        public void TryCombineHz_InvalidInput_ReturnsFalse()
        {
            Assert.IsFalse("abc".TryCombineHz("Hz", out _));
        }

        [TestMethod]
        public void TryCombineAmpere_ValidInput_ReturnsTrue()
        {
            Assert.IsTrue("0.001".TryCombineAmpere("mA", out string result));
            Assert.AreEqual("1E-06", result);
        }

        [TestMethod]
        public void TryCombineAmpere_InvalidInput_ReturnsFalse()
        {
            Assert.IsFalse("abc".TryCombineAmpere("mA", out _));
        }

        [TestMethod]
        public void TryCombineAmpere_PicoAmpere_ReturnsCorrectlyScaledValue()
        {
            Assert.IsTrue("1".TryCombineAmpere("pA", out string result));
            Assert.AreEqual("1E-12", result);
        }

        [TestMethod]
        public void TryCombineAmpere_FemtoAmpere_ReturnsCorrectlyScaledValue()
        {
            Assert.IsTrue("1".TryCombineAmpere("fA", out string result));
            Assert.AreEqual("1E-15", result);
        }
    }
}
