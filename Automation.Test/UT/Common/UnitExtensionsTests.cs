
using CommonLib.Extension;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Common
{

    [TestClass]
    public class UnitExtensionsTests
    {
        [DataTestMethod]
        [DataRow("1", "V", true, "1")]
        [DataRow("1500", "mV", true, "1.5")]
        [DataRow("1000000", "uV", true, "1")]
        [DataRow("1234", "mV", true, "1.234")]
        [DataRow("1.23456789", "mV", true, "0.001235")] // rounded to 6
        [DataRow("abc", "V", false, "")]
        [DataRow("1", "XYZ", false, "")]
        [DataRow("", "V", false, "")]
        [DataRow(null, "V", false, "")]

        public void TryCombineVolt_ShouldReturnExpectedResult(string source, string unit, bool expectedResult, string expectedOutput)
        {
            // Act
            bool result = source.TryCombineVolt(unit, out string output);

            // Assert
            Assert.AreEqual(expectedResult, result);
            Assert.AreEqual(expectedOutput, output);
        }

        [DataTestMethod]
        [DataRow("1", "Hz", true, "1")]
        [DataRow("1.5", "kHz", true, "1500")]
        [DataRow("2", "MHz", true, "2000000")]
        [DataRow("0.1", "GHz", true, "100000000")]
        [DataRow("abc", "Hz", false, "")]
        [DataRow("1", "XYZ", false, "")]
        [DataRow("", "Hz", false, "")]
        [DataRow(null, "Hz", false, "")]

        public void TryCombineHz_ShouldReturnExpectedResult(string source, string unit, bool expectedResult, string expectedOutput)
        {
            // Act
            bool result = source.TryCombineHz(unit, out string output);

            // Assert
            Assert.AreEqual(expectedResult, result);
            Assert.AreEqual(expectedOutput, output);
        }
    }
}
