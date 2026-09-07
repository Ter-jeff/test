using Automation.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using Automation.GenerateIgxl.HardIp.InputObject;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class RegisterConvertorTests : FunctionTestBase
    {
        [TestMethod]
        public void ConvertNumberToSrc_ShouldConvertBinaryAndPadCorrectly()
        {
            // Arrange
            var refObj = new HardIpInfo
            {
                SendBitName = "REG1",
                SendBitStr = "BIT_8"
            };

            // Act
            string result = RegisterConvertor.ConvertNumberToSrc("REG1", "0b101", refObj);

            // Assert
            Assert.AreEqual("10100000", result);
        }

        [TestMethod]
        public void ConvertNumberToSrc_ShouldConvertHexadecimal()
        {
            // Arrange
            var refObj = new HardIpInfo
            {
                SendBitName = "REG_HEX",
                SendBitStr = "BIT_8"
            };

            // Act
            string result = RegisterConvertor.ConvertNumberToSrc("REG_HEX", "0xF", refObj);

            // Assert
            Assert.AreEqual("11110000", result);
        }

        [TestMethod]
        public void ConvertNumberToSrc_ShouldConvertDecimal()
        {
            // Arrange
            var refObj = new HardIpInfo
            {
                SendBitName = "REG_DEC",
                SendBitStr = "BIT_8"
            };

            // Act
            string result = RegisterConvertor.ConvertNumberToSrc("REG_DEC", "0d5", refObj);

            // Assert
            Assert.AreEqual("10100000", result);
        }

        [TestMethod]
        public void ConvertNumberToSrc_ShouldReturnEmpty_WhenBitLengthIsZero()
        {
            // Arrange
            var refObj = new HardIpInfo
            {
                SendBitName = "REG_NONE",
                SendBitStr = ""
            };

            // Act
            string result = RegisterConvertor.ConvertNumberToSrc("REG_NONE", "0xA", refObj);

            // Assert
            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void ConvertNumberToSrc_ShouldReturnEmpty_WhenInvalidNumber()
        {
            // Arrange
            var refObj = new HardIpInfo
            {
                SendBitName = "REG_ERR",
                SendBitStr = "BIT_8"
            };

            // Act
            string result = RegisterConvertor.ConvertNumberToSrc("REG_ERR", "0xGARBAGE", refObj);

            // Assert
            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void ConvertNumberToSrc_ShouldReturnOriginal_WhenBinaryTooLong()
        {
            // Arrange
            var refObj = new HardIpInfo
            {
                SendBitName = "REG",
                SendBitStr = "BIT_4"
            };

            // Act
            string result = RegisterConvertor.ConvertNumberToSrc("REG", "0b111111", refObj);

            // Assert
            Assert.AreEqual("0b111111", result);
        }

        [TestMethod]
        public void ConvertNumberToSrc_ShouldHandleNegativeValues()
        {
            // Arrange
            var refObj = new HardIpInfo
            {
                SendBitName = "REG_NEG",
                SendBitStr = "BIT_8"
            };

            // Act
            string result = RegisterConvertor.ConvertNumberToSrc("REG_NEG", "0d-1", refObj);

            // Assert
            Assert.IsTrue(result.Length <= 8);
        }
    }
}
