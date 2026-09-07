using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class RegisterAssignParserTests
    {
        private RegisterAssignParser _parser = null!;

        [TestInitialize]
        public void Setup()
        {
            _parser = new RegisterAssignParser();
        }

        [TestMethod]
        public void GenRegisterAssign_ShouldReturnZeroBits_WhenLengthIsGiven()
        {
            // Arrange
            string sendBitName = "Reg1";
            int bitLength = 4;

            // Act
            string result = _parser.GenRegisterAssign(sendBitName, bitLength);

            // Assert
            Assert.AreEqual("Reg1=0000", result);
        }

        [TestMethod]
        public void GenRegisterAssign_ShouldReturnEmpty_WhenLengthIsZero()
        {
            string result = _parser.GenRegisterAssign("Reg1", 0);

            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void GenNestSweepRegisterAssign_ShouldReturnCorrectFormat()
        {
            // Act
            string result = _parser.GenNestSweepRegisterAssign("Reg1", 8, "00,11", "1");

            // Assert
            Assert.AreEqual("Reg1=nestsweep1(8@00,11)", result);
        }

        [TestMethod]
        public void GenRegisterAssign_WithAlgorithm_ShouldReturnExpected()
        {
            string result = _parser.GenRegisterAssign("Reg1", "00,11", "1", "BinToGray");

            Assert.AreEqual("Reg1=sweep1[00,11]@BinToGray", result);
        }

        [TestMethod]
        public void GenRegisterAssign_WithoutAlgorithm_ShouldReturnExpected()
        {
            string result = _parser.GenRegisterAssign("Reg1", "00,11", "1", "");

            Assert.AreEqual("Reg1=sweep1[00,11]", result);
        }

        [TestMethod]
        public void GetSweepCode_ShouldReturnBasicSweepAssignment()
        {
            // Arrange
            var pattern = new HardIpPattern();
            string sendBitName = "Reg1";
            string sweepCode = "sweep(0,3,1)";
            HardIpInfo patInfo = new HardIpInfo
            {
                Payload = "dp_brna0_c_fulp_an_aa22_frq_jtg_dyc_allfrv_si_ciopll_t2p1r",
                PatInfoExist = true,
                SendBitName = "Reg1",
                SendBitStr = "A_5",
                SeqInfo =
                [
                    new()
                    {
                        PinList = ["Pin1"]
                    }
                ]
            };

            // Act
            string result = _parser.GetSweepCode(sendBitName, sweepCode, pattern, patInfo);

            // Assert
            Assert.AreEqual("Reg1=00000", result);
        }

        [TestMethod]
        public void ParseSrcValue_ShouldReturnSameValue_WhenNoReverse()
        {
            // Arrange
            var pattern = new HardIpPattern();
            string sendBitName = "RegA";
            string sendValue = "1010";

            // Act
            string result = _parser.ParseSrcValue(sendBitName, sendValue, false, pattern);

            // Assert
            Assert.AreEqual("RegA=1010", result);
        }

        [TestMethod]
        public void ParseSrcValue_ShouldReverseBits_WhenReverseTrue()
        {
            // Arrange
            var pattern = new HardIpPattern();
            string sendBitName = "RegB";
            string sendValue = "1100";

            // Act
            string result = _parser.ParseSrcValue(sendBitName, sendValue, true, pattern);

            // Assert
            Assert.AreEqual("RegB=0011", result);
        }

        [TestMethod]
        public void ParseSrcValue_ShouldHandleEmptyValue()
        {
            // Arrange
            var pattern = new HardIpPattern();
            string sendBitName = "RegC";
            string sendValue = "";

            // Act
            string result = _parser.ParseSrcValue(sendBitName, sendValue, false, pattern);

            // Assert
            Assert.AreEqual("RegC=", result);
        }

        [TestMethod]
        public void ParseSrcValue_ShouldStoreValueInPatternSweepCodes()
        {
            // Arrange
            var pattern = new HardIpPattern();
            string sendBitName = "RegD";
            string sendValue = "1111";

            // Act
            string result = _parser.ParseSrcValue(sendBitName, sendValue, false, pattern);

            Assert.AreEqual("RegD=1111", result);
            Assert.AreNotEqual(null, pattern.SweepCodes);
        }
    }
}
