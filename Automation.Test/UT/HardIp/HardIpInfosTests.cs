using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.Static;

using CommonLib.Enums;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class HardIpInfosTests : FunctionTestBase
    {
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            LocalSpecs.Options.Device = EnumDevice.LCD;
        }

        [TestMethod]
        public void GetHardIpInfo_ShouldReturnMatchingInfo_WhenUseThisVersionIsTrue()
        {
            // Arrange
            var ref1 = new HardIpInfo
            {
                Payload = "PATTERN_A",
                UseThisVersion = true,
                CapBit = 300000,
                CapBitName = "CAP_A+Extra",
                CapBitStr = "A_123+Extra",
                DsscOut = "DSSC_OUT,32:TNAME:STORENAME"
            };
            var refs = new HardIpInfos(ref1);

            // Act
            HardIpInfo result = refs.GetHardIpInfo("PATTERN_A", useThisVersion: true);

            // Assert
            Assert.AreEqual(ref1, result);
            Assert.AreEqual("CAP_A+Extra", result.CapBitName);
            Assert.AreEqual("DSSC_OUT,32:TNAME:STORENAME", result.DsscOut);
        }

        [TestMethod]
        public void GetHardIpInfo_ShouldReturnEmpty_WhenNoMatch()
        {
            // Arrange
            var ref1 = new HardIpInfo
            {
                Payload = "PATTERN_X",
                UseThisVersion = true
            };
            var refs = new HardIpInfos(ref1);

            // Act
            HardIpInfo result = refs.GetHardIpInfo("PATTERN_Y", useThisVersion: true);

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.AreEqual(string.Empty, result.Payload);
        }

        [TestMethod]
        public void GetHardIpInfo_WithPattern_ShouldMergeInfos()
        {
            // Arrange
            var ref1 = new HardIpInfo { Payload = "AAA", MultiDsscOut = "OUT1,", SeqInfo = [new() { SeqName = "SEQ1" }], CapBit = 30000, DsscOut = "A", UseThisVersion = true };
            var ref2 = new HardIpInfo { Payload = "AAA", MultiDsscOut = "OUT2,", SeqInfo = [new() { SeqName = "SEQ2" }], CapBit = 30000, DsscOut = "B", UseThisVersion = true };
            var refs = new HardIpInfos(ref1, ref2);

            var pattern = new HardIpPattern
            {
                Pattern = new PatternClass("")
                {
                    PatternSetList =
                    [
                        ["AAA", "BBB"]
                    ]
                },
                MiscInfo = "Ignore_Patt_DigSrc"
            };

            // Act
            HardIpInfo result = refs.GetHardIpInfo(pattern);

            // Assert
            Assert.AreEqual("AAA+BBB", result.Payload);
            StringAssert.Contains(result.MultiDsscOut, "A|");
        }

        [TestMethod]
        public void GetHardIpInfo_WithIgnorePatInfo_ShouldReturnSimplifiedReference()
        {
            // Arrange
            var ref1 = new HardIpInfo
            {
                Payload = "PAT1",
                SeqInfo = [
                    new() { SeqName = "SEQ1" }
                ],
                CapBit = 123,
                UseThisVersion = true
            };
            var refs = new HardIpInfos(ref1);

            var pattern = new HardIpPattern
            {
                Pattern = new PatternClass("")
                {
                    PatternSetList = [["PAT1"]]
                },
                MiscInfo = "Ignore_Patt_Comment"
            };

            // Act
            HardIpInfo result = refs.GetHardIpInfo(pattern);

            // Assert
            Assert.IsTrue(result.IsIgnoreComment);
            Assert.AreEqual(ref1.CapBit, result.CapBit);
        }
    }
}
