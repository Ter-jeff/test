using System.Collections.Generic;
using System.IO;

using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.InputManager.Data;
using Automation.Static;
using Automation.Utility.HardIP;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Common
{
    [TestClass]
    public class HardIpServiceTests : FunctionTestBase
    {
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            string hardIpInfoFile = Path.Combine(InputPath, "HardIp", "HardIP_PatInfo_Default.log");
            List<HardIpInfo> patInfo = new PatInfoReader().ExtractHardIpInfos(hardIpInfoFile);
            LocalSpecs.HardIpInfos = new HardIpInfos(patInfo);
        }

        [TestMethod]
        public void GetFlowForLoopIntegerName_SinglePattern_ReturnsFormattedString()
        {
            // Arrange
            var sweepCode = new SweepCode();
            var pattern = new HardIpPattern
            {
                Pattern = new PatternClass("A"),
                SweepCodes = new Dictionary<string, List<SweepCode>> { { "A", new List<SweepCode> { sweepCode } } }
            };
            var inputData = new HardIpInputData(null);

            // Act
            string result = HardIpService.GetFlowForLoopIntegerName(pattern, inputData);

            // Assert
            Assert.AreEqual("SrcCodeIndxA;:0:0:1", result);
        }

        [TestMethod]
        public void GetFlowForLoopIntegerName_Exceeds6000Chars_AddsRegAssignToInputData()
        {
            // Arrange
            var inputData = new HardIpInputData(null);
            var pattern = new HardIpPattern
            {
                Pattern = new PatternClass("A"),
            };
            for (int i = 0; i < 300; i++)
            {
                pattern.SweepCodes.Add(i.ToString(), [new() { SendBitName = i.ToString() }]);
            }

            // Act
            string result = HardIpService.GetFlowForLoopIntegerName(pattern, inputData);

            // Assert
            Assert.IsTrue(result.StartsWith("Reg_assign:"), "Should return a register assignment reference.");
            Assert.AreEqual(1, inputData.HardIpRegAssigns.Count, "Should have added a RegAssign item to input data.");
            Assert.AreEqual(RegisterAssignType.DigSrc_FlowForLoopIntegerName, inputData.HardIpRegAssigns[0].Type);
        }

        [TestMethod]
        public void GetFlowForLoopIntegerName_MultipleBurstPatterns_JoinsWithPipe()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                Pattern = new PatternClass("A")
                {
                    PatternSetList =
                    [["A"], ["B"]]
                },
            };
            pattern.BurstPatterns.Add(CreateSimplePattern("A"));
            pattern.BurstPatterns.Add(CreateSimplePattern("B"));
            var inputData = new HardIpInputData(null);

            // Act
            string result = HardIpService.GetFlowForLoopIntegerName(pattern, inputData);

            // Assert
            StringAssert.Contains(result, "SrcCodeIndxA;:0:0:1");
            StringAssert.Contains(result, "SrcCodeIndxB;:0:0:1");
            Assert.IsTrue(result.Contains('|'));

        }

        [TestMethod]
        public void GetFlowForLoopIntegerName_NoSweepCodes_ReturnsEmpty()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                Pattern = new PatternClass("A")
            };
            var inputData = new HardIpInputData(null);

            // Act
            string result = HardIpService.GetFlowForLoopIntegerName(pattern, inputData);

            // Assert
            Assert.AreEqual(string.Empty, result);
        }

        private static HardIpPattern CreateSimplePattern(string key)
        {
            return new HardIpPattern
            {
                SweepCodes = new Dictionary<string, List<SweepCode>> { { key, new List<SweepCode> { new() } } }
            };
        }

        [TestMethod]
        public void GetHardIpInfo_WithPattern_ShouldMergeInfos()
        {
            // Arrange
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
            HardIpInfo result = HardIpService.GetHardIpInfo(pattern);

            // Assert
            Assert.AreEqual("AAA+BBB", result.Payload);
            Assert.AreNotEqual(null, result.MultiDsscOut);
        }

        [TestMethod]
        public void GetHardIpInfo_WithIgnorePatInfo_ShouldReturnSimplifiedReference()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                Pattern = new PatternClass("")
                {
                    PatternSetList = [["CZ_NVSA0_A_FULP_AN_AN00_MEA_JTG_DIO_ALLFRV_SI_ACTCONS_T7"]]
                },
                MiscInfo = "Ignore_Patt_Comment",
            };

            // Act
            HardIpInfo result = HardIpService.GetHardIpInfo(pattern);

            // Assert
            Assert.IsTrue(result.IsIgnoreComment);
            Assert.IsTrue(result.CapBit > 0);
        }

        [TestMethod]
        public void GetHardIpInfo_EmptyPatternSet_ReturnsEmptyPayload()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                Pattern = new PatternClass("")
                {
                    PatternSetList = []
                }
            };

            // Act
            HardIpInfo result = HardIpService.GetHardIpInfo(pattern);

            // Assert
            Assert.AreEqual(string.Empty, result.Payload);
        }

        [TestMethod]
        public void GetHardIpInfo_IgnorePattComment_ShouldNotBreakBasicInfo()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                Pattern = new PatternClass("")
                {
                    PatternSetList =
            [
                ["AAA"]
            ]
                },
                MiscInfo = HardIpConstData.IgnorePatInfo
            };

            // Act
            HardIpInfo result = HardIpService.GetHardIpInfo(pattern);

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.AreEqual("AAA", result.Payload);
        }

        [TestMethod]
        public void GetHardIpInfo_IgnorePattDigSrc_ShouldClearSendBit()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                Pattern = new PatternClass("")
                {
                    PatternSetList =
            [
                ["AAA"]
            ]
                },
                MiscInfo = HardIpConstData.IgnorePatDigSrc
            };

            // Act
            HardIpInfo result = HardIpService.GetHardIpInfo(pattern);

            // Assert
            Assert.AreEqual("0", result.SendBit);
            Assert.AreEqual(string.Empty, result.SendBitName);
            Assert.AreEqual(string.Empty, result.SendPinName);
        }

        [TestMethod]
        [DataRow("pattern_nan", true)]
        [DataRow("abc_ids_xyz", true)]
        [DataRow("nand_test", false)]
        [DataRow("", false)]
        public void IsNandPattern_ValidInputs_ReturnsExpected(string pattern, bool expected)
        {
            bool result = HardIpService.IsNandPattern(pattern);
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        [DataRow("pattern_spi", true)]
        [DataRow("abc_spi_xyz", true)]
        [DataRow("spi_test", false)]
        [DataRow("", false)]
        public void IsSpiPattern_ValidInputs_ReturnsExpected(string pattern, bool expected)
        {
            bool result = HardIpService.IsSpiPattern(pattern ?? string.Empty);
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        [DataRow("10", "1.5x", "15")]
        [DataRow("10", "2x", "20")]
        [DataRow("5", "", "5")]
        [DataRow("0", "1.5x", "0")]
        [DataRow("", "1.5x", "")]
        public void CalcuteLimit_ValidInputs_ReturnsExpectedProduct(string limit, string repeat, string expected)
        {
            // Act
            string result = HardIpService.CalcuteLimit(limit, repeat);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void ReverseRelaySetting_ValidInput_ShouldSwapOnOff()
        {
            // Arrange
            string input = "RelayOff:PIO1;RelayOn:PIO2";

            // Act
            string result = HardIpService.ReverseRelaySetting(input);

            // Assert
            StringAssert.Contains(result, "RelayOn:");
            StringAssert.Contains(result, "RelayOff:");
            Assert.IsTrue(result.Contains('_'));
        }

        [TestMethod]
        [DataRow("No_patt", true)]
        [DataRow("NO_PATT", true)]
        [DataRow("ABC", false)]
        public void IsNoPattern_ValidInputs(string pattern, bool expected)
        {
            Assert.AreEqual(expected, HardIpService.IsNoPattern(pattern));
        }

        [TestMethod]
        [DataRow("CZ2_Only", true)]
        [DataRow("cz2_only", true)]
        [DataRow("Other", false)]
        public void IsCz2Only_ValidInputs(string misc, bool expected)
        {
            Assert.AreEqual(expected, HardIpService.IsCz2Only(misc));
        }

        [TestMethod]
        public void ActualLabelVoltage_HvOnly_ReturnsHv()
        {
            var pattern = new HardIpPattern
            {
                MiscInfo = HardIpConstData.HvOnly
            };

            string result = HardIpService.ActualLabelVoltage(HardIpConstData.LabelNv, pattern);

            Assert.AreEqual(HardIpConstData.LabelHv, result);
        }

        [TestMethod]
        public void PatFlowNoNeedToGen_NV_RunHv_ShouldSkip()
        {
            bool result = HardIpService.PatFlowNoNeedToGen(HardIpConstData.LabelNv, HardIpConstData.RunHv);

            Assert.IsTrue(result);
        }

        [TestMethod]
        public void GetRepeatMapping_WithLimit_ReturnsValue()
        {
            // Act
            string result = HardIpService.GetRepeatMapping("Limit:1.5x;Other:AA");

            // Assert
            Assert.AreEqual("1.5x", result);
        }

        [TestMethod]
        public void GetRepeatMapping_WithLimitConst_ReturnsValue()
        {
            string misc = $"{HardIpConstData.Limit}:1.5x;Other:AAA";
            string result = HardIpService.GetRepeatMapping(misc);
            Assert.AreEqual("1.5x", result);
        }

        [TestMethod]
        public void CalcuteLimit_NoRepeat_ReturnsOriginal()
        {
            string result = HardIpService.CalcuteLimit("10", "");
            Assert.AreEqual("10", result);
        }

    }
}
