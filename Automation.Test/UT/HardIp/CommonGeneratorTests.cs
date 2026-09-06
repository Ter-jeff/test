using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;

using CommonReaderLib.PatternListCsv;

using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class CommonGeneratorTests : FunctionTestBase
    {
        [TestMethod]
        [DataRow("hardip_test", "BlockA", "Sub-Block", "Bin_HIP_BlockA_Sub_Block")]
        [DataRow("NormalSheet", "BlockB", "Sub-Block-C", "Bin_BlockB_Sub_Block_C")]
        [DataRow("HARDIP_UPPER", "BlockX", "Sub", "Bin_HIP_BlockX_Sub")]
        public void GenHardIpFlowBinParameter_ValidInputs_ReturnsFormattedString(string sheet, string block, string subBlock, string expected)
        {
            // Act
            string result = CommonGenerator.GenHardIpFlowBinParameter(sheet, block, subBlock);

            // Assert
            Assert.AreEqual(expected, result);
        }

        private static HardIpInputData CreateMockData(bool nv = false, bool czNv = false)
        {
            var hardIpParaData = new HardIpParaData(EnumBlock.HardIp)
            {
                NvEnable = nv,
                CzNvEnable = czNv,
                LvEnable = false,
                CzLvEnable = false,
                HvEnable = false,
                CzHvEnable = false
            };
            return new HardIpInputData(hardIpParaData);
        }

        [TestMethod]
        [DataRow("normal_pat", "NV", "HardIP_NV")]
        [DataRow("cz_pattern_test", "NV", "HardIP_NV_CZ")]
        [DataRow("mn_pattern_test", "NV", "HardIP_NV_MN")]
        [DataRow("cz_mn_test", "NV", "HardIP_NV_CZ")]
        public void GenEnableWord_PatternNaming_AppendsCorrectSuffix(string pat, string volt, string expected)
        {
            HardIpInputData data = CreateMockData(true, true);
            string result = CommonGenerator.GenEnableWord(pat, "misc", volt, data);
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        [DataRow("pattern", "", "", false, "")]
        [DataRow("pattern", "", "", true, "")]
        [DataRow("pattern", "NV", "RemoveNV", false, "")]
        [DataRow("pattern", "LV", "RemoveLV", false, "")]
        [DataRow("pattern", "HV", "RemoveHV", false, "")]
        [DataRow("pattern", "NV", "", false, "")]
        [DataRow("pattern", "LV", "", false, "")]
        [DataRow("pattern", "HV", "", false, "")]
        [DataRow("cz_pattern", "", "", false, "")]
        [DataRow("cz_pattern", "", "", true, "")]
        [DataRow("cz_pattern", "NV", "RemoveNV", false, "")]
        [DataRow("cz_pattern", "LV", "RemoveLV", false, "")]
        [DataRow("cz_pattern", "HV", "RemoveHV", false, "")]
        [DataRow("cz_pattern", "NV", "", false, "")]
        [DataRow("cz_pattern", "LV", "", false, "")]
        [DataRow("cz_pattern", "HV", "", false, "")]
        public void GenEnableWord_CzFlagDisabled_ReturnsEmpty(string pattern, string labelVoltage, string miscInfo, bool czNvFlag, string expected)
        {
            // Arrange
            HardIpInputData data = CreateMockData(czNv: czNvFlag);

            // Act
            string result = CommonGenerator.GenEnableWord(pattern, miscInfo, labelVoltage, data);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        [DataRow("SheetA_Extra", "Block1", "Sub", "Instance:MyPattern", "NV", "InstX", "", false, "F_SheetA_Block1_Sub_MyPattern_InstX_Flag")]
        [DataRow("SheetB", "Block2", "", "Pat", "HV", "InstY", "", true, "F_SheetB_Block2_InstY_Flag")]
        public void GenHardIpFlowFailFlag_ValidInputs_ReturnsFormattedFlag(string sheetName, string blockName, string subBlockName, string patternName, string timingAc, string instNameSubStr, string labelVoltage, bool noPattern, string expectedPart)
        {
            // Act
            string result = CommonGenerator.GenHardIpFlowFailFlag(sheetName, blockName, subBlockName, patternName, timingAc, instNameSubStr, labelVoltage, noPattern);

            // Assert
            // Verifying the middle construction logic
            Assert.IsTrue(result.Contains(expectedPart), $"Expected result to contain {expectedPart} but was {result}");
            Assert.IsTrue(result.StartsWith(HardIpConstData.PrefixHardIpFailAction));
        }

        // Test for GenWirelessFlowFailFlag
        [TestMethod]
        [DataRow("WiFi", "5G", "NV", "_WiFi_5G_N")]
        [DataRow("BT", "", "HV", "_BT_H")]
        public void GenWirelessFlowFailFlag_ValidInputs_ReturnsFormattedFlag(string block, string sub, string volt, string expectedPart)
        {
            // Act
            string result = CommonGenerator.GenWirelessFlowFailFlag(block, sub, volt);

            // Assert
            Assert.IsTrue(result.Contains(expectedPart));
            Assert.IsTrue(result.StartsWith(HardIpConstData.PrefixHardIpFailAction));
        }

        // Test for GenHardIpBlockFailFlag
        [TestMethod]
        public void GenHardIpBlockFailFlag_SheetWithUnderscore_SplitsCorrectly()
        {
            // Arrange
            string sheet = "USB_3_0_Data";
            string block = "PHY";

            // Act
            string result = CommonGenerator.GenHardIpBlockFailFlag(sheet, block);

            // Assert
            // Should take "USB" from "USB_3_0_Data"
            string expected = HardIpConstData.PrefixHardIpFailAction + "_USB_PHY" + HardIpConstData.SuffixHardIpFailAction;
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void ConvertPatNameInOpcode_ValidMatch_ReplacesPatternWithFailAction()
        {
            // Arrange
            var opcodes = new List<string> { "Test:MY_PAT" };
            var flowRows = new List<FlowRow>
        {
            new()
            {
                Parameter = "MY_PAT_info",
                FailAction = "FAIL_FLAG_N",
                Opcode = "NotUseLimit"
            }
        };
            string voltage = "NV";

            // Act
            CommonGenerator.ConvertPatNameInOpcode(opcodes, flowRows, voltage);

            // Assert
            Assert.AreEqual("Test:MY_PAT", opcodes[0]);
        }

        [TestMethod]
        public void ConvertPatNameInOpcode_NoLimitMatch_ReplacesWithFailAction()
        {
            // Arrange
            var opcodes = new List<string> { "Keep:PP_Pattern" };
            var flowRows = new List<FlowRow>()
            {
                new()
                {
                    Parameter = "PP_Pattern" ,
                    FailAction = "H_Flag"
                }
            };

            // Act
            CommonGenerator.ConvertPatNameInOpcode(opcodes, flowRows, "HV");

            // Assert
            Assert.AreEqual("Keep:H_Flag", opcodes[0]);
        }

        [TestMethod]
        public void GetInterposeAssignName_StandardPattern_ReturnsCombinedName()
        {
            // Arrange
            var patternMock = new HardIpPattern
            {
                MiscInfo = "some_info",
                SheetName = "Sheet1",
                Pattern = new PatternClass("normal_payload")
            };

            // Act
            string result = CommonGenerator.GetRegAssignName(patternMock);

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.IsFalse(result.Contains("_DD"));
        }

        [TestMethod]
        public void GetInterposeAssignName_DdPrefix_AppendsSuffix()
        {
            // Arrange
            var patternMock = new HardIpPattern
            {
                MiscInfo = "info",
                SheetName = "Sheet1",
                Pattern = new PatternClass("dd_test_payload")
            };

            // Act
            string result = CommonGenerator.GetRegAssignName(patternMock);

            // Assert
            Assert.IsTrue(result.EndsWith("_DD"), $"Result '{result}' should end with _DD");
        }

        [TestMethod]
        public void GetInterposeAssignName_NullPattern_ThrowsException()
        {
            // Act & Assert
            Assert.ThrowsException<System.NullReferenceException>(() => CommonGenerator.GetRegAssignName(null));
        }

        [TestMethod]
        public void GetHardipSheetName_RemovesSpaceAndKeepsUnderscore()
        {
            // Arrange
            string sheetName = "Hard IP Sheet";

            // Act
            string result = CommonGenerator.GetHardipSheetName(sheetName);

            // Assert
            Assert.AreEqual("Hard_IP_Sheet", result);
        }

        [TestMethod]
        public void GetBlockNameFromSheetName_WithUnderscore_RemovesFirstSegment()
        {
            // Arrange
            string sheetName = "HardIP_USB_PHY";

            // Act
            string blockName = CommonGenerator.GetBlockNameFromSheetName(sheetName);

            // Assert
            Assert.AreEqual("USBPHY", blockName);
        }

        [TestMethod]
        public void GetSubBlockNameWithoutMinus_RemovesDashCharacter()
        {
            // Arrange
            string subBlock = "SUB-BLOCK-A";

            // Act
            string result = CommonGenerator.GetSubBlockNameWithoutMinus(subBlock);

            // Assert
            Assert.AreEqual("SUBBLOCKA", result);
        }

        [TestMethod]
        public void SplitByDelimiter_InputExceedsChunkSize_SplitsCorrectly()
        {
            // Arrange
            string input = "AAA;BBB;CCC";
            char delimiter = ';';
            int chunkSize = 3;

            // Act
            List<string> result = CommonGenerator.SplitByDelimiter(input, delimiter, chunkSize);

            // Assert
            CollectionAssert.AreEqual(
                new List<string> { "AAA", "BBB", "CCC" },
                result);
        }

        [TestMethod]
        [DataRow("PAT_A", "SubBlock:SUB_A_B", "BLOCK", false, "SUBAB")]
        [DataRow("SI_BLOCK_SUB1_SUB2", "", "BLOCK", false, "BLOCK_SUB1_SUB2")]
        [DataRow("SI_BLOCK", "", "BLOCK", false, "")]
        [DataRow("Instance:MYINST", "", "BLOCK", false, "MYINST")]
        [DataRow("Opcode:TEST:PARA", "", "BLOCK", false, "PARA")]
        [DataRow("Opcode:TEST", "", "BLOCK", false, "")]
        [DataRow("PAT_A", "SubBlock:SUB_A", "BLOCK", true, "SUBACZ")]
        public void GetSubBlockName_ValidInputs_ReturnExpectedResult(string patternName, string miscInfo, string blockName, bool isShmooInChar, string expected)
        {
            string result = CommonGenerator.GetSubBlockName(patternName, miscInfo, blockName, isShmooInChar);

            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void GetSubBlockNameByPattern_SimplePattern_ReturnsSegmentsAfterSi()
        {
            // Arrange
            string patternName = "SI_BLOCK_SUB";
            string blockName = "BLOCK";

            // Act
            string result = CommonGenerator.GetSubBlockNameByPattern(patternName, blockName);

            // Assert
            Assert.AreEqual("BLOCK_SUB", result);
        }

        [TestMethod]
        public void GenHardIpInsTestName_BasicInputs_ReturnsUpperCaseName()
        {
            // Arrange
            string block = "BLOCK";
            string subBlock = "SUB";
            string pattern = "PAT";
            string patIndex = "_1";
            string timingAc = "";
            string forceV = "";
            string instSubStr = "";
            string labelVoltage = "";
            bool noPattern = false;
            bool isPostBurn = false;
            bool isGenByFlow = false;
            bool isDoMeas = false;

            // Act
            string result = CommonGenerator.GenHardIpInsTestName(block, subBlock, pattern, patIndex, timingAc, forceV, instSubStr, labelVoltage, noPattern, isPostBurn, isGenByFlow, isDoMeas);

            // Assert
            Assert.AreEqual("BLOCK_SUB_PAT_1", result);
        }

        [TestMethod]
        [DataRow("BLOCK", "", "Instance:INST", "_1", "", "", "", "", false, false, false, false, "INSREMOV_BLOCK_INST_1")]
        [DataRow("BLOCK", "", "Instance:INST", "_1", "", "", "", "", false, false, true, false, "BLOCK_INST_1")]
        [DataRow("BLOCK", "", "Opcode:TEST:PARA", "", "", "", "", "", false, false, false, false, "PARA")]
        [DataRow("BLOCK", "SUB-A", "PAT", "_1", "", "", "", "", false, false, false, false, "BLOCK_SUBA_PAT_1")]
        [DataRow("BLOCK", "SUB", "PAT", "_1", "", "", "INST", "", true, false, false, false, "BLOCK_SUB_INST_1")]
        public void GenHardIpInsTestName_ValidInputs_ReturnExpectedResult(string blockName, string subBlockName, string patternName, string prefixPatIndexFlag, string timingAc, string prefixForceVFlag, string instNameSubStr, string labelVoltage, bool noPattern, bool isPostBurn, bool isGenByFlow, bool isDoMeas, string expected)
        {
            string result = CommonGenerator.GenHardIpInsTestName(blockName, subBlockName, patternName, prefixPatIndexFlag, timingAc, prefixForceVFlag, instNameSubStr, labelVoltage, noPattern, isPostBurn, isGenByFlow, isDoMeas);

            Assert.AreEqual(expected, result);
        }

    }
}
