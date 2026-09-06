using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.SpecialSetting;

using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class OpCodeSettingMainTests
    {
        [TestMethod]
        public void GenOpcodeSetting_SimpleOpcode_SetsEnableCorrectly()
        {
            // Arrange
            var settings = new List<string> { "Call:MyParam" };
            string enableValue = "True";

            // Act
            List<FlowRow> result = OpCodeSettingMain.GenOpcodeSetting(settings, enable: enableValue);

            // Assert
            Assert.AreEqual("Call", result[0].Opcode);
            Assert.AreEqual("True", result[0].Enable);
        }

        [TestMethod]
        public void GenOpcodeSetting_ControlFlowOpcode_DoesNotSetEnable()
        {
            // Arrange: "If" should skip setting the Enable property based on regex
            var settings = new List<string> { "If:Condition" };
            string enableValue = "True";

            // Act
            List<FlowRow> result = OpCodeSettingMain.GenOpcodeSetting(settings, enable: enableValue);

            // Assert
            Assert.AreEqual("", result[0].Enable, "Control flow opcodes (If/ElseIf/EndIf) should not have Enable set.");
        }

        [TestMethod]
        public void GenOpcodeSetting_RegexReplacement_PrefixesCorrectly()
        {
            // Arrange
            var settings = new List<string> { "Op:pp_Test_Signal" };
            string blockName = "USB";

            // Act
            List<FlowRow> result = OpCodeSettingMain.GenOpcodeSetting(settings, blockName: blockName);

            // Assert
            // Logic: pp_ becomes F_{Block}_{Value}
            Assert.AreEqual("F_USB_pp_Test_Signal", result[0].Parameter);
        }

        [TestMethod]
        [DataRow("NV", "_N_Flag")]
        [DataRow("LV", "_L_Flag")]
        [DataRow("HV", "_H_Flag")]
        [DataRow("Unknown", "")]
        public void GenOpcodeSetting_VoltageFlag_AppendsCorrectSuffix(string voltageInput, string expectedSuffix)
        {
            // Arrange
            var settings = new List<string> { "Set:dd_Value" };
            string block = "IP";

            // Act
            List<FlowRow> result = OpCodeSettingMain.GenOpcodeSetting(settings, blockName: block, voltage: voltageInput);

            // Assert
            string expected = $"F_IP_dd_Value{expectedSuffix}";
            Assert.AreEqual(expected, result[0].Parameter);
        }

        [TestMethod]
        public void GenOpcodeSetting_MixedContent_OnlyReplacesMatches()
        {
            // Arrange
            var settings = new List<string> { "Calc:pp_Var + NormalVar" };

            // Act
            List<FlowRow> result = OpCodeSettingMain.GenOpcodeSetting(settings, blockName: "Core");

            // Assert
            // "pp_Var" is replaced, "NormalVar" stays the same
            Assert.AreEqual("F_Core_pp_Var + NormalVar", result[0].Parameter);
        }
    }
}
