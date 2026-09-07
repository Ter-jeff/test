using Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo.FunctionAssigner.Universal;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;

using CommonReaderLib.PatternListCsv;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class MiscAssignerTests
    {
        private static string InvokeGetInstSpecialSetup(string setup, HardIpInputData inputData)
        {
            return MiscAssigner.GetInstSpecialSetup(setup, inputData);
        }

        private static HardIpInputData NewInputData()
        {
            return new HardIpInputData(new HardIpParaData(EnumBlock.HardIp));
        }

        [TestMethod]
        public void GetInstSpecialSetup_NoColonInSetup_ReturnsEmpty()
        {
            // Act
            string result = InvokeGetInstSpecialSetup("NoColonHere", NewInputData());

            // Assert
            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void GetInstSpecialSetup_ValueParsesAsInt_ReturnsNumericString()
        {
            // Act
            string result = InvokeGetInstSpecialSetup("InstSpecialSetting:42", NewInputData());

            // Assert
            Assert.AreEqual("42", result);
        }

        [TestMethod]
        public void GetInstSpecialSetup_NonIntValueFoundInConfig_ReturnsMappedValue()
        {
            // Arrange
            HardIpInputData inputData = NewInputData();
            inputData.ConfigData.InstSpecialSetting["MyKey"] = "MappedVal";

            // Act
            string result = InvokeGetInstSpecialSetup("InstSpecialSetting:MyKey", inputData);

            // Assert
            Assert.AreEqual("MappedVal", result);
        }

        [TestMethod]
        public void GetInstSpecialSetup_NonIntValueNotFoundInConfig_ReturnsKeyNotFoundInSetting()
        {
            // Act
            string result = InvokeGetInstSpecialSetup("InstSpecialSetting:Unknown", NewInputData());

            // Assert
            Assert.AreEqual("KeyNotFoundInSetting", result);
        }
    }
}
