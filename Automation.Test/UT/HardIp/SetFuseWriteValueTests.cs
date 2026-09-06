using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;

using CommonReaderLib.PatternListCsv;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.VbtLib;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class SetFuseWriteValueTests
    {
        private SetFuseWriteValue _setValue = null!;

        [TestInitialize]
        public void Setup()
        {
            HardIpParaData paraData = new(EnumBlock.HardIp);
            var dummySheet = new HardIpSheet();
            var inputData = new HardIpInputData(paraData);
            _setValue = new SetFuseWriteValue(inputData, dummySheet);
        }

        private static HardIpPattern NewPattern(string miscInfo)
        {
            return new HardIpPattern { SheetName = "Sheet1", MiscInfo = miscInfo, Pattern = new PatternClass("") };
        }

        private static Function NewFunction()
        {
            return new Function
            {
                Parameters = "bankName,fieldName,dictionaryName,flagName,isBinToDecEnable",
                Type = ".NET"
            };
        }

        [TestMethod]
        public void SetArgsListValue_AllMiscInfoKeysPresent_SetsAllParams()
        {
            // Arrange
            Function function = NewFunction();

            // Act
            _setValue.SetArgsListValue(NewPattern("FuseType:BankA;m_catename:CatA;Dict_Store_Code_Name:DictA;Flag_Name:FlagA;Efuse_Binary_Write_Flag:TRUE"), ref function, "");

            // Assert
            Assert.AreEqual("BankA", function.GetParamValue("bankName"));
            Assert.AreEqual("CatA", function.GetParamValue("fieldName"));
            Assert.AreEqual("DictA", function.GetParamValue("dictionaryName"));
            Assert.AreEqual("FlagA", function.GetParamValue("flagName"));
            Assert.AreEqual("TRUE", function.GetParamValue("isBinToDecEnable"));
        }

        [TestMethod]
        public void SetArgsListValue_MissingMiscInfoKeys_SetsEmptyStrings()
        {
            // Arrange
            Function function = NewFunction();

            // Act
            _setValue.SetArgsListValue(NewPattern("Unrelated:Value"), ref function, "");

            // Assert
            Assert.AreEqual("", function.GetParamValue("bankName"));
            Assert.AreEqual("", function.GetParamValue("fieldName"));
            Assert.AreEqual("", function.GetParamValue("dictionaryName"));
            Assert.AreEqual("", function.GetParamValue("flagName"));
            Assert.AreEqual("", function.GetParamValue("isBinToDecEnable"));
        }

        [TestMethod]
        public void SetArgsListValue_NonDotNetFunctionType_NoOp()
        {
            // Arrange
            Function function = NewFunction();
            function.Type = "VBT";

            // Act
            _setValue.SetArgsListValue(NewPattern("FuseType:BankA"), ref function, "");

            // Assert
            Assert.AreEqual("", function.GetParamValue("bankName"));
        }

        [TestMethod]
        public void SetArgsListValue_MiscInfoKeyWithoutColon_TreatedAsMissing()
        {
            // Arrange - GetMiscArgContent requires a colon-separated value to extract
            Function function = NewFunction();

            // Act
            _setValue.SetArgsListValue(NewPattern("FuseType"), ref function, "");

            // Assert
            Assert.AreEqual("", function.GetParamValue("bankName"));
        }
    }
}
