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
    public class SetFuseReadValueTests
    {
        private SetFuseReadValue _setValue = null!;

        [TestInitialize]
        public void Setup()
        {
            HardIpParaData paraData = new(EnumBlock.HardIp);
            var dummySheet = new HardIpSheet();
            var inputData = new HardIpInputData(paraData);
            _setValue = new SetFuseReadValue(inputData, dummySheet);
        }

        private static HardIpPattern NewPattern(string miscInfo)
        {
            return new HardIpPattern { SheetName = "Sheet1", MiscInfo = miscInfo, Pattern = new PatternClass("") };
        }

        private static Function NewFunction()
        {
            return new Function
            {
                Parameters = "bankName,fieldName,dictionaryName,sampleSize,enableNonZeroCheck,isDecToBinEnable",
                Type = ".NET"
            };
        }

        [TestMethod]
        public void SetArgsListValue_AllMiscInfoKeysPresent_SetsAllParams()
        {
            // Arrange
            Function function = NewFunction();

            // Act
            _setValue.SetArgsListValue(NewPattern("FuseType:BankA;m_catename:CatA;Dict_Store_Code_Name:DictA;dspwavesize:32;NonZero_Val_Chk:TRUE"), ref function, "");

            // Assert
            Assert.AreEqual("BankA", function.GetParamValue("bankName"));
            Assert.AreEqual("CatA", function.GetParamValue("fieldName"));
            Assert.AreEqual("DictA", function.GetParamValue("dictionaryName"));
            Assert.AreEqual("32", function.GetParamValue("sampleSize"));
            Assert.AreEqual("TRUE", function.GetParamValue("enableNonZeroCheck"));
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
        }

        [TestMethod]
        public void SetArgsListValue_EfuseReadDecFlagTrue_SetsIsDecToBinEnableFalse()
        {
            // Arrange
            Function function = NewFunction();

            // Act
            _setValue.SetArgsListValue(NewPattern("Efuse_Read_Dec_Flag:true"), ref function, "");

            // Assert
            Assert.AreEqual("false", function.GetParamValue("isDecToBinEnable"));
        }

        [TestMethod]
        public void SetArgsListValue_EfuseReadDecFlagFalse_SetsIsDecToBinEnableTrue()
        {
            // Arrange
            Function function = NewFunction();

            // Act
            _setValue.SetArgsListValue(NewPattern("Efuse_Read_Dec_Flag:false"), ref function, "");

            // Assert
            Assert.AreEqual("true", function.GetParamValue("isDecToBinEnable"));
        }

        [TestMethod]
        public void SetArgsListValue_EfuseReadDecFlagOtherValue_DoesNotSetIsDecToBinEnable()
        {
            // Arrange
            Function function = NewFunction();

            // Act
            _setValue.SetArgsListValue(NewPattern("Efuse_Read_Dec_Flag:maybe"), ref function, "");

            // Assert
            Assert.AreEqual("", function.GetParamValue("isDecToBinEnable"));
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
    }
}
