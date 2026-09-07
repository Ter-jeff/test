using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz.GenFlowRow;
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
    public class RtosFlowRowGeneratorTests : FunctionTestBase
    {
        private static RtosFlowRowGenerator _generator = null!;

        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            HardIpParaData paraData = new HardIpParaData(EnumBlock.HardIp);
            HardIpInputData hardIpInputData = new HardIpInputData(paraData);
            _generator = new RtosFlowRowGenerator(hardIpInputData, "sheetName");
        }

        [TestInitialize]
        public void Setup()
        {
            _generator.Pat = new HardIpPattern();
            _generator.LabelVoltage = "LV";
        }

        private static string InvokeCreateUseLimitFailAction(string repeatStr)
        {
            return _generator.CreateUseLimitFailAction(repeatStr);
        }

        #region GetFlagName (via GenRunScenarioRows, since GetFlagName is private)

        [TestMethod]
        public void GenRunScenarioRows_NoScTokenInPattern_BuildsFlagFromDefaultScAndScenario()
        {
            _generator.Pat.Pattern = new PatternClass("PP_TEST");
            _generator.Pat.MiscInfoDict = new Dictionary<string, string> { ["CMD4"] = "3 " };

            _generator.GenRunScenarioRows();

            Assert.AreEqual("F_Rtos_func_SC_SCCMD3", _generator.Pat.Failflag);
        }

        [TestMethod]
        public void GenRunScenarioRows_BuildsTestRowParameterWithRtosPrefix()
        {
            _generator.Pat.Pattern = new PatternClass("PP_TEST");
            _generator.Pat.SubBlock = "RunSC";
            _generator.Pat.FunctionName = "RunScenario";
            List<FlowRow> result = _generator.GenRunScenarioRows();

            Assert.AreEqual("Rtos_RunSC_pp_test_LV", result[2].Parameter);
            Assert.IsTrue(result[2].Enable.StartsWith("Rtos_LV"));
        }

        [TestMethod]
        public void SetBasicInfoByPattern_DerivesBlockNameFromSheetNameSecondToken()
        {
            var pattern = new HardIpPattern { SheetName = "HARDIP_BLK", Pattern = new PatternClass("PP_TEST") };

            _generator.Pat = pattern;

            string blockName = _generator.BlockName;

            Assert.AreEqual("BLK", blockName);
        }

        [TestMethod]
        public void GenRunScenarioRows_ScTokenMatchesScenario_UsesTokenAsIs()
        {
            _generator.Pat.Pattern = new PatternClass("PP_TEST")
            {
                TestPlanPatternName = "SC5_TEST"
            };
            _generator.Pat.MiscInfoDict = new Dictionary<string, string> { ["CMD4"] = "5 " };

            _generator.GenRunScenarioRows();

            Assert.AreEqual("F_Rtos_func_SC5", _generator.Pat.Failflag);
        }

        [TestMethod]
        public void GenRunScenarioRows_ScTokenDiffersFromScenario_CombinesTokenAndScenario()
        {
            _generator.Pat.Pattern = new PatternClass("PP_TEST")
            {
                TestPlanPatternName = "SC7_TEST"
            };
            _generator.Pat.MiscInfoDict = new Dictionary<string, string> { ["CMD4"] = "5 " };

            _generator.GenRunScenarioRows();

            Assert.AreEqual("F_Rtos_func_SC7_SCCMD5", _generator.Pat.Failflag);
        }

        [TestMethod]
        public void GenRunScenarioRows_ExistingFailflag_IsNotOverwritten()
        {
            _generator.Pat.Pattern = new PatternClass("PP_TEST");
            _generator.Pat.Failflag = "F_Preset";

            _generator.GenRunScenarioRows();

            Assert.AreEqual("F_Preset", _generator.Pat.Failflag);
        }

        #endregion

        #region GetScenario / GetParaValue

        [TestMethod]
        public void GetScenario_Cmd4WithDigitsAndTrailingSpace_ReturnsDigits()
        {
            _generator.Pat.MiscInfoDict = new Dictionary<string, string> { ["CMD4"] = "42 " };

            string result = _generator.GetScenario();

            Assert.AreEqual("42", result);
        }

        [TestMethod]
        public void GetScenario_Cmd4Missing_ReturnsEmpty()
        {
            _generator.Pat.MiscInfoDict = [];

            string result = _generator.GetScenario();

            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void GetParaValue_KeyPresentUppercased_ReturnsValue()
        {
            _generator.Pat.MiscInfoDict = new Dictionary<string, string> { ["CMD4"] = "value" };

            string result = _generator.GetParaValue("cmd4");

            Assert.AreEqual("value", result);
        }

        [TestMethod]
        public void GetParaValue_KeyAbsent_ReturnsEmpty()
        {
            _generator.Pat.MiscInfoDict = [];

            string result = _generator.GetParaValue("cmd4");

            Assert.AreEqual("", result);
        }

        #endregion

        #region IdsNoFuse / IsNandPattern / IsSpiPattern

        [TestMethod]
        public void IdsNoFuse_MiscInfoWithoutEfusePowerPin_ReturnsTrue()
        {
            _generator.Pat.MiscInfo = "";

            bool result = _generator.IdsNoFuse;

            Assert.IsTrue(result);
        }

        [TestMethod]
        public void IdsNoFuse_MiscInfoContainsEfusePowerPin_ReturnsFalse()
        {
            _generator.Pat.MiscInfo = "EFUSEPOWER_Pin:PIN1";

            bool result = _generator.IdsNoFuse;

            Assert.IsFalse(result);
        }

        [TestMethod]
        public void IsNandPattern_PatternContainsNanToken_ReturnsTrue()
        {
            _generator.Pat.Pattern = new PatternClass("PP_TEST_nan_TEST");

            bool result = _generator.IsNandPattern;

            Assert.IsTrue(result);
        }

        [TestMethod]
        public void IsNandPattern_PatternWithoutNanToken_ReturnsFalse()
        {
            _generator.Pat.Pattern = new PatternClass("PP_TEST");

            bool result = _generator.IsNandPattern;

            Assert.IsFalse(result);
        }

        [TestMethod]
        public void IsSpiPattern_PatternContainsSpiToken_ReturnsTrue()
        {
            _generator.Pat.Pattern = new PatternClass("PP_TEST_spi_TEST");

            bool result = _generator.IsSpiPattern;

            Assert.IsTrue(result);
        }

        [TestMethod]
        public void IsSpiPattern_PatternWithoutSpiToken_ReturnsFalse()
        {
            _generator.Pat.Pattern = new PatternClass("PP_TEST");

            bool result = _generator.IsSpiPattern;

            Assert.IsFalse(result);
        }

        #endregion

        #region CreateTestFailAction

        [TestMethod]
        public void CreateTestFailAction_IgnorePatBinOutPresent_ReturnsEmpty()
        {
            _generator.Pat.MiscInfoDict = new Dictionary<string, string> { [HardIpConstData.IgnorePatBinOut] = "1" };
            _generator.Pat.Pattern = new PatternClass("PP_TEST");

            string result = _generator.CreateTestFailAction();

            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void CreateTestFailAction_OpcodePatternWithEmptyFailflag_ReturnsEmpty()
        {
            _generator.Pat.Pattern = new PatternClass("PP_TEST") { RealPatternName = "Opcode:Something" };
            _generator.Pat.Failflag = "";

            string result = _generator.CreateTestFailAction();

            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void CreateTestFailAction_FunctionNameSetWithExistingFailflag_ReturnsFailflagAsIs()
        {
            _generator.Pat.Pattern = new PatternClass("PP_TEST");
            _generator.Pat.FunctionName = "SomeFunction";
            _generator.Pat.Failflag = "F_EXISTING_FLAG";

            string result = _generator.CreateTestFailAction();

            Assert.AreEqual("F_EXISTING_FLAG", result);
        }

        [TestMethod]
        public void CreateTestFailAction_FunctionNameSetWithEmptyFailflag_BuildsFlagFromBlockAndSubBlock()
        {
            _generator.Pat.Pattern = new PatternClass("PP_TEST");
            _generator.Pat.FunctionName = "SomeFunction";
            _generator.Pat.Failflag = "";

            string result = _generator.CreateTestFailAction();

            Assert.IsTrue(result.StartsWith("F_"));
            Assert.IsTrue(result.EndsWith("_Flag"));
        }

        #endregion

        #region CreateUseLimitFailAction

        [TestMethod]
        public void CreateUseLimitFailAction_IgnorePatBinOutPresent_ReturnsEmpty()
        {
            _generator.Pat.MiscInfoDict = new Dictionary<string, string> { [HardIpConstData.IgnorePatBinOut] = "1" };

            string result = InvokeCreateUseLimitFailAction("1x");

            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void CreateUseLimitFailAction_IdsNoFuseWithExistingFailflag_ReturnsFailflagAsIs()
        {
            _generator.Pat.MiscInfo = "";
            _generator.Pat.Failflag = "F_EXISTING_FLAG";

            string result = InvokeCreateUseLimitFailAction("1x");

            Assert.AreEqual("F_EXISTING_FLAG", result);
        }

        [TestMethod]
        public void CreateUseLimitFailAction_NotIdsNoFuse_RepeatStrEmpty_ReturnsDefaultMain1x()
        {
            _generator.Pat.MiscInfo = "EFUSEPOWER_Pin:PIN1";

            string result = InvokeCreateUseLimitFailAction("");

            Assert.AreEqual("F_IDS_Current_Main_1x", result);
        }

        [TestMethod]
        public void CreateUseLimitFailAction_NotIdsNoFuse_RepeatStrWithDot_ReplacesDotWithP()
        {
            _generator.Pat.MiscInfo = "EFUSEPOWER_Pin:PIN1";

            string result = InvokeCreateUseLimitFailAction("1.5x");

            Assert.AreEqual("F_IDS_Current_Main_1p5x", result);
        }

        #endregion

        #region CreateTestParameter

        [TestMethod]
        public void CreateTestParameter_TestNameSet_ReturnsTestNamePlusActualLabelVoltage()
        {
            _generator.Pat.Pattern = new PatternClass("PP_TEST");
            _generator.Pat.TestName = "MyTest";
            _generator.Pat.MiscInfo = "";
            _generator.LabelVoltage = "LV";

            string result = _generator.CreateTestParameter(false);

            Assert.AreEqual("MyTest_LV", result);
        }

        #endregion
    }
}
