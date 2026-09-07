using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;
using Automation.Static;

using CommonReaderLib.PatternListCsv;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.Basic;
using TestPlanLib.VbtLib;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class SetIdsMathFuncValueTests
    {
        private SetIdsMathFuncValue _setValue = null!;

        [TestInitialize]
        public void Setup()
        {
            HardIpParaData paraData = new(EnumBlock.HardIp);
            var dummySheet = new HardIpSheet();
            var inputData = new HardIpInputData(paraData);
            _setValue = new SetIdsMathFuncValue(inputData, dummySheet);
            SetIdsMappingSheet(new IdsMappingSheet("IDS_Mapping_Table"));
        }

        [TestCleanup]
        public void Cleanup()
        {
            SetIdsMappingSheet(new IdsMappingSheet("IDS_Mapping_Table"));
        }

        private static void SetIdsMappingSheet(IdsMappingSheet sheet)
        {
            TestPlanStatic._idsMappingSheet = sheet;
        }

        private static HardIpPattern NewPattern(string miscInfo)
        {
            return new HardIpPattern { SheetName = "Sheet1", MiscInfo = miscInfo, Pattern = new PatternClass("") };
        }

        private static Function NewFunction()
        {
            return new Function { Parameters = "firstCalculatedData,secondCalculatedData,operationSymbol" };
        }

        [TestMethod]
        public void SetForDelta_BothCurrentDicsAndPowerRailPresent_SetsFirstAndSecondCalculatedData()
        {
            // Arrange
            Function function = NewFunction();

            // Act
            _setValue.SetArgsListValue(NewPattern("Func:IDS_Delta_calc;Current1_dic:C1;Current2_dic:C2;power_rail:PR"), ref function, "");

            // Assert
            Assert.AreEqual("C1:PR", function.GetParamValue("firstCalculatedData"));
            Assert.AreEqual("C2:PR", function.GetParamValue("secondCalculatedData"));
            Assert.AreEqual("-", function.GetParamValue("operationSymbol"));
        }

        [TestMethod]
        public void SetForDelta_MissingPowerRail_DoesNotSetCalculatedData()
        {
            // Arrange
            Function function = NewFunction();

            // Act
            _setValue.SetArgsListValue(NewPattern("Func:IDS_Delta_calc;Current1_dic:C1;Current2_dic:C2"), ref function, "");

            // Assert
            Assert.AreEqual("", function.GetParamValue("firstCalculatedData"));
            Assert.AreEqual("", function.GetParamValue("secondCalculatedData"));
            Assert.AreEqual("-", function.GetParamValue("operationSymbol"));
        }

        [TestMethod]
        public void SetForDelta_MissingCurrent1Dic_OnlySecondCalculatedDataSet()
        {
            // Arrange
            Function function = NewFunction();

            // Act
            _setValue.SetArgsListValue(NewPattern("Func:IDS_Delta_calc;Current2_dic:C2;power_rail:PR"), ref function, "");

            // Assert
            Assert.AreEqual("", function.GetParamValue("firstCalculatedData"));
            Assert.AreEqual("C2:PR", function.GetParamValue("secondCalculatedData"));
        }

        [TestMethod]
        public void SetForMath_MathFunctionAndCalcPin_SetsOperationSymbolAndMeasStrAsFirst()
        {
            // Arrange
            Function function = NewFunction();

            // Act
            _setValue.SetArgsListValue(NewPattern("Func:IDS_main_current_MathFunc;MathFunction:MAX;Calc_Pin:PIN1"), ref function, "");

            // Assert
            Assert.AreEqual("MAX", function.GetParamValue("operationSymbol"));
            Assert.AreEqual("ON:PIN1", function.GetParamValue("firstCalculatedData"));
            Assert.AreEqual("", function.GetParamValue("secondCalculatedData"));
        }

        [TestMethod]
        public void SetForMath_MathInverseTrue_SwapsFirstAndSecondCalculatedData()
        {
            // Arrange
            Function function = NewFunction();

            // Act
            _setValue.SetArgsListValue(NewPattern("Func:IDS_main_current_MathFunc;Calc_Pin:PIN1;MathInverse:true"), ref function, "");

            // Assert - with inverse, fuseStr (empty here) goes first and measStr goes second
            Assert.AreEqual("", function.GetParamValue("firstCalculatedData"));
            Assert.AreEqual("ON:PIN1", function.GetParamValue("secondCalculatedData"));
        }

        [TestMethod]
        public void SetForMath_FusedStageWithMatchingMappingRow_BuildsFuseStrFromTable()
        {
            // Arrange
            SetIdsMappingSheet(new IdsMappingSheet("IDS_Mapping_Table")
            {
                Rows = { new IdsMappingRow { Stage = "STG1", Pinname = "PINA", Efusefieldname = "EFUSE1" } }
            });
            Function function = NewFunction();

            // Act
            _setValue.SetArgsListValue(NewPattern("Func:IDS_main_current_MathFunc;Calc_Pin:PINA;FusedStage:STG1"), ref function, "");

            // Assert
            Assert.AreEqual("STG1:EFUSE1", function.GetParamValue("secondCalculatedData"));
        }

        [TestMethod]
        public void SetForMath_FusedStageWithCalcIdsSumMatch_UsesCalcIdsSumValueDirectly()
        {
            // Arrange - the Calc_IDS_Sum override should be used instead of the mapping table
            Function function = NewFunction();

            // Act
            _setValue.SetArgsListValue(NewPattern("Func:IDS_main_current_MathFunc;Calc_Pin:PINA;FusedStage:STG1;Calc_IDS_Sum(EFUSE@PINA#EFUSE_SUM)"), ref function, "");

            // Assert
            Assert.AreEqual("STG1:EFUSE_SUM", function.GetParamValue("secondCalculatedData"));
        }

        [TestMethod]
        public void SetForMath_FusedStageNoMatch_ProducesEmptyFuseList()
        {
            // Arrange - no mapping row and no Calc_IDS_Sum override for this pin
            Function function = NewFunction();

            // Act
            _setValue.SetArgsListValue(NewPattern("Func:IDS_main_current_MathFunc;Calc_Pin:PINZ;FusedStage:STG1"), ref function, "");

            // Assert
            Assert.AreEqual("STG1:", function.GetParamValue("secondCalculatedData"));
        }

        [TestMethod]
        public void SetArgsListValue_NeitherFuncPresent_NoOp()
        {
            // Arrange
            Function function = NewFunction();

            // Act
            _setValue.SetArgsListValue(NewPattern("SomeOtherInfo:X"), ref function, "");

            // Assert
            Assert.AreEqual("", function.GetParamValue("firstCalculatedData"));
            Assert.AreEqual("", function.GetParamValue("secondCalculatedData"));
            Assert.AreEqual("", function.GetParamValue("operationSymbol"));
        }
    }
}
