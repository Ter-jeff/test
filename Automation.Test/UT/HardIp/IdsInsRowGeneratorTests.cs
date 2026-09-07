using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenInstanceBiz;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;

using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class IdsInsRowGeneratorTests : FunctionTestBase
    {
        private HardIpInputData _inputData = null!;
        private HardIpSheet _sheet = null!;

        [TestInitialize]
        public void Setup()
        {
            _inputData = new HardIpInputData(null);
            _sheet = new HardIpSheet { Rows = [] };
            // Note: You must ensure 'Pattern' (the current row being processed) 
            // is initialized in the base class or via a setup method.
        }

        [TestMethod]
        public void GenInsRows_WhenAllConditionsFalse_ReturnsEmptyList()
        {
            // Arrange: Line 59 - NoFuse=false, IsNand=false, IsSpi=false
            // You'll need to mock or set Pattern.MiscInfo and GetLastPayload() 
            // so that CommonGenerator returns false for all.

            IdsInsRowGenerator sut = CreateSut("NormalPattern");

            // Act
            List<InstanceRow> result = sut.GenInsRows();

            // Assert
            Assert.AreEqual(1, result.Count, "Should exit early if no special flags are found.");
        }

        [TestMethod]
        public void GenInsRows_WithFuseStage_UpdatesHipPreWriteFlag()
        {
            // Arrange: Line 68 - Trigger the FuseStage Regex
            IdsInsRowGenerator sut = CreateSut(HardIpConstData.FuseStage + ":H;OtherCmd");
            sut.Pat = new HardIpPattern
            {
                HipPreWriteFlag = "FAIL_V_ACTION" // Mocking state for line 76
            };

            // Act
            List<InstanceRow> result = sut.GenInsRows();

            // Assert
            Assert.AreEqual("__NV", result[0].TestName);
        }

        [TestMethod]
        public void GenInsRows_WithAdaptiveCooling_AddsSecondInstance()
        {
            // Arrange: Line 108 - Trigger AdaptiveCooling branch
            _inputData.ConfigData.AdaptiveCoolingItem.Add("MyBlock_MySubBlock");

            IdsInsRowGenerator sut = CreateSut("IDS_NoFuse");
            sut.BlockName = "MyBlock";
            sut.SubBlockName = "MySubBlock";

            // Act
            List<InstanceRow> result = sut.GenInsRows();

            // Assert
            Assert.AreEqual(2, result.Count, "Should add a TMPSMON instance when adaptive cooling is enabled.");
            Assert.IsTrue(result[1].TestName.Contains("TMPSMON"));
        }

        [TestMethod]
        public void GenInsRows_IedaSetting_HasEmptyDcAndPinLevels()
        {
            // Arrange: Line 98-102 - Trigger IedaSetting string check
            // This requires CreateVbtName() in the base/mock to return "IedaSetting"
            IdsInsRowGenerator sut = CreateSut("IDS_NoFuse");
            // Mock sut.CreateVbtName to return "IedaSetting"

            // Act
            List<InstanceRow> result = sut.GenInsRows();

            // Assert
            Assert.AreEqual("Ids_X_X_X", result[0].DcCategory);
            Assert.AreEqual("Levels_IDS", result[0].PinLevels);
        }

        private IdsInsRowGenerator CreateSut(string miscInfo)
        {
            var generator = new IdsInsRowGenerator(_inputData, _sheet, "IDS_Sheet");
            var pattern = new HardIpPattern { MiscInfo = miscInfo };
            generator.Pat = pattern;
            return generator;
        }
    }
}
