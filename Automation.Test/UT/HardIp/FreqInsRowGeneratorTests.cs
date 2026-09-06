using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenInstanceBiz;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.Singleton;

using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class FreqInsRowGeneratorTests : FunctionTestBase
    {
        private HardIpInputData _inputData = null!;
        private HardIpSheet _sheet = null!;

        [TestInitialize]
        public void Setup()
        {
            _inputData = new HardIpInputData(null);
            _sheet = new HardIpSheet { Rows = [] };
            string timeSet = "Valid_TS";
            string expectedDomain = "DOM1";
            List<ComTimeSetBasicSheet> timeSetSheets =
            [
                new("SheetName")
                {
                    Name = timeSet,
                    TimeDomain = expectedDomain
                }
            ];
            AcTSetCategoryMapSingleton.Instance().SetMultiTimeSetSheet(timeSetSheets);
        }

        private FreqInsRowGenerator CreateSut(string miscInfo)
        {
            var generator = new FreqInsRowGenerator(_inputData, _sheet, "IDS_Sheet");
            var pattern = new HardIpPattern { MiscInfo = miscInfo };
            generator.Pat = pattern;
            return generator;
        }

        [TestMethod]
        public void CreateHardIpLevelConcurrent_TimeSetNotFound_ReturnsTBDError()
        {
            // Arrange
            string timeSet = "Unknown_TimeSet";
            string result = CreateSut("").CreateHardIpLevelConcurrent(timeSet);

            // Assert
            Assert.AreEqual("TBD(ConcurrentLevelError)", result);

            AcTSetCategoryMapSingleton.Instance().SetMultiTimeSetSheet([]);
        }

        [TestMethod]
        public void CreateHardIpLevelConcurrent_TimeSetNotFound_ReturnsTBDError_1()
        {
            // Arrange
            string timeSet = "Valid_TS";
            string result = CreateSut("").CreateHardIpLevelConcurrent(timeSet);

            // Assert
            Assert.AreEqual("Levels_Con_H_DOM1", result);

            AcTSetCategoryMapSingleton.Instance().SetMultiTimeSetSheet([]);
        }
    }
}
