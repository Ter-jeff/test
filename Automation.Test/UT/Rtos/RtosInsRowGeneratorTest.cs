using System.Linq;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;
using Automation.Static;

using CommonReaderLib.PatternListCsv;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Rtos
{
    [TestClass]
    public class RtosInsRowGeneratorTest : FunctionTestBase
    {
        [TestMethod]
        public void UpdateIdsMappingTest()
        {
            TestPlanStatic.IdsMappingSheet.Rows = [
                    new() { Stage = "CP1", SubBlock = "IDS_CP_RTOS", Pinname = "VDD_", Efusefieldname = "Vcc" }
                ];
            var hardIpParaData = new HardIpParaData(EnumBlock.Rtos);
            var hardIpInputData = new HardIpInputData(hardIpParaData);
            var generator = new GenerateIgxl.HardIp.AutoGenBusiness.GenInstanceBiz.RtosInsRowGenerator(hardIpInputData, new HardIpSheet(), "");
            generator.UpdateIdsMappingTable("IDS_CP_RTOS", "IDS_IDSCPRTOS_NO_PATT");
            string? result = TestPlanStatic.IdsMappingSheet.Rows.Where(x => x.InstanceName.Length > 0).Select(x => x.InstanceName).FirstOrDefault();
            Assert.AreEqual("IDS_IDSCPRTOS_NO_PATT", result);
        }
    }
}
