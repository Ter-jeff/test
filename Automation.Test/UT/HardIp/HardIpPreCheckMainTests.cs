using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.GenerateIgxl.HardIp.HardIpPreCheck;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;
using Automation.Test.Static;

using CommonLib.Extension;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.HardIpDc.BaseData;
using TestPlanLib.Singleton;
using TestPlanLib.VbtLib;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class HardIpPreCheckMainTests : FunctionTestBase
    {
        private static List<Function> _functions = null!;

        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            _functions = TestSuiteInitialize.Functions;
            LocalSpecs.TtrSummaryFileName = string.Join(",", new List<string>
            {
                Path.Combine(InputPath, "borneo_documents", "A0_V04A", "Borneo_A0_TTR_V04A_X_DigHardIP.xlsx")
            });
        }

        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void HardIpPreCheckMainTest()
        {
            string subName = "HardIpPreCheckMain";
            string outputPath = Path.Combine(OutputPath, "HardIp", subName);
            string expectPath = Path.Combine(ExpectPath, "HardIp", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }

            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            LocalSpecs.TarFolder = outputPath;
            LocalSpecs.HardIpInfoFileName = LocalSpecs.TimeSetFolder = Path.Combine(KPath, "borneo", "HARD_IP", "CSV_HardIP_Info", "HardIP_AutoGen_Info_All_borneo.txt");
            TestProgram.Clear();
            TestProgram.VbtFunctionLib.AddVbtFunctionRange([.. _functions.Where(x => x.Type.EqualsIgnoreCase("VBT"))]);
            BinNumberSingleton.Instance.Initialize(EpWorkbook.TestPlanWorkbook);
            Dictionary<string, HardIpSheet> planDic = [];
            HardIpDcSheet hardIpDcSheet = new HardIpDcSheet();
            new HardIpPreCheckMain().Check(planDic, hardIpDcSheet);

            // Assert
            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
