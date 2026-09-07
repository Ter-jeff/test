using System.IO;

using Automation.GenerateIgxl.BistBira;
using Automation.Static;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.BistBira
{
    [TestClass]
    public class BistBiraMainTests : FunctionTestBase
    {
        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void GenMbistFailBlockTest()
        {
            string subName = "MbistFailBlock";
            string outputPath = Path.Combine(OutputPath, "BistBira", subName);
            string expectPath = Path.Combine(ExpectPath, "BistBira", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            LocalSpecs.TarFolder = outputPath;

            string csv = Path.Combine(InputPath, "borneo_documents", "A0_V04A", "borneo_A0_pattern_dashboard_20251002_1524_15.csv");
            string hardipInfo = Path.Combine(KPath, "borneo", "HARD_IP", "CSV_HardIP_Info", "HardIP_AutoGen_Info_All_borneo.txt");
            string bistInfo = Path.Combine(KPath, "borneo", "HARD_IP", "CSV_Bist_Info", "borneo_Bist_Info_All.txt");
            var bistBiraMain = new BistBiraMain();
            bistBiraMain.GenMbistFailBlock(csv, hardipInfo, bistInfo);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
