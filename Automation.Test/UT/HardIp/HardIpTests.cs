using System.Collections.Generic;
using System.IO;

using Automation.Static;
using Automation.Utility.Pattern;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Newtonsoft.Json;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class HardIpTests : FunctionTestBase
    {
        [TestMethod]
        public void ReadHardIpInfoAllTest()
        {
            string subName = "PatInfo";
            string outputPath = Path.Combine(OutputPath, "HardIp", subName);
            string expectPath = Path.Combine(ExpectPath, "HardIp", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            string file = Path.Combine(KPath, "borneo", "A0_V04A", "HARD_IP", "CSV_HardIP_Info", "HardIP_AutoGen_Info_All_borneo.txt");
            var reader = new UpdateHardIpInfo(LocalSpecs.CurrentProject);
            Dictionary<string, Dictionary<string, string>> hardIpInfoAll = reader.ReadHardIpInfoAll(file);

            string json = JsonConvert.SerializeObject(hardIpInfoAll, Formatting.Indented);
            File.WriteAllText(Path.Combine(outputPath, "hardipInfo.json"), json);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
