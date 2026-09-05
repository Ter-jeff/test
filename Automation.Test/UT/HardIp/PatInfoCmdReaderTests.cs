using System.IO;

using Automation.GenerateIgxl.Basic.Business.GenPatSet.Business;
using Automation.Reader;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Newtonsoft.Json;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class PatInfoCmdReaderTests : FunctionTestBase
    {
        [TestMethod]
        public void PatInfoCmdReaderTest()
        {
            AssertOnlyWindowsOS("requires patinfo.exe via IGXLROOT environment variable");
            string subName = "PatInfoCmdReader";
            string inputPath = Path.Combine(InputPath, "HardIp");
            string outputPath = Path.Combine(OutputPath, "HardIp", subName);
            string expectPath = Path.Combine(ExpectPath, "HardIp", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            string atpContent = "";
            string value = Path.Combine(inputPath, "CZ_BRNA0_C_FULP_AN_AA00_DLL_JTG_VIX_ALLFRV_SI_CPLLDS_T6PD_1_A0_2504240522.pat.gz");
            _ = new PatInfoCmd().ConvertByArgs(value, ref atpContent, "-hdr -switches");
            PatternResult patternResult = new PatternResult();
            new PatInfoCmdReader().Read(ref patternResult, [.. atpContent.Split('\n')]);

            string json = JsonConvert.SerializeObject(patternResult.ModuleNameList, Formatting.Indented);
            File.WriteAllText(Path.Combine(outputPath, "result.json"), json);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
