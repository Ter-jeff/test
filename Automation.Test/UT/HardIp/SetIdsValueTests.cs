using System.IO;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo;
using Automation.InputManager.Data;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Newtonsoft.Json;

using TestPlanLib.VbtLib;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class SetIdsValueTests : FunctionTestBase
    {
        [TestMethod]
        public void SetArgsListValue()
        {
            string subName = "SetIdsValue_SetArgsListValue";
            string outputPath = Path.Combine(OutputPath, "HardIp", subName);
            string expectPath = Path.Combine(ExpectPath, "HardIp", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            var pattern = new HardIpPattern
            {
                Pattern = new PatternClass("OtherPattern") { },
                MeasPins =
                [
                    new()
                    {
                        MeasType = MeasType.MeasC,
                        CapBit = "8",
                        MeasLimitsN = [new("CP1") { LoLimit = "lo", HiLimit = "hi" }]
                    }
                ]
            };
            var function = new Function
            {
                Parameters = "DigSrc_Assignment",
                ArgList = [new('A', 6001), "A,B"],
                FunctionName = "VBT"
            };
            var dummySheet = new HardIpSheet();
            var inputData = new HardIpInputData(null) { HardIpRegAssigns = [] };
            new SetIdsValue(inputData, dummySheet).SetArgsListValue(pattern, ref function, "NV");

            string json = JsonConvert.SerializeObject(function, Formatting.Indented);
            File.WriteAllText(Path.Combine(outputPath, "hardipInfo.json"), json);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void SetArgsListValue_1()
        {
            string subName = "SetIdsValue_SetArgsListValue1";
            string outputPath = Path.Combine(OutputPath, "HardIp", subName);
            string expectPath = Path.Combine(ExpectPath, "HardIp", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            var pattern = new HardIpPattern
            {
                Pattern = new PatternClass("OtherPattern") { },
                MeasPins =
                [
                    new()
                    {
                        MeasType = MeasType.MeasC,
                        CapBit = "8",
                        MeasLimitsN = [new("CP1") { LoLimit = "lo", HiLimit = "hi" }]
                    }
                ]
            };
            var function = new Function
            {
                Parameters = "DigSrc_Assignment",
                ArgList = [new('A', 6001), "A,B"],
                FunctionName = ".NET"
            };
            var dummySheet = new HardIpSheet();
            var inputData = new HardIpInputData(null) { HardIpRegAssigns = [] };
            new SetIdsValue(inputData, dummySheet).SetArgsListValue(pattern, ref function, "NV");

            string json = JsonConvert.SerializeObject(function, Formatting.Indented);
            File.WriteAllText(Path.Combine(outputPath, "hardipInfo.json"), json);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
