using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.GenerateIgxl.HardIp.InstanceParameterSetting;
using Automation.InputManager.Data;
using Automation.Static;
using Automation.Test.Static;

using CommonLib.Enums;
using CommonLib.Extension;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Newtonsoft.Json;

using TestPlanLib.Static;
using TestPlanLib.VbtLib;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class SetArgsValueTests : FunctionTestBase
    {
        private static HardIpInputData _hardIpInputData = null!;
        private static HardIpSheet _hardIpSheet = null!;
        private static List<Function> _functions = null!;

        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            _hardIpInputData = new HardIpInputData(null);
            _hardIpSheet = new HardIpSheet
            {
                SheetName = "HardIp_1",
                PlanHeaderIdx =
                {
                    ["registerIndex"] = 5
                }
            };
            _functions = TestSuiteInitialize.Functions;
        }

        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void SetArgsValue_VBT()
        {
            string subName = "SetArgsValue_Vbt";
            string outputPath = Path.Combine(OutputPath, "HardIp", subName);
            string expectPath = Path.Combine(ExpectPath, "HardIp", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            TestProgram.Clear();
            TestProgram.VbtFunctionLib.AddVbtFunctionRange([.. _functions.Where(x => x.Type.EqualsIgnoreCase("VBT"))]);
            LocalSpecs.Options.Device = EnumDevice.AP;

            var functions = new List<Function>();
            var functionNames = new List<string>
            {
                VbtFunctionLibShared.FunctionalName,
                VbtFunctionLibShared.VifName,
                VbtFunctionLibShared.VifName,
                VbtFunctionLibShared.HardIpmtdTest,
                VbtFunctionLibShared.VirGpioTtrName,
                VbtFunctionLibShared.VirName,
                VbtFunctionLibShared.Ids,
                VbtFunctionLibShared.IdsMathFunc,
                VbtFunctionLibShared.LcdTrim,
                VbtFunctionLibShared.DvdcTrim,
                VbtFunctionLibShared.DvdcTrim3D,
                VbtFunctionLibShared.LcdMeas,
                VbtFunctionLibShared.RfTrim,
                VbtFunctionLibShared.RfTrim2D,
                VbtFunctionLibShared.RfHtolFunc,
                VbtFunctionLibShared.RfFunc,
                "HIP_eFuse_Read",
                "HardIPFuseRead",
                "HIP_eFuse_Write",
                "HardIPFuseWrite",
                "Unknown"
            };
            foreach (string functionName in functionNames)
            {
                // Arrange
                Function function = CommonGenerator.GetVbtFunctionBase(functionName);
                var pattern = new HardIpPattern
                {
                    Pattern = new PatternClass("P1")
                    {
                        PatternSetList = [["P1", "P2"]],
                        InstancePayloadName = ["P1", "P2"]
                    },
                    SheetName = "HardIP_1"
                };
                string voltage = "NV";
                function.ArgList[50] = new string('X', 6000);

                // Act
                new SetArgValueMain(_hardIpInputData, _hardIpSheet).SetArgsValue(pattern, ref function, voltage);
                functions.Add(function);
            }

            // Assert
            functions.ForEach(x => x.FileName = "");
            string json = JsonConvert.SerializeObject(functions, Formatting.Indented);
            File.WriteAllText(Path.Combine(outputPath, "result.json"), json);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void SetArgsValue_NET()
        {
            string subName = "SetArgsValue_Net";
            string outputPath = Path.Combine(OutputPath, "HardIp", subName);
            string expectPath = Path.Combine(ExpectPath, "HardIp", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            TestProgram.Clear();
            TestProgram.VbtFunctionLib.AddVbtFunctionRange([.. _functions.Where(x => x.Type.EqualsIgnoreCase(".NET"))]);
            LocalSpecs.Options.Device = EnumDevice.AP;

            var functions = new List<Function>();
            var functionNames = new List<string>
            {
                VbtFunctionLibShared.FunctionalName,
                VbtFunctionLibShared.VifName,
                VbtFunctionLibShared.VifName,
                VbtFunctionLibShared.HardIpmtdTest,
                VbtFunctionLibShared.VirGpioTtrName,
                VbtFunctionLibShared.VirName,
                VbtFunctionLibShared.Ids,
                VbtFunctionLibShared.IdsMathFunc,
                VbtFunctionLibShared.LcdTrim,
                VbtFunctionLibShared.DvdcTrim,
                VbtFunctionLibShared.DvdcTrim3D,
                VbtFunctionLibShared.LcdMeas,
                VbtFunctionLibShared.RfTrim,
                VbtFunctionLibShared.RfTrim2D,
                VbtFunctionLibShared.RfHtolFunc,
                VbtFunctionLibShared.RfFunc,
                "HIP_eFuse_Read",
                "HardIPFuseRead",
                "HIP_eFuse_Write",
                "HardIPFuseWrite",
                "Unknown"
            };
            foreach (string functionName in functionNames)
            {
                // Arrange
                Function function = CommonGenerator.GetVbtFunctionBase(functionName);
                var pattern = new HardIpPattern
                {
                    Pattern = new PatternClass("P1")
                    {
                        PatternSetList = [["P1", "P2"]],
                        InstancePayloadName = ["P1", "P2"]
                    },
                    SheetName = "HardIP_1"
                };
                string voltage = "NV";
                function.ArgList[50] = new string('X', 6000);

                // Act
                new SetArgValueMain(_hardIpInputData, _hardIpSheet).SetArgsValue(pattern, ref function, voltage);
                functions.Add(function);
            }

            // Assert
            functions.ForEach(x => x.FileName = "");
            string json = JsonConvert.SerializeObject(functions, Formatting.Indented);
            File.WriteAllText(Path.Combine(outputPath, "result.json"), json);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void CheckArgsExceedLimitationLcdTest()
        {
            string subName = "CheckArgsExceedLimitationLcd";
            string outputPath = Path.Combine(OutputPath, "HardIp", subName);
            string expectPath = Path.Combine(ExpectPath, "HardIp", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            TestProgram.Clear();
            TestProgram.VbtFunctionLib.AddVbtFunctionRange([.. _functions.Where(x => x.Type.EqualsIgnoreCase(".NET"))]);
            LocalSpecs.Options.Device = EnumDevice.LCD;

            var functions = new List<Function>();
            var functionNames = new List<string>
            {
                VbtFunctionLibShared.FunctionalName,
                VbtFunctionLibShared.VifName,
                VbtFunctionLibShared.VifName,
                VbtFunctionLibShared.HardIpmtdTest,
                VbtFunctionLibShared.VirGpioTtrName,
                VbtFunctionLibShared.VirName,
                VbtFunctionLibShared.Ids,
                VbtFunctionLibShared.IdsMathFunc,
                VbtFunctionLibShared.LcdTrim,
                VbtFunctionLibShared.DvdcTrim,
                VbtFunctionLibShared.DvdcTrim3D,
                VbtFunctionLibShared.LcdMeas,
                VbtFunctionLibShared.RfTrim,
                VbtFunctionLibShared.RfTrim2D,
                VbtFunctionLibShared.RfHtolFunc,
                VbtFunctionLibShared.RfFunc,
                "HIP_eFuse_Read",
                "HardIPFuseRead",
                "HIP_eFuse_Write",
                "HardIPFuseWrite",
                "Unknown"
            };
            foreach (string functionName in functionNames)
            {
                // Arrange
                Function function = CommonGenerator.GetVbtFunctionBase(functionName);
                var pattern = new HardIpPattern
                {
                    Pattern = new PatternClass("P1")
                    {
                        PatternSetList = [["P1", "P2"]],
                        InstancePayloadName = ["P1", "P2"]
                    },
                    SheetName = "HardIP_1"
                };
                string voltage = "NV";
                function.ArgList[50] = new string('X', 6000);
                // Act
                new SetArgValueMain(_hardIpInputData, _hardIpSheet).SetArgsValue(pattern, ref function, voltage);
                functions.Add(function);
            }

            // Assert
            functions.ForEach(x => x.FileName = "");
            string json = JsonConvert.SerializeObject(functions, Formatting.Indented);
            File.WriteAllText(Path.Combine(outputPath, "result.json"), json);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
