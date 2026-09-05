using System.Collections.Generic;
using System.IO;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class InterposeAssignGeneratorTests : TestBase
    {
        [TestMethod]
        public void InterposeAssignGeneratorTest()
        {
            string subName = "InterposeAssignGenerator";
            string outputPath = Path.Combine(OutputPath, "HardIp", subName);
            string expectPath = Path.Combine(ExpectPath, "HardIp", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            LocalSpecs.TarFolder = outputPath;
            var interposeAssigns = new List<InterposeAssign>
            {
                new()
                {
                    BlockName = "BlockA",
                    AssignName = "AssignA",
                    Type = InterposeAssignType.InterposePrePat,
                    InterposeAssignList = ["INT1", "INT2"]
                },
                new()
                {
                    BlockName = "BlockB",
                    AssignName = "AssignB",
                    Type = InterposeAssignType.InterposePreInit,
                    InterposeAssignList = ["INT3"]
                }
            };
            new InterposeAssignGenerator().WorkFlow(interposeAssigns);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
