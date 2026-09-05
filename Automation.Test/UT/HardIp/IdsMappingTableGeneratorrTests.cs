using System.IO;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness;
using Automation.Static;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.Basic;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class IdsMappingTableGeneratorrTests : FunctionTestBase
    {
        [TestMethod]
        public void WorkFlowTest()
        {
            string subName = "IdsMappingTableGeneratorWorkFlow";
            string outputPath = Path.Combine(OutputPath, "HardIp", subName);
            string expectPath = Path.Combine(ExpectPath, "HardIp", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            LocalSpecs.TarFolder = outputPath;
            IdsMappingSheet idsMappingSheet = TestPlanStatic.IdsMappingSheet;
            idsMappingSheet.IndexSubBlock = 999;

            // Act
            var gen = new IdsMappingTableGenerator(idsMappingSheet);
            gen.WorkFlow();

            // Assert
            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
