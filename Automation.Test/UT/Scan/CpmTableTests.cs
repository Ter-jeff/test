using System.Collections.Generic;
using System.IO;

using Automation.GenerateIgxl.Scan.CPM;
using Automation.Static;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.CPM;

namespace Automation.Test.UT.Scan
{
    [TestClass]
    public class CpmTableTests : FunctionTestBase
    {

        [TestMethod]
        public void CpmTableTest()
        {
            string subName = "CpmTable";
            string outputPath = Path.Combine(OutputPath, "Scan", subName);
            string expectPath = Path.Combine(ExpectPath, "Scan", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            LocalSpecs.TarFolder = outputPath;

            // Act
            var cpmInstanceFlow = new CpmTable();
            EfuseCpmSheet efuseCpmSheet = new EfuseCpmSheet("EfuseCPMSheet")
            {
                FlagColNumber = new Dictionary<string, int>
                {
                    { "FlagA", 1 },
                    { "FlagB", 2 }
                }
            };

            var row = new EfuseCpmSheetRow
            {
                EfuseBank = "bank_0",
                CpmEfuseName2 = "TestEfuse",
                FlagsValue = new Dictionary<string, string>
                {
                    { "FlagA", "High" },
                    { "FlagB", "Low" }
                }
            };
            efuseCpmSheet.Rows = [row];
            cpmInstanceFlow.WorkFlow(efuseCpmSheet);
            TestProgram.Print();

            // Assert
            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
