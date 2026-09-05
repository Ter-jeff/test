using System.IO;
using System.Linq;

using Automation.GenerateIgxl.Scan.CPM;
using Automation.Reader.ConfigFile.NamingRule.Base;
using Automation.Static;
using Automation.Test.Static;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace Automation.Test.UT.Scan
{
    [TestClass]
    public class CpmInstanceMainTests : FunctionTestBase
    {
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            TestProgram.VbtFunctionLib.AddVbtFunctionRange(TestSuiteInitialize.Functions);
        }

        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void CpmInstanceMainTest()
        {
            string subName = "CpmInstanceMain";
            string outputPath = Path.Combine(OutputPath, "Scan", subName);
            string expectPath = Path.Combine(ExpectPath, "Scan", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            LocalSpecs.TarFolder = outputPath;
            LocalSpecs.TestPlanFileName = Path.Combine(InputPath, "borneo_documents", "borneo_B0_TestPlan.xlsx");
            var package = new ExcelPackage(new FileInfo(LocalSpecs.TestPlanFileName));
            EpWorkbook.TestPlanWorkbook = package.Workbook;

            // Act
            ScanConfig config = SettingStatic.ScanConfig;
            var cpmInstanceFlow = new CpmInstanceMain(config);
            cpmInstanceFlow.WorkFlow();
            TestProgram.Print();

            // Assert
            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }

            EpWorkbook.TestPlanWorkbook = null;
        }

        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void CpmInstanceMainVbtTestV()
        {
            string subName = "CpmInstanceMainVbt";
            string outputPath = Path.Combine(OutputPath, "Scan", subName);
            string expectPath = Path.Combine(ExpectPath, "Scan", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            TestProgram.VbtFunctionLib.Clear();
            TestProgram.VbtFunctionLib.AddVbtFunctionRange([.. TestSuiteInitialize.Functions.Where(x => x.Type == "VBT")]);
            LocalSpecs.TarFolder = outputPath;
            LocalSpecs.TestPlanFileName = Path.Combine(InputPath, "borneo_documents", "borneo_B0_TestPlan.xlsx");
            var package = new ExcelPackage(new FileInfo(LocalSpecs.TestPlanFileName));
            EpWorkbook.TestPlanWorkbook = package.Workbook;

            // Act
            ScanConfig config = SettingStatic.ScanConfig;
            var cpmInstanceFlow = new CpmInstanceMain(config);
            cpmInstanceFlow.WorkFlow();
            TestProgram.Print();

            // Assert
            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }

            EpWorkbook.TestPlanWorkbook = null;
        }
    }
}
