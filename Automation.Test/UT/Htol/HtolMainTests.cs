using System.IO;

using Automation.GenerateIgxl.HTOL;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

using ProjectConfigLib.ProjectConfig;

using TestPlanLib.Singleton;

namespace Automation.Test.UT.Htol
{
    [TestClass]
    public class HtolMainTests
    {
        protected static readonly string InputPath = Path.Combine(Directory.GetCurrentDirectory(), "Input");
        protected static readonly string OutputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        protected static readonly string ExpectPath = Path.Combine(Directory.GetCurrentDirectory(), "Expected");
        protected static readonly string KPath = Path.Combine(Directory.GetCurrentDirectory(), "K");

        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            MockService.Mock();
            AllStatic.Clear();
            LocalSpecs.IsUnitTest = true;
            LocalSpecs.TestPlanFileName = Path.Combine(InputPath, "hidra_documents", "Hidra_Ext_ATE_TestPlan_649.xlsx");
            LocalSpecs.PatternFolder = Path.Combine(KPath, "Hidra", "patx");
            MultiTestSettingSheetsSingleton.Initialize(EpWorkbook.TestPlanWorkbook, "Hidra", LocalSpecs.AllJobs, EnumDevice.AP);
            BinNumberSingletonLegacy.Initialize(LocalSpecs.ConfigFiles.BinNumberConfig, LocalSpecs.CurrentProject);
            ProjectConfigSingleton.Instance().LoadProjectConfig(Path.Combine(InputPath, "hidra_documents", "ProjectConfig_hidra.ini"));
            LocalSpecs.Options = new Options(ProjectConfigSingleton.Instance());
        }

        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void HtolMainTest()
        {
            string subName = "HtolMain";
            string outputPath = Path.Combine(OutputPath, "Htol", subName);
            string expectPath = Path.Combine(ExpectPath, "Htol", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            LocalSpecs.TarFolder = outputPath;
            using (var htolMain = new HtolMain())
            {
                htolMain.Execute(null);
                htolMain.Print();
            }

            // Assert
            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
