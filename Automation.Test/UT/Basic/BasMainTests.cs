using System.IO;

using Automation.GenerateIgxl.PreAction.ReadBasLib;
using Automation.Static;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Basic
{
    [TestClass]
    public class BasMainTests : FunctionTestBase
    {
        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void BasMainTest()
        {
            string subName = "BasMain";
            string inputPath = Path.Combine(InputPath, "Basic");
            string outputPath = Path.Combine(OutputPath, "Basic", subName);
            string expectPath = Path.Combine(ExpectPath, "Basic", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            LocalSpecs.BasLibraryFolder = inputPath;
            LocalSpecs.TarFolder = outputPath;
            var basMain = new BasMain();
            basMain.WorkFlow(outputPath);
            string[] files = Directory.GetFiles(outputPath);
            foreach (string file in files)
            {
                basMain.AddComment(file);
            }

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
