using System.IO;

using Automation.GenerateIgxl.EVS;
using Automation.Static;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Evs
{
    [TestClass]
    public class EvsMainTests : FunctionTestBase
    {
        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void EvsMainTest()
        {
            string subName = "EvsMain";
            string outputPath = Path.Combine(OutputPath, "Evs", subName);
            string expectPath = Path.Combine(ExpectPath, "Evs", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            LocalSpecs.TarFolder = outputPath;

            using (var evsMain = new EvsMain())
            {
                evsMain.Execute(null);
                evsMain.Print();
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
