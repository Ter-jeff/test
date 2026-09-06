using System;
using System.Collections.Generic;
using System.IO;

using Automation.Utility.Pattern;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Basic
{
    [TestClass]
    public class WalkingZPatsTests : FunctionTestBase
    {
        [TestMethod]
        public void WorkFlowFromContiSheet_ShouldGenerateAtpFile_Correctly()
        {
            string subName = "WalkingZPatsTest";
            string outputPath = Path.Combine(OutputPath, "Basic", subName);
            string expectPath = Path.Combine(ExpectPath, "Basic", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            string outputDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(outputDir);

            var allIOs = new List<string> { "PIN_A", "PIN_B", "PIN_C", "PIN_D" };
            var testIOs = new List<string> { "PIN_A", "PIN_B" };
            string moduleName = "MyModule";
            var diffPairs = new Dictionary<string, string> { { "PIN_A", "PIN_B" } };
            int repeat = 2;

            var walkingZPats = new WalkingZPats(outputDir, repeat);

            // Act
            walkingZPats.WorkFlowFromContiSheet(allIOs, testIOs, moduleName, diffPairs);

            // Assert
            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void Constructor_ShouldCreateOutputDirectory_WhenNotExists()
        {
            // Arrange
            string outputDir = Path.Combine(Path.GetTempPath(), "WalkingZPats_" + Guid.NewGuid());

            // Act
            _ = new WalkingZPats(outputDir, 1);

            // Assert
            Assert.IsTrue(Directory.Exists(outputDir), "Constructor should create output directory if it does not exist.");

            // Cleanup
            Directory.Delete(outputDir, true);
        }
    }
}
