using System.Collections.Generic;
using System.IO;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.PostAction.GenIdsLeakagePatternSummary;
using Automation.Static;

using FileDiffLib;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.Basic;

namespace Automation.Test.UT.PostAction
{
    [TestClass]
    public class GenIdsLeakagePatSummaryTests : FunctionTestBase
    {
        [TestMethod]
        public void GenIdsLeakagePatSummaryTest()
        {
            string subName = "GenIdsLeakagePatSummary";
            string outputPath = Path.Combine(OutputPath, "PostAction", subName);
            string expectPath = Path.Combine(ExpectPath, "PostAction", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            LocalSpecs.TarFolder = outputPath;
            string patName = "PAT_IDS_EXAMPLE";
            string patIdsName = "IDS_PAT_IDS_EXAMPLE";
            var allPatterns = new Dictionary<string, PatternData>
            {
                { patName, new PatternData { PatternVersion = "V1" } }
            };
            var hardIpInfo = new List<HardIpInfo>
            {
                new()
                {
                    Payload = patName,
                    Xpins = "IO_PIN1,IO_PIN2",
                    UnusedIoPins = "IO_PIN3"
                }

            };

            var patSetSheet = new PatSetSheet("PatSets_All") { Name = "PatSets_All" };
            var patSetSheetHardIP = new PatSetSheet("PatSets_HardIP") { Name = "PatSets_HardIP" };
            patSetSheet.Rows.Add(new PatSet { PatSetName = patName, PatSetRows = [new() { PatternSet = patName }] });
            patSetSheetHardIP.Rows.Add(new PatSet { PatSetName = patIdsName });

            var patSetSheets = new Dictionary<string, PatSetSheet>
            {
                { "PatSets_All", patSetSheet },{ "patSetSheetHardIP", patSetSheetHardIP }
            };

            // Act
            new GenIdsLeakagePatSummary().WorkFlow(allPatterns, hardIpInfo, patSetSheets);

            // Assert
            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
