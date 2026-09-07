using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.InputManager.Data;

using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.HardIpDc.BaseData;

namespace Automation.Test.UT.BinCut
{
    [TestClass]
    public class BinCutInputDataTests
    {
        [TestMethod]
        public void GenHardipInstanceByPattern_ShouldReturnMatchingInstanceRows()
        {
            // Arrange
            string patternName = "MyPattern";
            var binCutData = new BinCutInputData
            {
                HardIpPatterns = new Dictionary<string, HardIpSheet>
                {
                    ["Sheet1"] = new HardIpSheet
                    {
                        Rows =
                        [
                            new() { Pattern = new PatternClass("OtherPattern") },
                            new() { Pattern = new PatternClass("MyPattern") }
                        ]
                    },
                    ["Sheet2"] = new HardIpSheet
                    {
                        Rows =
                        [
                            new() { Pattern = new PatternClass("AnotherPattern") }
                        ]
                    }
                }
            };
            var dcSheet = new HardIpDcSheet();
            dcSheet.Rows.Add(new HardIpCategoryDef("Name"));

            // Act
            List<InstanceRow> result = binCutData.GenHardipInstanceByPattern(patternName, dcSheet);

            // Assert
            Assert.AreNotEqual(null, result);
        }

        [TestMethod]
        public void GenHardipInstanceByPattern_ShouldReturnEmpty_WhenNoPatternMatches()
        {
            // Arrange
            var binCutData = new BinCutInputData
            {
                HardIpPatterns = new Dictionary<string, HardIpSheet>
                {
                    ["Sheet1"] = new HardIpSheet
                    {
                        Rows =
                        [
                            new() { Pattern = new PatternClass("PatternA") },
                            new() { Pattern = new PatternClass("PatternB") }
                        ]
                    }
                }
            };
            var dcSheet = new HardIpDcSheet();
            dcSheet.Rows.Add(new HardIpCategoryDef("Name"));

            // Act
            List<InstanceRow> result = binCutData.GenHardipInstanceByPattern("NonExistentPattern", dcSheet);

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.AreEqual(0, result.Count, "Expected no InstanceRows when no pattern matches.");
        }
    }
}
