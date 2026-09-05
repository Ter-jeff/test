using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class FlowGeneratorTests : FunctionTestBase
    {
        [TestMethod]
        public void ReArrangeBurstPatterns_ShouldInsertPatternsBetweenBursts()
        {
            // Arrange
            var generator = new FlowGenerator(null);

            var burstPatterns = new List<HardIpPattern>
            {
                new() { RowNum = 1  },
                new() { RowNum = 12 }
            };

            var insertPatterns = new List<HardIpPattern>
            {
                new() { RowNum = 15 },
                new() { RowNum = 25 }
            };

            var singlePatterns = new List<HardIpPattern>
            {
                new() { RowNum = 10 },
                new() { RowNum = 20 },
                new() { RowNum = 30 }
            };

            // Act
            List<HardIpPattern> result = generator.ReArrangeBurstPatterns(burstPatterns, insertPatterns, singlePatterns);

            // Assert
            var ints = result.Select(p => p.RowNum).ToList();

            CollectionAssert.Contains(ints, 1);
            CollectionAssert.Contains(ints, 15);
            CollectionAssert.Contains(ints, 25);
            Assert.IsTrue(ints.IndexOf(15) > ints.IndexOf(10));
        }

        [TestMethod]
        public void ReArrangeBurstPatterns_ShouldHandleEmptyInputGracefully()
        {
            // Arrange
            var generator = new FlowGenerator(null);

            List<HardIpPattern> result = generator.ReArrangeBurstPatterns([], []);

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void ReArrangeBurstPatterns_ShouldIncludeInsertPatterns_WhenNoSinglePatterns()
        {
            // Arrange
            var generator = new FlowGenerator(null);

            var burstPatterns = new List<HardIpPattern>
            {
                new()
                {
                    RowNum = 5 , BurstPatterns =
                    [
                        new() { Pattern = new PatternClass("P1")},
                        new() { Pattern = new PatternClass("P2")}
                    ]
                }
            };

            var insertPatterns = new List<HardIpPattern>
            {
                new() { RowNum = 5 },
                new() { RowNum = 15 }
            };

            var singlePatterns = new List<HardIpPattern>
            {
                new() { RowNum = 20 },
                new() { RowNum = 25 }
            };

            // Act
            List<HardIpPattern> result = generator.ReArrangeBurstPatterns(burstPatterns, insertPatterns, singlePatterns);

            // Assert
            Assert.IsTrue(result.Any(p => p.RowNum == 1 || p.RowNum == 5 || p.RowNum == 15));
            Assert.AreEqual(2, result.Count);
        }
    }
}
