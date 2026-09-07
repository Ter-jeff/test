using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class ParseTestPlanByConditionTests : FunctionTestBase
    {
        private ParseTestPlanByCondition _parser = null!;

        [TestInitialize]
        public void Setup()
        {
            _parser = new ParseTestPlanByCondition();
        }

        [TestMethod]
        public void ParseTestPlanPatternByCondition_MultipleConditions_SplitsIntoMultiplePatterns()
        {
            // Arrange
            var pin = new ForcePin { PinName = "VDD", ForceValue = "1.0" };
            var conditions = new List<ForceCondition>
            {
                new() { ForcePins = [pin] },
                new() { ForcePins = [pin] }
            };

            var pattern = new HardIpPattern
            {
                ForceConditionList = conditions
            };

            var sheet = new HardIpSheet { Rows = [pattern] };
            var testPlanDic = new Dictionary<string, HardIpSheet> { { "Sheet1", sheet } };

            // Act
            _parser.ParseTestPlanPatternByCondition(testPlanDic);

            // Assert
            List<HardIpPattern> resultRows = testPlanDic["Sheet1"].Rows;
            Assert.AreEqual(2, resultRows.Count, "The pattern should be split into 2 rows.");
            Assert.AreEqual(1, resultRows[0].ConditionIndex);
            Assert.AreEqual(2, resultRows[1].ConditionIndex);
            Assert.IsTrue(resultRows[0].ForceVoltageFlag.Contains("VDD1p0"), "Flag should be formatted correctly.");
        }

        [TestMethod]
        public void ParseTestPlanPatternByCondition_SingleCondition_DoesNotModifyList()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                ForceConditionList = [new()]
            };
            var sheet = new HardIpSheet { Rows = [pattern] };
            var testPlanDic = new Dictionary<string, HardIpSheet> { { "Sheet1", sheet } };

            // Act
            _parser.ParseTestPlanPatternByCondition(testPlanDic);

            // Assert
            Assert.AreEqual(1, testPlanDic["Sheet1"].Rows.Count, "List should remain unchanged.");
        }
    }
}
