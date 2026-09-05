using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class PatternRowTests : FunctionTestBase
    {
        [TestMethod]
        public void Copy()
        {
            // Arrange
            var pattern = new PatternRow
            {
                MiscInfo = "forloop:1,2,3,4;B",
                ForceCondition = new ForceClass(),
                SheetName = "Sheet1",
                RowNum = 1,
                Pattern = new PatternClass("PAT1") { RealPatternName = "PAT1" }
            };

            // Act
            PatternRow copy = pattern.Copy();

            // Assert
            Assert.AreEqual(copy.MiscInfo, pattern.MiscInfo);
        }

        [TestMethod]
        public void GetTestPlanRows()
        {
            // Arrange
            var pattern = new PatternRow
            {
                MiscInfo = "forloop:1,2,3,4;B",
                ForceCondition = new ForceClass(),
                SheetName = "Sheet1",
                RowNum = 1,
                PatChildRows = [new PatSubChildRow() { TpRows = [new()] }],
                Pattern = new PatternClass("PAT1") { RealPatternName = "PAT1" }
            };

            // Act
            List<TestPlanRow> rows = pattern.GetTestPlanRows();

            // Assert
            Assert.AreEqual(1, rows.Count);
        }
    }
}
