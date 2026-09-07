using System.Collections.Generic;

using Automation.GenerateIgxl.BistBira.Base;
using Automation.GenerateIgxl.BistBira.NewLogicData;
using Automation.Reader.ConfigFile.NamingRule.Base;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using ScghLib.Base;
using ScghLib.Enums;
using ScghLib.Reader;

using TestPlanLib.Basic;

namespace Automation.Test.UT.BistBira
{
    [TestClass]
    public class BistBurstTests : FunctionTestBase
    {
        [TestMethod]
        public void GenMbistFailBlockTest()
        {
            var mbistSheet = new MbistSheet { SheetName = "TestSheet" };
            var prodFlowSheet = new BistProdFlowSheet
            {
                MbistSheet = mbistSheet,
                Rows =
                [
                    new() { Label = "L1", Pattern = "Pat1", Voltage = "1.0V", FailBranch = "F1", TimeSet = "T1" },
                    new() { Label = "L1", Pattern = "Pat2", Voltage = "1.0V", FailBranch = "F1", TimeSet = "T1" },
                    new() { Label = "L2", Pattern = "Pat3", Voltage = "1.1V", FailBranch = "F2", TimeSet = "T2" }
                ]
            };

            var patternDatas = new Dictionary<string, PatternData>();
            var mbistDataStore = new MbistDataStore { DicPatSets = [] };
            BistNaming naming = new BistNaming(new MbistConfig());
            var target = new BistBurst(naming);
            // Act
            BistProdFlowSheet result = target.BurstLabel(prodFlowSheet, patternDatas, MbistPatSetType.BurstYes, mbistDataStore);

            // Assert 
            Assert.AreEqual(2, result.Rows.Count, "Should have combined first two rows and kept the third.");
            Assert.IsTrue(result.Rows[0].IsPatBurst, "First resulting row should be marked as a burst.");
            CollectionAssert.Contains(result.Rows[0].BurstPatterns, "Pat1");
            CollectionAssert.Contains(result.Rows[0].BurstPatterns, "Pat2");
        }
    }
}
