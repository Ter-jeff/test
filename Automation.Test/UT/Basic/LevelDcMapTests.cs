using Automation.GenerateIgxl.Basic.Business.GenLevel.BassData;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Basic
{
    [TestClass]
    public class LevelDcMapTests
    {
        [TestMethod]
        public void LevelDcMapTest()
        {
            string levelName = "level1";
            string specName = "spec1";
            LevelDcMap levelDcMap = new LevelDcMap(levelName, specName);

            Assert.AreEqual(levelName, levelDcMap.LevelSheetName);
            Assert.AreEqual(specName, levelDcMap.DcSpecName);
        }
    }
}
