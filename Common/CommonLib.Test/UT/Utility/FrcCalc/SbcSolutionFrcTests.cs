using System.Collections.Generic;

using CommonLib.Utility.FrcCalc;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test.UT.Utility.FrcCalc
{
    [TestClass]
    public class SbcSolutionFrcTests
    {
        [TestMethod]
        public void SbcSolutionFrc_DefaultConstructor_SbcFreqIsZero()
        {
            var solution = new SbcSolutionFrc();
            Assert.AreEqual(0.0, solution.SbcFreq);
        }

        [TestMethod]
        public void SbcSolutionFrc_DefaultConstructor_EngineListIsEmpty()
        {
            var solution = new SbcSolutionFrc();
            Assert.AreEqual(0, solution.EngineList.Count);
        }

        [TestMethod]
        public void SbcSolutionFrc_SetSbcFreq_ReturnsCorrectValue()
        {
            var solution = new SbcSolutionFrc { SbcFreq = 480000000.0 };
            Assert.AreEqual(480000000.0, solution.SbcFreq);
        }

        [TestMethod]
        public void SbcSolutionFrc_EngineListCanAddItems()
        {
            var solution = new SbcSolutionFrc();
            solution.EngineList.Add(new PaEngineItem { TargetFreq = 100000000.0 });
            solution.EngineList.Add(new PaEngineItem { TargetFreq = 200000000.0 });
            Assert.AreEqual(2, solution.EngineList.Count);
        }

        [TestMethod]
        public void SbcSolutionFrc_EngineListCanBeReplaced()
        {
            var solution = new SbcSolutionFrc();
            var newList = new List<PaEngineItem>
            {
                new() { TargetFreq = 300000000.0 }
            };
            solution.EngineList = newList;
            Assert.AreEqual(1, solution.EngineList.Count);
            Assert.AreEqual(300000000.0, solution.EngineList[0].TargetFreq);
        }
    }
}
