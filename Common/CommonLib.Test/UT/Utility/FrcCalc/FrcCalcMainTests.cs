using System.Collections.Generic;

using CommonLib.Utility.FrcCalc;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test.UT.Utility.FrcCalc
{
    [TestClass]
    public class FrcCalcMainTests
    {
        [TestMethod]
        public void CalculateFrcFreq_SingleFreq_ReturnsNonEmptyList()
        {
            List<double> result = FrcCalcMain.CalculateFrcFreq([400000000]);
            Assert.IsNotNull(result);
            Assert.IsTrue(result.Count > 0);
        }

        [TestMethod]
        public void CalculateFrcFreq_SingleFreq_ResultIsSorted()
        {
            List<double> result = FrcCalcMain.CalculateFrcFreq([400000000]);

            for (int i = 1; i < result.Count; i++)
            {
                Assert.IsTrue(result[i] >= result[i - 1], "Result should be sorted ascending");
            }
        }

        [TestMethod]
        public void CalculateFrcFreq_SingleFreq_AllResultsPositive()
        {
            List<double> result = FrcCalcMain.CalculateFrcFreq([400000000]);

            foreach (double freq in result)
            {
                Assert.IsTrue(freq > 0, "All SBC frequencies should be positive");
            }
        }

        [TestMethod]
        public void CalculateFrcFreq_MultipleFreqs_ReturnsIntersectionSolutions()
        {
            // Two frequencies that share compatible PLL input frequencies
            List<double> result = FrcCalcMain.CalculateFrcFreq([200000000, 400000000]);

            Assert.IsNotNull(result);
            // Result may be empty if no shared solution, but should not throw
        }

        [TestMethod]
        public void CalculateFrcFreq_SameFreqTwice_ReturnsSameAsOnce()
        {
            List<double> single = FrcCalcMain.CalculateFrcFreq([400000000]);
            List<double> doubled = FrcCalcMain.CalculateFrcFreq([400000000, 400000000]);

            // Intersection of same set with itself equals the same set
            Assert.AreEqual(single.Count, doubled.Count);
        }
    }
}
