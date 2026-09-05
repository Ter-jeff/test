using System.Collections.Generic;

using Automation.Utility;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Utility
{
    [TestClass]
    public class SbcFreqCalculatorTests
    {

        [TestMethod]
        public void SolveSbcFreq_EmptyTargetFreq_ReturnsEmptySolution()
        {
            // Arrange
            var calculator = new SbcFreqCalculator();

            // Act
            SbcSolution result = calculator.SolveSbcFreq();

            // Assert
            Assert.AreEqual(0, result.SbcFreq);
            Assert.AreEqual(0, result.EngineList.Count);
        }

        [TestMethod]
        public void SolveSbcFreq_NoValidClkD8RangeForTarget_ReturnsDefaultSolution()
        {
            // Arrange - a target of 300,000,000 puts the entire clkD8 search range at 0,
            // which always falls at or below the low limit, so no candidate is ever accepted.
            var calculator = new SbcFreqCalculator { TargetFreq = [300000000] };

            // Act
            SbcSolution result = calculator.SolveSbcFreq();

            // Assert
            Assert.AreEqual(62500000, result.SbcFreq);
            Assert.AreEqual(0, result.EngineList.Count);
        }

        [TestMethod]
        public void SolveSbcFreq_SingleTarget_FindsValidSolutionBelowThreshold()
        {
            // Arrange
            var calculator = new SbcFreqCalculator { TargetFreq = [100000000] };

            // Act
            SbcSolution result = calculator.SolveSbcFreq();

            // Assert
            Assert.AreEqual(20000000, result.SbcFreq);
            Assert.AreEqual(1, result.EngineList.Count);
            PaEngine engine = result.EngineList[0];
            Assert.AreEqual(200000000, engine.ClkD8Freq);
            Assert.AreEqual(2, engine.D2);
            Assert.AreEqual(5000000, engine.PdfFreq);
            Assert.AreEqual(40, engine.M);
            Assert.AreEqual(5000000, engine.PllInputFreq);
            Assert.AreEqual(1, engine.D1);
            Assert.AreEqual(100000000, engine.TargetFreq);
        }

        [TestMethod]
        public void SolveSbcFreq_MultipleTargetsAllSucceed_ReturnsSolutionWithAllEngines()
        {
            // Arrange
            var calculator = new SbcFreqCalculator { TargetFreq = [100000000, 50000000] };

            // Act
            SbcSolution result = calculator.SolveSbcFreq();

            // Assert
            Assert.AreEqual(20000000, result.SbcFreq);
            Assert.AreEqual(2, result.EngineList.Count);
            Assert.IsTrue(result.EngineList.Exists(e => e.TargetFreq == 100000000));
            Assert.IsTrue(result.EngineList.Exists(e => e.TargetFreq == 50000000));
        }

        [TestMethod]
        public void IsOk_AllTargetsSucceed_ReturnsTrueWithPopulatedEngines()
        {
            // Arrange
            var calculator = new SbcFreqCalculator();

            // Act
            bool result = calculator.IsOk(5000000, [100000000], out List<PaEngine> engines);

            // Assert
            Assert.IsTrue(result);
            Assert.AreEqual(1, engines.Count);
            Assert.AreEqual(100000000, engines[0].TargetFreq);
        }

        [TestMethod]
        public void IsOk_OneTargetFails_ReturnsFalseButKeepsSuccessfulEngines()
        {
            // Arrange - 999,999,999 is larger than the high clkD8 limit, so its search range
            // collapses to a single non-viable candidate and TrySingelValue always fails for it.
            var calculator = new SbcFreqCalculator();

            // Act
            bool result = calculator.IsOk(5000000, [100000000, 999999999], out List<PaEngine> engines);

            // Assert
            Assert.IsFalse(result);
            Assert.AreEqual(1, engines.Count);
            Assert.AreEqual(100000000, engines[0].TargetFreq);
        }

        [TestMethod]
        public void TrySingelValue_ValidCombination_ReturnsTrueWithEngineDetails()
        {
            // Arrange
            var calculator = new SbcFreqCalculator();

            // Act
            bool result = calculator.TrySingelValue(100000000, 5000000, out PaEngine engine);

            // Assert
            Assert.IsTrue(result);
            Assert.AreEqual(200000000, engine.ClkD8Freq);
            Assert.AreEqual(2, engine.D2);
            Assert.AreEqual(5000000, engine.PdfFreq);
            Assert.AreEqual(40, engine.M);
            Assert.AreEqual(1, engine.D1);
            Assert.AreEqual(5000000, engine.PllInputFreq);
            Assert.AreEqual(100000000, engine.TargetFreq);
        }

        [TestMethod]
        public void TrySingelValue_NoValidCombination_ReturnsFalse()
        {
            // Arrange
            var calculator = new SbcFreqCalculator();

            // Act
            bool result = calculator.TrySingelValue(100000000, 3000001, out PaEngine _);

            // Assert
            Assert.IsFalse(result);
        }
    }
}
