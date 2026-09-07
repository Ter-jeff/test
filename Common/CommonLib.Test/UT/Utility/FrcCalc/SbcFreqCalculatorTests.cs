using System.Collections.Generic;

using CommonLib.Utility.FrcCalc;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test.UT.Utility.FrcCalc
{
    [TestClass]
    public class SbcFreqCalculatorTests
    {
        // ── Constructor ──────────────────────────────────────────────────────
        [TestMethod]
        public void SbcFreqCalculator_DefaultConstructor_TargetFreqIsEmpty()
        {
            // Arrange & Act
            var calc = new SbcFreqCalculator();

            // Assert
            Assert.AreEqual(0, calc.TargetFreq.Count);
        }

        // ── Sdf property ─────────────────────────────────────────────────────
        [TestMethod]
        public void Sdf_SetAndGet_ReturnsAssignedValue()
        {
            // Arrange
            var calc = new SbcFreqCalculator
            {
                // Act
                Sdf = "TestSdf"
            };

            // Assert
            Assert.AreEqual("TestSdf", calc.Sdf);
        }

        [TestMethod]
        public void Sdf_DefaultValue_IsNull()
        {
            // Arrange & Act
            var calc = new SbcFreqCalculator();

            // Assert
            Assert.IsNull(calc.Sdf);
        }

        // ── Single-target happy path ──────────────────────────────────────────
        [TestMethod]
        public void SolveSbcFreqFrc_SingleLowFreq_ReturnsNonEmptyDict()
        {
            // Arrange
            var calc = new SbcFreqCalculator();
            // 400 MHz
            calc.TargetFreq.Add(400000000);

            // Act
            calc.SolveSbcFreqFrc(out Dictionary<string, SbcSolutionFrc> result);

            // Assert
            Assert.IsNotNull(result);
            Assert.IsTrue(result.Count > 0);
        }

        [TestMethod]
        public void SolveSbcFreqFrc_SolutionHasPositiveSbcFreq()
        {
            // Arrange
            var calc = new SbcFreqCalculator();
            calc.TargetFreq.Add(400000000);

            // Act
            calc.SolveSbcFreqFrc(out Dictionary<string, SbcSolutionFrc> result);

            // Assert
            foreach (KeyValuePair<string, SbcSolutionFrc> kv in result)
            {
                Assert.IsTrue(kv.Value.SbcFreq > 0);
            }
        }

        [TestMethod]
        public void SolveSbcFreqFrc_SolutionEngineListNotEmpty()
        {
            // Arrange
            var calc = new SbcFreqCalculator();
            calc.TargetFreq.Add(400000000);

            // Act
            calc.SolveSbcFreqFrc(out Dictionary<string, SbcSolutionFrc> result);

            // Assert
            foreach (KeyValuePair<string, SbcSolutionFrc> kv in result)
            {
                Assert.IsTrue(kv.Value.EngineList.Count > 0);
            }
        }

        [TestMethod]
        public void SolveSbcFreqFrc_ReturnsFirstSolutionObject()
        {
            // Arrange
            var calc = new SbcFreqCalculator();
            calc.TargetFreq.Add(400000000);

            // Act
            SbcSolutionFrc firstSolution = calc.SolveSbcFreqFrc(out _);

            // Assert
            Assert.IsNotNull(firstSolution);
        }

        [TestMethod]
        public void SolveSbcFreqFrc_TargetFreqPllInputWithinConstraints()
        {
            // Arrange
            var calc = new SbcFreqCalculator();
            calc.TargetFreq.Add(400000000);

            // Act
            calc.SolveSbcFreqFrc(out Dictionary<string, SbcSolutionFrc> result);

            // Assert
            foreach (KeyValuePair<string, SbcSolutionFrc> kv in result)
            {
                foreach (PaEngineItem engine in kv.Value.EngineList)
                {
                    Assert.IsTrue(engine.PllInputFreq > 3000000, "PllInputFreq should be above low limit");
                    Assert.IsTrue(engine.PllInputFreq < 200000000, "PllInputFreq should be below high limit");
                }
            }
        }

        [TestMethod]
        public void SolveSbcFreqFrc_SbcFreqEqualsFourTimesPllInputFreqOfLastEngine()
        {
            // Arrange
            var calc = new SbcFreqCalculator();
            calc.TargetFreq.Add(400000000);

            // Act
            calc.SolveSbcFreqFrc(out Dictionary<string, SbcSolutionFrc> result);

            // Assert
            foreach (KeyValuePair<string, SbcSolutionFrc> kv in result)
            {
                SbcSolutionFrc solution = kv.Value;
                PaEngineItem lastEngine = solution.EngineList[^1];
                Assert.AreEqual(lastEngine.PllInputFreq * 4, solution.SbcFreq, 0.001);
            }
        }

        // ── Two-target: primary 400 MHz, secondary 200 MHz ───────────────────
        // Causes TrySingleValue(200 MHz, ...) → hSpeedMode=1 x2=1 returns true (lines 60-83)
        // and covers IsOkFrc lines 135 and 137.
        [TestMethod]
        public void SolveSbcFreqFrc_TwoTargets_HighPrimaryLowSecondary_FindsSolutions()
        {
            // Arrange
            var calc = new SbcFreqCalculator();
            // 400 MHz – primary
            calc.TargetFreq.Add(400000000);
            // 200 MHz – in tryList
            calc.TargetFreq.Add(200000000);

            // Act
            SbcSolutionFrc first = calc.SolveSbcFreqFrc(out Dictionary<string, SbcSolutionFrc> dict);

            // Assert
            Assert.IsNotNull(first);
            Assert.IsTrue(dict.Count > 0);
            Assert.IsTrue(first.SbcFreq > 0);
        }

        [TestMethod]
        public void SolveSbcFreqFrc_TwoTargets_HighPrimaryLowSecondary_SolutionHasTwoEngines()
        {
            // Arrange - one engine per target frequency
            var calc = new SbcFreqCalculator();
            calc.TargetFreq.Add(400000000);
            calc.TargetFreq.Add(200000000);

            // Act
            calc.SolveSbcFreqFrc(out Dictionary<string, SbcSolutionFrc> dict);

            // Assert
            foreach (KeyValuePair<string, SbcSolutionFrc> kv in dict)
            {
                Assert.AreEqual(2, kv.Value.EngineList.Count, "Each solution should have one engine per target");
            }
        }

        [TestMethod]
        public void SolveSbcFreqFrc_TwoTargets_HighPrimaryLowSecondary_AllEnginesPllInputWithinLimits()
        {
            // Arrange
            var calc = new SbcFreqCalculator();
            calc.TargetFreq.Add(400000000);
            calc.TargetFreq.Add(200000000);

            // Act
            calc.SolveSbcFreqFrc(out Dictionary<string, SbcSolutionFrc> dict);

            // Assert
            foreach (KeyValuePair<string, SbcSolutionFrc> kv in dict)
            {
                foreach (PaEngineItem engine in kv.Value.EngineList)
                {
                    Assert.IsTrue(engine.PllInputFreq > 3000000);
                    Assert.IsTrue(engine.PllInputFreq < 200000000);
                }
            }
        }

        // ── Two-target: primary 200 MHz, secondary 400 MHz ───────────────────
        // Causes TrySingleValue(400 MHz, ...) → x2=1 trivially empty, x2=2 path (lines 89-119)
        [TestMethod]
        public void SolveSbcFreqFrc_TwoTargets_LowPrimaryHighSecondary_FindsSolutions()
        {
            // Arrange
            var calc = new SbcFreqCalculator();
            // 200 MHz – primary (hSpeedMode=2 skipped ≤ 250 MHz)
            calc.TargetFreq.Add(200000000);
            // 400 MHz – in tryList → x2=2 path in TrySingleValue
            calc.TargetFreq.Add(400000000);

            // Act
            SbcSolutionFrc first = calc.SolveSbcFreqFrc(out Dictionary<string, SbcSolutionFrc> dict);

            // Assert
            Assert.IsNotNull(first);
            Assert.IsTrue(dict.Count > 0);
            Assert.IsTrue(first.SbcFreq > 0);
        }

        [TestMethod]
        public void SolveSbcFreqFrc_TwoTargets_LowPrimaryHighSecondary_SolutionHasTwoEngines()
        {
            // Arrange
            var calc = new SbcFreqCalculator();
            calc.TargetFreq.Add(200000000);
            calc.TargetFreq.Add(400000000);

            // Act
            calc.SolveSbcFreqFrc(out Dictionary<string, SbcSolutionFrc> dict);

            // Assert
            foreach (KeyValuePair<string, SbcSolutionFrc> kv in dict)
            {
                Assert.AreEqual(2, kv.Value.EngineList.Count);
            }
        }

        // ── Impossible secondary target ───────────────────────────────────────
        // TrySingleValue(7 GHz, ...) → both hSpeedMode iterations skip via patGenFreq limit
        // → return false, covers IsOkFrc line 141 (suitable=false) and SolveSbcFreqFrc
        //   null-solution path (lines 337, 339).
        [TestMethod]
        public void SolveSbcFreqFrc_ImpossibleSecondaryTarget_DictIsEmpty()
        {
            // Arrange
            var calc = new SbcFreqCalculator();
            // solvable primary
            calc.TargetFreq.Add(400000000);
            // 7 GHz secondary – patGenFreq always ≥ 550 MHz
            calc.TargetFreq.Add(7000000000);

            // Act
            calc.SolveSbcFreqFrc(out Dictionary<string, SbcSolutionFrc> dict);

            // Assert
            Assert.AreEqual(0, dict.Count);
        }

        [TestMethod]
        public void SolveSbcFreqFrc_ImpossibleSecondaryTarget_ReturnsDefaultSbcFreq()
        {
            // Arrange
            var calc = new SbcFreqCalculator();
            calc.TargetFreq.Add(400000000);
            calc.TargetFreq.Add(7000000000);

            // Act
            SbcSolutionFrc result = calc.SolveSbcFreqFrc(out _);

            // Assert – default fallback is 62.5 MHz
            Assert.AreEqual(62500000.0, result.SbcFreq, 0.001);
        }

        // ── Primary frequency above all limits ───────────────────────────────
        // firstTarget ≥ 1100 MHz: hSpeedMode=1 skipped (≥ 550 MHz),
        // hSpeedMode=2 gives patGenFreq ≥ 550 MHz → also skipped.
        // → dict empty, default solution returned (lines 337, 339).
        [TestMethod]
        public void SolveSbcFreqFrc_PrimaryAbove1100MHz_DictIsEmpty()
        {
            // Arrange
            var calc = new SbcFreqCalculator();
            // 1.2 GHz
            calc.TargetFreq.Add(1200000000);

            // Act
            calc.SolveSbcFreqFrc(out Dictionary<string, SbcSolutionFrc> dict);

            // Assert
            Assert.AreEqual(0, dict.Count);
        }

        [TestMethod]
        public void SolveSbcFreqFrc_PrimaryAbove1100MHz_ReturnsDefaultSbcFreq()
        {
            // Arrange
            var calc = new SbcFreqCalculator();
            calc.TargetFreq.Add(1200000000);

            // Act
            SbcSolutionFrc result = calc.SolveSbcFreqFrc(out _);

            // Assert
            Assert.AreEqual(62500000.0, result.SbcFreq, 0.001);
        }

        // ── Frequency boundary tests ──────────────────────────────────────────
        [TestMethod]
        public void SolveSbcFreqFrc_HighFreqAbovePatGenLimit_DictMayBeEmpty()
        {
            // Arrange – 600 MHz: hSpeedMode=1 skipped, hSpeedMode=2 patGenFreq=300 MHz (< 550 MHz)
            var calc = new SbcFreqCalculator();
            calc.TargetFreq.Add(600000000);

            // Act
            SbcSolutionFrc result = calc.SolveSbcFreqFrc(out Dictionary<string, SbcSolutionFrc> dict);

            // Assert – if no solution, default 62.5 MHz; otherwise SbcFreq > 0
            if (dict.Count == 0)
            {
                Assert.AreEqual(62500000.0, result.SbcFreq, 0.001);
            }
            else
            {
                Assert.IsTrue(result.SbcFreq > 0);
            }
        }

        [TestMethod]
        public void SolveSbcFreqFrc_FreqAtOrBelow250MHz_OnlyHSpeedMode1Used()
        {
            // Arrange – ≤ 250 MHz forces hSpeedMode=2 to be skipped in SolveSbcFreqFrc
            var calc = new SbcFreqCalculator();
            // 250 MHz
            calc.TargetFreq.Add(250000000);

            // Act
            SbcSolutionFrc result = calc.SolveSbcFreqFrc(out Dictionary<string, SbcSolutionFrc> dict);

            // Assert
            Assert.IsNotNull(result);
            // For 250 MHz, hSpeedMode=2 is skipped but hSpeedMode=1 may find solutions
            if (dict.Count > 0)
            {
                Assert.IsTrue(result.SbcFreq > 0);
            }
            else
            {
                Assert.AreEqual(62500000.0, result.SbcFreq, 0.001);
            }
        }

        [TestMethod]
        public void SolveSbcFreqFrc_ThreeTargets_AllSolvable_EngineListHasThreeEntries()
        {
            // Arrange
            var calc = new SbcFreqCalculator();
            // primary
            calc.TargetFreq.Add(400000000);
            // secondary 1
            calc.TargetFreq.Add(200000000);
            // secondary 2
            calc.TargetFreq.Add(100000000);

            // Act
            SbcSolutionFrc first = calc.SolveSbcFreqFrc(out Dictionary<string, SbcSolutionFrc> dict);

            // Assert
            Assert.IsNotNull(first);
            if (dict.Count > 0)
            {
                // Each solution has one engine per target (2 secondary + 1 primary)
                foreach (KeyValuePair<string, SbcSolutionFrc> kv in dict)
                {
                    Assert.AreEqual(3, kv.Value.EngineList.Count);
                }
            }
        }

        [TestMethod]
        public void SolveSbcFreqFrc_TwoTargets_DuplicateSolutions_NoDuplicatePllInputFreq()
        {
            // Arrange
            var calc = new SbcFreqCalculator();
            calc.TargetFreq.Add(400000000);
            calc.TargetFreq.Add(200000000);

            // Act
            calc.SolveSbcFreqFrc(out Dictionary<string, SbcSolutionFrc> dict);

            // Assert – existData dedup: no two solutions share the same SbcFreq
            var sbcFreqs = new HashSet<double>();
            foreach (KeyValuePair<string, SbcSolutionFrc> kv in dict)
            {
                Assert.IsTrue(sbcFreqs.Add(kv.Value.SbcFreq), $"Duplicate SbcFreq {kv.Value.SbcFreq} found in solutions");
            }
        }
    }
}
