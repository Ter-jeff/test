using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class PatternClassTests
    {
        #region Constructor parsing

        [TestMethod]
        public void Constructor_SinglePattern_ParsesNamesAndSets()
        {
            var pc = new PatternClass("PAT1");

            Assert.AreEqual("pat1", pc.TestPlanPatternName);
            Assert.AreEqual("PAT1", pc.RealPatternName);
            CollectionAssert.AreEqual(new List<string> { "pat1" }, pc.InstancePatternName);
            Assert.AreEqual(1, pc.PatternSetList.Count);
            CollectionAssert.AreEqual(new List<string> { "pat1" }, pc.PatternSetList[0]);
        }

        [TestMethod]
        public void Constructor_CommaSeparatedPatterns_CreatesMultiplePatternSets()
        {
            var pc = new PatternClass("PAT1,PAT2");

            CollectionAssert.AreEqual(new List<string> { "pat1", "pat2" }, pc.InstancePatternName);
            Assert.AreEqual(2, pc.PatternSetList.Count);
        }

        [TestMethod]
        public void Constructor_PlusSeparatedPatterns_GroupedInSamePatternSet()
        {
            var pc = new PatternClass("PAT1+PAT2");

            Assert.AreEqual(1, pc.PatternSetList.Count);
            CollectionAssert.AreEqual(new List<string> { "pat1", "pat2" }, pc.PatternSetList[0]);
            CollectionAssert.AreEqual(new List<string> { "pat1", "pat2" }, pc.InstancePatternName);
        }

        [TestMethod]
        public void Constructor_HashSeparatedPatterns_TreatedLikeComma()
        {
            var pc = new PatternClass("PAT1#PAT2");

            Assert.AreEqual(2, pc.PatternSetList.Count);
        }

        [TestMethod]
        public void Constructor_LastPayloadStartsWithBinOutPrefix_SetsIsDefaultNobinoutTrue()
        {
            var pc = new PatternClass("FA_TEST");
            Assert.IsTrue(pc.IsDefaultNobinout);
        }

        [TestMethod]
        public void Constructor_LastPayloadDoesNotStartWithBinOutPrefix_SetsIsDefaultNobinoutFalse()
        {
            var pc = new PatternClass("XX_TEST");
            Assert.IsFalse(pc.IsDefaultNobinout);
        }

        #endregion

        #region Copy constructor / Copy()

        [TestMethod]
        public void CopyConstructor_NullOther_LeavesDefaults()
        {
            var pc = new PatternClass((PatternClass)null);

            Assert.AreEqual("", pc.TestPlanPatternName);
            Assert.AreEqual("", pc.RealPatternName);
            Assert.AreEqual(0, pc.InstancePatternName.Count);
            Assert.AreEqual(0, pc.PatternSetList.Count);
            Assert.IsFalse(pc.IsDefaultNobinout);
        }

        [TestMethod]
        public void CopyConstructor_ValidOther_DeepCopiesLists()
        {
            var original = new PatternClass("PAT1+PAT2");
            var copy = new PatternClass(original);

            copy.InstancePatternName.Add("extra");
            copy.PatternSetList[0].Add("extra2");

            Assert.AreEqual(2, original.InstancePatternName.Count);
            Assert.AreEqual(3, copy.InstancePatternName.Count);
            Assert.AreEqual(2, original.PatternSetList[0].Count);
            Assert.AreEqual(3, copy.PatternSetList[0].Count);
        }

        [TestMethod]
        public void Copy_ReturnsIndependentDeepCopy()
        {
            var original = new PatternClass("PAT1");
            PatternClass copy = original.Copy();

            Assert.AreNotSame(original, copy);
            Assert.AreEqual(original.RealPatternName, copy.RealPatternName);
        }

        #endregion

        #region GetPatternName

        [TestMethod]
        public void GetPatternName_Multiple_ReturnsMultipleConstPrefixedLastPayload()
        {
            var pc = new PatternClass("PAT1,PAT2");
            Assert.AreEqual(PatternClass.MultipleConst + "pat2", pc.GetPatternName());
        }

        [TestMethod]
        public void GetPatternName_MultiTimeDomain_ReturnsMtdConstPrefixedLastInstancePattern()
        {
            var pc = new PatternClass("PAT1") { RealPatternName = "SOMETHING#ELSE" };
            Assert.AreEqual(PatternClass.MtdConst + "pat1", pc.GetPatternName());
        }

        [TestMethod]
        public void GetPatternName_MatchesInstancePattern_ReturnsRealPatternNameAsIs()
        {
            var pc = new PatternClass("PAT1") { RealPatternName = "Instance:ABC" };
            Assert.AreEqual("Instance:ABC", pc.GetPatternName());
        }

        [TestMethod]
        public void GetPatternName_MatchesOpcodePattern_ReturnsRealPatternNameAsIs()
        {
            var pc = new PatternClass("PAT1") { RealPatternName = "Opcode:XYZ" };
            Assert.AreEqual("Opcode:XYZ", pc.GetPatternName());
        }

        [TestMethod]
        public void GetPatternName_FallbackNoSpecialMatch_ReturnsLastPayload()
        {
            var pc = new PatternClass("PAT1");
            Assert.AreEqual("pat1", pc.GetPatternName());
        }

        #endregion

        #region IsMultiple / IsMultiTimeDomain / IsFullMtdPattern

        [TestMethod]
        public void IsMultiple_SinglePattern_ReturnsFalse()
        {
            Assert.IsFalse(new PatternClass("PAT1").IsMultiple());
        }

        [TestMethod]
        public void IsMultiple_MultiplePatterns_ReturnsTrue()
        {
            Assert.IsTrue(new PatternClass("PAT1,PAT2").IsMultiple());
        }

        [TestMethod]
        public void IsMultiTimeDomain_ContainsHash_ReturnsTrue()
        {
            Assert.IsTrue(new PatternClass("PAT1#PAT2").IsMultiTimeDomain());
        }

        [TestMethod]
        public void IsMultiTimeDomain_NoHash_ReturnsFalse()
        {
            Assert.IsFalse(new PatternClass("PAT1").IsMultiTimeDomain());
        }

        [TestMethod]
        public void IsFullMtdPattern_LengthMatchesInstanceCount_ReturnsTrue()
        {
            Assert.IsTrue(new PatternClass("PAT1+PAT2").IsFullMtdPattern());
        }

        [TestMethod]
        public void IsFullMtdPattern_LengthMismatch_ReturnsFalse()
        {
            var pc = new PatternClass("PAT1+PAT2");
            pc.InstancePatternName.Add("extra");
            Assert.IsFalse(pc.IsFullMtdPattern());
        }

        #endregion

        #region GetAliasPatternList

        [TestMethod]
        public void GetAliasPatternList_SplitsOnDelimitersKeepingSeparators()
        {
            var pc = new PatternClass("PAT1,PAT2;PAT3&PAT4");
            List<string> expected = ["pat1", ",", "pat2", ";", "pat3", "&", "pat4"];
            CollectionAssert.AreEqual(expected, pc.GetAliasPatternList());
        }

        #endregion

        #region GetLastPayload

        [TestMethod]
        public void GetLastPayload_MultiTimeDomain_DelegatesToMtdLastPayload()
        {
            // PatternSetList is deliberately different from InstancePatternName's last entry, so the
            // MTD branch (InstancePatternName.Last()) is distinguishable from the non-MTD fallback
            // (PatternSetList.Last().Last()).
            var pc = new PatternClass("PAT1")
            {
                RealPatternName = "X#Y",
                PatternSetList = [["different_last_value"]]
            };
            Assert.AreEqual("pat1", pc.GetLastPayload());
        }

        [TestMethod]
        public void GetLastPayload_EmptyPatternSetList_ReturnsEmptyString()
        {
            var pc = new PatternClass("PAT1") { PatternSetList = [] };
            Assert.AreEqual("", pc.GetLastPayload());
        }

        [TestMethod]
        public void GetLastPayload_LastSetEmpty_ReturnsEmptyString()
        {
            var pc = new PatternClass("PAT1")
            {
                PatternSetList = [[]]
            };
            Assert.AreEqual("", pc.GetLastPayload());
        }

        [TestMethod]
        public void GetLastPayload_NormalCase_ReturnsLastItemOfLastSet()
        {
            var pc = new PatternClass("PAT1+PAT2,PAT3");
            Assert.AreEqual("pat3", pc.GetLastPayload());
        }

        [TestMethod]
        public void GetLastPayload_FirstSetEmptyButLastSetNonEmpty_ReturnsLastValue()
        {
            // Guards against checking PatternSetList.First() instead of .Last().
            var pc = new PatternClass("PAT1")
            {
                PatternSetList = [[], ["pat2"]]
            };
            Assert.AreEqual("pat2", pc.GetLastPayload());
        }

        #endregion

        #region GetMtdLastPayLoad / GetMtdEachLastPayLoad

        [TestMethod]
        public void GetMtdLastPayLoad_ReturnsLastInstancePatternName()
        {
            Assert.AreEqual("pat2", new PatternClass("PAT1+PAT2").GetMtdLastPayLoad());
        }

        [TestMethod]
        public void GetMtdEachLastPayLoad_SplitsOnHashAndPlusTakingLast()
        {
            var pc = new PatternClass("PAT1") { RealPatternName = "A+B#C+D" };
            CollectionAssert.AreEqual(new List<string> { "B", "D" }, pc.GetMtdEachLastPayLoad());
        }

        #endregion

        #region GetInstancePatternName

        [TestMethod]
        public void GetInstancePatternName_NormalPattern_ReturnsJoinedNames()
        {
            Assert.AreEqual("pat1;pat2", new PatternClass("PAT1+PAT2").GetInstancePatternName());
        }

        [TestMethod]
        public void GetInstancePatternName_NoPattPattern_ReturnsEmpty()
        {
            Assert.AreEqual("", new PatternClass("No_patt").GetInstancePatternName());
        }

        [TestMethod]
        public void GetInstancePatternName_InstancePattern_ReturnsEmpty()
        {
            var pc = new PatternClass("PAT1") { TestPlanPatternName = "instance:abc" };
            Assert.AreEqual("", pc.GetInstancePatternName());
        }

        #endregion

        #region ConvertRealPatternName

        [TestMethod]
        public void ConvertRealPatternName_NullScghData_LeavesRealPatternNameUnchanged()
        {
            var pc = new PatternClass("PAT1");
            string before = pc.RealPatternName;

            pc.ConvertRealPatternName(null);

            Assert.AreEqual(before, pc.RealPatternName);
        }

        [TestMethod]
        public void ConvertRealPatternName_EmptyConvertedPatternRowList_LeavesRealPatternNameUnchanged()
        {
            var pc = new PatternClass("PAT1");
            string before = pc.RealPatternName;

            pc.ConvertRealPatternName(new ScghData());

            Assert.AreEqual(before, pc.RealPatternName);
        }

        #endregion
    }
}
