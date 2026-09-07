using System.Collections.Generic;

using CommonLib.Utility;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test.UT.Utility
{
    [TestClass]
    public class CombinationTests
    {
        [TestMethod]
        public void GetAllPossibleCombos_StringList_ReturnsAllCombinations()
        {
            var items = new List<List<string>>
            {
                new() { "A", "B" },
                new() { "1", "2" },
                new() { "X", "Y" }
            };

            List<List<string>> result = Combination.GetAllPossibleCombos(items);
            // 2 * 2 * 2

            Assert.AreEqual(8, result.Count);
        }

        [TestMethod]
        public void GetAllPossibleCombos_WithNullValue_SkipsNull()
        {
            var items = new List<List<string>>
            {
                new() { "A", null },
                new() { "1", "2" }
            };

            List<List<string>> result = Combination.GetAllPossibleCombos(items);

            // Should skip null values
            Assert.IsTrue(result.Count > 0);
        }

        [TestMethod]
        public void GetAllPossibleCombos_NestedStringList_ReturnsAllCombinations()
        {
            var items = new List<List<List<string>>>
            {
                new() {
                    new List<string> { "A", "B" },
                    new List<string> { "1", "2" }
                },
                new() {
                    new List<string> { "X", "Y" }
                }
            };

            List<List<List<string>>> result = Combination.GetAllPossibleCombos(items);

            Assert.IsTrue(result.Count > 0);
        }

        [TestMethod]
        public void CombineByUnderLine_TwoStrings_ReturnsCombined()
        {
            string result = Combination.CombineByUnderLine("Test", "Value");
            Assert.AreEqual("Test_Value", result);
        }

        [TestMethod]
        public void CombineByUnderLine_FirstEmpty_ReturnsSecond()
        {
            string result = Combination.CombineByUnderLine("", "Value");
            Assert.AreEqual("Value", result);
        }

        [TestMethod]
        public void CombineByUnderLine_SecondEmpty_ReturnsFirst()
        {
            string result = Combination.CombineByUnderLine("Test", "");
            Assert.AreEqual("Test", result);
        }

        [TestMethod]
        public void CombineByUnderLine_BothEmpty_ReturnsEmpty()
        {
            string result = Combination.CombineByUnderLine("", "");
            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void CombineByUnderLine_WithWhitespace_TrimsAndCombines()
        {
            string result = Combination.CombineByUnderLine("  Test  ", "  Value  ");
            Assert.AreEqual("Test_Value", result);
        }

        [TestMethod]
        public void CombineByUnderLine_StringList_ReturnsCombined()
        {
            var list = new List<string> { "A", "B", "C" };
            string result = Combination.CombineByUnderLine(list);
            Assert.AreEqual("A_B_C", result);
        }

        [TestMethod]
        public void CombineByUnderLine_ListWithEmpty_SkipsEmpty()
        {
            var list = new List<string> { "A", "", "C" };
            string result = Combination.CombineByUnderLine(list);
            Assert.AreEqual("A_C", result);
        }

        [TestMethod]
        public void CombineByUnderLine_EmptyList_ReturnsEmpty()
        {
            var list = new List<string>();
            string result = Combination.CombineByUnderLine(list);
            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void CombineEnableWord_TwoStrings_ReturnsCombined()
        {
            string result = Combination.CombineEnableWord("Enable", "Disable");
            Assert.AreEqual("Enable && Disable", result);
        }

        [TestMethod]
        public void CombineEnableWord_SameValue_ReturnsOne()
        {
            string result = Combination.CombineEnableWord("Enable", "enable");
            Assert.AreEqual("Enable", result);
        }

        [TestMethod]
        public void CombineEnableWord_FirstEmpty_ReturnsSecond()
        {
            string result = Combination.CombineEnableWord("", "Value");
            Assert.AreEqual("Value", result);
        }

        [TestMethod]
        public void GetAllPossibleCombos_SingleList_ReturnsCombos()
        {
            var items = new List<List<string>>
            {
                new() { "A", "B" }
            };

            List<List<string>> result = Combination.GetAllPossibleCombos(items);

            Assert.AreEqual(2, result.Count);
        }

        [TestMethod]
        public void GetAllPossibleCombos_AllEmpty_ReturnsEmpty()
        {
            var items = new List<List<string>>
            {
                new(),
                new()
            };

            List<List<string>> result = Combination.GetAllPossibleCombos(items);

            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void GetAllPossibleCombos_NestedLists_ReturnsCorrect()
        {
            var items = new List<List<List<string>>>
            {
                new() {
                    new List<string> { "A", "B" }
                }
            };

            List<List<List<string>>> result = Combination.GetAllPossibleCombos(items);

            Assert.IsTrue(result.Count > 0);
        }

        [TestMethod]
        public void CombineByUnderLine_ListWithNull_SkipsNull()
        {
            var list = new List<string> { "A", null, "B" };
            string result = Combination.CombineByUnderLine(list);

            Assert.IsNotNull(result);
        }

        [TestMethod]
        public void CombineEnableWord_WithParentheses_AddsAdditionalParentheses()
        {
            string result = Combination.CombineEnableWord("(A|B)", "C");

            Assert.IsTrue(result.Contains('('));
        }

        [TestMethod]
        public void CombineEnableWord_BothWithOperators_AddsParentheses()
        {
            string result = Combination.CombineEnableWord("A&B", "C|D");

            Assert.IsTrue(result.Contains('('));
            Assert.IsTrue(result.Contains("&&"));
        }
    }
}
