using CommonLib.Utility;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test.UT.Utility
{
    [TestClass]
    public class CellDiffTests
    {
        [TestMethod]
        public void IsLiked_ExactMatch_ReturnsTrue()
        {
            Assert.IsTrue(CellDiff.IsLiked("Test", "test"));
            Assert.IsTrue(CellDiff.IsLiked("TEST", "test"));
        }

        [TestMethod]
        public void IsLiked_DifferentText_ReturnsFalse()
        {
            Assert.IsFalse(CellDiff.IsLiked("Test", "Other"));
        }

        [TestMethod]
        public void IsLiked_WithWhitespace_TrimsBefore()
        {
            Assert.IsTrue(CellDiff.IsLiked("  Test  ", "test"));
        }

        [TestMethod]
        public void IsLiked_WithDoubleSpaces_NormalizesSpaces()
        {
            Assert.IsTrue(CellDiff.IsLiked("Test  Value", "test value"));
        }

        [TestMethod]
        public void IsLiked_RegexDotStar_MatchesPattern()
        {
            Assert.IsTrue(CellDiff.IsLiked("Test123", "test.*"));
            Assert.IsTrue(CellDiff.IsLiked("TestAny", "test.*"));
        }

        [TestMethod]
        public void IsLiked_RegexDotPlus_MatchesPattern()
        {
            Assert.IsTrue(CellDiff.IsLiked("Test1", "test.+"));
            Assert.IsFalse(CellDiff.IsLiked("Test", "test.+"));
        }

        [TestMethod]
        public void IsLiked_RegexOrPattern_MatchesAlternatives()
        {
            Assert.IsTrue(CellDiff.IsLiked("Cat", "cat|dog"));
            Assert.IsTrue(CellDiff.IsLiked("Dog", "cat|dog"));
            Assert.IsFalse(CellDiff.IsLiked("Bird", "cat|dog"));
        }

        [TestMethod]
        public void FormatStringForCompare_TrimsAndUppercase()
        {
            string result = CellDiff.FormatStringForCompare("  test  ");
            Assert.AreEqual("TEST", result);
        }

        [TestMethod]
        public void FormatStringForCompare_NormalizesDoubleSpaces()
        {
            string result = CellDiff.FormatStringForCompare("Test  Value");
            Assert.AreEqual("TEST VALUE", result);
        }

        [TestMethod]
        public void FormatStringForCompare_TripleSpaces_BecomesOne()
        {
            string result = CellDiff.FormatStringForCompare("Test   Value");
            Assert.AreEqual("TEST VALUE", result);
        }
    }
}
