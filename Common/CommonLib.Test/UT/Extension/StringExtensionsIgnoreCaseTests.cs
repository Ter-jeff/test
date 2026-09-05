using CommonLib.Extension;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test.UT.Extension
{
    [TestClass]
    public class StringExtensionsIgnoreCaseTests
    {
        [TestMethod]
        public void EqualsIgnoreCase_SameValues_ReturnsTrue()
        {
            Assert.IsTrue("ABC".EqualsIgnoreCase("abc"));
            Assert.IsTrue("ABC".EqualsIgnoreCase("ABC"));
            Assert.IsTrue("Test".EqualsIgnoreCase("test"));
        }

        [TestMethod]
        public void EqualsIgnoreCase_DifferentValues_ReturnsFalse()
        {
            Assert.IsFalse("ABC".EqualsIgnoreCase("XYZ"));
            Assert.IsFalse("Test".EqualsIgnoreCase("Testing"));
        }

        [TestMethod]
        public void EqualsIgnoreCase_NullValues_ReturnsCorrect()
        {
            Assert.IsTrue(((string)null).EqualsIgnoreCase(null));
            Assert.IsFalse(((string)null).EqualsIgnoreCase("abc"));
            Assert.IsFalse("abc".EqualsIgnoreCase(null));
        }

        [TestMethod]
        public void ContainsIgnoreCase_ValidInput_ReturnsTrue()
        {
            Assert.IsTrue("HELLO WORLD".ContainsIgnoreCase("world"));
            Assert.IsTrue("Hello World".ContainsIgnoreCase("HELLO"));
            Assert.IsTrue("Test123".ContainsIgnoreCase("test"));
        }

        [TestMethod]
        public void ContainsIgnoreCase_InvalidInput_ReturnsFalse()
        {
            Assert.IsFalse("HELLO".ContainsIgnoreCase("XYZ"));
            Assert.IsFalse("Test".ContainsIgnoreCase("testing"));
        }

        [TestMethod]
        public void ContainsIgnoreCase_NullSource_ReturnsFalse()
        {
            Assert.IsFalse(((string)null).ContainsIgnoreCase("test"));
            Assert.IsTrue(((string)null).ContainsIgnoreCase(null));
        }

        [TestMethod]
        public void StartsWithIgnoreCase_ValidInput_ReturnsTrue()
        {
            Assert.IsTrue("HELLO WORLD".StartsWithIgnoreCase("hello"));
            Assert.IsTrue("Test".StartsWithIgnoreCase("TEST"));
        }

        [TestMethod]
        public void StartsWithIgnoreCase_InvalidInput_ReturnsFalse()
        {
            Assert.IsFalse("HELLO".StartsWithIgnoreCase("WORLD"));
        }

        [TestMethod]
        public void EndsWithIgnoreCase_ValidInput_ReturnsTrue()
        {
            Assert.IsTrue("HELLO WORLD".EndsWithIgnoreCase("world"));
            Assert.IsTrue("Test".EndsWithIgnoreCase("TEST"));
        }

        [TestMethod]
        public void EndsWithIgnoreCase_InvalidInput_ReturnsFalse()
        {
            Assert.IsFalse("HELLO WORLD".EndsWithIgnoreCase("HELLO"));
        }

        [TestMethod]
        public void IndexOfIgnoreCase_ValidInput_ReturnsCorrectIndex()
        {
            Assert.AreEqual(0, "HELLO".IndexOfIgnoreCase("hello"));
            Assert.AreEqual(6, "HELLO WORLD".IndexOfIgnoreCase("world"));
        }

        [TestMethod]
        public void IndexOfIgnoreCase_InvalidInput_ReturnsMinus1()
        {
            Assert.AreEqual(-1, "HELLO".IndexOfIgnoreCase("XYZ"));
        }

        [TestMethod]
        public void IndexOfIgnoreCase_NullSource_ReturnsMinus1()
        {
            Assert.AreEqual(-1, ((string)null).IndexOfIgnoreCase("test"));
        }

        [TestMethod]
        public void ReplaceStartsWith_ValidInput_ReturnsModified()
        {
            string result = "HELLO WORLD".ReplaceStartsWith("HELLO", "HI");
            Assert.AreEqual("HI WORLD", result);

            result = "TestCase".ReplaceStartsWith("test", "NEW");
            Assert.AreEqual("NEWCase", result);
        }

        [TestMethod]
        public void ReplaceStartsWith_NoMatch_ReturnsOriginal()
        {
            string result = "HELLO".ReplaceStartsWith("WORLD", "NEW");
            Assert.AreEqual("HELLO", result);
        }

        [TestMethod]
        public void ReplaceEndsWith_ValidInput_ReturnsModified()
        {
            string result = "HELLO WORLD".ReplaceEndsWith("WORLD", "THERE");
            Assert.AreEqual("HELLO THERE", result);
        }

        [TestMethod]
        public void ReplaceEndsWith_NoMatch_ReturnsOriginal()
        {
            string result = "HELLO".ReplaceEndsWith("WORLD", "NEW");
            Assert.AreEqual("HELLO", result);
        }

        [TestMethod]
        public void SplitStartsWith_WithDelimiter_ReturnsSplit()
        {
            string[] result = "_TEST_".SplitStartsWith("_");
            Assert.AreEqual(2, result.Length);
            Assert.AreEqual("", result[0]);
            Assert.AreEqual("TEST_", result[1]);
        }

        [TestMethod]
        public void SplitStartsWith_WithoutDelimiter_ReturnsOriginal()
        {
            string[] result = "TEST_".SplitStartsWith("_");
            Assert.AreEqual(1, result.Length);
            Assert.AreEqual("TEST_", result[0]);
        }
    }
}
