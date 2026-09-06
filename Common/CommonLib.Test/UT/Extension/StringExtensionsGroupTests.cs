using CommonLib.Extension;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test.UT.Extension
{
    [TestClass]
    public class StringExtensionsGroupTests
    {
        [TestMethod]
        public void GetIterates_ValidRangeFormat_ReturnsGroupItems()
        {
            string input = "[1:5]";
            GroupItem result = input.GetIterates();

            Assert.IsNotNull(result);
            Assert.AreEqual(input, result.OriText);
        }

        [TestMethod]
        public void GetIterates_NoGroupBrackets_ReturnsEmptyGroups()
        {
            string input = "NoGroups";
            GroupItem result = input.GetIterates();

            Assert.AreEqual(input, result.OriText);
            Assert.AreEqual(0, result.Groups.Count);
        }

        [TestMethod]
        public void GetIterates_MultipleRanges_ReturnsGroupItems()
        {
            string input = "[0:2],[1:3]";
            GroupItem result = input.GetIterates();

            Assert.IsNotNull(result);
        }

        [TestMethod]
        public void GetIterates_ReverseRange_ReturnsGroupItems()
        {
            string input = "[5:1]";
            GroupItem result = input.GetIterates();

            Assert.IsNotNull(result);
        }

        [TestMethod]
        public void GetIterates_DotDotRange_ReturnsGroupItems()
        {
            string input = "[1..5]";
            GroupItem result = input.GetIterates();

            Assert.IsNotNull(result);
        }

        [TestMethod]
        public void GetIterates_SingleNumber_ReturnsGroupItems()
        {
            string input = "[3]";
            GroupItem result = input.GetIterates();

            Assert.IsNotNull(result);
        }
    }
}
