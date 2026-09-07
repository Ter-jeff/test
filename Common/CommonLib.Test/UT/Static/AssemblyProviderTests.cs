using CommonLib.Static;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test.UT.Static
{
    [TestClass]
    public class AssemblyProviderTests
    {
        [TestMethod]
        public void GetFileVersion_WithVersion_ReturnsVersion()
        {
            var provider = new AssemblyProvider();
            string result = provider.GetFileVersion("1.0.0.0");

            Assert.AreEqual("1.0.0.0", result);
        }

        [TestMethod]
        public void GetFileVersion_WithDifferentVersion_ReturnsCorrect()
        {
            var provider = new AssemblyProvider();
            string result = provider.GetFileVersion("2.5.10.3");

            Assert.AreEqual("2.5.10.3", result);
        }

        [TestMethod]
        public void GetFileVersion_WithEmptyString_ReturnsEmpty()
        {
            var provider = new AssemblyProvider();
            string result = provider.GetFileVersion("");

            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void Current_CanBeSet()
        {
            AssemblyProvider originalProvider = AssemblyProvider.Current;
            var newProvider = new AssemblyProvider();

            AssemblyProvider.Current = newProvider;
            Assert.AreEqual(newProvider, AssemblyProvider.Current);

            AssemblyProvider.Current = originalProvider;
        }
    }
}
