using CommonLib.Static;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test.UT.Static
{
    [TestClass]
    public class MockAssemblyProviderTests
    {
        [TestMethod]
        public void GetFileVersion_AlwaysReturnsFixedVersion()
        {
            MockAssemblyProvider provider = MockAssemblyProvider.Instance;
            string result = provider.GetFileVersion("any.version");

            Assert.AreEqual("2022.12.19.1", result);
        }

        [TestMethod]
        public void GetFileVersion_WithDifferentInput_StillReturnsMockVersion()
        {
            MockAssemblyProvider provider = MockAssemblyProvider.Instance;

            string result1 = provider.GetFileVersion("1.0.0.0");
            string result2 = provider.GetFileVersion("9.9.9.9");

            Assert.AreEqual("2022.12.19.1", result1);
            Assert.AreEqual("2022.12.19.1", result2);
        }

        [TestMethod]
        public void Instance_IsSingleton()
        {
            MockAssemblyProvider instance1 = MockAssemblyProvider.Instance;
            MockAssemblyProvider instance2 = MockAssemblyProvider.Instance;

            Assert.AreSame(instance1, instance2);
        }

        [TestMethod]
        public void Instance_HasCorrectMockVersion()
        {
            MockAssemblyProvider provider = MockAssemblyProvider.Instance;
            Assert.IsTrue(provider.GetFileVersion("test").Contains("2022.12.19"));
        }
    }
}
