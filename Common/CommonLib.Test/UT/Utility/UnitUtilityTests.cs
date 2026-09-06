using CommonLib.Utility;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test.UT.Utility
{
    [TestClass]
    public class UnitUtilityTests
    {
        [TestMethod]
        public void GetScale_ValidUnit_ReturnsCorrectScale()
        {
            _ = UnitUtility.Instance;

            Assert.AreEqual(-15, UnitUtility.GetScale("f"));
            Assert.AreEqual(-12, UnitUtility.GetScale("p"));
            Assert.AreEqual(-9, UnitUtility.GetScale("n"));
            Assert.AreEqual(-6, UnitUtility.GetScale("u"));
            Assert.AreEqual(-3, UnitUtility.GetScale("m"));
            Assert.AreEqual(-2, UnitUtility.GetScale("%"));
            Assert.AreEqual(3, UnitUtility.GetScale("k"));
            Assert.AreEqual(3, UnitUtility.GetScale("K"));
            Assert.AreEqual(6, UnitUtility.GetScale("M"));
            Assert.AreEqual(9, UnitUtility.GetScale("G"));
            Assert.AreEqual(12, UnitUtility.GetScale("T"));
        }

        [TestMethod]
        public void GetScale_InvalidUnit_ReturnsZero()
        {
            _ = UnitUtility.Instance;
            Assert.AreEqual(0, UnitUtility.GetScale("Q"));
            Assert.AreEqual(0, UnitUtility.GetScale(""));
        }

        [TestMethod]
        public void GetScale_NullUnit_ReturnsZero()
        {
            _ = UnitUtility.Instance;
            Assert.AreEqual(0, UnitUtility.GetScale(null));
        }

        [TestMethod]
        public void Instance_IsSingleton_ReturnsSameInstance()
        {
            UnitUtility instance1 = UnitUtility.Instance;
            UnitUtility instance2 = UnitUtility.Instance;

            Assert.AreSame(instance1, instance2);
        }
    }
}
