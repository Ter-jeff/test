using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class TSetTests
    {
        [TestMethod]
        public void TSet_DefaultConstructor_CreatesEmptyInstance()
        {
            // Arrange & Act
            var tSet = new TSet();

            // Assert
            Assert.IsInstanceOfType(tSet, typeof(IgxlRow));
        }

        [TestMethod]
        public void TSet_Inherits_FromIgxlRow()
        {
            // Arrange & Act
            var tSet = new TSet();

            // Assert
            Assert.IsInstanceOfType(tSet, typeof(IgxlRow));
        }
    }
}
