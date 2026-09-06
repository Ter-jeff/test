using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class PSetTests
    {
        [TestMethod]
        public void PSet_DefaultConstructor_CreatesEmptyInstance()
        {
            // Arrange & Act
            var pSet = new PSet();

            // Assert
            Assert.IsInstanceOfType(pSet, typeof(IgxlRow));
        }

        [TestMethod]
        public void PSet_Inherits_FromIgxlRow()
        {
            // Arrange & Act
            var pSet = new PSet();

            // Assert
            Assert.IsInstanceOfType(pSet, typeof(IgxlRow));
        }
    }
}
