using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class PatSetSubRowTests
    {
        [TestMethod]
        public void PatSetSubRow_DefaultConstructor_CreatesEmptyInstance()
        {
            // Arrange & Act
            var patSetSubRow = new PatSetSubRow();

            // Assert
            Assert.IsInstanceOfType(patSetSubRow, typeof(IgxlRow));
        }

        [TestMethod]
        public void PatSetSubRow_Inherits_FromIgxlRow()
        {
            // Arrange & Act
            var patSetSubRow = new PatSetSubRow();

            // Assert
            Assert.IsInstanceOfType(patSetSubRow, typeof(IgxlRow));
        }
    }
}
