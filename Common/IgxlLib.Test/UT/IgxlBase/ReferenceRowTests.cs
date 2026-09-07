using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class ReferenceRowTests
    {
        [TestMethod]
        public void ReferenceRow_DefaultConstructor_CreatesEmptyInstance()
        {
            // Arrange & Act
            var referenceRow = new ReferenceRow();

            // Assert
            Assert.IsInstanceOfType(referenceRow, typeof(IgxlRow));
        }

        [TestMethod]
        public void ReferenceRow_Inherits_FromIgxlRow()
        {
            // Arrange & Act
            var referenceRow = new ReferenceRow();

            // Assert
            Assert.IsInstanceOfType(referenceRow, typeof(IgxlRow));
        }
    }
}
