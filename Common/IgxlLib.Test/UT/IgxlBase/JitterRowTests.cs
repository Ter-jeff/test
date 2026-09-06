using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{

    [TestClass]
    public class JitterRowTests
    {
        [TestMethod]
        public void JitterRow_DefaultConstructor_CreatesEmptyInstance()
        {
            // Arrange & Act
            var jitterRow = new JitterRow();

            // Assert
            Assert.IsInstanceOfType(jitterRow, typeof(IgxlRow));
        }

        [TestMethod]
        public void JitterRow_Inherits_FromIgxlRow()
        {
            // Arrange & Act
            var jitterRow = new JitterRow();

            // Assert
            Assert.IsInstanceOfType(jitterRow, typeof(IgxlRow));
        }
    }
}
