using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class MixedSigRowTests
    {
        [TestMethod]
        public void MixedSigRow_DefaultConstructor_CreatesEmptyInstance()
        {
            // Arrange & Act
            var mixedSigRow = new MixedSigRow();

            // Assert
            Assert.IsInstanceOfType(mixedSigRow, typeof(IgxlRow));
        }

        [TestMethod]
        public void MixedSigRow_Inherits_FromIgxlRow()
        {
            // Arrange & Act
            var mixedSigRow = new MixedSigRow();

            // Assert
            Assert.IsInstanceOfType(mixedSigRow, typeof(IgxlRow));
        }
    }
}
