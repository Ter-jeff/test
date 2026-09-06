using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class WaveDefRowTests
    {
        [TestMethod]
        public void WaveDefRow_DefaultConstructor_CreatesEmptyInstance()
        {
            // Arrange & Act
            var waveDefRow = new WaveDefRow();

            // Assert
            Assert.IsInstanceOfType(waveDefRow, typeof(IgxlRow));
        }

        [TestMethod]
        public void WaveDefRow_Inherits_FromIgxlRow()
        {
            // Arrange & Act
            var waveDefRow = new WaveDefRow();

            // Assert
            Assert.IsInstanceOfType(waveDefRow, typeof(IgxlRow));
        }
    }
}
