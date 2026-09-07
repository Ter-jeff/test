using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class PortSetTests
    {
        [TestMethod]
        public void PortSet_DefaultConstructor_CreatesEmptyInstance()
        {
            // Arrange & Act
            var portSet = new PortSet();

            // Assert
            Assert.IsInstanceOfType(portSet, typeof(IgxlRow));
        }

        [TestMethod]
        public void PortSet_Inherits_FromIgxlRow()
        {
            // Arrange & Act
            var portSet = new PortSet();

            // Assert
            Assert.IsInstanceOfType(portSet, typeof(IgxlRow));
        }
    }
}
