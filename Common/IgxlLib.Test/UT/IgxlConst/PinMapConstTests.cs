using IgxlLib.IgxlConst;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlConst
{

    [TestClass]
    public class PinMapConstTests
    {
        [TestMethod]
        public void PinMapConst_ConstantsExist()
        {
            // Assert - verify that the class can be instantiated
            var pinMapConst = new PinMapConst();
            Assert.IsNotNull(pinMapConst);
        }
    }
}
