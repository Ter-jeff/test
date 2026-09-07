using IgxlLib.IgxlConst;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlConst
{
    [TestClass]
    public class OpCodeTests
    {
        [TestMethod]
        public void OpCode_ConstantsExist()
        {
            // Assert - verify that the class can be instantiated
            var opCode = new OpCode();
            Assert.IsNotNull(opCode);
        }
    }
}
