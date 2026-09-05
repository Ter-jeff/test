using IgxlLib.Const;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.Const
{
    [TestClass]
    public class IgxlSheetNamesTests
    {
        [TestMethod]
        public void IgxlSheetNames_ConstantsExist()
        {
            // Assert - verify that the class can be instantiated
            var sheetNames = new IgxlSheetNames();
            Assert.IsNotNull(sheetNames);
        }
    }
}
