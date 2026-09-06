using Automation.Const;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Const
{
    [TestClass]
    public class DcContiConstTests
    {
        [TestMethod]
        public void FailFlagRegex_LowerCasePrefix_IsCaseInsensitiveMatch()
        {
            Assert.IsTrue(DcContiConst.FailFlagRegex.IsMatch("f_test"));
        }
    }
}
