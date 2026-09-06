using System.Collections.Generic;

using Automation.GenerateIgxl.Basic.Business.GenConti.Base.DcContiStrategy;
using Automation.Reader;

using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Basic
{
    [TestClass]
    public class ContiUserFunctionTests : FunctionTestBase
    {
        [TestMethod]
        public void NullInputs_Returns0_NoThrow()
        {

            var grpTNameLimit = new List<DcTestContiSheetLimit>
            {
                null!,
                new() { LimitHeader = "TN_1", LimitValue = "0.1", HiLimitValue = "0.9", LimitUnit = "mA" },
                null!,
                new() { LimitHeader = "TN_2", LimitValue = "1.0", HiLimitValue = "2.0", LimitUnit = "" }
            };

            var flowRows = new List<FlowRow>();

            int added = ContiUserFunction.AddUseLimitRows(grpTNameLimit, "PARAM_X", flowRows);

            Assert.AreEqual(2, added);
            Assert.AreEqual(2, flowRows.Count);
        }

    }
}
