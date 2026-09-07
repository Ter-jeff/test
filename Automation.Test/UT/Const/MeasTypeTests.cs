using System.Collections.Generic;

using Automation.Const;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Const
{
    [TestClass]
    public class MeasTypeTests
    {
        [TestMethod]
        public void MeasTypes_Always_ContainsExpectedTypesInOrder()
        {
            List<string> expected =
            [
                MeasType.MeasV,
                MeasType.MeasI,
                MeasType.MeasC,
                MeasType.MeasF,
                MeasType.MeasIdiff,
                MeasType.MeasVdiff,
                MeasType.MeasVdiff2,
                MeasType.MeasFdiff,
                MeasType.MeasVocm,
                MeasType.MeasR1,
                MeasType.MeasR2,
                MeasType.MeasCalc,
                MeasType.MeasLimit,
                MeasType.MeasCalcLimit,
                MeasType.WiMeas,
                MeasType.WiSrc,
                MeasType.MeasDutyCycle,
            ];

            CollectionAssert.AreEqual(expected, MeasType.MeasTypes);
        }
    }
}
