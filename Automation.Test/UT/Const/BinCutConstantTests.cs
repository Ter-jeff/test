using System.Collections.Generic;
using System.Linq;

using Automation.Const;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Const
{
    [TestClass]
    public class BinCutConstantTests
    {
        private static readonly string[] _expected = ["FT1", "FT2"];

        [TestMethod]
        public void GradeJobMap_Always_ContainsExpectedJobsPerGrade()
        {
            Dictionary<string, HashSet<string>> gradeJobMap = BinCutConstant.GradeJobMap;

            CollectionAssert.AreEquivalent(_expected, gradeJobMap.Keys.ToArray());

            HashSet<string> expectedFt1 = ["FT1", "FT2_25C", "WLFT", "WLFT1", "FT_ROOM", "RMA_ROOM"];
            HashSet<string> expectedFt2 = ["FT2", "FT2_85C", "WLFT2", "FT_HOT", "RMA_HOT"];

            CollectionAssert.AreEquivalent(expectedFt1.ToArray(), gradeJobMap["FT1"].ToArray());
            CollectionAssert.AreEquivalent(expectedFt2.ToArray(), gradeJobMap["FT2"].ToArray());
        }
    }
}
