using System.Collections.Generic;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.Parser;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class MiscInfoParserTests
    {
        private readonly MiscInfoParser _parser = new();

        [DataTestMethod]
        [DataRow(";", "", "")]
        [DataRow(":", "", "")]
        [DataRow("A:B", "A", "B")]
        [DataRow("A:B:C:D", "A", "B:C:D")]
        [DataRow("B", "B", "")]
        [DataRow("B:\"value\"", "B", "value")]
        [DataRow("C:A;C:B;C:\"D\"", "C", "A;B;D")]
        public void MiscInfoParserTest(string miscInfo, string miscKey, string miscValue)
        {
            Dictionary<string, string> dict = _parser.ParseKeyValueToDictionary(miscInfo, out string _);
            if (dict.TryGetValue(miscKey, out string? misc))
            {
                Assert.AreEqual(misc, miscValue);
            }
            else
            {
                if (dict.Count == 0)
                {
                    Assert.IsTrue(string.IsNullOrEmpty(miscKey) && string.IsNullOrEmpty(miscValue));
                }
                else
                {
                    Assert.Fail();
                }
            }
        }
    }
}
