using System.Collections.Generic;
using System.Linq;

using Automation.Utility.Basic;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Basic
{
    [TestClass]
    public class DifferentialServiceTests
    {
        [DataTestMethod]
        [DataRow("A_P,A_N")]
        [DataRow("A_N,A_P")]
        public void DifferentialPair_RuleBasedPairs(string pinCsv)
        {
            List<string> pins = [.. pinCsv.Split(',')];
            Dictionary<string, string> pairs = DifferentialService.DifferentialPair(pins);

            Assert.AreEqual("A_P", pairs.Keys.ElementAt(0));
            Assert.AreEqual("A_N", pairs.Keys.ElementAt(1));
        }

        [DataTestMethod]
        [DataRow("TX_P,TX_N,B_P,B_N")]
        [DataRow("TX_N,TX_P,B_P,B_N")]
        public void GroupDiffPairs_ShouldReturnExpectedPairGrouping(string pinCsv)
        {
            List<string> pins = [.. pinCsv.Split(',')];
            List<string> groups = DifferentialService.GroupDiffPairs(pins);

            Assert.IsTrue(groups.Any(g => g.Contains("TX_P::TX_N")));
            Assert.IsTrue(groups.Any(g => g.Contains("B_P::B_N")));
        }

        [DataTestMethod]
        [DataRow("TX_P::TX_N")]
        [DataRow("TX_N::TX_P")]
        public void DiffPinPosAndNeg_ShouldIdentifyDiffPair_Correctly(string diffPins)
        {
            bool result = DifferentialService.DiffPinPosAndNeg(diffPins, out string pos, out string neg, out string group);

            Assert.IsTrue(result);
            Assert.IsTrue(pos.EndsWith("_P") || pos.EndsWith("PN"));
            Assert.IsTrue(neg.EndsWith("_N") || neg.EndsWith("NP"));
            StringAssert.Contains(group, "Diff");
        }

        [DataTestMethod]
        [DataRow("TX_P::TX_N", false, "TX_P,TX_N")]
        [DataRow("AB", false, "AB")]
        public void GenDiffGroupName_ShouldReturnCommaSeparatedPins_WhenNoGroupNeeded(string diffName, bool isNeedGenPinGroup, string expected)
        {
            string result = DifferentialService.GenDiffGroupName(null, diffName, isNeedGenPinGroup);
            Assert.AreEqual(expected, result);
        }

        [DataTestMethod]
        [DataRow("TX_P::TX_N", true, "All_DiffPairs")]
        public void GenDiffGroupName_ShouldReturnCommaSeparatedPins_WhenNoGroupNeeded1(string diffName, bool isNeedGenPinGroup, string expected)
        {
            var pinMapSheet = new PinMapSheet("");
            var pinGroup = new PinGroup("All_DiffPairs", "I/O");
            pinGroup.AddPin(new Pin("TX_P", "I/O"));
            pinGroup.AddPin(new Pin("TX_N", "I/O"));
            pinMapSheet.AddGroup(pinGroup);
            string result = DifferentialService.GenDiffGroupName(pinMapSheet, diffName, isNeedGenPinGroup);
            Assert.AreEqual(expected, result);
        }

        [DataTestMethod]
        [DataRow("TX_P::TX_N", true)]
        [DataRow("RX_P::RX_N", true)]

        public void GenDiffGroupName_ShouldReturnGeneratedGroupName_WhenNoPinMapSheet(string diffName, bool isNeedGenPinGroup)
        {
            string result = DifferentialService.GenDiffGroupName(null, diffName, isNeedGenPinGroup);
            StringAssert.Contains(result, "Diff");
        }

        [DataTestMethod]
        [DataRow("A::B")]
        [DataRow("TX_P_TX_N")]
        [DataRow("RX_P-RX_N")]
        public void DiffPinPosAndNeg_ShouldReturnFalse_ForInvalidString(string diffPins)
        {
            bool result = DifferentialService.DiffPinPosAndNeg(diffPins, out string pos, out string neg, out string group);

            Assert.IsFalse(result);
            Assert.AreEqual("", pos);
            Assert.AreEqual("", neg);
            Assert.AreEqual("", group);
        }
    }
}
