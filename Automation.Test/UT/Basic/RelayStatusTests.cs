using Automation.GenerateIgxl.Basic.Business.GenConti.Base;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Basic
{

    [TestClass]
    public class RelayStatusTests
    {
        [TestMethod]
        public void IsEqualStatus_BothListsEmpty_ReturnsTrue()
        {
            // Arrange
            var status1 = new RelayStatus();
            var status2 = new RelayStatus();

            // Act
            bool result = status1.IsEqualStatus(status2);

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void IsEqualStatus_SameItemsDifferentOrder_ReturnsTrue()
        {
            // Arrange
            var status1 = new RelayStatus { OpenRelayList = ["R1", "R2"] };
            var status2 = new RelayStatus { OpenRelayList = ["R2", "R1"] };

            // Act
            bool result = status1.IsEqualStatus(status2);

            // Assert
            Assert.IsTrue(result, "Lists with same items in different order should be equal.");
        }

        [TestMethod]
        [DataRow("R1", "R2", false)] // Different items
        [DataRow("R1", "R1,R2", false)] // One is a subset
        public void IsEqualStatus_MismatchedLists_ReturnsFalse(string list1, string list2, bool expected)
        {
            // Arrange
            var status1 = new RelayStatus { OpenRelayList = [.. list1.Split(',')] };
            var status2 = new RelayStatus { OpenRelayList = [.. list2.Split(',')] };

            // Act
            bool result = status1.IsEqualStatus(status2);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void IsEqualStatus_OffRelayListIsIgnored_VerifyCurrentBehavior()
        {
            // Arrange - Open lists match, but Off lists differ
            var status1 = new RelayStatus
            {
                OpenRelayList = ["R1"],
                OffRelayList = ["Off1"]
            };
            var status2 = new RelayStatus
            {
                OpenRelayList = ["R1"],
                OffRelayList = ["Off2"]
            };

            // Act
            bool result = status1.IsEqualStatus(status2);

            // Assert
            // This passes because your current code doesn't check OffRelayList
            Assert.IsTrue(result);
        }
    }

}
