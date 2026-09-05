using System;
using System.Linq;

using CommonLib.Enums;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test.UT.Enums
{
    [TestClass]
    public class EnumDeviceTests
    {
        [TestMethod]
        public void EnumDevice_HasExpectedCount_FiveValues()
        {
            // Act
            Array values = Enum.GetValues(typeof(EnumDevice));

            // Assert
            Assert.AreEqual(5, values.Length);
        }

        [TestMethod]
        public void EnumDevice_AllValuesAreUnique()
        {
            // Act
            var values = Enum.GetValues(typeof(EnumDevice)).Cast<int>().ToList();
            var uniqueValues = values.Distinct().ToList();

            // Assert
            Assert.AreEqual(values.Count, uniqueValues.Count);
        }
    }
}
