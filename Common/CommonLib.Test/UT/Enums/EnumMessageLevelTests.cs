using System;
using System.Linq;

using CommonLib.Enums;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test.UT.Enums
{
    [TestClass]
    public class EnumMessageLevelTests
    {
        [TestMethod]
        public void MessageLevel_HasExpectedCount_EightValues()
        {
            // Act
            Array values = Enum.GetValues(typeof(EnumMessageLevel));

            // Assert
            Assert.AreEqual(8, values.Length);
        }

        [TestMethod]
        public void MessageLevel_AllValuesAreUnique()
        {
            // Act
            var values = Enum.GetValues(typeof(EnumMessageLevel)).Cast<int>().ToList();
            var uniqueValues = values.Distinct().ToList();

            // Assert
            Assert.AreEqual(values.Count, uniqueValues.Count);
        }
    }
}
