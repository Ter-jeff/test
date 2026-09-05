using System;
using System.Linq;

using CommonLib.Enums;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test.UT.Enums
{
    [TestClass]
    public class EnumPatternTypeTests
    {
        [TestMethod]
        public void EnumPatternType_HasExpectedCount_SixValues()
        {
            // Act
            Array values = Enum.GetValues(typeof(EnumPatternType));

            // Assert
            Assert.AreEqual(6, values.Length);
        }

        [TestMethod]
        public void EnumPatternType_AllValuesAreUnique()
        {
            // Act
            var values = Enum.GetValues(typeof(EnumPatternType)).Cast<int>().ToList();
            var uniqueValues = values.Distinct().ToList();

            // Assert
            Assert.AreEqual(values.Count, uniqueValues.Count);
        }
    }
}
