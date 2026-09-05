using System;
using System.Linq;

using CommonLib.Enums;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test.UT.Enums
{
    [TestClass]
    public class EnumEquipmentTests
    {
        [TestMethod]
        public void EnumEquipment_HasExpectedCount_TwoValues()
        {
            // Act
            Array values = Enum.GetValues(typeof(EnumEquipment));

            // Assert
            Assert.AreEqual(2, values.Length);
        }

        [TestMethod]
        public void EnumEquipment_AllValuesAreUnique()
        {
            // Act
            var values = Enum.GetValues(typeof(EnumEquipment)).Cast<int>().ToList();
            var uniqueValues = values.Distinct().ToList();

            // Assert
            Assert.AreEqual(values.Count, uniqueValues.Count);
        }

        [TestMethod]
        public void EnumEquipment_CanParseFromString()
        {
            // Act
            object value = Enum.Parse(typeof(EnumEquipment), "UltraFlex");

            // Assert
            Assert.AreEqual(EnumEquipment.UltraFlex, value);
        }

        [TestMethod]
        public void EnumEquipment_CanParseUltraFlexPlusFromString()
        {
            // Act
            object value = Enum.Parse(typeof(EnumEquipment), "UltraFlexPlus");

            // Assert
            Assert.AreEqual(EnumEquipment.UltraFlexPlus, value);
        }
    }
}
