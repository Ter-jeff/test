using System;
using System.Linq;

using CommonLib.Enums;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test.UT.Enums
{
    [TestClass]
    public class EnumInstrumentTests
    {
        [TestMethod]
        public void EnumInstrument_HasExpectedCount_FourValues()
        {
            // Act
            Array values = Enum.GetValues(typeof(EnumInstrument));

            // Assert
            Assert.AreEqual(4, values.Length);
        }

        [TestMethod]
        public void EnumInstrument_AllValuesAreUnique()
        {
            // Act
            var values = Enum.GetValues(typeof(EnumInstrument)).Cast<int>().ToList();
            var uniqueValues = values.Distinct().ToList();

            // Assert
            Assert.AreEqual(values.Count, uniqueValues.Count);
        }

        [TestMethod]
        public void EnumInstrument_CanParseFromString()
        {
            // Act
            object value = Enum.Parse(typeof(EnumInstrument), "UFlex");

            // Assert
            Assert.AreEqual(EnumInstrument.UFlex, value);
        }
    }
}
