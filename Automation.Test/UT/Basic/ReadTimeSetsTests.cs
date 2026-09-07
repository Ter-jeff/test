using System;
using System.Collections.Generic;

using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Business;

using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Basic
{
    [TestClass]
    public class ReadTimeSetsTests
    {
        private TimeSetGenerator _reader = null!;

        [TestInitialize]
        public void Init()
        {
            _reader = new TimeSetGenerator();
        }

        [TestMethod]
        public void GetEgValueInDecimal_ShouldParseNormalNumber()
        {
            decimal result = _reader.GetEgValueInDecimal("123.45");
            Assert.AreEqual(123.45m, result);
        }

        [TestMethod]
        public void GetEgValueInDecimal_Exception()
        {
            bool isException = false;
            try
            {
                decimal result = _reader.GetEgValueInDecimal("xxx");
            }
            catch (Exception)
            {
                isException = true;
            }
            Assert.IsTrue(isException);
        }

        [TestMethod]
        public void GetEgValueInDecimal_ShouldParseScientificNotation()
        {
            decimal result = _reader.GetEgValueInDecimal("1.23E3");
            Assert.AreEqual(1230m, result);
        }

        [TestMethod]
        public void IgnoredFileName_ShouldReturnTrue_WhenNAorNAString()
        {
            Assert.IsTrue(_reader.IgnoredFileName("NA"));
            Assert.IsTrue(_reader.IgnoredFileName("n/a"));
            Assert.IsTrue(_reader.IgnoredFileName("N/A"));
        }

        [TestMethod]
        public void IgnoredFileName_ShouldReturnFalse_ForNormalNames()
        {
            Assert.IsFalse(_reader.IgnoredFileName("MyTimeSet"));
        }

        [TestMethod]
        public void IsContextVariable_ShouldDetectVariablePattern()
        {
            bool result = _reader.IsContextVariable("=_Cycle_S_VAR+0.1/_ShiftIn_Freq_VAR");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void IsContextVariable_ShouldReturnFalse_ForPlainNumbers()
        {
            bool result = _reader.IsContextVariable("123.45");
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void GetContextVariable_ShouldExtractUniqueVars()
        {
            var list = new List<string>();
            _reader.GetContextVariable("=_Cycle_S_VAR+0.1/_ShiftIn_Freq_VAR+_Strobe_VAR", ref list);

            Assert.AreEqual(3, list.Count);
            CollectionAssert.Contains(list, "Cycle_S_VAR");
            CollectionAssert.Contains(list, "ShiftIn_Freq_VAR");
            CollectionAssert.Contains(list, "Strobe_VAR");
        }

        [TestMethod]
        public void Converter_ShouldReturnTimeRow1P4Converter_ByDefault()
        {
            TimeRow1P4Converter converter = TimeSetGenerator.Converter("Dummy header");
            Assert.AreEqual("TimeRow1P4Converter", converter.GetType().Name);
        }

        [TestMethod]
        public void Converter_ShouldReturnTimeRow2P3Converter_WhenHeaderContainsVersion23()
        {
            TimeRow1P4Converter converter = TimeSetGenerator.Converter("DTTimesetBasicSheet,version=2.3");
            Assert.AreEqual("TimeRow2P3Converter", converter.GetType().Name);
        }
    }
}
