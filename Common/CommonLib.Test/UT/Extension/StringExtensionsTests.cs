using System;
using System.Collections.Generic;
using System.IO;

using CommonLib.Extension;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test.UT.Extension
{
    [TestClass]
    public class StringExtensionsTests
    {
        private string _tempDir;

        [TestInitialize]
        public void Setup()
        {
            _tempDir = Path.Combine(Path.GetTempPath(), "StringExtensionsTests_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_tempDir);
        }

        [TestCleanup]
        public void Teardown()
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, recursive: true);
            }
        }

        [TestMethod]
        public void CalBitWidth_ValidInputs_ReturnsCorrectWidth()
        {
            Assert.AreEqual("8", "0".CalBitWidth("7"));
            Assert.AreEqual("10", "1".CalBitWidth("10"));
            Assert.AreEqual("16", "15".CalBitWidth("0"));
        }

        [TestMethod]
        public void IsHexValue_ValidHexInput_ReturnsTrue()
        {
            Assert.IsTrue("0x1A2B".IsHexValue());
            Assert.IsTrue("0xFFFF".IsHexValue());
            Assert.IsTrue("0x0".IsHexValue());
        }

        [TestMethod]
        public void IsHexValue_InvalidHexInput_ReturnsFalse()
        {
            Assert.IsFalse("1A2B".IsHexValue());
            Assert.IsFalse("0xG123".IsHexValue());
            Assert.IsFalse("x123".IsHexValue());
        }

        [TestMethod]
        public void IsHexOrOctValue_HexInput_ReturnsTrue()
        {
            Assert.IsTrue("0x1A".IsHexOrOctValue());
            Assert.IsTrue("x1A".IsHexOrOctValue());
            Assert.IsTrue("1A".IsHexOrOctValue());
            Assert.IsTrue("h1A".IsHexOrOctValue());
        }

        [TestMethod]
        public void IsHexOrOctValue_OctInput_ReturnsTrue()
        {
            Assert.IsTrue("0b1010".IsHexOrOctValue());
            Assert.IsTrue("b1010".IsHexOrOctValue());
            Assert.IsTrue("1010".IsHexOrOctValue());
        }

        [TestMethod]
        public void IsHexOrOctValue_InvalidInput_ReturnsFalse()
        {
            Assert.IsFalse("xyz".IsHexOrOctValue());
            Assert.IsFalse("zzzzz".IsHexOrOctValue());
        }

        [TestMethod]
        public void IsNumber_ValidNumber_ReturnsTrue()
        {
            Assert.IsTrue("123".IsNumber());
            Assert.IsTrue("123.456".IsNumber());
            Assert.IsTrue("-123".IsNumber());
            Assert.IsTrue("+123.456".IsNumber());
        }

        [TestMethod]
        public void IsNumber_InvalidNumber_ReturnsFalse()
        {
            Assert.IsFalse("abc".IsNumber());
            Assert.IsFalse("".IsNumber());
            Assert.IsFalse("abc123".IsNumber());
        }

        [TestMethod]
        public void EFuseBitToInt_BinaryInput_ReturnsCorrectInt()
        {
            Assert.AreEqual("10", "b1010".EFuseBitToInt());
            Assert.AreEqual("15", "b1111".EFuseBitToInt());
        }

        [TestMethod]
        public void EFuseBitToInt_HexInput_ReturnsCorrectInt()
        {
            Assert.AreEqual("0xA", "0xA".EFuseBitToInt());
            Assert.AreEqual("15", "xF".EFuseBitToInt());
            Assert.AreEqual("255", "hFF".EFuseBitToInt());
        }

        [TestMethod]
        public void EFuseBitToInt_InvalidInput_ReturnsOriginalText()
        {
            Assert.AreEqual("0xyz", "xyz".EFuseBitToInt());
        }

        [TestMethod]
        public void EFuseBitToInt_UnderscoreInput_ReturnsCorrectValue()
        {
            Assert.AreEqual("0b1010_1100", "0b1010_1100".EFuseBitToInt());
        }

        [TestMethod]
        public void Join_ValidConnectorAndComponents_ReturnsJoined()
        {
            Assert.AreEqual("A-B-C", "-".Join("A", "B", "C"));
        }

        [TestMethod]
        public void Join_EmptyConnector_ReturnsConcatenated()
        {
            Assert.AreEqual("ABC", "".Join("A", "B", "C"));
        }

        [TestMethod]
        public void Join_NoComponents_ReturnsEmpty()
        {
            Assert.AreEqual("", "-".Join());
        }

        [TestMethod]
        public void BitToInt_BinaryWithPrefix_ReturnsInteger()
        {
            Assert.AreEqual("10", "0b1010".BitToInt());
        }

        [TestMethod]
        public void BitToInt_HexWithPrefix_ReturnsInteger()
        {
            Assert.AreEqual("255", "0xFF".BitToInt());
        }

        [TestMethod]
        public void BitToInt_NumericString_ReturnsSameNumber()
        {
            Assert.AreEqual("1010", "1010".BitToInt());
        }

        [TestMethod]
        public void BitToInt_InvalidString_ReturnsEmpty()
        {
            Assert.AreEqual("", "test".BitToInt());
        }

        [TestMethod]
        public void BitToInt_WithUnderscores_StripsAndConverts()
        {
            Assert.AreEqual("172", "0b1010_1100".BitToInt());
        }

        [TestMethod]
        public void BitToInt_AlreadyNumber_ReturnsAsIs()
        {
            Assert.AreEqual("123.45", "123.45".BitToInt());
        }

        [TestMethod]
        public void ToLines_MultilineText_ReturnsList()
        {
            List<string> result = "Line1\nLine2\nLine3".ToLines();
            Assert.AreEqual(3, result.Count);
            Assert.AreEqual("Line1", result[0]);
            Assert.AreEqual("Line2", result[1]);
            Assert.AreEqual("Line3", result[2]);
        }

        [TestMethod]
        public void ToLines_EmptyString_ReturnsSingleEmptyLine()
        {
            List<string> result = "".ToLines();
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("", result[0]);
        }

        [TestMethod]
        public void PadCenter_ShortText_PadsCorrectly()
        {
            string result = "Hi".PadCenter(10);
            Assert.AreEqual(10, result.Length);
            Assert.IsTrue(result.Contains("Hi"));
        }

        // ─── CsvConvertToLists (string extension) ─────────────────────────────
        [TestMethod]
        public void CsvConvertToLists_SimpleCsv_ReturnsCorrectLists()
        {
            string csvFile = Path.Combine(_tempDir, "simple.csv");
            File.WriteAllText(csvFile, "a,b,c\n1,2,3");

            List<List<string>> result = csvFile.CsvConvertToLists();

            Assert.AreEqual(2, result.Count);
            Assert.AreEqual("a", result[0][0]);
            Assert.AreEqual("b", result[0][1]);
            Assert.AreEqual("c", result[0][2]);
            Assert.AreEqual("1", result[1][0]);
        }

        [TestMethod]
        public void CsvConvertToLists_CommaInQuotes_PreservesComma()
        {
            string csvFile = Path.Combine(_tempDir, "quoted.csv");
            File.WriteAllText(csvFile, "\"a,b\",c\n1,2");

            List<List<string>> result = csvFile.CsvConvertToLists();

            Assert.AreEqual(2, result.Count);
            Assert.AreEqual("a,b", result[0][0]);
            Assert.AreEqual("c", result[0][1]);
        }

        [TestMethod]
        public void CsvConvertToLists_SingleRow_ReturnsOneList()
        {
            string csvFile = Path.Combine(_tempDir, "single.csv");
            File.WriteAllText(csvFile, "x,y,z");

            List<List<string>> result = csvFile.CsvConvertToLists();

            Assert.AreEqual(1, result.Count);
            Assert.AreEqual(3, result[0].Count);
        }
    }
}
