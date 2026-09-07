using Automation.Singleton;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Singleton
{
    [TestClass]
    public class SelSramWriterSingletonTests
    {
        private SelSramWriterSingleton _instance = null!;

        [TestInitialize]
        public void Setup()
        {
            SelSramWriterSingleton.Initialize();
            _instance = SelSramWriterSingleton.GetInstance();
        }

        [DataTestMethod]
        [DataRow("Soc", "S")]
        [DataRow("Cpu", "C")]
        [DataRow("Gfx", "L")]
        [DataRow("Scan", "SC")]
        [DataRow("Mbist", "BI")]
        [DataRow("Other", "Other")]
        public void GetAbbreviation_MapsKnownDomainsToAbbreviation(string input, string expected)
        {
            // Act
            string result = _instance.GetAbbreviation(input);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void GetSubName_FullRule_ReturnsWholeName()
        {
            // Act
            string result = _instance.GetSubName("A_B_C", "full");

            // Assert
            Assert.AreEqual("A_B_C", result);
        }

        [TestMethod]
        public void GetSubName_IndexListRule_SelectsWordsByIndex()
        {
            // Act
            string result = _instance.GetSubName("A_B_C", "0,2");

            // Assert
            Assert.AreEqual("A_C", result);
        }

        [TestMethod]
        public void GetSubName_EmptyRule_ReturnsEmpty()
        {
            // Act
            string result = _instance.GetSubName("A_B_C", "");

            // Assert
            Assert.AreEqual(string.Empty, result);
        }

        [TestMethod]
        public void GetSubName_IndexOutOfRange_SkipsThatIndex()
        {
            // Act
            string result = _instance.GetSubName("A_B_C", "0,5");

            // Assert
            Assert.AreEqual("A", result);
        }

        [TestMethod]
        public void GetPayloadType_EmptyPattern_ReturnsEmpty()
        {
            // Act
            string result = _instance.GetPayloadType("");

            // Assert
            Assert.AreEqual(string.Empty, result);
        }

        [TestMethod]
        public void GetPayloadType_NoRowsInTable_ReturnsEmpty()
        {
            // Arrange - the singleton constructs PayloadTypeTable as a fresh empty DataTable
            string result = _instance.GetPayloadType("SomePattern");

            // Assert
            Assert.AreEqual(string.Empty, result);
        }

    }
}
