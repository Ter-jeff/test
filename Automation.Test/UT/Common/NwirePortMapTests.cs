using Automation.GenerateIgxl.Basic.Business.GenNwire.Business;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Common
{
    [TestClass]
    public class NwirePortMapTests
    {
        [TestMethod]
        public void Protocol_With_Valid_A_B_Format_Should_Parse_Correctly()
        {
            // Arrange
            string protocol = "FRC:RefClock";
            string defaultA = "A";
            string defaultB = "B";

            // Act
            (string family, string type) = NwirePortMap.ParseProtocolOrDefault(protocol, defaultA, defaultB);

            // Assert
            Assert.AreEqual("FRC", family);
            Assert.AreEqual("RefClock", type);
        }

        [TestMethod]
        public void Protocol_Invalid_Format_Should_Return_Default()
        {
            // Arrange
            string protocol = "INVALID";
            string defaultA = "A";
            string defaultB = "B";

            // Act
            (string family, string type) = NwirePortMap.ParseProtocolOrDefault(protocol, defaultA, defaultB);

            // Assert
            Assert.AreEqual(defaultA, family);
            Assert.AreEqual(defaultB, type);
        }

        [TestMethod]
        public void Protocol_With_Empty_A_Should_Return_Default()
        {
            // Arrange
            string protocol = ":RefClock";
            string defaultA = "A";
            string defaultB = "B";

            // Act
            (string family, string type) = NwirePortMap.ParseProtocolOrDefault(protocol, defaultA, defaultB);
            Assert.AreEqual(defaultA, family);
            Assert.AreEqual(defaultB, type);
        }

        [TestMethod]
        public void Protocol_With_Empty_B_Should_Return_Default()
        {
            // Arrange
            string protocol = "FRC:";
            string defaultA = "A";
            string defaultB = "B";

            // Act
            (string family, string type) = NwirePortMap.ParseProtocolOrDefault(protocol, defaultA, defaultB);
            Assert.AreEqual(defaultA, family);
            Assert.AreEqual(defaultB, type);
        }

    }
}
