using System.Collections.Generic;
using System.Data;

using Automation.Reader;

using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Common
{
    [TestClass]
    public class NwireSettingTests
    {
        private DataTable _mockTable = null!;

        [TestInitialize]
        public void Setup()
        {
            // Initialize a common table structure for tests
            _mockTable = new DataTable();
            _mockTable.Columns.Add("SettingValue", typeof(string));
        }

        [TestMethod]
        public void ReferenceFlow_WhenNoRegexMatch_ReturnsAllRows()
        {
            // Arrange
            _mockTable.Rows.Add("Flow1");
            _mockTable.Rows.Add("Flow2");
            var sut = new NwireSetting();

            // Act
            List<string> result = sut.ReferenceFlow(_mockTable);

            // Assert
            Assert.AreEqual(2, result.Count);
            Assert.AreEqual("Flow1", result[0]);
        }

        [TestMethod]
        public void GetNwireCall_WhenMatchFound_ReturnsPrefixedParameter()
        {
            // Arrange
            string keyword = "LoginFlow";
            _mockTable.Rows.Add(keyword);
            var sut = new NwireSetting
            {
                SettingTable = _mockTable
            };

            // Act
            FlowRow result = sut.GetNwireCall("LoginFlow");

            // Assert
            Assert.AreEqual("Call", result.Opcode);
            Assert.IsTrue(result.Parameter.EndsWith(keyword));
        }

        [TestMethod]
        [DataRow("MissingFlow", "Flow_nWire_Default")]
        public void GetNwireCall_WhenNoMatch_ReturnsDefaultParameter(string input, string expected)
        {
            // Arrange
            _mockTable.Rows.Add("OtherFlow");
            var sut = new NwireSetting
            {
                SettingTable = _mockTable
            };

            // Act
            FlowRow result = sut.GetNwireCall(input);

            // Assert
            Assert.AreEqual(expected, result.Parameter);
        }
    }
}
