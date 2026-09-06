using System;

using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class PortRowTests
    {
        [TestMethod]
        public void PortRow_SetProperties_UpdatesValuesCorrectly()
        {
            // Arrange
            var portRow = new PortRow
            {
                // Act
                PortName = "Port1",
                ProtocolFamily = "SerDes",
                ProtocolType = "8B10B",
                FunctionName = "HighSpeedIO",
                FunctionPin = "Pin1",
                Comment = "High-speed port"
            };

            // Assert
            Assert.AreEqual("Port1", portRow.PortName);
            Assert.AreEqual("SerDes", portRow.ProtocolFamily);
            Assert.AreEqual("8B10B", portRow.ProtocolType);
            Assert.AreEqual("HighSpeedIO", portRow.FunctionName);
        }

        [TestMethod]
        public void PortRow_AddSetting_AddsValueToList()
        {
            // Arrange
            var portRow = new PortRow();

            // Act
            portRow.AddSetting("Setting1");
            portRow.AddSetting("Setting2");

            // Assert
            Assert.AreEqual(2, portRow.ProtocolSettingValues.Count);
            Assert.AreEqual("Setting1", portRow.ProtocolSettingValues[0]);
            Assert.AreEqual("Setting2", portRow.ProtocolSettingValues[1]);
        }

        [TestMethod]
        public void PortRow_AddSetting_ExceedsMaxNumber_ThrowsException()
        {
            // Arrange
            var portRow = new PortRow();

            // Act & Assert
            // Add 11 settings - should throw on the 11th one
            try
            {
                for (int i = 0; i < 11; i++)
                {
                    portRow.FunctionPropertyValues.Add(i.ToString());
                    portRow.AddSetting($"Setting{i}");
                }
                Assert.Fail("Should have thrown an exception");
            }
            catch (Exception ex)
            {
                Assert.IsTrue(ex.Message.Contains("exceed"));
            }
        }

        [TestMethod]
        public void PortRow_InitializeCollections_HasEmptyLists()
        {
            // Arrange & Act
            var portRow = new PortRow();

            // Assert
            Assert.AreEqual(0, portRow.ProtocolSettingValues.Count);
            Assert.AreEqual(0, portRow.FunctionPropertyValues.Count);
        }

        [TestMethod]
        public void PortRow_Inherits_FromIgxlRow()
        {
            // Arrange & Act
            var portRow = new PortRow();

            // Assert
            Assert.IsInstanceOfType(portRow, typeof(IgxlRow));
        }
    }
}
