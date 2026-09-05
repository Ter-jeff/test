using System.Data;
using System.IO;
using System.Xml;

using Automation.Reader.ConfigFile.NamingRule.Base;
using Automation.Reader.ConfigFile.NamingRule.Business;
using Automation.Static;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class HardIpConfigFileReaderTests
    {
        private string _testDir = null!;
        private string _configPath = null!;

        [TestInitialize]
        public void Setup()
        {
            _testDir = Path.Combine(Path.GetTempPath(), "HardIpConfigTest");
            Directory.CreateDirectory(_testDir);

            _configPath = Path.Combine(_testDir, "HardIP_Config_Default.xml");
            LocalSpecs.CurrentProject = "UT";
            LocalSpecs.SettingFolder = _testDir;
            LocalSpecs.SettingFiles = new SettingFiles();
        }

        [TestCleanup]
        public void Cleanup()
        {
            if (Directory.Exists(_testDir))
            {
                Directory.Delete(_testDir, true);
            }
        }

        [TestMethod]
        public void ReadConfig_WhenFileNotExists_ShouldReturnEmptyConfig()
        {
            // Arrange
            var reader = new HardIpConfigFileReader();

            // Act
            HardIpConfig config = reader.ReadConfig();

            // Assert
            Assert.AreNotEqual(null, config);
            Assert.AreNotEqual(null, config.NamingRulesList);
            Assert.AreEqual(1, config.NamingRulesList.Count);
        }

        [TestMethod]
        public void ReadConfig_WithSplitRuleAndPayload_ShouldParseSuccessfully()
        {
            // Arrange
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>
<HardIPConfig>
  <NamingRule>
    <HardIP>
      <init1></init1>
      <init2></init2>
      <init3>9</init3>
      <init4></init4>
      <init5></init5>
      <init6></init6>
      <init7></init7>
      <init8></init8>
      <init9></init9>
      <init10></init10>
      <payload>full</payload>
    </HardIP>
  </NamingRule>
  <SplitRule>SplitRule</SplitRule>
  <PayloadType>
   <KeyPosition>KeyPosition</KeyPosition>
   <Key>Key</Key>
  </PayloadType>
  <SpecialFlagSetting>SpecialFlagSetting</SpecialFlagSetting>
</HardIPConfig>";
            File.WriteAllText(_configPath, xml);

            // Act
            LocalSpecs.SettingFiles.HardipConfig = _configPath;
            var reader = new HardIpConfigFileReader();
            HardIpConfig config = reader.ReadConfig();

            // Assert
            Assert.AreNotEqual(null, config);
            Assert.AreEqual(1, config.SplitRuleList.Count);
            Assert.AreEqual("#text", config.SplitRuleList[0].SubName);
            Assert.AreEqual("SplitRule", config.SplitRuleList[0].SubRule);

            Assert.AreNotEqual(null, config.PayloadType);
            Assert.AreEqual(1, config.PayloadType.Rows.Count);
            Assert.AreEqual("#text", config.PayloadType.Rows[0]["key"]);
        }

        [TestMethod]
        public void ReadPayloadType_ShouldParseCorrectly()
        {
            // Arrange
            string xml = @"
<PayloadType>
    <KeyPosition>pos1,pos2</KeyPosition>
    <Key>
        <Alpha>v1,v2</Alpha>
        <Beta>v3,v4</Beta>
    </Key>
</PayloadType>";

            var doc = new XmlDocument();
            doc.LoadXml(xml);
            XmlElement? pNode = doc.DocumentElement;

            var reader = new HardIpConfigFileReader();

            // Act
            DataTable result = reader.ReadPayloadType(pNode);

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.AreEqual(3, result.Columns.Count);
            Assert.AreEqual(2, result.Rows.Count);

            Assert.AreEqual("Alpha", result.Rows[0]["key"]);
            Assert.AreEqual("v1", result.Rows[0]["pos1"]);
            Assert.AreEqual("v2", result.Rows[0]["pos2"]);

            Assert.AreEqual("Beta", result.Rows[1]["key"]);
            Assert.AreEqual("v3", result.Rows[1]["pos1"]);
            Assert.AreEqual("v4", result.Rows[1]["pos2"]);
        }

        [TestMethod]
        public void ReadPayloadType_WhenNoKeyPosition_ShouldReturnEmptyTable()
        {
            // Arrange
            string xml = @"
<PayloadType>
    <Key>
        <Alpha>v1,v2</Alpha>
    </Key>
</PayloadType>";

            var doc = new XmlDocument();
            doc.LoadXml(xml);
            XmlElement? pNode = doc.DocumentElement;

            var reader = new HardIpConfigFileReader();

            // Act
            DataTable result = reader.ReadPayloadType(pNode);

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.AreEqual(0, result.Columns.Count);
            Assert.AreEqual(0, result.Rows.Count);
        }

        [TestMethod]
        public void ReadPayloadType_WhenKeyPositionEmpty_ShouldReturnEmptyColumns()
        {
            // Arrange
            string xml = @"
<PayloadType>
    <KeyPosition></KeyPosition>
</PayloadType>";

            var doc = new XmlDocument();
            doc.LoadXml(xml);
            XmlElement? pNode = doc.DocumentElement;

            var reader = new HardIpConfigFileReader();

            // Act
            DataTable result = reader.ReadPayloadType(pNode);

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.AreEqual(2, result.Columns.Count);
            Assert.AreEqual("key", result.Columns[0].ColumnName);
            Assert.AreEqual(0, result.Rows.Count);
        }
    }
}
