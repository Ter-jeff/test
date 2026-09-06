using System.Collections.Generic;
using System.IO;

using FileDiffLib;

using IgxlLib.Utility.Tester;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

using Newtonsoft.Json;

namespace IgxlLib.Test.UT.Utility
{
    [TestClass]
    public class TesterConfigTests
    {
        public string InputPath = Path.Combine(Directory.GetCurrentDirectory(), "Input");
        public string OutputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        public string ExpectPath = Path.Combine(Directory.GetCurrentDirectory(), "Expected");

        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void TesterConfigTest()
        {
            string subName = "TesterConfigTest";
            string outputPath = Path.Combine(OutputPath, subName);
            string expectPath = Path.Combine(ExpectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            string file = Path.Combine(InputPath, "TesterConfig_AP.xml");
            Dictionary<string, TesterConfig> testerConfigs = TesterConfigReader.GetTesterConfigs(file);

            string instrumentType = TesterConfigManager.GetToolTypeByChannelAssignments(testerConfigs, ["14.CH85"], "Name");

            Assert.AreEqual("I/O", instrumentType);

            string json = JsonConvert.SerializeObject(testerConfigs, Formatting.Indented);
            File.WriteAllText(Path.Combine(outputPath, "result.json"), json);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void TesterConfig_DefaultConstructor_InitializesWithEmptyStrings()
        {
            // Arrange & Act
            var testerConfig = new TesterConfig();

            // Assert
            Assert.AreEqual("", testerConfig.Vsm);
            Assert.AreEqual("", testerConfig.Io);
            Assert.AreEqual("", testerConfig.HexVs);
            Assert.AreEqual("", testerConfig.Uvs64);
            Assert.AreEqual("", testerConfig.Uvs256);
        }

        [TestMethod]
        public void TesterConfig_SetProperties_UpdatesValues()
        {
            // Arrange
            var testerConfig = new TesterConfig
            {
                // Act
                Vsm = "VSM1,VSM2",
                Io = "IO1,IO2",
                HexVs = "HexVS1",
                Uvs256 = "UVS256_1,UVS256_2"
            };

            // Assert
            Assert.AreEqual("VSM1,VSM2", testerConfig.Vsm);
            Assert.AreEqual("IO1,IO2", testerConfig.Io);
            Assert.AreEqual("HexVS1", testerConfig.HexVs);
        }

        [TestMethod]
        public void TesterConfig_GetToolType_WithVsmChannel_ReturnsVSM()
        {
            // Arrange
            var testerConfig = new TesterConfig
            {
                Vsm = "VSM1,VSM2,VSM3"
            };

            // Act
            string result = testerConfig.GetToolType("VSM1");

            // Assert
            Assert.AreEqual("VSM", result);
        }

        [TestMethod]
        public void TesterConfig_GetToolType_WithIoChannel_ReturnsIO()
        {
            // Arrange
            var testerConfig = new TesterConfig
            {
                Io = "IO1,IO2"
            };

            // Act
            string result = testerConfig.GetToolType("IO1");

            // Assert
            Assert.AreEqual("I/O", result);
        }

        [TestMethod]
        public void TesterConfig_GetToolType_WithHexVsChannel_ReturnsHexVS()
        {
            // Arrange
            var testerConfig = new TesterConfig
            {
                HexVs = "HexVS1,HexVS2"
            };

            // Act
            string result = testerConfig.GetToolType("HexVS1");

            // Assert
            Assert.AreEqual("HexVS", result);
        }

        [TestMethod]
        public void TesterConfig_GetToolType_WithUvs256Channel_ReturnsUVS256()
        {
            // Arrange
            var testerConfig = new TesterConfig
            {
                Uvs256 = "UVS256_1,UVS256_2"
            };

            // Act
            string result = testerConfig.GetToolType("UVS256_1");

            // Assert
            Assert.AreEqual("UVS256", result);
        }

        [TestMethod]
        public void TesterConfig_GetToolType_CaseInsensitive()
        {
            // Arrange
            var testerConfig = new TesterConfig
            {
                Vsm = "VSM1,VSM2"
            };

            // Act
            string result = testerConfig.GetToolType("vsm1");

            // Assert
            Assert.AreEqual("VSM", result);
        }

        [TestMethod]
        public void TesterConfig_AllProperties_CanBeSet()
        {
            // Arrange
            var testerConfig = new TesterConfig
            {
                // Act
                Vsm = "VSM",
                Io = "IO",
                HexVs = "HexVS",
                Uvs64 = "UVS64",
                Uvs256 = "UVS256",
                Uvs256Hp = "UVS256HP",
                Uvi80 = "UVI80",
                UltraPac = "UltraPac",
                Dc30 = "DC30",
                Us10G = "US10G",
                Uw24Source = "UW24Source",
                Uw24Measure = "UW24Measure",
                Mx8 = "MX8",
                Support = "Support"
            };

            // Assert
            Assert.AreEqual("VSM", testerConfig.Vsm);
            Assert.AreEqual("IO", testerConfig.Io);
            Assert.AreEqual("HexVS", testerConfig.HexVs);
            Assert.AreEqual("UVS256", testerConfig.Uvs256);
            Assert.AreEqual("UVS256HP", testerConfig.Uvs256Hp);
        }

        [TestMethod]
        public void TesterConfig_GetToolType_WithUvs256HpChannel()
        {
            // Arrange
            var testerConfig = new TesterConfig { Uvs256Hp = "UVS256HP1" };

            // Act
            string result = testerConfig.GetToolType("UVS256HP1");

            // Assert
            Assert.AreEqual("UVS256HP", result);
        }

        [TestMethod]
        public void TesterConfig_GetToolType_WithUvs64Channel()
        {
            // Arrange
            var testerConfig = new TesterConfig { Uvs64 = "UVS64_1" };

            // Act
            string result = testerConfig.GetToolType("UVS64_1");

            // Assert
            Assert.AreEqual("UVS64", result);
        }

        [TestMethod]
        public void TesterConfig_GetToolType_WithUvi80Channel()
        {
            // Arrange
            var testerConfig = new TesterConfig { Uvi80 = "UVI80_1" };

            // Act
            string result = testerConfig.GetToolType("UVI80_1");

            // Assert
            Assert.AreEqual("UVI80", result);
        }

        [TestMethod]
        public void TesterConfig_GetToolType_WithUltraPacChannel()
        {
            // Arrange
            var testerConfig = new TesterConfig { UltraPac = "UltraPAC1" };

            // Act
            string result = testerConfig.GetToolType("UltraPAC1");

            // Assert
            Assert.AreEqual("UltraPAC", result);
        }

        [TestMethod]
        public void TesterConfig_GetToolType_WithDc30Channel()
        {
            // Arrange
            var testerConfig = new TesterConfig { Dc30 = "DC30_1" };

            // Act
            string result = testerConfig.GetToolType("DC30_1");

            // Assert
            Assert.AreEqual("DC30", result);
        }

        [TestMethod]
        public void TesterConfig_GetToolType_WithUs10GChannel()
        {
            // Arrange
            var testerConfig = new TesterConfig { Us10G = "US10G_1" };

            // Act
            string result = testerConfig.GetToolType("US10G_1");

            // Assert
            Assert.AreEqual("US10G", result);
        }

        [TestMethod]
        public void TesterConfig_GetToolType_WithSupportChannel()
        {
            // Arrange
            var testerConfig = new TesterConfig { Support = "Support1" };

            // Act
            string result = testerConfig.GetToolType("Support1");

            // Assert
            Assert.AreEqual("Support", result);
        }

        [TestMethod]
        public void TesterConfig_GetToolType_WithMultipleChannels_ReturnsFirstMatch()
        {
            // Arrange
            var testerConfig = new TesterConfig
            {
                Vsm = "VSM1,VSM2,VSM3"
            };

            // Act
            string result1 = testerConfig.GetToolType("VSM1");
            string result2 = testerConfig.GetToolType("VSM2");
            string result3 = testerConfig.GetToolType("VSM3");

            // Assert
            Assert.AreEqual("VSM", result1);
            Assert.AreEqual("VSM", result2);
            Assert.AreEqual("VSM", result3);
        }

        [TestMethod]
        public void TesterConfig_GetToolType_WithUnknownChannel_ReturnsEmpty()
        {
            // Arrange
            var testerConfig = new TesterConfig { Vsm = "VSM1" };

            // Act
            string result = testerConfig.GetToolType("UnknownChannel");

            // Assert
            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void TesterConfig_MultipleProperties_CanBeTestedConcurrently()
        {
            // Arrange
            var config1 = new TesterConfig { Vsm = "VSM1", Io = "IO1" };
            var config2 = new TesterConfig { HexVs = "HexVS1", Uvs256 = "UVS256_1" };

            // Act
            string type1Vsm = config1.GetToolType("VSM1");
            string type1Io = config1.GetToolType("IO1");
            string type2Hex = config2.GetToolType("HexVS1");
            string type2Uvs = config2.GetToolType("UVS256_1");

            // Assert
            Assert.AreEqual("VSM", type1Vsm);
            Assert.AreEqual("I/O", type1Io);
            Assert.AreEqual("HexVS", type2Hex);
            Assert.AreEqual("UVS256", type2Uvs);
        }

        [TestMethod]
        public void TesterConfig_Uw24_Properties()
        {
            // Arrange
            var testerConfig = new TesterConfig
            {
                Uw24Source = "Source1",
                Uw24Measure = "Measure1"
            };

            // Assert
            Assert.AreEqual("Source1", testerConfig.Uw24Source);
            Assert.AreEqual("Measure1", testerConfig.Uw24Measure);
        }

        [TestMethod]
        public void TesterConfig_Mx8_Property()
        {
            // Arrange
            var testerConfig = new TesterConfig { Mx8 = "MX8_1" };

            // Assert
            Assert.AreEqual("MX8_1", testerConfig.Mx8);
        }
    }
}
