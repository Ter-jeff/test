using System.Collections.Generic;
using System.IO;

using Automation.GenerateIgxl.BistBira;
using Automation.Reader.ConfigFile.NamingRule.Base;
using Automation.Static;

using CommonLib.Enums;

using FileDiffLib;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Newtonsoft.Json;

using ScghLib.Enums;
using ScghLib.Reader;

namespace Automation.Test.UT.BistBira
{
    [TestClass]
    public class MbistFlowGeneratorTests : FunctionTestBase
    {
        [TestMethod]
        public void CreateBinTable()
        {
            string subName = "CreateBinTable";
            string outputPath = Path.Combine(OutputPath, "BistBira", subName);
            string expectPath = Path.Combine(ExpectPath, "BistBira", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            LocalSpecs.TarFolder = outputPath;
            var config = new MbistConfig();
            var prodFlowSheet = new BistProdFlowSheet
            {
                SheetName = "Mbist_Test",
            };
            var needBinoutList = new List<BinTableRow>
            {
                new() { Name = "binName", ItemList = "flag", Items = ["T"], Op = "AND" }
            };

            // Act
            var mbistFlowGenerator = new MbistFlowGenerator(prodFlowSheet, config, MbistBinTableType.Burst, false);
            List<BinTableRow> result = mbistFlowGenerator.CreateBinTable(needBinoutList, true);

            // Assert
            string json = JsonConvert.SerializeObject(result, Formatting.Indented);
            File.WriteAllText(Path.Combine(outputPath, "result.json"), json);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void GenerateSubflow()
        {
            string subName = "GenerateSubflow";
            string outputPath = Path.Combine(OutputPath, "BistBira", subName);
            string expectPath = Path.Combine(ExpectPath, "BistBira", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            LocalSpecs.TarFolder = outputPath;
            var config = new MbistConfig
            {
                RepairFlagSetting = new Dictionary<string, List<KeyAndPosition>>()
                {
                    { "" , new List<KeyAndPosition>() { new() { KeyPositions = "0", Keys = "P1" } } }
                }
            };
            var prodFlowSheet = new BistProdFlowSheet
            {
                SheetName = "Mbist_Test",
                Rows =
                [
                    new()
                    {
                        RowNum = 101,
                        Label = "AA",
                        Pattern = "",
                        Note = "",
                        Voltage = "Bincut_X_X_X NV",
                        SheetName = "Test",
                        PassBranch = "B1",
                        FailBranch = "B2",
                        IsDsscRow = true,
                        OriPerformance = "MS001,MS002",
                        BurstPatterns = ["P1", "P2"]
                    },
                    new()
                    {
                        RowNum = 101,
                        Label = "AA",
                        Pattern = "",
                        Note = "",
                        Voltage = "Bincut_X_X_X NV",
                        SheetName = "Test",
                        PassBranch = "B1",
                        FailBranch = "B2",
                        IsDsscRow = true,
                        OriPerformance = "MS001,MS002",
                        Action = "PASS"
                    }
                ],
            };
            prodFlowSheet.CreateLabelDic();

            // Act
            var mbistFlowGenerator = new MbistFlowGenerator(prodFlowSheet, config, MbistBinTableType.Burst, true);
            (SubFlowSheet, List<BinTableRow>) result = mbistFlowGenerator.GenerateSubflow();

            // Assert
            string json = JsonConvert.SerializeObject(result, Formatting.Indented);
            File.WriteAllText(Path.Combine(outputPath, "result.json"), json);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestCleanup]
        public void ResetDevice()
        {
            LocalSpecs.Options.Device = EnumDevice.AP;
        }

        [TestMethod]
        public void CheckBranchType_BranchNotFound_ReturnsNull()
        {
            // Arrange
            var config = new MbistConfig();
            var prodFlowSheet = new BistProdFlowSheet { SheetName = "Mbist_Test" };
            prodFlowSheet.CreateLabelDic();
            var generator = new MbistFlowGenerator(prodFlowSheet, config, MbistBinTableType.Burst, false);

            // Act
            BistActionType result = generator.CheckBranchType("MissingBranch");

            // Assert
            Assert.AreEqual(BistActionType.Null, result);
        }

        [TestMethod]
        public void CheckBranchType_RunPatternBranch_ReturnsRunPattern()
        {
            // Arrange
            var config = new MbistConfig();
            var prodFlowSheet = new BistProdFlowSheet
            {
                SheetName = "Mbist_Test",
                Rows =
                [
                    new() { Label = "L1", Note = "", Pattern = "PAT1", Action = "NORMAL", PassBranch = "", FailBranch = "" }
                ]
            };
            prodFlowSheet.CreateLabelDic();
            var generator = new MbistFlowGenerator(prodFlowSheet, config, MbistBinTableType.Burst, false);

            // Act
            BistActionType result = generator.CheckBranchType("L1");

            // Assert
            Assert.AreEqual(BistActionType.RunPattern, result);
        }

        [TestMethod]
        public void GetBranchByLabel_RfDevice_ReturnsLabelUnchanged()
        {
            // Arrange
            LocalSpecs.Options.Device = EnumDevice.RF;
            var config = new MbistConfig();
            var prodFlowSheet = new BistProdFlowSheet { SheetName = "Mbist_Test" };
            prodFlowSheet.CreateLabelDic();
            var generator = new MbistFlowGenerator(prodFlowSheet, config, MbistBinTableType.Burst, false);

            // Act
            string result = generator.GetBranchByLabel("SomeLabel");

            // Assert
            Assert.AreEqual("SomeLabel", result);
        }

        [TestMethod]
        public void GetBranchByLabel_ApDeviceLabelNotFound_ReturnsControlPrefixedLabel()
        {
            // Arrange
            LocalSpecs.Options.Device = EnumDevice.AP;
            var config = new MbistConfig();
            var prodFlowSheet = new BistProdFlowSheet { SheetName = "Mbist_Test" };
            prodFlowSheet.CreateLabelDic();
            var generator = new MbistFlowGenerator(prodFlowSheet, config, MbistBinTableType.Burst, false);

            // Act
            string result = generator.GetBranchByLabel("SomeLabel");

            // Assert
            Assert.AreEqual("Control_SomeLabel", result);
        }
    }
}
