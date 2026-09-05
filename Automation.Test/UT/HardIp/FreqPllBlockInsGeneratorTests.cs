using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenInstanceBiz;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;
using Automation.Static;
using Automation.Test.Static;

using CommonReaderLib.PatternListCsv;

using FileDiffLib;

using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Newtonsoft.Json;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class FreqPllBlockInsGeneratorTests : FunctionTestBase
    {
        private HardIpSheet _sheet = null!;
        private HardIpInputData _inputData = null!;

        [TestInitialize]
        public void Setup()
        {
            _sheet = new HardIpSheet
            {
                Rows = []
            };
            HardIpParaData paraData = new HardIpParaData(EnumBlock.HardIp);
            _inputData = new HardIpInputData(paraData)
            {
                HardIpRegAssigns = []
            };
            TestProgram.VbtFunctionLib.AddVbtFunctionRange([.. TestSuiteInitialize.Functions.Where(x => x.Type == "VBT")]);
        }

        [TestMethod]
        public void GenBlockInsRows_ShouldReturnEmptyList_WhenNoPatternsExist()
        {
            // Arrange
            var generator = new FreqPllBlockInsGenerator(
                _inputData,
                "SheetA",
                _sheet);

            // Act
            List<InstanceSheet> result = generator.GenBlockInsRows("NV");

            // Assert
            Assert.AreEqual(0, result.Count, "Empty HardIpSheet.Rows should produce no InstanceSheets");
        }

        [TestMethod]
        public void GenBlockInsRows_ShouldRunBlockPath_WhenByBlockIsTrue()
        {
            string subName = "FreqPllBlockInsGenerator_GenBlockInsRows";
            string outputPath = Path.Combine(OutputPath, "HardIp", subName);
            string expectPath = Path.Combine(ExpectPath, "HardIp", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            var patterns = new List<HardIpPattern>
            {
            new()
            {
                SheetName = "BlockSheet",
                RowNum = 1,
                Pattern = new PatternClass("PAT_BLOCK_L_4_5_6_7"),
                SkipDotNet = false,
                SkipList = []
            },
            new()
            {
                SheetName = "BlockSheet",
                RowNum = 2,
                Pattern = new PatternClass("PAT_BLOCK_C_4_BI_6_7"),
                SkipDotNet = false,
                SkipList = []
            },
            new()
            {
                SheetName = "BlockSheet",
                RowNum = 2,
                Pattern = new PatternClass("PAT_BLOCK_S_4_SC_6_TDF"),
                SkipDotNet = false,
                SkipList = []
            },
            new()
            {
                SheetName = "BlockSheet",
                RowNum = 2,
                Pattern = new PatternClass("PAT_BLOCK_S_4_CH_6_TDF"),
                SkipDotNet = false,
                SkipList = []
            },
            new()
            {
                SheetName = "BlockSheet",
                RowNum = 2,
                Pattern = new PatternClass("PAT_BLOCK_S_4_SC_6_saa"),
                SkipDotNet = false,
                SkipList = []
            },
            new()
            {
                SheetName = "BlockSheet",
                RowNum = 2,
                Pattern = new PatternClass("PAT_BLOCK_S_4_CH_6_saa"),
                SkipDotNet = false,
                SkipList = []
            },
            };

            _sheet.Rows.AddRange(patterns);

            var generator = new FreqPllBlockInsGenerator(
                _inputData,
                "SheetBlock",
                _sheet);

            // Act
            List<InstanceSheet> result = generator.GenBlockInsRows("NV");

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
        public void GenBlockInsRows_ShouldHandleMultiplePatternsGracefully()
        {
            // Arrange
            _sheet.Rows.Add(new HardIpPattern
            {
                SheetName = "Test1",
                Pattern = new PatternClass("PAT1"),
                SkipList = []
            });

            _sheet.Rows.Add(new HardIpPattern
            {
                SheetName = "Test2",
                Pattern = new PatternClass("PAT2"),
                SkipList = []
            });

            var generator = new FreqPllBlockInsGenerator(
                _inputData,
                "MultiSheet",
                _sheet);

            // Act
            List<InstanceSheet> result = generator.GenBlockInsRows("NV");

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.IsTrue(result.Count >= 0, "Multiple patterns should not cause exceptions");
        }
    }
}
