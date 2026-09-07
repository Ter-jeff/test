using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz.GenFlow;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;

using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.BinNumberLegacy;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class HardIpFlowSheetGeneratorTests
    {
        private HardIpFlowSheetGenerator _generator = null!;

        [TestInitialize]
        public void Setup()
        {
            _generator = new HardIpFlowSheetGenerator(null, "TestSheet", null);
        }

        [DataTestMethod]
        [DataRow("NV", "NV")]
        [DataRow("nv", "NV")]
        [DataRow("LV", "LV")]
        [DataRow("lv", "LV")]
        [DataRow("HV", "HV")]
        [DataRow("hv", "HV")]
        [DataRow("unknown", "")]
        public void GetFlowSequence_ShouldReturnCorrectConstant(string input, string expected)
        {
            string result = _generator.GetFlowSequence(input);
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void RowByVoltage_ShouldGenerateBinRowsForAllVoltages()
        {
            // Arrange
            var flowRow = new FlowRow { Parameter = "BinBase" };
            var binData = new SoftBinRangeData
            {
                HardHlvBin = "X",
                HardHvBin = "Y",
                HardLvBin = "Z",
                HardNvBin = "N"
            };

            // Act
            List<FlowRow> result = _generator.SplitBinRowByVoltage(flowRow, binData);

            // Assert
            Assert.AreEqual(4, result.Count);
            Assert.IsTrue(result.Exists(r => r.Parameter.EndsWith("_HLV")));
            Assert.IsTrue(result.Exists(r => r.Parameter.EndsWith("_HV")));
            Assert.IsTrue(result.Exists(r => r.Parameter.EndsWith("_LV")));
            Assert.IsTrue(result.Exists(r => r.Parameter.EndsWith("_NV")));
        }

        [TestMethod]
        public void RowByVoltage_ShouldIncludeOriginalRow_WhenSomeBinsMissing()
        {
            var flowRow = new FlowRow { Parameter = "BinBase" };
            var binData = new SoftBinRangeData
            {
                HardHvBin = "HVOnly"
            };

            List<FlowRow> result = _generator.SplitBinRowByVoltage(flowRow, binData);

            Assert.AreEqual(2, result.Count);
            Assert.IsTrue(result.Exists(r => r.Parameter == "BinBase_HV"));
            Assert.IsTrue(result.Exists(r => r.Parameter == "BinBase"));
        }

        [TestMethod]
        public void IsNeedGenerate_ShouldReturnFalse_WhenBlockTypeAndEmptyMeasPinsAndPayloads()
        {
            var pattern = new HardIpPattern
            {
                BlockType = "SomeBlock",
                MeasPins = [],
                Pattern = new PatternClass("") { InstancePayloadName = [] }
            };

            bool result = _generator.IsNeedGenerate(pattern);
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void IsNeedGenerate_ShouldReturnTrue_WhenNoBlockType()
        {
            var pattern = new HardIpPattern
            {
                BlockType = "",
                MeasPins = [],
                Pattern = new PatternClass("") { InstancePayloadName = [] }
            };

            bool result = _generator.IsNeedGenerate(pattern);
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void IsNeedGenerate_ShouldReturnTrue_WhenHasMeasPins()
        {
            var pattern = new HardIpPattern
            {
                BlockType = "Block",
                MeasPins = [new()],
                Pattern = new PatternClass("") { InstancePayloadName = [] }
            };

            bool result = _generator.IsNeedGenerate(pattern);
            Assert.IsTrue(result);
        }
    }
}
