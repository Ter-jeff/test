using System.Collections.Generic;

using Automation.GenerateIgxl.Basic.Business.GenDc;
using Automation.Singleton;
using Automation.Static;

using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.DataStruct;

namespace Automation.Test.UT.Basic
{
    [TestClass]
    public class DcSpecGeneratorTests : FunctionTestBase
    {
        private IoInfoSheet _realIoInfo = null!;
        private IoInfoSheet _realIoInfoConcurrent = null!;
        private DcSpecGenerator _generator = null!;
        private MultiTestSettingSheetsSingleton _settings = null!;

        [TestInitialize]
        public void Setup()
        {
            _settings = MultiTestSettingSheetsSingleton.Instance();

            // Use the real class with its data-driven constructor
            _realIoInfo = new IoInfoSheet("Main", []);
            _realIoInfoConcurrent = new IoInfoSheet("Concurrent", []);

            _generator = new DcSpecGenerator(
                _settings,
                _realIoInfo,
                _realIoInfoConcurrent
            );
        }

        [TestMethod]
        public void GetIoDcSpecsByBlock_SeparatesScanAndCommonCategories()
        {
            // Arrange
            _realIoInfo.BlockIoInfo["Common"] = [new() { PinGrpName = "PinA", Type = "Common" }];
            _realIoInfo.BlockIoInfo["HardIP"] = [new() { PinGrpName = "PinB", Type = "HardIP" }];
            _realIoInfo.BlockIoInfo["DCTEST_Continuity"] = [new() { PinGrpName = "PinC", Type = "DCTEST_Continuity" }];
            _realIoInfo.BlockIoInfo["Scan"] = [new() { PinGrpName = "PinD", Type = "Scan" }];

            // Act
            Dictionary<string, List<DcSpec>> result = _generator.GetIoDcSpecsByBlock();

            // Assert
            Assert.IsTrue(result.ContainsKey("FT1"));
        }

        [TestMethod]
        public void GetIoDcSpecs_AggregatesAllBlocksCorrectly()
        {
            // Arrange
            _realIoInfo.BlockIoInfo["Common"] = [new() { PinGrpName = "P1", Type = "Common" }];
            _realIoInfo.BlockIoInfo["HardIP"] = [new() { PinGrpName = "Pin2", Type = "HardIP" }];
            _realIoInfo.BlockIoInfo["DCTEST_Continuity"] = [new() { PinGrpName = "Pin3", Type = "DCTEST_Continuity" }];
            _realIoInfo.BlockIoInfo["Scan"] = [new() { PinGrpName = "P1" }];

            // Act
            Dictionary<string, List<DcSpec>> result = _generator.GetIoDcSpecs();

            // Assert
            Assert.AreEqual(3, result.Count);
        }

        [TestMethod]
        public void GetIoDcSpecsByBlock_HandlesNullConcurrentSheet()
        {
            // Arrange
            _generator = new DcSpecGenerator(_settings, _realIoInfo, null);

            _realIoInfo.BlockIoInfo["Common"] = [];

            // Act
            Dictionary<string, List<DcSpec>> result = _generator.GetIoDcSpecsByBlock();

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.AreEqual(6, result.Count);
        }

        [TestMethod]
        [DataRow("0.75V", 6)]
        [DataRow("0.75%", 6)]
        public void GetPowerDcSpecsByBlock(string value, int expected)
        {
            // Arrange
            _settings.TestSettingSheetsList[0].DataRows[0].DcCategoryValues[0].Hv.Value = "0.75V";
            _settings.TestSettingSheetsList[0].DataRows[0].DcCategoryValues[0].Lv.Value = "0.75V";
            _settings.TestSettingSheetsList[0].DataRows[0].DcCategoryValues[0].Nv.Value = "0.75V";
            _realIoInfo.BlockIoInfo["Common"] = [new() { PinGrpName = "P1" }];
            _realIoInfo.BlockIoInfo["Scan"] = [new() { PinGrpName = "P1" }];

            // Act
            Dictionary<string, List<DcSpec>> result = _generator.GetPowerDcSpecsByBlock();

            // Assert
            Assert.AreEqual(expected, result.Count);
        }

        [TestCleanup]
        public void ResetPwrSupplyRes()
        {
            LocalSpecs.PwrSupplyRes = null;
        }

        [TestMethod]
        public void PercentageConvertToString_ValidPercentage_ComputesRatioAgainstNv()
        {
            // Arrange - percent = (5 * 0.01) + 1 = 1.05, prefixed by the nv argument
            string result = _generator.PercentageConvertToString("5%", "2");

            // Assert
            Assert.IsTrue(result.StartsWith('2'));
            Assert.IsTrue(result.Contains("1.05"));
        }

        [TestMethod]
        public void PercentageConvertToString_InvalidFormula_ReturnsError()
        {
            // Act
            string result = _generator.PercentageConvertToString("NotANumber%", "1");

            // Assert
            Assert.AreEqual("error", result);
        }

        [TestMethod]
        public void MakeCeilingFormula_ResolutionAboveThreshold_WrapsWithCeilingFunction()
        {
            // Arrange
            LocalSpecs.PwrSupplyRes = "0.01";

            // Act
            string result = _generator.MakeCeilingFormula("SomeCell");

            // Assert
            Assert.AreEqual("CEILING(SomeCell,0.01)", result);
        }

        [TestMethod]
        public void MakeCeilingFormula_ResolutionBelowThreshold_ReturnsCellUnchanged()
        {
            // Arrange
            LocalSpecs.PwrSupplyRes = "0";

            // Act
            string result = _generator.MakeCeilingFormula("SomeCell");

            // Assert
            Assert.AreEqual("SomeCell", result);
        }

        [TestMethod]
        public void MakeFloorFormula_ResolutionAboveThreshold_WrapsWithFloorFunction()
        {
            // Arrange
            LocalSpecs.PwrSupplyRes = "0.01";

            // Act
            string result = _generator.MakeFloorFormula("SomeCell");

            // Assert
            Assert.AreEqual("FLOOR(SomeCell,0.01)", result);
        }

        [TestMethod]
        public void MakeFloorFormula_ResolutionBelowThreshold_ReturnsCellUnchanged()
        {
            // Arrange
            LocalSpecs.PwrSupplyRes = "0";

            // Act
            string result = _generator.MakeFloorFormula("SomeCell");

            // Assert
            Assert.AreEqual("SomeCell", result);
        }

        [TestMethod]
        public void GetSelectorList_ReturnsMinTypMaxInOrder()
        {
            // Act
            List<Selector> result = _generator.GetSelectorList();

            // Assert
            Assert.AreEqual(3, result.Count);
            Assert.AreEqual("Min", result[0].SelectorName);
            Assert.AreEqual("Typ", result[1].SelectorName);
            Assert.AreEqual("Max", result[2].SelectorName);
        }

        [TestMethod]
        public void ConvertSpecialCategoryValue_ValueContainsLetterWithNullVrs_ReplacesWithZero()
        {
            // Arrange
            var dcValue = new DcCategoryValue("Cat1");
            dcValue.Hv.Value = "SomeLetterValue";
            dcValue.Lv.Value = "0.5";
            dcValue.Nv.Value = "0.5";

            // Act
            DcCategoryValue result = _generator.ConvertSpecialCategoryValue(dcValue, null!);

            // Assert
            Assert.AreEqual("0", result.Hv.Value);
            Assert.AreEqual("0.5", result.Lv.Value);
        }

        [TestMethod]
        public void ConvertSpecialCategoryValue_ValueContainsLetterWithVrsProvided_UsesVrsValue()
        {
            // Arrange
            var dcValue = new DcCategoryValue("Cat1");
            dcValue.Hv.Value = "SomeLetterValue";
            var vrsDcValue = new DcCategoryValue("Cat1");
            vrsDcValue.Hv.Value = "1.5";

            // Act
            DcCategoryValue result = _generator.ConvertSpecialCategoryValue(dcValue, vrsDcValue);

            // Assert
            Assert.AreEqual("1.5", result.Hv.Value);
        }

        [TestMethod]
        public void ConvertSpecialCategoryValue_NoLetterValues_ReturnsSameInstance()
        {
            // Arrange
            var dcValue = new DcCategoryValue("Cat1");
            dcValue.Hv.Value = "0.5";
            dcValue.Lv.Value = "0.5";
            dcValue.Nv.Value = "0.5";

            // Act
            DcCategoryValue result = _generator.ConvertSpecialCategoryValue(dcValue, null!);

            // Assert
            Assert.AreSame(dcValue, result);
        }
    }
}
