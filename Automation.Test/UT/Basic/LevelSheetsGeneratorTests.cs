using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.GenerateIgxl.Basic.Business.GenLevel.BassData;
using Automation.GenerateIgxl.Basic.Business.GenLevel.Business;
using Automation.Singleton;
using Automation.Static;
using Automation.Static.Result;

using FileDiffLib;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.DataStruct;

namespace Automation.Test.UT.Basic
{
    [TestClass]
    public class LevelSheetsGeneratorTests : FunctionTestBase
    {
        private LevelSheetsGenerator _generator = null!;
        private IoInfoSheet _ioInfoSheet = null!;
        private IoInfoSheet _concurrentIoInfoSheet = null!;

        [TestInitialize]
        public void Setup()
        {
            PowerInfoSheet powerInfo = new PowerInfoSheet();
            (_ioInfoSheet, _concurrentIoInfoSheet) = GenIoInfoSheets();
            PinMapSheet pinMapSheet = new PinMapSheet("");
            pinMapSheet.AddPins([.. new List<string> { "VDD_PCPU", "VDD_CPU_SRAM", "VDD_ECPU", "VDD_GPU", "VDD_SRAM_GPU", "VDD_SRAM_SOC", "VDD_SOC", "VDD_DCS_DDR", "VDD_AVE", "VDD_DISP", "VDD_LOW", "VDD_FIXED", "VDD_SRAM_ULPPLL_FLPPLL_SLC", "VDD_SRAM_VID_PLL", "VDD_SRAM_VIDSEC_PLL", "VDD_SRAM_USB_DEBUG" }.Select(x => new Pin(x, "Power"))]);
            pinMapSheet.AddGroups([.. new List<string> { "VDDIO12_AOP", "VDDIO12_AOP_2", "VDD_FIXED_AMUX", "VDD_FIXED_CIO", "VDD_FIXED_CPU", "VDD_FIXED_ECPU_MTR", "VDD_FIXED_LPDP_RX", "VDD_FIXED_LPDP_TX_INT", "VDD_FIXED_LPDP_TX_SEC", "VDD_FIXED_MTR_CPM_PCPU", "VDD_FIXED_PCIE", "VDD_FIXED_PCIE_REFBUF", "VDD_FIXED_PLL", "VDD_FIXED_PLL_SOC", "VDD_FIXED_USB", "VDD_FIXED_XTAL", "VDD_AMPH_DDR", "VDD10_ADC_SOC", "VDD10_AMUX_FMON", "VDDA_EFUSE0", "VDDA_EFUSE1", "VDDA_EFUSE2", "VDDA_EFUSE3", "VDDA_EFUSE4", "VDD10_LPDP_RX", "VDD10_LPDP_TX_INT", "VDD10_LPDP_TX_SEC", "VDD10_PCIE", "VDD10_PCIE_REFBUF", "VDD10_PLL", "VDD10_ULPPLL_FLPPLL", "VDD12_USB", "VDD12_USB_DEBUG", "VDD10_XTAL", "VDDIO12_GRP", "VDD10_PLL_DDR", "VDD10_VID_PLL", "VDD10_VIDSEC_PLL", "VDD10_CIO", "VDD10_CPLL_SOC", "VDD10_HSCDFT0", "VDD10_HSCDFT1", "VDD10_SLC_PLL", "VDDQL_DDR", "VDD1", "VDD2" }.Select(x => new PinGroup(x, "Power") { PinList = [new("A", "Power")] })]);
            pinMapSheet.AddGroup(new PinGroup("WalkingZ", "IO") { PinList = [new("nWire", "IO")] });
            pinMapSheet.AddGroup(new PinGroup("nWire", "IO") { PinList = [new("XI0_Diff_PA", "IO"), new("XI0_Diff", "IO")] });
            _generator = new LevelSheetsGenerator(MultiTestSettingSheetsSingleton.Instance(), powerInfo, pinMapSheet);

            TestProgram.Clear();
        }

        [TestMethod]
        public void GenerateFlow_ShouldCreateLevelSheets()
        {
            string subName = "LevelGenerator";
            string outputPath = Path.Combine(OutputPath, "Basic", subName);
            string expectPath = Path.Combine(ExpectPath, "Basic", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            List<string> block =
            [
                "IDS",
                "Scan",
                "Rtos",
                "Mbist",
                "Efuse",
                "HardIP",
                "BinCut",
                "EVS"
            ];
            LocalSpecs.SetEnableModules(block);
            List<string> module =
            [
                "Basic",
                "Scan",
                "Rtos",
                "Mbist",
                "Efuse",
                "HardIP",
                "BinCut",
                "EVS"
            ];
            BlockStatus.Create();
            module.ForEach(x => BlockStatus.GetAutomationBlockStatus(x).Down = true);

            LevelInitial levelInitial = new LevelInitial(_ioInfoSheet, _concurrentIoInfoSheet, false);
            Dictionary<string, LevelData> levels = levelInitial.InitialLevelDatas();

            // Act
            MultiLevelSheets sheets = _generator.GenerateFlow(levels);

            foreach (LevelSheet sheet in sheets.Values)
            {
                sheet.Write(Path.Combine(outputPath, sheet.Name + ".txt"));
            }

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void CreateLevelSheet_ShouldReturnCorrectLevelSheet()
        {
            LevelSheet sheet = _generator.CreateLevelSheet("Levels_Test");
            Assert.AreNotEqual(null, sheet);
            Assert.AreEqual("Levels_Test", sheet.Name);
        }

        [TestMethod]
        public void CreateSpiRomLevelSheet()
        {
            string subName = "CreateSpiRomLevelSheet";
            string outputPath = Path.Combine(OutputPath, "Basic", subName);
            string expectPath = Path.Combine(ExpectPath, "Basic", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            LevelSheet sheet = _generator.CreateSpiRomLevelSheet();

            sheet.Write(Path.Combine(outputPath, sheet.Name + ".txt"));

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void LevelGeneratorByBlockTest()
        {
            string subName = "LevelGeneratorByBlock";
            string outputPath = Path.Combine(OutputPath, "Basic", subName);
            string expectPath = Path.Combine(ExpectPath, "Basic", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            BasicResult.Level = true;
            GenIoInfoSheets();
            List<string> block =
            [
                "IDS",
                "Scan",
                "Rtos",
                "Mbist",
                "Efuse",
                "HardIP",
                "BinCut",
                "EVS"
            ];
            LocalSpecs.SetEnableModules(block);
            List<string> module =
            [
                "Basic",
                "Scan",
                "Rtos",
                "Mbist",
                "Efuse",
                "HardIP",
                "BinCut",
                "EVS"
            ];
            BlockStatus.Create();
            module.ForEach(x => BlockStatus.GetAutomationBlockStatus(x).Down = true);

            LevelInitial levelInitial = new LevelInitial(_ioInfoSheet, _concurrentIoInfoSheet, true);
            Dictionary<string, LevelData> levels = levelInitial.InitialLevelDatas();

            // Act
            MultiLevelSheets multiLevelSheets = _generator.GenerateFlow(levels, TestPlanStatic.HardIpDcSheet);
            foreach (LevelSheet multiLevelSheet in multiLevelSheets.Values)
            {
                multiLevelSheet.Write(Path.Combine(outputPath, multiLevelSheet.Name + ".txt"));
            }

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
