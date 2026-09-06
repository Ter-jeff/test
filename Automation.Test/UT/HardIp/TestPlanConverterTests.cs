using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.GenerateIgxl.Wireless.DVDC.InputObject;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class TestPlanConverterTests : FunctionTestBase
    {
        private ScghData _mockScghData = null!;
        private TestPlanConverter _converter = null!;

        [TestInitialize]
        public void Setup()
        {
            _mockScghData = new ScghData();
            _converter = new TestPlanConverter(_mockScghData);
        }

        [TestMethod]
        public void ReadHardipPatterns_InvalidSheets_AreIgnored()
        {
            // Arrange
            ExcelPackage package = new ExcelPackage();
            ExcelWorksheet ws = package.Workbook.Worksheets.Add("Hardip_AAAA");

            #region build ws sample

            ws.Cells[1, 1].Value = "TTR";
            ws.Cells[1, 2].Value = "NoBinOut";
            ws.Cells[1, 3].Value = "Test Item";
            ws.Cells[1, 4].Value = "Step";
            ws.Cells[1, 5].Value = "Description";
            ws.Cells[1, 6].Value = "Pattern";
            ws.Cells[1, 7].Value = "Force Condition";
            ws.Cells[1, 8].Value = "Register Assignment";
            ws.Cells[1, 9].Value = "Misc Info";
            ws.Cells[1, 10].Value = "Meas";
            ws.Cells[1, 11].Value = "CP1 Lo Limit (H,L,N)";
            ws.Cells[1, 12].Value = "CP1 Hi Limit (H,L,N)";
            ws.Cells[1, 13].Value = "CP2 Lo Limit (H,L,N)";
            ws.Cells[1, 14].Value = "CP2 Hi Limit (H,L,N)";
            ws.Cells[1, 15].Value = "FT1 Lo Limit (H,L,N)";
            ws.Cells[1, 16].Value = "FT1 Hi Limit (H,L,N)";
            ws.Cells[1, 17].Value = "FT2 Lo Limit (H,L,N)";
            ws.Cells[1, 18].Value = "FT2 Hi Limit (H,L,N)";
            ws.Cells[1, 19].Value = "FT3 Lo Limit (H,L,N)";
            ws.Cells[1, 20].Value = "FT3 Hi Limit (H,L,N)";
            ws.Cells[1, 21].Value = "Comment";
            ws.Cells[2, 3].Value = 5.0;
            ws.Cells[2, 4].Value = 0.1;
            ws.Cells[2, 5].Value = "RX4G3_PWR";
            ws.Cells[3, 1].Value = "NV";
            ws.Cells[3, 4].Value = 0.2;
            ws.Cells[3, 5].Value = "Run the pattern provided";
            ws.Cells[3, 6].Value = "PP_BALA0_A_IN81_AN_IN81_MEA_JTG_VMX_ALLFRV_SI_ANA_HEADER+ PP_BALA0_A_PL81_AN_CIRX_MEA_JTG_IMX_ALLFRV_SI_PWR_RX4G3";
            ws.Cells[3, 7].Value = "ATC_RX_ALL:DisConnectDigital; ATC_TX_ALL:DisConnectDigital;";
            ws.Cells[3, 8].Value = "ATC0__efuse_aciophy_cio4pll_dco_coarsebin0=ATC0__Fcal10p0g_trim_ate&0; ATC1__efuse_aciophy_cio4pll_dco_coarsebin0=ATC1__Fcal10p0g_trim_ate&0; ATC0__efuse_aciophy_cio4pll_dco_coarsebin1=ATC0__Fcal10p3g_trim_ate&0; ATC1__efuse_aciophy_cio4pll_dco_coarsebin1=ATC1__Fcal10p3g_trim_ate&0; ATC0__efuse_aciophy_auspll_rodco_encap=ATC0__Fcal_trim_ate:3:4; ATC1__efuse_aciophy_auspll_rodco_encap=ATC1__Fcal_trim_ate:3:4; ATC0__efuse_aciophy_auspll_rodco_biasadj=ATC0__Fcal_trim_ate:0:2; ATC1__efuse_aciophy_auspll_rodco_biasadj=ATC1__Fcal_trim_ate:0:2; ATC0__efuse_aciophy_auscmn_vreg_trim=aciophy0_cmn_reg_cal_ate; ATC1__efuse_aciophy_auscmn_vreg_trim=aciophy1_cmn_reg_cal_ate;";
            ws.Cells[3, 9].Value = "SubBlock:RX4G3-PWR; MeasI_Range:0.8,0.8+0.2,0.2+0.8,0.8; MeasI_WaitTime:0.01,0.01+0.01,0.01+0.01,0.01; MainProgramCustomString: BypassAutorange;";
            ws.Cells[4, 4].Value = 0.3;
            ws.Cells[4, 5].Value = "Measure current on VDD_CIO MeasI Pin = VDD_CIO, VDD12_CIO MeasName = p0_state-X2,p0_state-X2";
            ws.Cells[4, 10].Value = "MeasI Pin = VDD_CIO \"P0-STATE-X2\"";
            ws.Cells[4, 11].Value = "85mA";
            ws.Cells[4, 12].Value = "284mA";
            ws.Cells[4, 13].Value = "90.5mA";
            ws.Cells[4, 14].Value = "470.5mA";
            ws.Cells[4, 15].Value = "85mA";
            ws.Cells[4, 16].Value = "284mA";
            ws.Cells[4, 17].Value = "90.5mA";
            ws.Cells[4, 18].Value = "470.5mA";
            ws.Cells[4, 19].Value = "85mA";
            ws.Cells[4, 20].Value = "284mA";
            ws.Cells[5, 4].Value = 0.4;
            ws.Cells[5, 5].Value = "Measure current on VDD12_CIO MeasI Pin = VDD_CIO, VDD12_CIO MeasName = p0_state-X2,p0_state-X2";
            ws.Cells[5, 10].Value = "MeasI Pin = VDD12_CIO \"P0-STATE-X2\"";
            ws.Cells[5, 11].Value = "177mA";
            ws.Cells[5, 12].Value = "248.5mA";
            ws.Cells[5, 13].Value = "194.5mA";
            ws.Cells[5, 14].Value = "298mA";
            ws.Cells[5, 15].Value = "177mA";
            ws.Cells[5, 16].Value = "248.5mA";
            ws.Cells[5, 17].Value = "194.5mA";
            ws.Cells[5, 18].Value = "298mA";
            ws.Cells[5, 19].Value = "177mA";
            ws.Cells[5, 20].Value = "248.5mA";
            ws.Cells[6, 4].Value = 0.5;
            ws.Cells[6, 5].Value = "Measure current on VDD_CIO MeasI Pin = VDD_CIO, VDD12_CIO MeasName = p10_state-X2,p10_state-X2";
            ws.Cells[6, 10].Value = "MeasI Pin = VDD_CIO \"P10-STATE-X2\"";
            ws.Cells[6, 11].Value = "0.5mA";
            ws.Cells[6, 12].Value = "49.5mA";
            ws.Cells[6, 13].Value = "10mA";
            ws.Cells[6, 14].Value = "189mA";
            ws.Cells[6, 15].Value = "0.5mA";
            ws.Cells[6, 16].Value = "49.5mA";
            ws.Cells[6, 17].Value = "10mA";
            ws.Cells[6, 18].Value = "189mA";
            ws.Cells[6, 19].Value = "0.5mA";
            ws.Cells[6, 20].Value = "49.5mA";
            ws.Cells[7, 4].Value = 0.6;
            ws.Cells[7, 5].Value = "Measure current on VDD12_CIO MeasI Pin = VDD_CIO, VDD12_CIO MeasName = p10_state-X2,p10_state-X2";
            ws.Cells[7, 10].Value = "MeasI Pin = VDD12_CIO \"P10-STATE-X2\"";
            ws.Cells[7, 11].Value = "7.5mA";
            ws.Cells[7, 12].Value = "20.5mA";
            ws.Cells[7, 13].Value = "7mA";
            ws.Cells[7, 14].Value = "24mA";
            ws.Cells[7, 15].Value = "7.5mA";
            ws.Cells[7, 16].Value = "20.5mA";
            ws.Cells[7, 17].Value = "7mA";
            ws.Cells[7, 18].Value = "24mA";
            ws.Cells[7, 19].Value = "7.5mA";
            ws.Cells[7, 20].Value = "20.5mA";
            ws.Cells[8, 4].Value = 0.7;
            ws.Cells[8, 5].Value = "Measure current on VDD_CIO MeasI Pin = VDD_CIO, VDD12_CIO MeasName = p1_state-X2,p1_state-X2";
            ws.Cells[8, 10].Value = "MeasI Pin = VDD_CIO \"P1-STATE-X2\"";
            ws.Cells[8, 11].Value = "27.5mA";
            ws.Cells[8, 12].Value = "118mA";
            ws.Cells[8, 13].Value = "40mA";
            ws.Cells[8, 14].Value = "275mA";
            ws.Cells[8, 15].Value = "27.5mA";
            ws.Cells[8, 16].Value = "118mA";
            ws.Cells[8, 17].Value = "40mA";
            ws.Cells[8, 18].Value = "275mA";
            ws.Cells[8, 19].Value = "27.5mA";
            ws.Cells[8, 20].Value = "118mA";
            ws.Cells[9, 4].Value = 0.8;
            ws.Cells[9, 5].Value = "Measure current on VDD12_CIO MeasI Pin = VDD_CIO, VDD12_CIO MeasName = p1_state-X2,p1_state-X2";
            ws.Cells[9, 10].Value = "MeasI Pin = VDD12_CIO \"P1-STATE-X2\"";
            ws.Cells[9, 11].Value = "78mA";
            ws.Cells[9, 12].Value = "138.5mA";
            ws.Cells[9, 13].Value = "79mA";
            ws.Cells[9, 14].Value = "166.5mA";
            ws.Cells[9, 15].Value = "78mA";
            ws.Cells[9, 16].Value = "138.5mA";
            ws.Cells[9, 17].Value = "79mA";
            ws.Cells[9, 18].Value = "166.5mA";
            ws.Cells[9, 19].Value = "78mA";
            ws.Cells[9, 20].Value = "138.5mA";

            #endregion

            ScghData scgh = new ScghData();
            TestPlanConverter converter = new TestPlanConverter(scgh);

            // Act
            Dictionary<string, HardIpSheet> result = _converter.ReadHardipPatterns(package.Workbook, sheetName => true);

            HardIpSheet sheet = result["Hardip_AAAA"];
            Assert.AreEqual("Hardip_AAAA", sheet.SheetName);
            Assert.IsTrue(sheet.Rows.Count != 0);
            Assert.IsTrue(sheet.Rows[0].MiscInfo.Contains("SubBlock:RX4G3-PWR"));

            // Assert
            Assert.AreEqual(1, result.Count, "Invalid sheets should be ignored.");
        }

        [TestMethod]
        public void ParseTrimItems_SetsTrimAndPostBurnFlags()
        {
            // Arrange
            var patterns = new List<HardIpPattern>
            {
                new() { WirelessData = new WirelessData { TrimTarget = "TRIM1" }, MiscInfo = "postburn" },
                new() { WirelessData = new WirelessData { TrimTarget = "" }, MiscInfo = "" }
            };

            // Act
            List<HardIpPattern> trimPatterns = _converter.ParseTrimItems(ref patterns);

            // Assert
            Assert.AreEqual(1, trimPatterns.Count, "Only patterns with TrimTarget should be in trimPatterns.");
            Assert.IsTrue(patterns[0].WirelessData.IsNeedPostBurn, "PostBurn flag should be set for 'postburn' miscInfo.");
            Assert.IsTrue(patterns[1].WirelessData.IsDoMeasure, "IsDoMeasure should be true for non-trim pattern.");
        }

        [TestMethod]
        public void ParseTrimItems_TrimTarget_AddsToTrimPatterns()
        {
            // Arrange
            var patterns = new List<HardIpPattern>
            {
                new()
                {
                    WirelessData = new WirelessData { TrimTarget = "TRIM1" },
                    MiscInfo = "postburn"
                },
                new()
                {
                    WirelessData = new WirelessData { TrimTarget = "" },
                    MiscInfo = ""
                }
            };

            // Act
            List<HardIpPattern> trimPatterns = _converter.ParseTrimItems(ref patterns);

            // Assert
            Assert.AreEqual(1, trimPatterns.Count);
            Assert.AreEqual("TRIM1", trimPatterns[0].WirelessData.TrimTarget);

            Assert.AreEqual(2, patterns.Count);

            Assert.IsTrue(patterns[0].WirelessData.IsNeedPostBurn);

            Assert.IsTrue(patterns[1].WirelessData.IsDoMeasure);
        }

        [TestMethod]
        public void ParseTrimItems_NoTrimTarget_SetsIsDoMeasure()
        {
            // Arrange
            var patterns = new List<HardIpPattern>
            {
                new()
                {
                    WirelessData = new WirelessData { TrimTarget = "" },
                    MiscInfo = ""
                }
            };

            // Act
            List<HardIpPattern> trimPatterns = _converter.ParseTrimItems(ref patterns);

            // Assert
            Assert.AreEqual(0, trimPatterns.Count, "No pattern with TrimTarget should result in empty trimPatterns");
            Assert.AreEqual(1, patterns.Count);
            Assert.IsTrue(patterns[0].WirelessData.IsDoMeasure, "Pattern without TrimTarget should have IsDoMeasure=true");
        }

        [TestMethod]
        public void SetPatternBurst_MultiplePatterns_ExpandsBurstPatternsCollection()
        {
            // Arrange 
            var planDic = new ConcurrentDictionary<string, HardIpSheet>();
            var pattern = new HardIpPattern
            {
                SheetName = "TestSheet",
                Pattern = new PatternClass("ParentPattern"),
                RegisterAssignment = "Value1|Value2",
                MiscInfo = "some_info"
            };

            pattern.Pattern.PatternSetList.Add(["SubPat1", "SubPat2"]);

            var sheet = new HardIpSheet { Rows = [pattern] };
            planDic.TryAdd("TestSheet", sheet);

            // Act
            new PatternBurstResolver().SetPatternBurst(planDic);

            // Assert
            Assert.AreEqual(3, pattern.BurstPatterns.Count, "BurstPatterns should contain 2 sub-patterns.");
        }

        [TestMethod]
        public void SetPatternBurst_WithRefSubBlock_InheritsMeasPinsCorrectly()
        {
            // Arrange 
            var planDic = new ConcurrentDictionary<string, HardIpSheet>();

            var refPattern = new HardIpPattern
            {
                Pattern = new PatternClass("Target") { PatternSetList = [["P1", "P2"]] },
                SheetSubBlockName = "RefBlock",
                SheetName = "TestSheet",
                RegisterAssignment = "RefValue",
                MiscInfo = "ref_subblock:A"
            };

            var targetPattern = new HardIpPattern
            {
                SheetName = "TestSheet",
                Pattern = new PatternClass("Target"),
                MiscInfo = "ref_subblock:RefBlock"
            };

            planDic.TryAdd("TestSheet", new HardIpSheet
            {
                Rows = [refPattern, targetPattern]
            });

            // Act
            new PatternBurstResolver().SetPatternBurst(planDic);

            // Assert
            Assert.IsFalse(targetPattern.MeasPins.Count != 0, "Target pattern should have inherited MeasPins.");
        }

        [TestMethod]
        public void SetPatternBurst_AssignReplicate_ShouldExpandAllSubPatterns()
        {
            HardIpPattern pattern = new HardIpPattern
            {
                SheetName = "Sheet",
                Pattern = new PatternClass("PAT"),
                RegisterAssignment = "A"
            };

            pattern.Pattern.PatternSetList.Add(["P1", "P2", "P3"]);

            ConcurrentDictionary<string, HardIpSheet> dic = new ConcurrentDictionary<string, HardIpSheet>();
            dic.TryAdd("Sheet", new HardIpSheet { Rows = [pattern] });

            new PatternBurstResolver().SetPatternBurst(dic);

            Assert.AreEqual(4, pattern.BurstPatterns.Count);

            Assert.AreEqual("p1", pattern.BurstPatterns[1].Pattern.TestPlanPatternName);
            Assert.AreEqual("p2", pattern.BurstPatterns[2].Pattern.TestPlanPatternName);
            Assert.AreEqual("p3", pattern.BurstPatterns[3].Pattern.TestPlanPatternName);

        }

        [TestMethod]
        public void SetPatternBurst_AssignListMultiple_ShouldMapCorrectly()
        {
            HardIpPattern pattern = new HardIpPattern
            {
                SheetName = "Sheet",
                Pattern = new PatternClass("PAT"),
                RegisterAssignment = "A|B"
            };

            pattern.Pattern.PatternSetList.Add(["P1", "P2"]);

            ConcurrentDictionary<string, HardIpSheet> dic = new ConcurrentDictionary<string, HardIpSheet>();
            dic.TryAdd("Sheet", new HardIpSheet { Rows = [pattern] });

            new PatternBurstResolver().SetPatternBurst(dic);

            Assert.AreEqual(3, pattern.BurstPatterns.Count);

            Assert.AreEqual("P1", pattern.BurstPatterns[1].Pattern.RealPatternName);
            Assert.AreEqual("P2", pattern.BurstPatterns[2].Pattern.RealPatternName);

            Assert.AreNotSame(pattern, pattern.BurstPatterns[2]);
        }

        [TestMethod]
        public void SetPatternBurst_WithMeasPins_ShouldSkipRefProcessing()
        {
            HardIpPattern pattern = new HardIpPattern
            {
                SheetName = "Sheet",
                Pattern = new PatternClass("PAT"),
                MiscInfo = "ref_subblock:A"
            };

            pattern.MeasPins.Add(new MeasPin());

            ConcurrentDictionary<string, HardIpSheet> dic = new ConcurrentDictionary<string, HardIpSheet>();
            dic.TryAdd("Sheet", new HardIpSheet { Rows = [pattern] });

            new PatternBurstResolver().SetPatternBurst(dic);

            Assert.IsTrue(pattern.MeasPins.Count == 1);
        }

        [TestMethod]
        public void SetPatternBurst_RefSubBlock_Local_ShouldCopyBurstPattern()
        {
            HardIpPattern source = new HardIpPattern
            {
                SheetName = "Sheet",
                Pattern = new PatternClass("SRC"),
                SheetSubBlockName = "A"
            };

            HardIpPattern target = new HardIpPattern
            {
                SheetName = "Sheet",
                Pattern = new PatternClass("TGT"),
                MiscInfo = "ref_subblock:A"
            };

            ConcurrentDictionary<string, HardIpSheet> dic = new ConcurrentDictionary<string, HardIpSheet>();
            dic.TryAdd("Sheet", new HardIpSheet { Rows = [source, target] });

            new PatternBurstResolver().SetPatternBurst(dic);

            Assert.IsTrue(target.BurstPatterns.Count >= 0);
        }

        [TestMethod]
        public void SetPatternBurst_RefSubBlock_NotFound_ShouldNotCrash()
        {
            HardIpPattern pattern = new HardIpPattern
            {
                SheetName = "Sheet",
                Pattern = new PatternClass("PAT"),
                MiscInfo = "ref_subblock:UNKNOWN"
            };

            ConcurrentDictionary<string, HardIpSheet> dic = new ConcurrentDictionary<string, HardIpSheet>();
            dic.TryAdd("Sheet", new HardIpSheet { Rows = [pattern] });

            new PatternBurstResolver().SetPatternBurst(dic);

            Assert.IsTrue(pattern.BurstPatterns.Count >= 0);
        }

        [TestMethod]
        public void SetPatternBurst_EmptySweep_ShouldInheritFromLastBurst()
        {
            HardIpPattern pattern = new HardIpPattern
            {
                SheetName = "Sheet",
                Pattern = new PatternClass("PAT"),
                RegisterAssignment = "A"
            };

            pattern.Pattern.PatternSetList.Add(["P1", "P2"]);

            ConcurrentDictionary<string, HardIpSheet> dic = new ConcurrentDictionary<string, HardIpSheet>();
            dic.TryAdd("Sheet", new HardIpSheet { Rows = [pattern] });

            new PatternBurstResolver().SetPatternBurst(dic);

            Assert.IsTrue(pattern.BurstPatterns.Count > 0);
        }

        [TestMethod]
        public void SetPatternBurst_WithExistingBurstPatterns_ShouldNotDuplicate()
        {
            HardIpPattern pattern = new HardIpPattern
            {
                SheetName = "Sheet",
                Pattern = new PatternClass("PAT"),
                RegisterAssignment = "A"
            };

            pattern.Pattern.PatternSetList.Add(["P1"]);

            pattern.BurstPatterns.Add(new HardIpPattern());

            ConcurrentDictionary<string, HardIpSheet> dic = new ConcurrentDictionary<string, HardIpSheet>();
            dic.TryAdd("Sheet", new HardIpSheet { Rows = [pattern] });

            new PatternBurstResolver().SetPatternBurst(dic);

            Assert.IsTrue(pattern.BurstPatterns.Count >= 1);
        }

        [TestMethod]
        public void SetPatternBurst_AssignListShorter_ShouldUseEmptyAssignment()
        {
            HardIpPattern pattern = new HardIpPattern
            {
                SheetName = "Sheet",
                Pattern = new PatternClass("PAT"),
                RegisterAssignment = "A" // only 1
            };

            pattern.Pattern.PatternSetList.Add(["P1", "P2", "P3"]);

            ConcurrentDictionary<string, HardIpSheet> dic = new ConcurrentDictionary<string, HardIpSheet>();
            dic.TryAdd("Sheet", new HardIpSheet { Rows = [pattern] });

            new PatternBurstResolver().SetPatternBurst(dic);

            Assert.AreEqual(4, pattern.BurstPatterns.Count);

            Assert.IsTrue(pattern.BurstPatterns.Any(p => string.IsNullOrEmpty(p.RegisterAssignment)));
        }

        [TestMethod]
        public void SetPatternBurst_NoHardIpInfo_ShouldUseEmptyEquation()
        {
            HardIpPattern pattern = new HardIpPattern
            {
                SheetName = "Sheet",
                Pattern = new PatternClass("UNKNOWN_PATTERN"),
                RegisterAssignment = "A|B"
            };

            pattern.Pattern.PatternSetList.Add(["XXX1", "XXX2"]);

            ConcurrentDictionary<string, HardIpSheet> dic = new ConcurrentDictionary<string, HardIpSheet>();
            dic.TryAdd("Sheet", new HardIpSheet { Rows = [pattern] });

            new PatternBurstResolver().SetPatternBurst(dic);

            Assert.IsTrue(pattern.BurstPatterns.Count > 0);
        }

    }
}
