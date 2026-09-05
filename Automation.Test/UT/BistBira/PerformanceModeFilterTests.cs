using Automation.GenerateIgxl.BistBira.BistInputLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using ScghLib.Base;
using ScghLib.Reader;

namespace Automation.Test.UT.BistBira
{
    [TestClass]
    public class PerformanceModeFilterTests
    {
        private PerformanceModeFilter _filter = null!;

        [TestInitialize]
        public void Setup()
        {
            _filter = new PerformanceModeFilter();
        }

        [TestMethod]
        [Description("Verify that WorkFlow executes without exception for mixed DSSC, INIT, and PAYLOAD patterns.")]
        public void WorkFlow_ShouldProcessProdFlow_WithVariousPatternTypes()
        {
            // Arrange
            var prodFlow = new BistProdFlowSheet
            {
                MbistSheet = new MbistSheet(),
                Rows =
                [
                    new()
                    {
                        Pattern = "MBIST_DSSC_01",
                        Voltage = "VDST"
                    },
                    new()
                    {
                        Pattern = "F00_TEST_CON_INIT_MODEX",
                        Voltage = "VMARGIN1"
                    },
                    new()
                    {
                        Pattern = "MBIST_PAYLOAD_PATTERN",
                        Voltage = "VMARGIN3"
                    },
                    new()
                    {
                        Pattern = "MBIST_PAYLOAD_NORMAL",
                        Voltage = "1.0V"
                    }
                ]
            };

            // Act
            _filter.GetPerformanceMode(prodFlow);

            // Assert
            Assert.IsTrue(prodFlow.Rows.Count > 0, "Should process rows");
        }

        [TestMethod]
        [Description("Verify IsTargetPattern identifies correct keyword position.")]
        public void IsTargetPattern_ShouldReturnTrue_WhenKeywordMatches()
        {
            // Act
            bool result = _filter.IsTargetPattern("A_B_C_D_E_F_EFU_X", 6, "EFU");

            // Assert
            Assert.IsTrue(result, "Pattern with EFU at position 6 should return true");
        }

        [DataTestMethod]
        [DataRow("PAT_EFU_001", "PAT_X_X_X_X_X_EFU_X_X_MS001", "", "", "", DisplayName = "01_EFU_Branch")]
        [DataRow("X_X_IN01_SRMDSSC-IN01_X", "PAT_VMARGIN1", "VMARGIN1", "VMARGIN1", "VMARGIN1", DisplayName = "02_VMARGIN1_Branch")]
        [DataRow("X_X_IN01_SRMDSSC-IN01_X", "PAT_VMARGIN3", "VMARGIN3", "VMARGIN3", "VMARGIN3", DisplayName = "03_VMARGIN3_Branch")]
        [DataRow("X_X_IN01_SRMDSSC-IN01_X", "PAT_VMARGIN", "VMARGIN", "VMARGIN", "VMARGIN", DisplayName = "04_VMARGIN_Branch")]
        [DataRow("X_X_IN01_SRMDSSC-IN01_X", "PAT_VDST", "VDST", "VDST", "VDST", DisplayName = "05_VDST_Branch")]
        [DataRow("X_X_IN01_SRMDSSC-IN01_X", "PAT_VDISTURB", "VDISTURB", "VDST", "VDST", DisplayName = "06_VDISTURB_Branch")]
        [DataRow("X_X_IN01_SRMDSSC-IN01_X", "PAT_MHV1", "MHV1", "MHV1", "MHV1", DisplayName = "07_MHV1_Branch")]
        [DataRow("X_X_IN01_SRMDSSC-IN01_X", "PAT_MHV2", "MHV2", "MHV2", "MHV2", DisplayName = "08_MHV2_Branch")]
        [DataRow("X_X_IN01_SRMDSSC-IN01_X", "PAT_MHV3", "MHV3", "MHV3", "MHV3", DisplayName = "09_MHV3_Branch")]
        [DataRow("X_X_IN01_SRMDSSC-IN01_X", "PAT_MHV", "MHV", "VMARGIN1", "VMARGIN1", DisplayName = "10_MHV_Branch")]
        [DataRow("X_X_IN01_SRMDSSC-IN01_X", "PAT_MLV1", "MLV1", "MLV1", "MLV1", DisplayName = "11_MLV1_Branch")]
        [DataRow("X_X_IN01_SRMDSSC-IN01_X", "PAT_MLV2", "MLV2", "MLV2", "MLV2", DisplayName = "12_MLV2_Branch")]
        [DataRow("X_X_IN01_SRMDSSC-IN01_X", "PAT_MLV3", "MLV3", "MLV3", "MLV3", DisplayName = "13_MLV3_Branch")]
        [DataRow("X_X_IN01_SRMDSSC-IN01_X", "PAT_MLV", "MLV", "VMARGIN3", "VMARGIN3", DisplayName = "14_MLV_Branch")]
        [DataRow("X_X_IN01_SRMDSSC-IN01_X", "PAT_VRS", "VRS", "VRS", "VRS", DisplayName = "15_VRS")]
        [DataRow("X_X_IN01_SRMDSSC-IN01_X", "PAT_Colon", "VH1:EXTRA", "VH1", "VH1", DisplayName = "16_Colon_Branch")]
        [DataRow("X_X_IN01_SRMDSSC-IN01_X", "PAT_Comma", "VH1,VH2", "VH1,VH2", "VH1,VH2", DisplayName = "17_Comma_Branch")]
        [DataRow("X_X_IN01_SRMDSSC-IN01_X", "PAT_NormalVoltage", "1.0V", "1.0V", "1.0V", DisplayName = "18_NormalVoltage_Branch")]
        [DataRow("X_X_IN01_SRMDSSC-IN01_X", "PAT_BlankVoltage", "", "", "", DisplayName = "19_BlankVoltage_Branch")]
        [DataRow("X_X_IN01_SRMDSSC-IN01_X", "PAT_VMAX", "VMAX", "", "", DisplayName = "20_VMAX")]
        [DataRow("X_X_IN01_SRMDSSC-IN01_X", "PAT_X_X_X_X_X_EXX_X_X_MS001", "VMAX,VMIN", "VMAX,VMIN", "VMAX,VMIN", DisplayName = "21_Other")]
        public void GetPerformanceMode_AllBranchesCovered(string pattern1, string pattern2, string voltage, string expectedPerformance, string expectedMode)
        {
            // Arrange
            var prodFlow = new BistProdFlowSheet
            {
                MbistSheet = new MbistSheet { SheetName = "SheetA" },
                Rows =
        [
            new() { Pattern = pattern1, Voltage = voltage },
            new() { Pattern = "X_X_X_IN01" },
            new() { Pattern = pattern2, Voltage = voltage }
        ]
            };

            // Act
            _filter.GetPerformanceMode(prodFlow);

            // Assert
            BistProdFlowRow? testRow1 = prodFlow.Rows.Find(r => r.Pattern == pattern1);
            BistProdFlowRow? testRow2 = prodFlow.Rows.Find(r => r.Pattern == pattern2);

            Assert.AreEqual(expectedMode.ToUpper(), (testRow2!.VoltageMode ?? "").ToUpper(), $"VoltageMode mismatch for {pattern2}");

            Assert.AreEqual(expectedPerformance.ToUpper(), (testRow1!.OriPerformance ?? "").ToUpper(), $"OriPerformance mismatch for {pattern1}");
        }

    }
}
