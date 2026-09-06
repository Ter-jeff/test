using System.Globalization;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.Static;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class TimeSetSheetGeneratorTests : FunctionTestBase
    {
        private readonly TimeSetSheetGenerator _generator = new();

        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            var timeSetBasicSheet = new TimeSetBasicSheet("TIMESET_BRNA0_S_AN_SI");
            var data = new TSet
            {
                Name = "clk_tset"
            };
            data.TimingRows.Add(new TimingRow
            {
                PinGrpName = "PIN_A",
                PinGrpClockPeriod = "5"
            });
            timeSetBasicSheet.AddRow(data);
            TestProgram.IgxlWorkBk.AddTimeSetSheet("", timeSetBasicSheet);
        }

        [TestMethod]
        public void ConvertUnits_ShouldSortAndConvertFrequenciesCorrectly()
        {
            // Arrange
            string input = ";AC:PINB:100MHz;AC:PINA:200MHz;";

            // Act
            string result = _generator.ConvertUnits(input);

            // Assert
            StringAssert.StartsWith(result, "PINA:");
            StringAssert.Contains(result, "PINB");
            Assert.IsTrue(result.Contains("100000000"));
            Assert.IsTrue(result.Contains("200000000"));
        }

        [TestMethod]
        public void GetPinClockPeriod_ShouldReturnPeriodFromSheet_WhenTimeSetExists()
        {
            // Arrange
            var sheet = new TimeSetBasicSheet("tset1");
            var data = new TSet
            {
                Name = "clk_tset"
            };
            data.TimingRows.Add(new TimingRow
            {
                PinGrpName = "PIN_A",
                PinGrpClockPeriod = "5"
            });
            sheet.Rows.Add(data);

            // Act
            string result = _generator.GetPinClockPeriod("PIN_A", "clk_tset", sheet);

            // Assert
            Assert.AreEqual("5", result);
        }

        [TestMethod]
        public void GetPinClockPeriod_ShouldComputeInverseOfFrequency()
        {
            // Arrange
            var sheet = new TimeSetBasicSheet("tset2");
            string newValue = "100MHz";

            // Act
            string result = _generator.GetPinClockPeriod("PIN_A", newValue, sheet);

            // Assert
            double parsed = double.Parse(result, CultureInfo.InvariantCulture);
            Assert.IsTrue(parsed > 0 && parsed < 1e-6, $"Unexpected period {result}");
        }

        [TestMethod]
        public void SetNewTimeSetValue_ShouldApplyClock2X_ForHighFrequency()
        {
            // Arrange
            var row = new TimingRow();

            // Act
            _generator.SetNewTimeSetValue(row, "1e-9", false, "1e-8", "SRC");

            // Assert
            Assert.AreEqual("clock_2X", row.PinGrpSetup);
            Assert.AreEqual("RL-2X", row.DataFmt);
            Assert.AreEqual("=2E-09", row.DriveOff);
        }

        [TestMethod]
        public void SetNewTimeSetValue_ShouldHandleNormalFrequency()
        {
            // Arrange
            var row = new TimingRow();

            // Act
            _generator.SetNewTimeSetValue(row, "0.00001", false, "0.00001", "SRC");

            // Assert
            Assert.AreEqual("clock", row.PinGrpSetup);
            Assert.IsTrue(row.DataFmt.StartsWith("RL"));
            Assert.AreEqual("SRC", row.DataSrc);
        }

        [TestMethod]
        public void Should_DoNothing_WhenMcgSettingIsEmpty()
        {
            var pattern = new HardIpPattern
            {
                TimeSetUsed = new TimeSetUsed { McgSetting = string.Empty }
            };
            var timeSet = new TimeSetBasicSheet("") { Name = "TS1" };

            _generator.UpdateTimeSetSheetWithMcg(timeSet, pattern);

            Assert.AreEqual(null, pattern.TimeSetUsed.TimeSetMcg);
        }

        [TestMethod]
        public void Should_CreateNewTimeSetMcg()
        {
            var pattern = new HardIpPattern
            {
                TimeSetUsed = new TimeSetUsed { McgSetting = "PIN1:PIN_A:5mHz;" },
                Pattern = new PatternClass("cz_brna0_c_fulp_an_aa00_dll_jtg_vix_allfrv_si_cpllds_t10pd")
            };
            var timeSet = new TimeSetBasicSheet("")
            {
                Name = "TS1",
                Rows =
                [
                    new()
                    {
                        CyclePeriod = "10",
                        TimingRows =
                        [
                            new() { PinGrpName = "PIN_A" },
                        ]
                    }
                ]
            };

            _generator.UpdateTimeSetSheetWithMcg(timeSet, pattern);

            Assert.AreEqual("TS1_MCG", pattern.TimeSetUsed.TimeSetMcg);
        }

        [TestMethod]
        public void Should_CreateNewTimeSetMcg_1()
        {
            var pattern = new HardIpPattern
            {
                TimeSetUsed = new TimeSetUsed { McgSetting = "PIN1:PIN_A&PIN_B:5mHz;" },
                Pattern = new PatternClass("cz_brna0_c_fulp_an_aa00_dll_jtg_vix_allfrv_si_cpllds_t10pd")
            };
            var timeSet = new TimeSetBasicSheet("")
            {
                Name = "TS1",
                Rows =
                [
                    new()
                    {
                        CyclePeriod = "10",
                        TimingRows =
                        [
                            new() { PinGrpName = "PIN_A" },
                        ]
                    }
                ]
            };

            _generator.UpdateTimeSetSheetWithMcg(timeSet, pattern);

            Assert.AreEqual("TS1_MCG", pattern.TimeSetUsed.TimeSetMcg);
        }

        [TestMethod]
        public void Should_NotDuplicateTimeSetMcg()
        {
            var pattern = new HardIpPattern
            {
                TimeSetUsed = new TimeSetUsed { McgSetting = "PIN1:PIN_A:5ns;" }
            };
            var timeSet = new TimeSetBasicSheet("")
            {
                Name = "TS_DUP",
                Rows =
                [
                    new()
                    {
                        CyclePeriod = "10",
                        TimingRows =
                        [
                            new() { PinGrpName = "PIN_A" }
                        ]
                    }
                ]
            };

            _generator.UpdateTimeSetSheetWithMcg(timeSet, pattern);

            Assert.AreEqual(null, pattern.TimeSetUsed.TimeSetMcg);
        }
    }
}
