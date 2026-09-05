using System.Collections.Generic;
using System.Linq;

using Automation.Reader;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Extension;

using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

using TestPlanLib.Static;

namespace Automation.Test.UT.Common
{
    [TestClass]
    public class NwireSingletonTests : FunctionTestBase
    {
        private NwireSingleton _sut = null!;

        [TestMethod]
        public void GetOutClkVoltage_Set_WhenCurrentIsZero()
        {
            // Act
            double result = NwireSingleton.GetOutClkVoltage("1.5", 0);

            // Assert
            Assert.AreEqual(1.5, result);
        }

        [TestMethod]
        public void GetOutClkVoltage_WhenCurrentIsNotZero_ShouldKeepCurrentValue()
        {
            // Act
            double result = NwireSingleton.GetOutClkVoltage("1.5", 2.0);

            // Assert
            Assert.AreEqual(2.0, result);
        }

        [TestMethod]
        public void GetOutClkVoltage_WhenTargetValueInvalid_ShouldReturnCurrentValue()
        {
            // Act
            double result = NwireSingleton.GetOutClkVoltage("ABC", 3.3);

            // Assert
            Assert.AreEqual(3.3, result);
        }

        [TestMethod]
        public void ConvertPogo2Single_WithValidMapping_ShouldReplaceChannel()
        {
            // Arrange
            var sut = new NwireSingleton();

            var rows = new List<ChannelMapRow>
            {
                new()
                {
                    Sites = ["SITE1.CH1"]
                }
            };

            var pogoMapping = new Dictionary<string, string>
            {
                { "ch1", "A1" }
            };

            // Act
            List<ChannelMapRow> result = sut.ConvertPogo2Single(rows, pogoMapping);

            // Assert
            Assert.AreEqual("SITE1.A1", result[0].Sites[0]);
        }

        private static (NwireSingleton sut, ExcelWorksheet ws) CreateSut(string[] headers, params (string pin, string up, string down)[] powerRows)
        {
            var package = new ExcelPackage();
            ExcelWorkbook wb = package.Workbook;

            // NWire
            ExcelWorksheet ws = wb.Worksheets.Add("NWire");
            for (int i = 0; i < headers.Length; i++)
            {
                ws.Cells[1, i + 2].Value = headers[i].Trim();
            }

            // PowerInfo
            if (powerRows.Length > 0)
            {
                ExcelWorksheet pws = wb.Worksheets.Add(NeededSheets.PowerInfo);
                pws.Cells[1, 1].Value = "PinName";
                pws.Cells[1, 2].Value = "PowerSequence";
                pws.Cells[1, 3].Value = "PowerDownSequence";

                for (int i = 0; i < powerRows.Length; i++)
                {
                    pws.Cells[i + 2, 1].Value = powerRows[i].pin?.Trim();
                    pws.Cells[i + 2, 2].Value = powerRows[i].up;
                    pws.Cells[i + 2, 3].Value = powerRows[i].down;
                }
            }

            EpWorkbook.TestPlanWorkbook = wb;

            var sut = new NwireSingleton
            {
                SettingInfo = new NwireSetting()
            };

            return (sut, ws);
        }

        // 1. Workbook = null
        [TestMethod]
        public void AddFrcPin_WorkbookNull()
        {
            var package = new ExcelPackage();
            ExcelWorksheet ws = package.Workbook.Worksheets.Add("NWire");
            ws.Cells[1, 2].Value = "CLK1";

            var sut = new NwireSingleton();

            sut.AddFrcPinFromPowerInfo(ws);
        }

        // 2. PowerInfo sheet
        [TestMethod]
        public void AddFrcPin_NoPowerSheet()
        {
            (NwireSingleton sut, ExcelWorksheet ws) = CreateSut(["CLK1"]);

            sut.AddFrcPinFromPowerInfo(ws);
        }

        // 3. pin no header
        [TestMethod]
        public void AddFrcPin_NoHeaderMatch()
        {
            (NwireSingleton sut, ExcelWorksheet ws) = CreateSut(["CLK1"], ("CLK2", "SEQ", "DOWN"));

            sut.SettingInfo.NwirePins.Add(new ProtocolAwarePin { OutClk = "CLK2" });

            sut.AddFrcPinFromPowerInfo(ws);

            Assert.AreEqual("99", sut.SettingInfo.NwirePins.First().PowerUpSeq);
        }

        // 4. target no exist
        [TestMethod]
        public void AddFrcPin_NoTarget()
        {
            (NwireSingleton sut, ExcelWorksheet ws) = CreateSut(["CLK1"], ("CLK1", "SEQ", "DOWN"));

            sut.AddFrcPinFromPowerInfo(ws);

            Assert.AreEqual(0, sut.SettingInfo.NwirePins.Count);
        }

        [TestMethod]
        public void ResolveVoltage_WithUnit_Should_ReturnCombined()
        {
            string result = NwireSingleton.ResolveVoltage("3.3V");
            Assert.AreEqual("3.3", result);
        }

        [TestMethod]
        public void ResolveVoltage_NoUnit_Should_ReturnValue()
        {
            string result = NwireSingleton.ResolveVoltage("3.3");
            Assert.AreEqual("3.3", result);
        }

        [TestMethod]
        public void TryAddPattern_LengthLessOrEqual20_ShouldReturnFalse()
        {
            var dic = new Dictionary<string, string>(StringExtensions.IgnoreCase);
            int cnt = 1;

            bool result = NwireSingleton.TryAddPattern(dic, "SHORT_NAME", ref cnt);

            Assert.IsFalse(result);
            Assert.AreEqual(0, dic.Count);
            Assert.AreEqual(1, cnt);
        }

        [TestMethod]
        public void TryAddPattern_ValidNewLongPattern_ShouldAddAndIncreaseCounter()
        {
            var dic = new Dictionary<string, string>(StringExtensions.IgnoreCase);
            int cnt = 1;

            bool result = NwireSingleton.TryAddPattern(dic, "THIS_IS_A_VERY_LONG_PATTERN_NAME", ref cnt);

            Assert.IsTrue(result);
            Assert.AreEqual(1, dic.Count);
            Assert.AreEqual("Pattern_1", dic["THIS_IS_A_VERY_LONG_PATTERN_NAME"]);
            Assert.AreEqual(2, cnt);
        }

        [TestMethod]
        public void TryAddPattern_DuplicatePattern_ShouldNotAddAgain()
        {
            var dic = new Dictionary<string, string>(StringExtensions.IgnoreCase)
        {
            { "THIS_IS_A_VERY_LONG_PATTERN_NAME", "Pattern_1" }
        };
            int cnt = 2;

            bool result = NwireSingleton.TryAddPattern(dic, "THIS_IS_A_VERY_LONG_PATTERN_NAME", ref cnt);

            Assert.IsFalse(result);
            Assert.AreEqual(1, dic.Count);
            Assert.AreEqual(2, cnt);
        }

        [TestMethod]
        public void ParseChannel_With_Valid_Data_Should_Return_True_And_Parse_Correctly()
        {
            // Arrange
            var target = new NwireSingleton();
            string? slot = null;
            int ch = 0;

            // Act
            bool result = target.ParseChannel("SLOT1.ch3", ref slot, ref ch);

            // Assert
            Assert.IsTrue(result);
            Assert.AreEqual("SLOT1", slot);
            Assert.AreEqual(3, ch);
        }

        [TestMethod]
        public void ParseChannel_With_Invalid_Data_Should_Return_False()
        {
            // Arrange
            var target = new NwireSingleton();
            string slot = "INIT";
            int ch = 99;

            // Act
            bool result = target.ParseChannel("INVALID_FORMAT", ref slot, ref ch);

            // Assert
            Assert.IsFalse(result);
            Assert.AreEqual("INIT", slot);
            Assert.AreEqual(99, ch);
        }

        [TestMethod]
        public void ShouldAddReferenceFlow_When_NotMatchRegex_Should_Return_True()
        {
            // Arrange
            var target = new NwireSingleton();

            // Act
            bool result = target.ShouldAddReferenceFlow("FLOW_A");

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void ShouldAddReferenceFlow_When_MatchRegex_Should_Return_False()
        {
            // Arrange
            var target = new NwireSingleton();

            // Act
            bool result = target.ShouldAddReferenceFlow("t1:abc");

            // Assert
            Assert.IsFalse(result);
        }

        [TestInitialize]
        public void Init()
        {
            _sut = new NwireSingleton();
        }

        [TestMethod]
        public void IsInTheSameChannelBlockCore_SameSlot_SameBlock_ShouldReturnTrue()
        {
            bool result = _sut.IsInTheSameChannelBlockCore("S1", 0, "S1", 31);
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void IsInTheSameChannelBlockCore_SameSlot_DifferentBlock_ShouldReturnFalse()
        {
            bool result = _sut.IsInTheSameChannelBlockCore("S1", 31, "S1", 32);
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void IsInTheSameChannelBlockCore_DifferentSlot_SameBlock_ShouldReturnFalse()
        {
            bool result = _sut.IsInTheSameChannelBlockCore("S1", 0, "S2", 0);
            Assert.IsFalse(result);
        }
    }
}
