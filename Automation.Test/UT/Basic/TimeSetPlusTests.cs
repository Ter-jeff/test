using System.Collections.Generic;
using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Business;
using Automation.Singleton;
using Automation.Static;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.Settings;

namespace Automation.Test.UT.Basic
{
    [TestClass]
    public class TimeSetPlusTests : FunctionTestBase
    {
        [TestInitialize]
        public void Setup()
        {
            NwireSingleton.Initialize();
            NwireSingleton.Instance().SettingInfo.NwirePins =
            [
                new() { OutClk = "NWPIN1" }
            ];
            TestProgram.Clear();
        }

        [TestMethod]
        public void PlusFlow_ShouldAddNwireTimeSet_WhenHasNwirePins()
        {
            var multiTimeSet = new TimeSetSheets()
            {
                new ComTimeSetBasicSheet("SheetName"){ }
            };

            var nonFrcNWiress = new List<NonFrcNWires>()
            {
                new()
                {
                    ProtocalType = "ProtocalType" , FunctionName = "SCLK"
                },
                new()
                {
                    ProtocalType = "ProtocalType" , FunctionName = "SYNCN"
                },
                new()
                {
                    ProtocalType = "ProtocalType" , FunctionName = "RESETN"
                }
            };
            var tsPlus = new TimeSetPlus(nonFrcNWiress);

            tsPlus.PlusFlow(multiTimeSet);

            Assert.AreEqual("SheetName", multiTimeSet[0].Name);
            Assert.AreEqual("TIMESET_nWire", multiTimeSet[1].Name);
            Assert.AreEqual(3, multiTimeSet.Count);
        }

        [TestMethod]
        public void PlusFlow_ShouldAddSpiRomTimeSet_WhenDeviceNeedsSpiRomPower()
        {
            var multiTimeSet = new TimeSetSheets();
            var tsPlus = new TimeSetPlus([]);

            tsPlus.PlusFlow(multiTimeSet, deviceNeedSpiRomPower: true);

            Assert.IsTrue(multiTimeSet.Any(s => s.Name == "TSB_Write_SPIROM"));
        }

        [TestMethod]
        public void AddNwirePort_ShouldNotDuplicateTimingRows()
        {
            var tsSheet = new ComTimeSetBasicSheet("TEST");
            var tsPlus = new TimeSetPlus([]);

            tsPlus.AddNwirePort(tsSheet);
            int firstCount = tsSheet.Rows.Count;

            tsPlus.AddNwirePort(tsSheet);
            int secondCount = tsSheet.Rows.Count;

            Assert.AreEqual(firstCount, secondCount);
        }

        [TestMethod]
        public void GenerateSpiRomTimeSet_ShouldCreateCorrectTimeSet()
        {
            // Act
            ComTimeSetBasicSheet result = new TimeSetPlus([]).GenerateSpiRomTimeSet();

            // Assert
            Assert.AreNotEqual(null, result, "TimeSet should not be null");
            Assert.AreEqual("Single", result.TimingMode, "TimingMode should be 'Single'");
            Assert.AreEqual(2, result.Rows.Count, "There should be 2 TimeSet");

            TSet tSet = result.Rows.First();
            Assert.AreEqual("TS1", tSet.Name, "TimeSet name should be TS1");
            Assert.AreEqual("=1/_" + AcConst.TckFreq + "_VAR", tSet.CyclePeriod, "CyclePeriod formula mismatch");

            TimingRow? ssinRow = tSet.TimingRows.FirstOrDefault(r => r.PinGrpName == "SPI0_SSIN");
            Assert.AreNotEqual(null, ssinRow, "SPI0_SSIN row should exist");
            Assert.AreEqual("=1/_" + AcConst.TckFreq + "_VAR*0.25", ssinRow!.DriveData);

            TimingRow? sclkRow = tSet.TimingRows.FirstOrDefault(r => r.PinGrpName == "SPI0_SCLK");
            Assert.AreNotEqual(null, sclkRow, "SPI0_SCLK row should exist");
            Assert.AreEqual("=1/_" + AcConst.TckFreq + "_VAR*0.5", sclkRow!.DriveData);
            Assert.AreEqual("=1/_" + AcConst.TckFreq + "_VAR", sclkRow.DriveReturn);

            TimingRow? mosiRow = tSet.TimingRows.FirstOrDefault(r => r.PinGrpName == "SPI0_MOSI");
            Assert.AreNotEqual(null, mosiRow, "SPI0_MOSI row should exist");

            TimingRow? misoRow = tSet.TimingRows.FirstOrDefault(r => r.PinGrpName == "SPI0_MISO");
            Assert.AreNotEqual(null, misoRow, "SPI0_MISO row should exist");
            Assert.AreEqual("0", misoRow!.DriveData);
            Assert.AreEqual("=1/_" + AcConst.TckFreq + "_VAR*0.5", misoRow.CompareOpen);
        }

        private static void InvokeAddNonFrcPort(TimeSetPlus tsPlus, ComTimeSetBasicSheet timeSet)
        {
            tsPlus.AddNonFrcPort(timeSet);
        }

        [TestMethod]
        public void AddNonFrcPort_NullSetting_NoRowsAdded()
        {
            // Arrange
            var tsSheet = new ComTimeSetBasicSheet("TEST");
            var tsPlus = new TimeSetPlus(null);

            // Act
            InvokeAddNonFrcPort(tsPlus, tsSheet);

            // Assert
            Assert.AreEqual(0, tsSheet.Rows.Count);
        }

        [TestMethod]
        public void AddNonFrcPort_SclkFunction_SetsScaledDriveDataAndOffCompareMode()
        {
            // Arrange
            var tsSheet = new ComTimeSetBasicSheet("TEST");
            var settings = new List<NonFrcNWires>
            {
                new() { PortName = "P1", ProtocalType = "TS1", DeviecPinName = "Pin1", FunctionName = "SCLK" }
            };
            var tsPlus = new TimeSetPlus(settings);

            // Act
            InvokeAddNonFrcPort(tsPlus, tsSheet);

            // Assert
            Assert.AreEqual(1, tsSheet.Rows.Count);
            TSet row = tsSheet.Rows.First();
            Assert.AreEqual("TS1", row.Name);
            TimingRow timingRow = row.TimingRows.First();
            Assert.AreEqual((0.000001 * 0.2).ToString(), timingRow.DriveData);
            Assert.AreEqual("Off", timingRow.CompareMode);
            Assert.AreEqual("RTN", timingRow.DataFmt);
            Assert.AreEqual((0.000001 * 0.5).ToString(), timingRow.DriveReturn);
        }

        [TestMethod]
        public void AddNonFrcPort_SyncnFunction_SetsScaledDriveDataAndPaCompareMode()
        {
            // Arrange
            var tsSheet = new ComTimeSetBasicSheet("TEST");
            var settings = new List<NonFrcNWires>
            {
                new() { PortName = "P1", ProtocalType = "TS1", DeviecPinName = "Pin1", FunctionName = "SYNCN" }
            };
            var tsPlus = new TimeSetPlus(settings);

            // Act
            InvokeAddNonFrcPort(tsPlus, tsSheet);

            // Assert
            TimingRow timingRow = tsSheet.Rows.First().TimingRows.First();
            Assert.AreEqual((0.000001 * 0.1).ToString(), timingRow.DriveData);
            Assert.AreEqual("PA", timingRow.CompareMode);
            Assert.AreEqual(0.000001.ToString(), timingRow.CompareOpen);
        }

        [TestMethod]
        public void AddNonFrcPort_ResetnFunction_SetsFixedDriveData()
        {
            // Arrange
            var tsSheet = new ComTimeSetBasicSheet("TEST");
            var settings = new List<NonFrcNWires>
            {
                new() { PortName = "P1", ProtocalType = "TS1", DeviecPinName = "Pin1", FunctionName = "RESETN" }
            };
            var tsPlus = new TimeSetPlus(settings);

            // Act
            InvokeAddNonFrcPort(tsPlus, tsSheet);

            // Assert
            TimingRow timingRow = tsSheet.Rows.First().TimingRows.First();
            Assert.AreEqual(0.000000003125.ToString(), timingRow.DriveData);
            Assert.AreEqual("Off", timingRow.CompareMode);
        }

        [TestMethod]
        public void AddNonFrcPort_DinFunction_SetsPaCompareModeWithoutSpecialDriveData()
        {
            // Arrange
            var tsSheet = new ComTimeSetBasicSheet("TEST");
            var settings = new List<NonFrcNWires>
            {
                new() { PortName = "P1", ProtocalType = "TS1", DeviecPinName = "Pin1", FunctionName = "DIN" }
            };
            var tsPlus = new TimeSetPlus(settings);

            // Act
            InvokeAddNonFrcPort(tsPlus, tsSheet);

            // Assert
            TimingRow timingRow = tsSheet.Rows.First().TimingRows.First();
            Assert.AreEqual(string.Empty, timingRow.DriveData);
            Assert.AreEqual("PA", timingRow.CompareMode);
            Assert.AreEqual(0.000001.ToString(), timingRow.CompareOpen);
        }

        [TestMethod]
        public void AddNonFrcPort_UnknownFunction_DefaultsToOffCompareMode()
        {
            // Arrange
            var tsSheet = new ComTimeSetBasicSheet("TEST");
            var settings = new List<NonFrcNWires>
            {
                new() { PortName = "P1", ProtocalType = "TS1", DeviecPinName = "Pin1", FunctionName = "OTHER" }
            };
            var tsPlus = new TimeSetPlus(settings);

            // Act
            InvokeAddNonFrcPort(tsPlus, tsSheet);

            // Assert
            TimingRow timingRow = tsSheet.Rows.First().TimingRows.First();
            Assert.AreEqual(string.Empty, timingRow.DriveData);
            Assert.AreEqual("Off", timingRow.CompareMode);
        }

        [TestMethod]
        public void AddNonFrcPort_MultipleItemsInSameGroup_SecondItemGetsNrFormatAndNoDriveReturn()
        {
            // Arrange
            var tsSheet = new ComTimeSetBasicSheet("TEST");
            var settings = new List<NonFrcNWires>
            {
                new() { PortName = "P1", ProtocalType = "TS1", DeviecPinName = "Pin1", FunctionName = "SCLK" },
                new() { PortName = "P1", ProtocalType = "TS1", DeviecPinName = "Pin2", FunctionName = "SCLK" }
            };
            var tsPlus = new TimeSetPlus(settings);

            // Act
            InvokeAddNonFrcPort(tsPlus, tsSheet);

            // Assert - both items share PortName "P1", so they're grouped into a single TSet row
            Assert.AreEqual(1, tsSheet.Rows.Count);
            List<TimingRow> timingRows = tsSheet.Rows.First().TimingRows;
            Assert.AreEqual(2, timingRows.Count);
            Assert.AreEqual("RTN", timingRows[0].DataFmt);
            Assert.AreEqual((0.000001 * 0.5).ToString(), timingRows[0].DriveReturn);
            Assert.AreEqual("NR", timingRows[1].DataFmt);
            Assert.AreEqual(string.Empty, timingRows[1].DriveReturn);
        }

        [TestMethod]
        public void AddNonFrcPort_DuplicateExistingName_SkipsBuildingTimingRowsForThatItem()
        {
            // Arrange - a row named "TS1" already exists in the sheet; the matching item is
            // skipped via `continue` before any TimingRow is built, but the outer AddRow call
            // still unconditionally appends a (now-empty) TSet for the group.
            var tsSheet = new ComTimeSetBasicSheet("TEST");
            tsSheet.AddRow(new ComTimeSetBasic { Name = "TS1" });
            var settings = new List<NonFrcNWires>
            {
                new() { PortName = "P1", ProtocalType = "TS1", DeviecPinName = "Pin1", FunctionName = "SCLK" }
            };
            var tsPlus = new TimeSetPlus(settings);

            // Act
            InvokeAddNonFrcPort(tsPlus, tsSheet);

            // Assert
            Assert.AreEqual(2, tsSheet.Rows.Count);
            Assert.AreEqual(0, tsSheet.Rows.Last().TimingRows.Count);
        }
    }
}
