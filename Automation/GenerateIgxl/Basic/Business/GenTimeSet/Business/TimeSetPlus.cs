using System;
using System.Collections.Generic;
using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.Basic.Business.GenAc;
using Automation.Reader;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;
using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

using TestPlanLib.Settings;

using TimeSetBasic = IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet.ComTimeSetBasic;
using TimeSetBasicSheet = IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet.ComTimeSetBasicSheet;

namespace Automation.GenerateIgxl.Basic.Business.GenTimeSet.Business
{
    public class TimeSetPlus
    {
        public const string NwireTimeSetSheetName = "TIMESET_nWire";
        public const string TsbWriteSpirom = "TSB_Write_SPIROM";
        public const string TsbContiPat = "TSB_Conti_pat";
        private const string OutClkSetUp = "PA_FRC";
        private const string RefClkSetUp = "PA_DUT_CLK";
        private const string RtnFmt = "RTN";
        private const string RefClkFmt = "STAY";

        private const string Spi0Ssin = "SPI0_SSIN";
        private const string Spi0Sclk = "SPI0_SCLK";
        private const string Spi0Mosi = "SPI0_MOSI";
        private const string Spi0Miso = "SPI0_MISO";

        private const string SrcPat = "PAT";
        private const string Io = "i/o";
        private const string FmtNr = "NR";
        private const string FmtRl = "RL";
        private const string ModeOff = "Off";
        private const string ModeEdgeCap = "Edge Capture";
        private const string ModeMachine = "Machine";
        private const string ModeEdge = "Edge";
        private readonly List<NonFrcNWires> _nonFrcsetting;

        public TimeSetPlus(List<NonFrcNWires> nonFrcsetting)
        {
            _nonFrcsetting = nonFrcsetting;
        }

        public virtual void PlusFlow(TimeSetSheets timeSet, bool deviceNeedSpiRomPower = false)
        {
            // add nwire port to all timeset
            foreach (TimeSetBasicSheet timeSetSheet in timeSet)
            {
                AddNwirePort(timeSetSheet);
                AddNonFrcPort(timeSetSheet);
            }

            bool hasUltraFlex = TestPlanStatic.Equipments.Exists(x => x.Equals(EnumEquipment.UltraFlex));
            if (hasUltraFlex)
            {
                // create nwire timeset sheet
                if (NwireSingleton.Instance().HasNwirePin)
                {
                    TimeSetBasicSheet nWireTimeSet = GenerateNwireTimeSet();
                    timeSet.AddTimeSetSheet(nWireTimeSet);
                }
            }

            //create workingZ timeset sheet
            TimeSetBasicSheet workingzTimeSet = GenerateWalkingZTimeSet();
            timeSet.AddTimeSetSheet(workingzTimeSet);

            //create Rtos/SPIRom timeset sheet
            bool hasSpiPwr = false;
            PinMapSheet pinMap = TestProgram.IgxlWorkBk.PinMapPair.Value;
            if (pinMap != null)
            {
                if (pinMap.PinList != null && pinMap.PinList.Any()) //LocalSpec.LocalSpecs.IgxlWorkBk!= null)
                {
                    if (pinMap.IsPinExist(CommonConst.SpiRomPwr) || pinMap.IsPinExist(CommonConst.RtosRomPwr))
                    {
                        hasSpiPwr = true;
                    }
                }
                else
                {
                    hasSpiPwr = deviceNeedSpiRomPower;
                }
            }
            else
            {
                hasSpiPwr = deviceNeedSpiRomPower;
            }
            if (hasSpiPwr)
            {
                TimeSetBasicSheet spiRomTimeSet = GenerateSpiRomTimeSet();
                timeSet.AddTimeSetSheet(spiRomTimeSet);
            }
        }

        /// <summary>
        /// Add Nwire port tset for every timeset
        /// </summary>
        /// <param name="timeSet"></param>
        internal void AddNwirePort(TimeSetBasicSheet timeSet)
        {
            List<ProtocolAwarePin> nWirePin = NwireSingleton.Instance().SettingInfo.NwirePins;
            bool isUltraFlexPlusOnly = !TestPlanStatic.Equipments.First().Equals(EnumEquipment.UltraFlex);
            if (!isUltraFlexPlusOnly)
            {
                foreach (ProtocolAwarePin awarePin in nWirePin)
                {
                    var timeSetBasic = new TimeSetBasic { Name = awarePin.CreatePortName(EnumEquipment.UltraFlex) };
                    if (timeSet.Rows.Any(p => p.Name == timeSetBasic.Name))
                    {
                        continue;
                    }

                    //timeSetBasic.CyclePeriod = awarePin.IsSuperClock ? "=2/_" + awarePin.CreateFreqSpecName() : "=1/_" + awarePin.CreateFreqSpecName();
                    timeSetBasic.CyclePeriod = "=1/_" + awarePin.CreateFreqVarName();

                    var clkRow = new TimingRow
                    {
                        PinGrpName = awarePin.CreatePaClkPinName(EnumEquipment.UltraFlex),
                        PinGrpSetup = OutClkSetUp,
                        DataSrc = "PA",
                        DataFmt = RtnFmt,
                        DriveData = "0",
                        DriveReturn = "=0.5/_" + awarePin.CreateAcSymbol() + "_VAR",
                        CompareMode = "Off",
                        EdgeMode = "auto"
                    };

                    var refClkRow = new TimingRow
                    {
                        PinGrpName = awarePin.RefClk,
                        PinGrpClockPeriod = "=1/_" + NwireSingleton.SbcGlbSpecName,
                        PinGrpSetup = RefClkSetUp,
                        DataSrc = "PA",
                        DataFmt = RefClkFmt,
                        DriveOff = "0",
                        CompareMode = "Off",
                        EdgeMode = "auto"
                    };

                    timeSetBasic.TimingRows.Add(clkRow);
                    timeSetBasic.TimingRows.Add(refClkRow);
                    timeSet.AddRow(timeSetBasic);
                }
            }
        }

        internal void AddNonFrcPort(TimeSetBasicSheet timeSet)
        {
            const double cyclePeriod = 0.000001;

            if (_nonFrcsetting == null)
            {
                return;
            }

            var list = _nonFrcsetting.GroupBy(x => x.PortName).ToList();
            foreach (IGrouping<string, NonFrcNWires> set in list)
            {
                var timeSetBasic = new TimeSetBasic();
                int cnt = 1;
                foreach (NonFrcNWires item in set)
                {
                    timeSetBasic.Name = item.ProtocalType;
                    if (timeSet.Rows.Any(p => p.Name == timeSetBasic.Name))
                    {
                        continue;
                    }

                    timeSetBasic.CyclePeriod = cyclePeriod.ToString();

                    var clkRow = new TimingRow
                    {
                        PinGrpName = item.DeviecPinName,
                        PinGrpSetup = OutClkSetUp,
                        DataSrc = "PA",
                        DataFmt = cnt == 1 ? RtnFmt : FmtNr
                    };
                    if (item.FunctionName.Equals("SCLK", StringComparison.OrdinalIgnoreCase))
                    {
                        clkRow.DriveData = (cyclePeriod * 0.2).ToString();
                    }
                    else if (item.FunctionName.Equals("SYNCN", StringComparison.OrdinalIgnoreCase))
                    {
                        clkRow.DriveData = (cyclePeriod * 0.1).ToString();
                    }
                    else if (item.FunctionName.Equals("RESETN", StringComparison.OrdinalIgnoreCase))
                    {
                        clkRow.DriveData = 0.000000003125.ToString();
                    }

                    clkRow.DriveReturn = cnt == 1 ? (cyclePeriod * 0.5).ToString() : "";

                    if (item.FunctionName.Equals("SYNCN", StringComparison.OrdinalIgnoreCase) || item.FunctionName.Equals("DIN", StringComparison.OrdinalIgnoreCase))
                    {
                        clkRow.CompareMode = "PA";
                        clkRow.CompareOpen = cyclePeriod.ToString();
                    }
                    else
                    {
                        clkRow.CompareMode = "Off";
                    }

                    timeSetBasic.TimingRows.Add(clkRow);
                    cnt++;
                }
                timeSet.AddRow(timeSetBasic);
            }
        }

        /// <summary>
        /// Generate a Timeset for Free running clock enable
        /// </summary>
        /// <returns></returns>
        protected TimeSetBasicSheet GenerateNwireTimeSet()
        {
            var timeSet = new TimeSetBasicSheet(NwireTimeSetSheetName);

            AddNwirePort(timeSet);

            return timeSet;
        }

        /// <summary>
        /// Generate a Timeset for SpiRom
        /// </summary>
        /// <returns></returns>
        internal TimeSetBasicSheet GenerateSpiRomTimeSet()
        {
            var timeSet = new TimeSetBasicSheet(TsbWriteSpirom) { TimingMode = "Single" };
            //TS1 for SpiRom I/O Pins
            var tset = new TimeSetBasic { CyclePeriod = "=1/_" + BasicInitial.TckFreqVar + "_VAR", Name = "TS1" };

            var row = new TimingRow
            {
                PinGrpName = Spi0Ssin,
                PinGrpSetup = Io,
                DataSrc = SrcPat,
                DataFmt = FmtNr,
                DriveData = "=1/_" + BasicInitial.TckFreqVar + "_VAR*0.25",
                DriveOff = "0",
                CompareMode = ModeOff,
                EdgeMode = ModeMachine
            };
            tset.AddTimingRow(row);

            row = new TimingRow
            {
                PinGrpName = Spi0Sclk,
                PinGrpSetup = Io,
                DataSrc = SrcPat,
                DataFmt = FmtRl,
                DriveData = "=1/_" + BasicInitial.TckFreqVar + "_VAR*0.5",
                DriveReturn = "=1/_" + BasicInitial.TckFreqVar + "_VAR",
                DriveOff = "0",
                CompareMode = ModeOff,
                EdgeMode = ModeMachine
            };
            tset.AddTimingRow(row);

            row = new TimingRow
            {
                PinGrpName = Spi0Mosi,
                PinGrpSetup = Io,
                DataSrc = SrcPat,
                DataFmt = FmtNr,
                DriveData = "=1/_" + BasicInitial.TckFreqVar + "_VAR*0.25",
                DriveOff = "0",
                CompareMode = ModeOff,
                EdgeMode = ModeMachine
            };
            tset.AddTimingRow(row);

            row = new TimingRow
            {
                PinGrpName = Spi0Miso,
                PinGrpSetup = Io,
                DataSrc = SrcPat,
                DataFmt = FmtNr,
                DriveData = "0",
                DriveOff = "0",
                CompareMode = ModeEdgeCap,
                CompareOpen = "=1/_" + BasicInitial.TckFreqVar + "_VAR*0.5",
                EdgeMode = ModeMachine
            };
            tset.AddTimingRow(row);
            timeSet.AddRow(tset);

            //Don't add Blank row 2017 0831 by JN
            //TimeSetBasic tsetblank = new TimeSetBasic();
            //tsetblank.AddTimingRow(new TimingRow());
            //timeSet.AddTimeSet(tsetblank);

            AddNwirePort(timeSet);

            return timeSet;
        }

        private TimeSetBasicSheet GenerateWalkingZTimeSet()
        {
            var timeSet = new TimeSetBasicSheet(TsbContiPat) { TimingMode = "Single" };
            var tset = new TimeSetBasic { CyclePeriod = "0.000001", Name = "Tset" };

            var row = new TimingRow
            {
                PinGrpName = "All_Digital",
                PinGrpSetup = Io,
                DataSrc = SrcPat,
                DataFmt = FmtNr,
                DriveData = "0",
                DriveOff = "0",
                CompareMode = ModeEdge,
                CompareOpen = "0.0000009",
                EdgeMode = "auto"
            };
            tset.AddTimingRow(row);

            timeSet.AddRow(tset);
            return timeSet;
        }
    }
}
