using System;
using System.Collections.Generic;
using System.Linq;

using BinCutScriptLib.Base;
using BinCutScriptLib.Comparer;

using CommonLib.Extension;

using IgxlLib.Enums;
using IgxlLib.IgxlSheets;

using TestPlanLib;
using TestPlanLib.Basic;
using TestPlanLib.BinCut;
using TestPlanLib.BinCut.Binning;
using TestPlanLib.BinCut.Flow;
using TestPlanLib.BinCut.NonIgxlSheet.HarvestDssc;
using TestPlanLib.BinCut.PowerBinning;

using PinInfo = TestPlanLib.BinCut.Binning.PinInfo;

namespace BinCutScriptLib.Static
{
    public class BinCutData
    {
        public static Job Job { get; set; } = new Job(nameof(EnumJob.None));
        public static BinCutFlowTables BinCutFlowTables { get; set; } = [];
        public static BinCutFlowSheets PostFlowSheets { get; set; } = [];
        public static BinningTables BinningTables { get; set; } = [];
        public static BinningTables OtherRailTables { get; set; } = [];
        public static IdsDistributionTable IdsdistributionTable { get; set; } = new IdsDistributionTable();
        public static SelsrmMappingSheet SelsrmMappingSheet { get; set; } = new SelsrmMappingSheet();
        public static Dictionary<string, PowerBinningSheet> PowerBinningSheetList { get; set; } = new Dictionary<string, PowerBinningSheet>(StringExtensions.IgnoreCase);
        public static PwrbinSeqSheet? PwrbinSeqSheet { get; set; }
        public static PowerBinningHavrSheet? PowerBinHavrSheet { get; set; }
        //public static SepvmfusingSheet SepvmfusingSheet { get; set; }
        public static BinCutPatternReport BinCutPatternReport { get; set; } = new();
        public static bool IsSepvmfusing { get; set; }

        public static List<string> AllPowers { get; set; } = [];
        public static List<string> PowerPins { get; set; } = [];
        //public static List<List<string>> PostPowerPins { get; set; } = new List<List<string>>();
        public static Dictionary<string, string> ModeVsCorePowerName { get; set; } = new Dictionary<string, string>(StringExtensions.IgnoreCase);
        public static List<PinInfo> PinInfos { get; set; } = [];

        public static VoltageTable? NvTable { get; set; } //_VAR
        public static VoltageTable? VrsTable { get; set; } //VOP_VAR
        public static VoltageTable? NvAllTable { get; set; }
        public static VoltageTable? VrsAllTable { get; set; }
        public static InstanceSheet? TestInstanceSheet { get; set; }
        public static bool HasBinCutInstance { get; set; }
        public static PinMapSheet? PinMapData { get; set; }
        public static JobListSheet? JoblistSheet { get; set; }
        public static List<PatternData> PatList { get; set; } = [];
        public static List<string> GoodBins { get; set; } = [];

        public static bool HasPatList { get; set; }

        //public static HarvPmodeTableModeSheet HarvPmodeTableModeSheet { get; set; }
        public static HarvPmodeTableHarvCheckSheet? HarvPmodeTableHarvCheckSheet { get; set; }
        public static HarvPmodeTableGroupSheet? HarvPmodeTableGroupSheet { get; set; }
        public static HarvPmodeTableOffsetSheet? HarvPmodeTableOffsetSheet { get; set; }
        public static HarvPmodeTableFstpSheet? HarvPmodeTableFstpSheet { get; set; }
        public static HarvMappingTableSheet? HarvMappingTableSheet { get; set; }
        //public static HarvPinGroupTableReader.HarvPinGroup HarvPinGroupsheet { get; set; }
        public static DigSrcInstructionSheet? DigSrcInstructionSheet { get; set; }
        //public static bool HasRunBinCutValidation { get; set; }

        public static void Initialize()
        {
            BinCutFlowTables = [];
            PostFlowSheets = [];
            BinningTables = [];
            OtherRailTables = [];
            IdsdistributionTable = new IdsDistributionTable();
            SelsrmMappingSheet = new SelsrmMappingSheet();
            PowerBinningSheetList = new Dictionary<string, PowerBinningSheet>(StringExtensions.IgnoreCase);
            PwrbinSeqSheet = null;
            IsSepvmfusing = false;
            ModeVsCorePowerName = new Dictionary<string, string>(StringExtensions.IgnoreCase);
            PowerPins = [];
            AllPowers = [];
            PinInfos = [];
        }

        public static string GetNewTestSettingString(List<Tuple<string, string>> list)
        {
            var pins = new Dictionary<string, double>();
            for (int i = 0; i < PowerPins.Count; i++)
            {
                foreach (Tuple<string, string> item in list)
                {
                    if (PowerPins[i] == item.Item1)
                    {
                        _ = double.TryParse(item.Item2, out double value);
                        pins.Add(PowerPins[i], value);
                    }
                }
            }
            return GetComposeStr(pins);
        }

        private static string GetComposeStr(Dictionary<string, double> pins)
        {
            var list = pins.Select(x => $"{x.Key}=" + $"{x.Value:F3}").ToList();
            return string.Join(",", list);
        }
    }
}
