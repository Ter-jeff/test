using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using CommonReaderLib;

using TestPlanLib.Static;

namespace TestPlanLib.BinCut.BinCutInstance
{
    public class BinCutInstanceRow : MyRow
    {
        private readonly FlowRowData _flow = new();
        private readonly PatternRowData _pattern = new();
        private readonly MiscRowData _misc = new();

        public string FlowNameOri { get => _flow.FlowNameOri; set => _flow.FlowNameOri = value; }
        public string FlowName
        {
            get { return Reg._regex.Replace(FlowNameOri.Trim(), " "); }
            set { FlowNameOri = value; }
        }
        public string Instance { get => _flow.Instance; set => _flow.Instance = value; }
        public string Opcode { get => _flow.Opcode; set => _flow.Opcode = value; }
        public string EnableAndDevice { get => _flow.EnableAndDevice; set => _flow.EnableAndDevice = value; }
        public string BintableEnableWd { get => _flow.BintableEnableWd; set => _flow.BintableEnableWd = value; }
        public string SubFlow { get => _flow.SubFlow; set => _flow.SubFlow = value; }
        public string EnableFlow { get => _flow.EnableFlow; set => _flow.EnableFlow = value; }
        public string JobTestStage { get => _flow.JobTestStage; set => _flow.JobTestStage = value; }
        public string SiteVar { get => _flow.SiteVar; set => _flow.SiteVar = value; }
        public string FailFlag { get => _flow.FailFlag; set => _flow.FailFlag = value; }
        public string PassFlag { get => _flow.PassFlag; set => _flow.PassFlag = value; }
        public string BinOutStage { get => _flow.BinOutStage; set => _flow.BinOutStage = value; }
        public string TempMon { get => _flow.TempMon; set => _flow.TempMon = value; }

        public bool IsBypassBinOut
        {
            get
            {
                if (BinOutStage.EqualsIgnoreCase("x") || BinOutStage.Replace(" ", "").EqualsIgnoreCase("NoBinOut"))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }
        public bool IsOpcode => !string.IsNullOrEmpty(Opcode);
        public string DCcategory { get => _flow.DCcategory; set => _flow.DCcategory = value; }
        public string Levels { get => _flow.Levels; set => _flow.Levels = value; }
        public string HarvPinGrpEnable { get => _flow.HarvPinGrpEnable; set => _flow.HarvPinGrpEnable = value; }
        public string HarvestBinningFlag { get => _flow.HarvestBinningFlag; set => _flow.HarvestBinningFlag = value; }
        public EvsRowData Evs = new();
        public string Burst { get => _flow.Burst; set => _flow.Burst = value; }
        public string TimeSet { get => _flow.TimeSet; set => _flow.TimeSet = value; }
        public string ShiftSpeed { get => _flow.ShiftSpeed; set => _flow.ShiftSpeed = value; }
        public string PatternPinGroup { get => _pattern.PatternPinGroup; set => _pattern.PatternPinGroup = value; }
        public string PinGroupBinoutFlag { get => _pattern.PinGroupBinoutFlag; set => _pattern.PinGroupBinoutFlag = value; }
        public string IsBinout { get => _misc.IsBinout; set => _misc.IsBinout = value; }
        public string MultiFstpEnable { get => _pattern.MultiFstpEnable; set => _pattern.MultiFstpEnable = value; }
        public string UserFunction { get => _pattern.UserFunction; set => _pattern.UserFunction = value; }
        public string SplitKey { get => _misc.SplitKey; set => _misc.SplitKey = value; }
        public string AcSpec { get => _misc.AcSpec; set => _misc.AcSpec = value; }
        public string PatSetNameOrange { get => _flow.PatSetNameOrange; set => _flow.PatSetNameOrange = value; }
        public string RetentionWaitTime { get => _misc.RetentionWaitTime; set => _misc.RetentionWaitTime = value; }
        public List<int> RetentionWaitIdx { get => _misc.RetentionWaitIdx; set => _misc.RetentionWaitIdx = value; }
        public BincutInstanceType Type { get => _misc.Type; set => _misc.Type = value; }
        #region CPM
        public string SpecialSetting { get => _misc.SpecialSetting; set => _misc.SpecialSetting = value; }
        public string Overlay { get => _misc.Overlay; set => _misc.Overlay = value; }
        public string Char { get => _misc.Char; set => _misc.Char = value; }
        #endregion
        #region Efuse
        public bool IsCallFlow { get { return Instance.StartsWithIgnoreCase("Flow_"); } }
        public string FunctionName { get => _misc.FunctionName; set => _misc.FunctionName = value; }
        public string PayloadValue
        {
            get
            {
                return PayloadList.Count == 0 ? "" : PayloadList[0];
            }
        }
        #endregion

        public string IsHarvesting
        {
            get
            {
                if (string.IsNullOrEmpty(_misc.IsHarvestingCache))
                {
                    string[] patternPinGroupItems = PatternPinGroup.Split(';');
                    foreach (string item in patternPinGroupItems)
                    {
                        if (!item.Contains("EnableCoreHarvest"))
                        {
                            continue;
                        }

                        _misc.IsHarvestingCache = item.Split(':').LastOrDefault()!.EqualsIgnoreCase("TRUE") ? "TRUE" : "FALSE";
                    }

                    if (string.IsNullOrEmpty(_misc.IsHarvestingCache))
                    {
                        _misc.IsHarvestingCache = "FALSE";
                    }
                }

                return _misc.IsHarvestingCache;
            }
            set => _misc.IsHarvestingCache = value;
        }
        public string EnableCoreMask
        {
            get
            {
                if (string.IsNullOrEmpty(_misc.EnableCoreMaskCache))
                {
                    string[] patternPinGroupItems = PatternPinGroup.Split(';');
                    foreach (string item in patternPinGroupItems)
                    {
                        if (!item.Contains("EnableCoreMask"))
                        {
                            continue;
                        }

                        _misc.EnableCoreMaskCache = item.Split(':').LastOrDefault()!.EqualsIgnoreCase("TRUE") ? "TRUE" : "FALSE";
                    }
                    if (string.IsNullOrEmpty(_misc.EnableCoreMaskCache))
                    {
                        _misc.EnableCoreMaskCache = "TRUE";
                    }
                }
                return _misc.EnableCoreMaskCache;
            }
            set => _misc.EnableCoreMaskCache = value;
        }

        public List<string> PatternList { get => _pattern.PatternList; set => _pattern.PatternList = value; }
        public List<string> InitList { get => _pattern.InitList; set => _pattern.InitList = value; }
        public Dictionary<int, string> PatWithIndex { get => _pattern.PatWithIndex; set => _pattern.PatWithIndex = value; }
        public List<string> PayloadList { get => _pattern.PayloadList; set => _pattern.PayloadList = value; }

        //Optional
        public string PerfMode { get => _pattern.PerfMode; set => _pattern.PerfMode = value; }
        public float PmOrder { get => _pattern.PmOrder; set => _pattern.PmOrder = value; }
        public string CallFlow { get => _pattern.CallFlow; set => _pattern.CallFlow = value; }

        public BinCutInstanceRow(string sheetName = "")
        {
            SheetName = sheetName;
        }

        public BinCutInstanceRow(BinCutInstanceRow binCutInstanceRow) : base(binCutInstanceRow)
        {
            if (binCutInstanceRow == null)
            {
                return;
            }

            _flow = binCutInstanceRow._flow.Copy();
            _pattern = binCutInstanceRow._pattern.Copy();
            _misc = binCutInstanceRow._misc.Copy();
            Evs = binCutInstanceRow.Evs.Copy();
        }

        public BinCutInstanceRow Copy()
        {
            return new BinCutInstanceRow(this);
        }

        public string GetColumnA()
        {
            return string.IsNullOrEmpty(FlowName) ? SheetName + ",Row" + RowNum : SheetName + ",Row" + RowNum + "," + FlowName;
        }

        public Dictionary<string, string>? GetSpecialSettings()
        {
            List<string> specialSettings = [.. SpecialSetting.Replace("\n", "").Replace("\r", "").Trim().Split(';')];
            if (string.IsNullOrEmpty(SpecialSetting))
            {
                return null;
            }

            var dic = new Dictionary<string, string>(StringExtensions.IgnoreCase);
            if (specialSettings.Exists(x => x.Contains("FUNCTION", StringComparison.OrdinalIgnoreCase)) || specialSettings.Exists(x => x.Contains("FAIL", StringComparison.OrdinalIgnoreCase) || x.Contains("PASS", StringComparison.OrdinalIgnoreCase)))
            {
                string regexPattern = @"(?<key>[^=:;,\s]+)\s*[=:]\s*(?<value>(?:""[^""]*"")|[^;,\s]+)|(?<key>\w+)";
                MatchCollection matches = Regex.Matches(SpecialSetting, regexPattern);
                foreach (Match match in matches)
                {
                    if (match.Success)
                    {
                        string inputKey = match.Groups["key"].Value.Trim();
                        string inputValue = match.Groups["value"].Success ? match.Groups["value"].Value.Trim().Trim('"').Trim() : "";
                        if (!match.Groups["value"].Success)
                        {
                            inputValue = inputKey;
                        }
                        if (!dic.TryAdd(inputKey, inputValue))
                        {
                            dic[inputKey] += ";" + inputValue;
                        }
                    }
                }
            }
            else if (specialSettings.Exists(x => x.Contains("WRITEEFUSE", StringComparison.OrdinalIgnoreCase)))
            {
                dic.Add("WriteEfuse", "WriteEfuse");
            }
            return dic.Count != 0 ? dic : null;
        }
    }
}
