using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.GenerateIgxl.Wireless.DVDC.InputObject;
using Automation.Reader;
using Automation.Singleton;
using Automation.Static;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

namespace Automation.GenerateIgxl.HardIp.InputObject
{
    public class HardIpPattern
    {
        #region Const string of Header
        public const string TtrHeader = "TTR";
        public const string PartHeader = "Part";
        public const string NoBinOutHeader = "NoBinout";
        public const string EnableHeader = @"^Enable\s*\(all\s*sites\)";
        public const string SiteFlagHeader = @"SiteFlag\s*\(per\s*site\)";
        public const string FailFlagHeader = "FailFlag";
        public const string TestItemHeader = @"Test\s*Item";
        public const string StepHeader = "Step";
        public const string DescriptionHeader = "Description";
        public const string PatternHeader = @"Pattern[\s]*([\/]Instance)?";
        public const string ForceConditionHeader = "Condition";
        public const string RegisterAssignmentHeader = "Register";
        public const string MiscInfoHeader = "Misc.*Info";
        public const string MeasHeader = "^Meas";
        public const string LoHeader = @"\s*Lo\s*Limit.*";
        public const string HiHeader = @"\s*Hi\s*Limit.*";
        public const string CommentHeader = "Comment";
        //Add new header pattern due to ADC autogen
        public const string AnalogSetupHeader = @"Analog\s*Instrument\s*Setup";
        public const string TestNameHeader = @"^Test_*\s*Name";
        //Analog Instrument Setup	Test Name	Discard and Selected Cycle
        public static readonly List<string> KnownHeaders = new List<string>
        {
            TtrHeader, NoBinOutHeader,EnableHeader,SiteFlagHeader,FailFlagHeader, PartHeader, TestItemHeader, StepHeader, DescriptionHeader, PatternHeader, ForceConditionHeader,
            RegisterAssignmentHeader, MiscInfoHeader, MeasHeader, LoHeader, HiHeader, CommentHeader,
            AnalogSetupHeader, TestNameHeader
        };

        #endregion

        private string _functionName;
        private string _miscInfo;

        #region Property
        public int RowNum { get; set; }
        public int ColumnNum { get; set; }

        public string Failflag { get; set; }
        public string SiteFlag { get; set; }
        public string Enable { get; set; }
        public string SheetName { get; set; }
        public string SubBlock { get; set; }
        public string FunctionName { set { _functionName = value; } get { return _functionName.ToLower(); } }
        public string PatternType { get; set; }
        public string MiscInfo
        {
            get { return _miscInfo; }
            set
            {
                MiscInfoDict = TestPlanStatic.MiscInfoParser.ParseKeyValueToDictionary(value, out string newValue);
                _miscInfo = newValue;
            }
        }
        public Dictionary<string, string> MiscInfoDict { get; set; }
        public ForceClass ForceCondition { get; set; }
        public HardipCharSetup Shmoo { get; set; }
        public string TestName { get; set; }
        public List<MeasPin> OriMeasPins { get; set; }
        public List<MeasPin> MeasPins { get; set; }
        public List<MeasPin> UseLimitsH { get; set; }
        public List<MeasPin> UseLimitsL { get; set; }
        public List<MeasPin> UseLimitsN { get; set; }
        public List<ForceCondition> ForceConditionList { get; set; }

        public string TtrStr { get; set; }
        public string PartStr { get; set; }
        public HashSet<string> NoBinoutStr { get; set; }
        public Dictionary<string, string> RelaySetting { get; set; }
        public Dictionary<string, string> NewRelaySetting { get; set; }
        public PatternClass Pattern { get; set; }
        public int DupIndex { get; set; }
        public int ConditionIndex { get; set; }
        public Dictionary<string, List<SweepVData>> SweepVoltage { get; set; }
        public string RegisterAssignment { get; set; }
        public string DigSrcEquation { get; set; }
        public string CalcEqn { get; set; }
        public List<string> DspFunction { get; set; }
        public string LevelUsed { get { return ForceCondition.GetLevelSetting(); } }
        public string AcUsed { get { return ForceCondition.GetAcSetting(); } }
        public string AcSelectorUsed { get { return ForceCondition.GetAcSelector(); } }
        public string DcCategory { get { return ForceCondition.GetDcCategory(); } }
        public string AcCategory { get { return ForceCondition.GetAcCategory(); } }
        public string DcSelectorUsed { get { return ForceCondition.GetDcSelector(); } }
        public string RtosIdsTimeSetUsed { get { return ForceCondition.GetRtosIdsTimeSet(); } }
        public TimeSetUsed TimeSetUsed { get; set; }
        public Dictionary<string, List<SweepCode>> SweepCodes { get; set; }
        public List<TestPlanSequence> TestPlanSequences { get; set; }
        public List<TestPlanSequence> TestPlanSequencesRf { get; set; }
        public string SpecialMeasType { get; set; }
        public string InterposePostTest { get; set; }
        public int FlowControlFlag { get; set; }
        public bool IsFlowInsRepeat { get; set; }
        public string MixSigUsed { get; set; }
        public WirelessData WirelessData { get; set; }
        public List<PlanType> VbtTypes { get; set; }
        public string InstrumentSetup { get; set; }
        public string RfInterPose { get; set; }
        public string BlockType { get; set; }
        public string Chiplet { get; set; }
        public string RegAssignName { get; set; }
        public List<HardIpPattern> BurstPatterns { get; set; }
        public bool IsBurst { get; set; }
        public List<string> SkipList { get; set; }
        public string ForceVoltageFlag { get; set; }
        public string HipPreWriteFlag { get; set; }
        public bool IsVtShmooUsed { get; set; }
        public bool SkipDotNet { get; set; }

        public string PatternIndexFlag
        {
            get
            {
                string flag = "";

                if (DupIndex > 0)
                {
                    flag += "_" + DupIndex;
                }

                if (ConditionIndex > 0)
                {
                    flag += "_" + ConditionIndex;
                }

                return flag;
            }
        }

        public string SubBlockCopy { get; set; }
        public string SheetSubBlockName { get; set; }
        public (string ateTestCondition, string overlayName) BinCutOverlay { get; set; }
        #endregion

        public HardIpPattern()
        {
            RowNum = 0;
            ColumnNum = 0;
            SheetName = "";
            SubBlock = "";
            Pattern = new PatternClass("");
            ForceCondition = new ForceClass();
            TtrStr = "";
            PartStr = "";
            NoBinoutStr = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            RelaySetting = new Dictionary<string, string>();
            NewRelaySetting = new Dictionary<string, string>();
            DupIndex = 0;
            ConditionIndex = 0;
            MeasPins = new List<MeasPin>();
            OriMeasPins = new List<MeasPin>();
            UseLimitsH = new List<MeasPin>();
            UseLimitsL = new List<MeasPin>();
            UseLimitsN = new List<MeasPin>();
            MiscInfo = "";
            ForceConditionList = new List<ForceCondition>();
            TimeSetUsed = new TimeSetUsed();
            RegisterAssignment = "";
            DigSrcEquation = "";
            FunctionName = "";
            ForceVoltageFlag = "";
            HipPreWriteFlag = "";
            PatternType = "";
            SweepCodes = new Dictionary<string, List<SweepCode>>();
            Shmoo = new HardipCharSetup();
            CalcEqn = "";
            DspFunction = new List<string>();
            TestPlanSequences = new List<TestPlanSequence>();
            TestPlanSequencesRf = new List<TestPlanSequence>();
            SpecialMeasType = "";
            InterposePostTest = "";
            SweepVoltage = new Dictionary<string, List<SweepVData>>();
            FlowControlFlag = -1;
            IsFlowInsRepeat = false;
            MixSigUsed = "";
            WirelessData = new WirelessData();
            VbtTypes = new List<PlanType> { PlanType.Default };
            InstrumentSetup = "";
            BlockType = "";
            RfInterPose = "";
            Chiplet = "";
            RegAssignName = "";
            BurstPatterns = new List<HardIpPattern>();
            SkipList = new List<string>();
            IsVtShmooUsed = false;
            SkipDotNet = false;
        }

        public HardIpPattern(HardIpPattern other)
        {
            if (other == null)
            {
                return;
            }

            // Value Types and Strings (Shallow Copy is sufficient)
            RowNum = other.RowNum;
            ColumnNum = other.ColumnNum;
            Failflag = other.Failflag;
            SiteFlag = other.SiteFlag;
            Enable = other.Enable;
            SheetName = other.SheetName;
            SubBlock = other.SubBlock;
            _functionName = other._functionName;
            PatternType = other.PatternType;
            _miscInfo = other._miscInfo;
            TestName = other.TestName;
            TtrStr = other.TtrStr;
            PartStr = other.PartStr;
            DupIndex = other.DupIndex;
            ConditionIndex = other.ConditionIndex;
            RegisterAssignment = other.RegisterAssignment;
            DigSrcEquation = other.DigSrcEquation;
            CalcEqn = other.CalcEqn;
            SpecialMeasType = other.SpecialMeasType;
            InterposePostTest = other.InterposePostTest;
            FlowControlFlag = other.FlowControlFlag;
            IsFlowInsRepeat = other.IsFlowInsRepeat;
            MixSigUsed = other.MixSigUsed;
            InstrumentSetup = other.InstrumentSetup;
            RfInterPose = other.RfInterPose;
            BlockType = other.BlockType;
            Chiplet = other.Chiplet;
            RegAssignName = other.RegAssignName;
            IsBurst = other.IsBurst;
            ForceVoltageFlag = other.ForceVoltageFlag;
            HipPreWriteFlag = other.HipPreWriteFlag;
            IsVtShmooUsed = other.IsVtShmooUsed;
            SkipDotNet = other.SkipDotNet;
            SubBlockCopy = other.SubBlockCopy;
            SheetSubBlockName = other.SheetSubBlockName;
            BinCutOverlay = other.BinCutOverlay;

            // Deep Copy Custom Classes (Assumes these classes have copy constructors)
            Pattern = other.Pattern != null ? new PatternClass(other.Pattern) : null;
            ForceCondition = other.ForceCondition != null ? new ForceClass(other.ForceCondition) : null;
            Shmoo = other.Shmoo != null ? new HardipCharSetup(other.Shmoo) : null;
            TimeSetUsed = other.TimeSetUsed != null ? new TimeSetUsed(other.TimeSetUsed) : null;
            WirelessData = other.WirelessData != null ? new WirelessData(other.WirelessData) : null;

            // Deep Copy Lists (Assumes MeasPin, ForceCondition, etc. have Copy() or copy constructors)
            MeasPins = other.MeasPins?.Select(x => x.Copy()).ToList();
            OriMeasPins = other.OriMeasPins?.Select(x => x.Copy()).ToList();
            UseLimitsH = other.UseLimitsH?.Select(x => x.Copy()).ToList();
            UseLimitsL = other.UseLimitsL?.Select(x => x.Copy()).ToList();
            UseLimitsN = other.UseLimitsN?.Select(x => x.Copy()).ToList();
            ForceConditionList = other.ForceConditionList?.Select(x => x.Copy()).ToList();
            DspFunction = other.DspFunction != null ? new List<string>(other.DspFunction) : null;
            VbtTypes = other.VbtTypes != null ? new List<PlanType>(other.VbtTypes) : null;
            SkipList = other.SkipList != null ? new List<string>(other.SkipList) : null;
            BurstPatterns = other.BurstPatterns?.Select(x => new HardIpPattern(x)).ToList();
            TestPlanSequences = other.TestPlanSequences?.Select(x => x.Copy()).ToList();
            TestPlanSequencesRf = other.TestPlanSequencesRf?.Select(x => x.Copy()).ToList();

            // Deep Copy Dictionaries and HashSets
            NoBinoutStr = other.NoBinoutStr != null ? new HashSet<string>(other.NoBinoutStr, StringComparer.OrdinalIgnoreCase) : null;
            RelaySetting = other.RelaySetting != null ? new Dictionary<string, string>(other.RelaySetting) : null;
            NewRelaySetting = other.NewRelaySetting != null ? new Dictionary<string, string>(other.NewRelaySetting) : null;

            SweepVoltage = other.SweepVoltage?.ToDictionary(k => k.Key, v => v.Value.Select(x => x.Copy()).ToList());

            SweepCodes = other.SweepCodes?.ToDictionary(k => k.Key, v => v.Value.Select(x => x.Copy()).ToList());
        }

        public HardIpPattern Copy()
        {
            return new HardIpPattern(this);
        }

        public void Copy(HardIpPattern pattern)
        {
            SheetName = pattern.SheetName;
            SubBlock = pattern.SubBlock;
            RowNum = pattern.RowNum;
            ColumnNum = pattern.ColumnNum;
            Pattern = pattern.Pattern;
            ConditionIndex = pattern.ConditionIndex;
            TtrStr = pattern.TtrStr;
            PartStr = pattern.PartStr;
            NoBinoutStr = pattern.NoBinoutStr;
            RelaySetting = pattern.RelaySetting;
            NewRelaySetting = pattern.NewRelaySetting;
            DupIndex = pattern.DupIndex;
            _miscInfo = pattern.MiscInfo;
            MiscInfoDict = pattern.MiscInfoDict;
            ForceConditionList = pattern.ForceConditionList;
            RegisterAssignment = pattern.RegisterAssignment;
            DigSrcEquation = pattern.DigSrcEquation;
            FunctionName = pattern.FunctionName;
            PatternType = pattern.PatternType;
            TimeSetUsed = pattern.TimeSetUsed;
            SweepCodes = pattern.SweepCodes;
            Shmoo = pattern.Shmoo;
            CalcEqn = pattern.CalcEqn;
            DspFunction = pattern.DspFunction;
            TestPlanSequences = pattern.TestPlanSequences;
            TestPlanSequencesRf = pattern.TestPlanSequencesRf;
            InterposePostTest = pattern.InterposePostTest;
            SweepVoltage = pattern.SweepVoltage;
            FlowControlFlag = pattern.FlowControlFlag;
            IsFlowInsRepeat = pattern.IsFlowInsRepeat;
            ForceCondition = pattern.ForceCondition;
            MixSigUsed = pattern.MixSigUsed;
            WirelessData = new WirelessData(pattern.WirelessData);
            TestName = pattern.TestName;
            VbtTypes = pattern.VbtTypes;
            InstrumentSetup = pattern.InstrumentSetup;
            BlockType = pattern.BlockType;
            RfInterPose = pattern.RfInterPose;
            OriMeasPins = pattern.OriMeasPins;
            Chiplet = pattern.Chiplet;
            SkipList = pattern.SkipList;
            ForceVoltageFlag = pattern.ForceVoltageFlag;
            HipPreWriteFlag = pattern.HipPreWriteFlag;
            IsVtShmooUsed = pattern.IsVtShmooUsed;
            SkipDotNet = pattern.SkipDotNet;
            Failflag = pattern.Failflag;
            SiteFlag = pattern.SiteFlag;
            foreach (HardIpPattern pat in pattern.BurstPatterns)
            {
                BurstPatterns.Add(pat);
            }
        }

        public int CapBitsInTp()
        {
            int capBits = 0;
            foreach (MeasPin measPin in MeasPins.Where(p => p.MeasType.Equals(MeasType.MeasC)))
            {
                if (!string.IsNullOrEmpty(measPin.CapBit))
                {
                    capBits += int.Parse(measPin.CapBit);
                }
                else
                {
                    ErrorReportManager.AddError(
                        HardIpErrorType.E_WrongMeasC_02,
                        SheetName,
                        measPin.RowNum,
                        0,
                        []
                    );
                }
            }
            return capBits;
        }

        public string GetDigCapNameByMiscInfo()
        {
            string digCapNameReg = HardIpConstData.DigCapName + ":(?<name>.*)";
            foreach (string item in MiscInfo.Split(';'))
            {
                if (Regex.IsMatch(item, digCapNameReg, RegexOptions.IgnoreCase))
                {
                    return Regex.Match(item, digCapNameReg, RegexOptions.IgnoreCase).Groups["name"].ToString();
                }
            }
            return "";
        }

        public List<Timing> GetTimingsByAc()
        {
            string acInfo = AcUsed;
            var timings = new List<Timing>();
            foreach (string info in acInfo.Split(';'))
            {
                string[] acArray = info.Split(':');
                if (!info.Contains("AC") && !(acArray.Length == 3 && acArray.Length == 4))
                {
                    continue;
                }

                foreach (string acpin in acArray[1].Split(','))
                {
                    var timing = new Timing { Name = acpin };
                    if (acArray.Length == 3)
                    {
                        timing.SuffixAcSpecName = acArray[2];
                    }
                    else if (acArray.Length == 4)
                    {
                        timing.SuffixAcSpecName = acArray[3];
                    }

                    ProtocolAwarePin pin = NwireSingleton.Instance().SettingInfo.NwirePins.Find(s => s.OutClk.Equals(timing.Name, StringComparison.OrdinalIgnoreCase));
                    if (pin != null)
                    {
                        timing.Name = pin.CreatePinNameWithDiff();
                    }
                    string acValue = "";
                    acValue = acArray[2];

                    if (acValue.Contains(","))
                    {
                        string[] data = acValue.Split(',');
                        string value = GetFreq(data[0].ToString(CultureInfo.InvariantCulture));
                        string catagory = data[1].ToString(CultureInfo.InvariantCulture);
                        if (catagory.Equals("Typ", StringComparison.OrdinalIgnoreCase))
                        {
                            timing.Type = value;
                        }
                        else if (catagory.Equals("Min", StringComparison.OrdinalIgnoreCase))
                        {
                            timing.Min = value;
                        }
                        else if (catagory.Equals("Max", StringComparison.OrdinalIgnoreCase))
                        {
                            timing.Max = value;
                        }
                    }
                    else
                    {
                        string value = GetFreq(acValue);
                        timing.Type = value;
                        timing.Min = value;
                        timing.Max = value;
                    }

                    timings.Add(timing);
                }


            }
            return timings;
        }

        public static string GetFreq(string srcFreq)
        {
            string result = srcFreq;
            const string lStrValuePattern = @"^(?<str>\d*[.]?\d*)[a-zA-Z]+$";
            const string lStrUnitPattern = @"^\d*[.]?\d*(?<str>[a-zA-Z]+)$";
            if (Regex.IsMatch(result, lStrUnitPattern))
            {
                string lStrUnit = Regex.Match(result, lStrUnitPattern).Groups["str"].ToString().Trim().ToUpper();
                string lStrValue = Regex.Match(result, lStrValuePattern).Groups["str"].ToString().Trim();
                lStrValue.TryCombineHz(lStrUnit, out result);
            }
            return result;
        }

        public string ProcessSweepData(ref bool isAddSweep, bool isBurst = false)
        {
            //reg1=[1000,0010,0001];reg2=[1011,0010,1111]
            var assignList = new List<string>();
            int cnt = 1;
            var sweepSrc = new List<string>();
            bool isAddSweepBurst = false;

            if (Pattern.IsMultiple())
            {
                foreach (HardIpPattern burPat in BurstPatterns)
                {
                    assignList.Add(Regex.IsMatch(burPat.RegisterAssignment, "Reg_assign", RegexOptions.IgnoreCase)
                        ? burPat.RegisterAssignment
                        : burPat.ProcessSweepData(ref isAddSweepBurst, true));
                }

                RegisterAssignment = string.Join("|", assignList);
                return RegisterAssignment;
            }

            string sweepCodeOrder = "";
            var nestSweepVoltageList = new List<SweepVData>();
            if (SweepVoltage.Any() && SweepVoltage.TryGetValue("NestSweep", out List<SweepVData> value))
            { //check the first voltage order
                nestSweepVoltageList = value;
            }
            string sweepsrckey = "sweepsrc_singleloop:";
            ParseSweepAssignments(assignList, sweepSrc, ref cnt, ref sweepCodeOrder, ref sweepsrckey);
            if (sweepSrc.Count > 0)
            {
                isAddSweep = true;
                string ret = !string.IsNullOrEmpty(sweepCodeOrder) && nestSweepVoltageList.Any()
                    ? BuildMergedSweepString(nestSweepVoltageList, sweepSrc, sweepCodeOrder, sweepsrckey)
                    : sweepsrckey + string.Join(";", sweepSrc.Distinct());
                DspFunction.Add(ret);
                if (isBurst)
                {
                    RegisterAssignment = string.Join(";", assignList);
                }
            }
            else
            {
                if (nestSweepVoltageList.Any())
                {
                    foreach (SweepVData nestSweepVoltage in nestSweepVoltageList)
                    {
                        sweepSrc.Add(BuildSweepVoltageString(nestSweepVoltage));
                    }
                    DspFunction.Add(string.Join(";", sweepSrc.Distinct()));
                }
            }

            return assignList.Any() ? string.Join(";", assignList) : RegisterAssignment;
        }

        private void ParseSweepAssignments(List<string> assignList, List<string> sweepSrc, ref int cnt, ref string sweepCodeOrder, ref string sweepsrckey)
        {
            foreach (string assignment in RegisterAssignment.Split(';').ToList())
            {
                if (!Regex.IsMatch(assignment, @".*\=.*"))
                {
                    continue;
                }

                if (Regex.IsMatch(assignment, "nestsweep", RegexOptions.IgnoreCase))
                {
                    AddNestSweepAssignment(assignment, assignList, sweepSrc, ref cnt, ref sweepCodeOrder, ref sweepsrckey);
                }
                else if (Regex.IsMatch(assignment, @".*\=.*\[.*\]"))
                {
                    AddSweepRegisterAssignment(assignment, assignList, sweepSrc, ref cnt, ref sweepCodeOrder, ref sweepsrckey);
                }
                else if (Regex.IsMatch(assignment, @".*\=Sweep.*_Singleloop\:\:.*"))
                {
                    AddSweepSingleloopAssignment(assignment, assignList, sweepSrc, ref cnt);
                }
                else
                {
                    assignList.Add(assignment);
                }
            }
        }

        private void AddNestSweepAssignment(string assignment, List<string> assignList, List<string> sweepSrc, ref int cnt, ref string sweepCodeOrder, ref string sweepsrckey)
        {
            string args = Regex.Match(assignment, @"\((?<value>.*)\)", RegexOptions.IgnoreCase).Groups["value"].ToString().Replace(",", "@");
            string order = Regex.Match(assignment, @"nestsweep(?<order>\w*)\((?<format>.*)\)", RegexOptions.IgnoreCase).Groups["order"].ToString();
            sweepCodeOrder = order;
            sweepsrckey = "sweepsrc:";
            if (args.Contains(":"))
            {
                return;
            }

            sweepSrc.Add(string.IsNullOrEmpty(order)
                ? "sweep" + cnt + "=" + args
                : "sweep" + order + "=" + args);
            assignList.Add(
                $"{assignment.Split('=')[0]}={(string.IsNullOrEmpty(order) ? "sweep" + cnt : "sweep" + order)}");

            cnt++;
        }

        private void AddSweepRegisterAssignment(string assignment, List<string> assignList, List<string> sweepSrc, ref int cnt, ref string sweepCodeOrder, ref string sweepsrckey)
        {
            string reg = Regex.Match(assignment, @"\[(?<value>.*)\]", RegexOptions.IgnoreCase).Groups["value"].ToString();
            string order = Regex.Match(assignment, @"sweep(?<order>\w*)_singleloop\[(?<format>.*)\]", RegexOptions.IgnoreCase).Groups["order"].ToString();
            string alg = "";
            if (string.IsNullOrEmpty(order))
            {
                order = Regex.Match(assignment, @"sweep(?<order>\w*)\[(?<format>.*)\]", RegexOptions.IgnoreCase).Groups["order"].ToString();
                sweepsrckey = "sweepsrc:";
                alg = Regex.Match(assignment, "@(?<alg>.*)", RegexOptions.IgnoreCase).Groups["alg"].ToString();
            }
            sweepCodeOrder = order;
            if (reg.Contains(":"))
            {
                return;
            }

            sweepSrc.Add(string.IsNullOrEmpty(order)
                ? "sweep" + cnt + "=" + reg + (string.IsNullOrEmpty(alg) ? "" : "@" + alg)
                : "sweep" + order + "=" + reg + (string.IsNullOrEmpty(alg) ? "" : "@" + alg));
            assignList.Add(
                $"{assignment.Split('=')[0]}={(string.IsNullOrEmpty(order) ? "sweep" + cnt : "sweep" + order)}");

            cnt++;
        }

        private void AddSweepSingleloopAssignment(string assignment, List<string> assignList, List<string> sweepSrc, ref int cnt)
        {
            Match regassignment = Regex.Match(assignment, @"(?<regname>\w*)\=Sweep(?<order>.*)_Singleloop\:\:(?<table>.*)", RegexOptions.IgnoreCase);

            string assiOrder = regassignment.Groups["order"].ToString();
            string assiTable = regassignment.Groups["table"].ToString();

            if (!EpWorkbook.TestPlanWorkbook.Worksheets.Any(ws => ws.Name.Equals(assiTable, StringComparison.OrdinalIgnoreCase)))
            {
                ErrorReportManager.AddError(
                    HardIpErrorType.E_MissingSweepSingleloopTable_01,
                    SheetName,
                    RowNum,
                    0,
                    [assiTable]
                );
            }

            sweepSrc.Add(assiTable);
            assignList.Add(
                $"{assignment.Split('=')[0]}={(string.IsNullOrEmpty(assiOrder) ? "sweep" + cnt : "sweep" + assiOrder)}");

            cnt++;
        }

        private string BuildMergedSweepString(List<SweepVData> nestSweepVoltageList, List<string> sweepSrc, string sweepCodeOrder, string sweepsrckey)
        {
            string sweepCodeStr = sweepsrckey + string.Join(";", sweepSrc.Distinct());
            var sweepVoltageStrList1 = new List<string>();
            var sweepVoltageStrList2 = new List<string>();
            foreach (SweepVData nestSweepVoltage in nestSweepVoltageList)
            {
                if (string.IsNullOrEmpty(nestSweepVoltage.Order))
                {
                    nestSweepVoltage.Order = "1";
                }

                string sweepVoltageStr = BuildSweepVoltageString(nestSweepVoltage);

                if (int.Parse(sweepCodeOrder) < int.Parse(nestSweepVoltage.Order))
                {
                    //sweep code first, add sweep voltage
                    sweepVoltageStrList2.Add(sweepVoltageStr);
                }
                else
                {
                    sweepVoltageStrList1.Add(sweepVoltageStr);
                }
            }
            var mergedSweepStrList = new List<string>();
            mergedSweepStrList.AddRange(sweepVoltageStrList1);
            mergedSweepStrList.Add(sweepCodeStr);
            mergedSweepStrList.AddRange(sweepVoltageStrList2);
            return string.Join(";", mergedSweepStrList.Distinct());
        }

        private string BuildSweepVoltageString(SweepVData nestSweepVoltage)
        {
            string sweepVoltageStr = !string.IsNullOrEmpty(nestSweepVoltage.VoltageList) ?
                    $"sweepvoltage{nestSweepVoltage.Order}={nestSweepVoltage.VoltageList}" :
                    $"sweepvoltage{nestSweepVoltage.Order}={nestSweepVoltage.Start}@{nestSweepVoltage.Stop}@{nestSweepVoltage.Step}";

            if (!string.IsNullOrEmpty(nestSweepVoltage.InstanceVoltage))
            {
                sweepVoltageStr = $"{nestSweepVoltage.InstanceVoltage}@" + sweepVoltageStr;
            }

            return sweepVoltageStr;
        }

        public bool IsIgnorePatBinOut()
        {
            return MiscInfoDict.ContainsKey(HardIpConstData.IgnorePatBinOut) ||
                (Pattern.IsDefaultNobinout && !MiscInfoDict.ContainsKey(HardIpConstData.EnablePattBinout));
        }
    }

    public enum PlanType
    {
        Default, Trim, Rf
    }
}
