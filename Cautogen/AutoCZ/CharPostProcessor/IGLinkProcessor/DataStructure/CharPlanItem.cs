using System.Collections.Generic;
using System.Linq;

using Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor.DataStructure.ShmooData;

namespace Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor.DataStructure
{
    public class CharPlanItem
    {
        /* Property */
        private string _voltage = "";
        private string _initPatternsCell = string.Empty;
        private string _payloadPatternsCell = string.Empty;

        public int RowNum { get; set; }
        public bool Visit { get; set; }
        public bool Select { get; set; }
        public string Item { get; set; }
        public string BlockName { get; set; }
        public string Description { get; set; }
        public string Type { get; set; }
        public string MeasType { get; set; }
        public string TestInstanceName { get; set; }
        public string InstanceName { get; set; }
        public string NamingSelection { get; set; }
        public string Voltage
        {
            get
            {
                if (IsOverWriteVoltage && _voltage.Split(' ').Length > 1)
                {
                    return _voltage.Split(' ')[1];
                }
                return _voltage;
            }
            set
            {
                _voltage = value;
                if (_voltage.Split(' ').Length > 1)
                {
                    IsOverWriteVoltage = true;
                }
            }
        }

        public bool IsOverWriteVoltage { get; set; }

        public string PowerRunScenario { get; set; }
        public string Wait { get; set; }
        public ShmooSetup CharShmooSetup { get; set; }
        public string DcCategory { get; set; }
        public string DcSelector { get; set; }
        public string AcCategory { get; set; }
        public string AcSelector { get; set; }
        public bool IsUseRtosCmd { get; set; }
        public string Levels { get; set; }
        public string Timeset { get; set; }
        public string InitPatternsCell
        {
            get
            {
                return _initPatternsCell;
            }
            set
            {
                _initPatternsCell = value;
                var patternList = new List<string>();
                value.Split(';').ToList().ForEach(patternList.Add);
                var patternDict = new Dictionary<string, string>();
                patternList.Where(x => x.Split('=').Length > 1).ToList().ForEach(x => patternDict[x.Split('=')[0]] = x.Split('=')[1]);
                InitPatternsDict = patternDict;
            }
        }
        public string PayloadPatternsCell
        {
            get
            {
                return _payloadPatternsCell;
            }
            set
            {
                _payloadPatternsCell = value;
                var patternList = new List<string>();
                value.Split(';').ToList().ForEach(patternList.Add);
                var patternDict = new Dictionary<string, string>();
                patternList.Where(x => x.Split('=').Length > 1).ToList().ForEach(x => patternDict[x.Split('=')[0]] = x.Split('=')[1]);
                PayloadPatternsDict = patternDict;
            }
        }
        public Dictionary<string, string> InitPatternsDict { get; set; } = new Dictionary<string, string>();
        public Dictionary<string, string> PayloadPatternsDict { get; set; } = new Dictionary<string, string>();
        public Dictionary<string, string> AllPatternsDict
        {
            get
            {
                var result = new Dictionary<string, string>();
                foreach (KeyValuePair<string, string> item in InitPatternsDict)
                {
                    result.Add(item.Key, item.Value);
                }

                foreach (KeyValuePair<string, string> item in PayloadPatternsDict)
                {
                    result.Add(item.Key, item.Value);
                }
                return result;
            }
        }
        public string Payload1 { get; set; } = string.Empty;
        public string Payload2 { get; set; } = string.Empty;
        public string Retention { get; set; }
        public string IpUse1 { get; set; }
        public string IpUse2 { get; set; }
        public string IpUse3 { get; set; }
        public string IpUse4 { get; set; }
        public string IpUse5 { get; set; }
        public string IpUse6 { get; set; }
        public string SearchMethod { get; set; }
        public string CharCondition { get; set; }
        public string SelSrmSendbit { get; set; }
        public string ProgramTestName1 { get; set; }
        public string ProgramTestName2 { get; set; }
        public string DigSrc { get; set; }
        public string DigSrcAssignment { get; set; }
        public bool Htol { get; set; }
        public bool Use { get; set; }
        public bool Ttr { get; set; }
        public bool InProgFlow { get; set; }
        public bool InProgInstance { get; set; }
        public string IsNeedMask { get; set; }
        public string SheetName { get; set; }
        public string SuspendDatalog { get; set; }
        public string HarvFstp { get; set; }
        public string SiteFlag { get; set; }
        public string FailFlag { get; set; }
        public string ManualAc { get; set; }
        public bool IsFreeRunClk = false;
        public string FailInfo { get; set; }
        public string Environment { get; set; }
        public string Burst { get; set; }
        public bool IsDateNeedReverse { get; set; }
        public string EnableWord { get; set; }
        public string ManualACfromTimeset { get; set; }
        public string ShiftFreq { get; set; }
        public string AcCategoryOri { get; set; }
        public string BypassShmooHole { get; set; }
        public string DigSrcBitSize { get; set; }
        public string DigSrcSeg { get; set; }
        public string DigSrcPin { get; set; }
        public string DigSrcEq { get; set; }

        public Dictionary<string, string> ArgPatternIndexConvertedRetDict
        {
            get
            {
                var dict = new Dictionary<string, string>();
                int initIndex = 0;
                foreach (KeyValuePair<string, string> init in InitPatternsDict)
                {
                    initIndex += init.Value.Split(',').Length;
                    dict[init.Key] = $"INIT{initIndex}";
                }
                int payloadIndex = 0;
                foreach (KeyValuePair<string, string> payload in PayloadPatternsDict)
                {
                    payloadIndex += payload.Value.Split(',').Length;
                    dict[payload.Key] = $"PL{payloadIndex}";
                }
                return dict;
            }
        }
        public Dictionary<string, string> ArgPatternIndexConvertedDsscDict
        {
            get
            {
                var dict = new Dictionary<string, string>();
                int initIndex = 1;
                foreach (KeyValuePair<string, string> init in InitPatternsDict)
                {
                    dict[init.Key] = $"INIT{initIndex}";
                    initIndex += init.Value.Split(',').Length;
                }
                int payloadIndex = 1;
                foreach (KeyValuePair<string, string> payload in PayloadPatternsDict)
                {
                    dict[payload.Key] = $"PL{payloadIndex}";
                    payloadIndex += payload.Value.Split(',').Length;
                }
                return dict;
            }
        }
        public Dictionary<string, List<string>> ArgPatternIndexConvertedPrsDict
        {
            get
            {
                var dict = new Dictionary<string, List<string>>();
                int initIndex = 1;
                foreach (KeyValuePair<string, string> init in InitPatternsDict)
                {
                    dict[init.Key] = new List<string>();
                    foreach (string pattern in init.Value.Split(','))
                    {
                        dict[init.Key].Add($"INIT{initIndex}");
                        initIndex++;
                    }
                }
                int payloadIndex = 1;
                foreach (KeyValuePair<string, string> payload in PayloadPatternsDict)
                {
                    dict[payload.Key] = new List<string>();
                    foreach (string pattern in payload.Value.Split(','))
                    {
                        dict[payload.Key].Add($"PL{payloadIndex}");
                        payloadIndex++;
                    }
                }
                return dict;
            }
        }
        public List<string> UsedPatterns
        {
            get
            {
                var patterns = new List<string>();
                patterns.AddRange(InitPatternsDict.Values.SelectMany(x => x.Split(',')).Select(x => x.Split(':')[0]));
                patterns.AddRange(PayloadPatternsDict.Values.SelectMany(x => x.Split(',')).Select(x => x.Split(':')[0]));
                return patterns;
            }
        }

        public List<string> UsedInits
        {
            get
            {
                var patterns = new List<string>();
                patterns.AddRange(InitPatternsDict.Values.SelectMany(x => x.Split(',')).Select(x => x.Split(':')[0]));
                return patterns;
            }
        }

        public List<string> UsedPayloads
        {
            get
            {
                var patterns = new List<string>();
                patterns.AddRange(PayloadPatternsDict.Values.SelectMany(x => x.Split(',')).Select(x => x.Split(':')[0]));
                return patterns;
            }
        }

        public string GetPatternByKey(string key)
        {
            string result = "";
            if (AllPatternsDict.ContainsKey(key.ToUpper()))
            {
                result = AllPatternsDict[key.ToUpper()];
            }
            return result;
        }

        public string GetPatternsStartByKey(string prefix, int index)
        {
            var result = new List<string>();
            while (AllPatternsDict.ContainsKey(prefix + index))
            {
                result.Add(AllPatternsDict[prefix + index]);
                index++;
            }

            return string.Join(",", result);
        }

        public string RetentionRamp { get; set; }
        public string OneTimeInit { get; set; }
        public string UserFunction { get; set; }
        public string HarvestPinGrpOtherFail { get; set; }
        public string EnableCoreHarvest { get; set; }
        public string EnableCoreMask { get; set; }
        public string PinGrpSpecifyMask { get; set; }
        public string SsnSpecifyMask { get; set; }
        public string MappingPatternSet { get; set; }
        public string ApplyVoltageFromBinCut { get; set; }
        public string FreeRunningClock { get; set; }
        public string SelsramPatternIdx { get; set; }
        public int RunPayloadAfterSelsramShiftCount { get; set; } = 0;
        public bool ExtendInit { get; set; } = false;

        public string AdaptiveCooling { get; set; }
        public string SelsramUserDef9 { get; set; }
        public string StageCp1 { get; set; }
        public string StageCp2 { get; set; }
        public string StageFt1 { get; set; }
        public string StageFt2 { get; set; }
        public string Die { get; set; }

        public CharPlanItem(CharPlanItem charRow)
        {
            Copy(charRow);
        }

        public CharPlanItem()
        {
        }

        public void Copy(CharPlanItem item)
        {
            _initPatternsCell = item._initPatternsCell;
            _payloadPatternsCell = item._payloadPatternsCell;
            InitPatternsDict = item.InitPatternsDict;
            PayloadPatternsDict = item.PayloadPatternsDict;
            RowNum = item.RowNum;
            Visit = item.Visit;
            Select = item.Select;
            Item = item.Item;
            BlockName = item.BlockName;
            Description = item.Description;
            Type = item.Type;
            MeasType = item.MeasType;
            TestInstanceName = item.TestInstanceName;
            InstanceName = item.InstanceName;
            NamingSelection = item.NamingSelection;
            Voltage = item.Voltage;
            PowerRunScenario = item.PowerRunScenario;
            Wait = item.Wait;
            CharShmooSetup = item.CharShmooSetup;
            DcCategory = item.DcCategory;
            DcSelector = item.DcSelector;
            AcCategory = item.AcCategory;
            AcSelector = item.AcSelector;
            IsUseRtosCmd = item.IsUseRtosCmd;
            Levels = item.Levels;
            Timeset = item.Timeset;
            Retention = item.Retention;
            IpUse1 = item.IpUse1;
            IpUse2 = item.IpUse2;
            IpUse3 = item.IpUse3;
            IpUse4 = item.IpUse4;
            IpUse5 = item.IpUse5;
            IpUse6 = item.IpUse6;
            SearchMethod = item.SearchMethod;
            CharCondition = item.CharCondition;
            SelSrmSendbit = item.SelSrmSendbit;
            ProgramTestName1 = item.ProgramTestName1;
            ProgramTestName2 = item.ProgramTestName2;
            DigSrc = item.DigSrc;
            DigSrcAssignment = item.DigSrcAssignment;
            Htol = item.Htol;
            Use = item.Use;
            Ttr = item.Ttr;
            InProgFlow = item.InProgFlow;
            InProgInstance = item.InProgInstance;
            IsNeedMask = item.IsNeedMask;
            SheetName = item.SheetName;
            SuspendDatalog = item.SuspendDatalog;
            HarvFstp = item.HarvFstp;
            SiteFlag = item.SiteFlag;
            ManualAc = item.ManualAc;
            FailInfo = item.FailInfo;
            Environment = item.Environment;
            Burst = item.Burst;
            IsOverWriteVoltage = item.IsOverWriteVoltage;
            IsDateNeedReverse = item.IsDateNeedReverse;
            EnableWord = item.EnableWord;
            FailFlag = item.FailFlag;
            ManualACfromTimeset = item.ManualACfromTimeset;
            ShiftFreq = item.ShiftFreq;
            AcCategoryOri = item.AcCategoryOri;
            BypassShmooHole = item.BypassShmooHole;
            DigSrcBitSize = item.DigSrcBitSize;
            DigSrcSeg = item.DigSrcSeg;
            DigSrcPin = item.DigSrcPin;
            DigSrcEq = item.DigSrcEq;
            OneTimeInit = item.OneTimeInit;
            MappingPatternSet = item.MappingPatternSet;
            ApplyVoltageFromBinCut = item.ApplyVoltageFromBinCut;
            FreeRunningClock = item.FreeRunningClock;
            SelsramPatternIdx = item.SelsramPatternIdx;
            RunPayloadAfterSelsramShiftCount = item.RunPayloadAfterSelsramShiftCount;
            UserFunction = item.UserFunction;
            HarvestPinGrpOtherFail = item.HarvestPinGrpOtherFail;
            EnableCoreHarvest = item.EnableCoreHarvest;
            EnableCoreMask = item.EnableCoreMask;
            PinGrpSpecifyMask = item.PinGrpSpecifyMask;
            SsnSpecifyMask = item.SsnSpecifyMask;
            AdaptiveCooling = item.AdaptiveCooling;
            StageCp1 = item.StageCp1;
            StageCp2 = item.StageCp2;
            StageFt1 = item.StageFt1;
            StageFt2 = item.StageFt2;
            Die = item.Die;
        }
    }
}
