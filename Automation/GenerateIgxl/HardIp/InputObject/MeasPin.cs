using System;
using System.Collections.Generic;
using System.Linq;

using Automation.Static;

namespace Automation.GenerateIgxl.HardIp.InputObject
{
    public class MeasPin : IDisposable
    {
        private string _miscInfo;
        public int RowNum { get; set; }
        public string PinName { get; set; }
        public string CusStr { get; set; }
        public string CapBit { get; set; }
        public string MeasType { get; set; }
        public int RowNumForMergeMeas { get; set; }
        public List<ForceCondition> ForceConditions { get; set; }
        public string Job { get; set; }
        public int PinCount { get; set; }
        public int RepeatCount { get; set; }
        public List<MeasLimit> MeasLimitsH { get; set; }
        public List<MeasLimit> MeasLimitsL { get; set; }
        public List<MeasLimit> MeasLimitsN { get; set; }
        public string LowLimit { get; set; }
        public string HighLimit { get; set; }
        public List<CurrentRange> CurrentRangeList { get; set; }
        public List<CurrentRange> CurrentRangeListH { get; set; }
        public List<CurrentRange> CurrentRangeListL { get; set; }
        public List<CurrentRange> CurrentRangeListN { get; set; }
        public int SequenceIndex { get; set; }
        public int SubSequenceIndex { get; set; }
        public int VisitedTime { get; set; }// For sort use-limit
        public string OriPinName { get; set; }

        //start: For forceCondition Merged 
        public string PinType { get; set; }
        //end: For forceCondition Merged 
        public string TestName { get; set; }
        //For Calc, Limit row
        public string CalcEqn { get; set; }
        //For Calc, Limit row
        public bool IsUsedPin { get; set; }
        //For FW, Interpose
        public string InterPoseFunc;
        public string RfInterPose;
        public string RfInstrumentSetup;
        public string MeasWaitTime;
        public string MeasRange;
        public string SkipUnit;

        public string PatternName;
        public int PatternIndex;
        public string MiscInfo
        {
            get { return _miscInfo; }
            set
            {
                MiscInfoDict = TestPlanStatic.MiscInfoParser.ParseKeyValueToDictionary(value, out string _);
                _miscInfo = value;
            }
        }
        public Dictionary<string, string> MiscInfoDict { get; private set; }

        public MeasPin()
        {
            PinCount = 0;
            PinName = "";
            CusStr = "";
            MeasType = "";
            Job = "";
            ForceConditions = new List<ForceCondition>();
            LowLimit = "";
            HighLimit = "";
            CurrentRangeList = new List<CurrentRange>();
            CurrentRangeListH = new List<CurrentRange>();
            CurrentRangeListL = new List<CurrentRange>();
            CurrentRangeListN = new List<CurrentRange>();
            MeasLimitsH = new List<MeasLimit>();
            MeasLimitsL = new List<MeasLimit>();
            MeasLimitsN = new List<MeasLimit>();
            VisitedTime = 1;
            SequenceIndex = 0;
            PinType = "";
            TestName = "";
            CalcEqn = "";
            IsUsedPin = false;
            MeasWaitTime = "";
            MeasRange = "";
            RfInstrumentSetup = "";
            RepeatCount = 0;
            SkipUnit = "";
            MiscInfo = "";
            InterPoseFunc = "";
            OriPinName = "";
            PatternName = "";
            PatternIndex = -1;
        }

        public MeasPin(string pinName, string measType)
        {
            PinCount = 0;
            PinName = pinName;
            CusStr = "";
            MeasType = measType;
            Job = "";
            ForceConditions = new List<ForceCondition>();
            LowLimit = "";
            HighLimit = "";
            CurrentRangeList = new List<CurrentRange>();
            CurrentRangeListH = new List<CurrentRange>();
            CurrentRangeListL = new List<CurrentRange>();
            CurrentRangeListN = new List<CurrentRange>();
            MeasLimitsH = new List<MeasLimit>();
            MeasLimitsL = new List<MeasLimit>();
            MeasLimitsN = new List<MeasLimit>();
            VisitedTime = 1;
            SequenceIndex = 0;
            PinType = "";
            TestName = "";
            CalcEqn = "";
            MeasWaitTime = "";
            MeasRange = "";
            RfInstrumentSetup = "";
            RepeatCount = 0;
            RfInterPose = "";
            MiscInfo = "";
            InterPoseFunc = "";
            OriPinName = "";
            PatternName = "";
            PatternIndex = -1;
        }

        public void Copy(MeasPin pin)
        {
            //PinName = pin.PinName;
            PinCount = pin.PinCount;
            MeasType = pin.MeasType;
            RowNumForMergeMeas = pin.RowNumForMergeMeas;
            CusStr = pin.CusStr;
            Job = pin.Job;
            RowNum = pin.RowNum;
            ForceConditions = pin.ForceConditions;
            LowLimit = pin.LowLimit;
            HighLimit = pin.HighLimit;
            CurrentRangeList = pin.CurrentRangeList;
            CurrentRangeListH = pin.CurrentRangeListH;
            CurrentRangeListL = pin.CurrentRangeListL;
            CurrentRangeListN = pin.CurrentRangeListN;
            MeasLimitsH = pin.MeasLimitsH;
            MeasLimitsL = pin.MeasLimitsL;
            MeasLimitsN = pin.MeasLimitsN;
            VisitedTime = pin.VisitedTime;
            PinType = pin.PinType;
            TestName = pin.TestName;
            CalcEqn = pin.CalcEqn;
            InterPoseFunc = pin.InterPoseFunc;
            RfInterPose = pin.RfInterPose;
            MeasWaitTime = pin.MeasWaitTime;
            MeasRange = "";
            RfInstrumentSetup = pin.RfInstrumentSetup;
            RepeatCount = pin.RepeatCount;
            _miscInfo = pin.MiscInfo;
            MiscInfoDict = pin.MiscInfoDict;
            OriPinName = pin.OriPinName;
            PatternName = pin.PatternName;
            PatternIndex = pin.PatternIndex;
        }

        public MeasPin Copy()
        {
            var clone = new MeasPin
            {
                RowNum = RowNum,
                PinName = PinName,
                CusStr = CusStr,
                CapBit = CapBit,
                MeasType = MeasType,
                RowNumForMergeMeas = RowNumForMergeMeas,
                Job = Job,
                PinCount = PinCount,
                RepeatCount = RepeatCount,
                LowLimit = LowLimit,
                HighLimit = HighLimit,
                SequenceIndex = SequenceIndex,
                VisitedTime = VisitedTime,
                OriPinName = OriPinName,
                PinType = PinType,
                TestName = TestName,
                CalcEqn = CalcEqn,
                IsUsedPin = IsUsedPin,

                InterPoseFunc = InterPoseFunc,
                RfInterPose = RfInterPose,
                RfInstrumentSetup = RfInstrumentSetup,
                MeasWaitTime = MeasWaitTime,
                MeasRange = MeasRange,
                SkipUnit = SkipUnit,
                PatternName = PatternName,
                PatternIndex = PatternIndex,
                _miscInfo = MiscInfo,
                MiscInfoDict = MiscInfoDict,

                ForceConditions = ForceConditions != null
                    ? ForceConditions.Select(fc => new ForceCondition
                    {
                        ForcePins = fc?.ForcePins != null
                            ? fc.ForcePins.Select(fp => fp?.Copy()).ToList()
                            : new List<ForcePin>()
                    }).ToList()
                    : new List<ForceCondition>(),

                MeasLimitsH = CloneMeasLimitList(MeasLimitsH),
                MeasLimitsL = CloneMeasLimitList(MeasLimitsL),
                MeasLimitsN = CloneMeasLimitList(MeasLimitsN),

                // CurrentRange lists
                CurrentRangeList = CloneCurrentRangeList(CurrentRangeList),
                CurrentRangeListH = CloneCurrentRangeList(CurrentRangeListH),
                CurrentRangeListL = CloneCurrentRangeList(CurrentRangeListL),
                CurrentRangeListN = CloneCurrentRangeList(CurrentRangeListN)
            };

            return clone;
        }

        private static List<MeasLimit> CloneMeasLimitList(List<MeasLimit> src)
        {
            var dst = new List<MeasLimit>();
            if (src == null)
            {
                return dst;
            }

            foreach (MeasLimit m in src)
            {
                if (m == null)
                { dst.Add(null); continue; }

                var nm = new MeasLimit(m.JobName)
                {
                    LoLimit = m.LoLimit,
                    HiLimit = m.HiLimit,
                    LoHeaderIndex = m.LoHeaderIndex,
                    HiHeaderIndex = m.HiHeaderIndex
                };
                dst.Add(nm);
            }
            return dst;
        }


        private static List<CurrentRange> CloneCurrentRangeList(List<CurrentRange> src)
        {
            var dst = new List<CurrentRange>();
            if (src == null)
            {
                return dst;
            }

            foreach (CurrentRange cr in src)
            {
                if (cr == null)
                { dst.Add(null); continue; }
                var ncr = new CurrentRange
                {
                    JobName = cr.JobName,
                    Value = cr.Value
                };
                dst.Add(ncr);
            }
            return dst;
        }

        public void Dispose()
        {
            GC.Collect();
        }

    }
}
