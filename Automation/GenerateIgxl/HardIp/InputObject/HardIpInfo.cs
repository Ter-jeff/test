using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Automation.GenerateIgxl.HardIp.InputObject
{
    public class HardIpInfo
    {
        private readonly List<string> _trimfuseNameStr = new List<string>();
        private readonly List<string> _miscinfo = new List<string>();
        public string VmVector = "";
        public string Version = "";
        public string SsnCoreName = "";
        public string Payload { get; set; } = "";
        public string TimeDomain { get; set; } = "";
        public string TimeSet { get; set; } = "";
        public string Subroutine { get; set; } = "";

        public List<string> SubrList
        {
            get
            {
                if (!string.IsNullOrEmpty(Subroutine))
                {
                    return Subroutine.Split(',').ToList();
                }

                return null;
            }
        }
        //Call_Subrs
        public string CallSubrs { get; set; } = "";
        public string FullPattern
        {
            get { return Payload + "_" + Version; }
        }
        public int CallSubrsCnt { get; set; }
        public string CallSubRoutinesCntDiff { get; set; }
        public string MeasIStr { get; set; } = "";
        public string MeasVStr { get; set; } = "";
        public string DisConnectPins { get; set; } = "";
        public string MeasFStr { get; set; } = "";
        public string MeasVdiffStr { get; set; } = "";
        public string MeasVdiff2Str { get; set; } = "";
        public string MeasIdiffStr { get; set; } = "";
        public string MeasFdiffStr { get; set; } = "";
        public string MeasVocmStr { get; set; } = "";
        public string MeasR1Str { get; set; } = "";
        public string MeasR2Str { get; set; } = "";
        public string MeasDutyCycleStr { get; set; } = "";
        public string MeasVdmStr { get; set; } = "";
        public string MeasSeqStr { get; set; } = "";
        public string DigSrcAssign { get; set; } = "";
        public string DigSrcDataWidth { get; set; } = "";
        public string TrimTarget { get; set; } = "";

        public string TrimFuseName
        {
            get
            {
                return string.Join(",", _trimfuseNameStr);
            }
            set
            {
                _trimfuseNameStr.Add(value);
            }
        }

        public List<string> TrimRegName = new List<string>();
        public string TestName { get; set; } = "";
        public string TrimType { get; set; } = "";
        public string CalcStoreName { get; set; } = "";
        public string PsetUseStr { set; get; }
        public string BestCodeCalcFunc { get; set; }
        public string EfuseRepeatCount { get; set; }
        public bool UseThisVersion { get; set; }

        #region Digital Capture
        public int CapBit { get; set; }

        public string CapBitStr { get; set; } = "";

        public string CapBitName { get; set; } = "";

        public string CapPinName { get; set; } = "";

        public string DsscOut { get; set; } = "";

        public string MultiDsscOut { get; set; } = "";

        #endregion

        #region Digital Source
        public string SendBit { get; set; }

        public string SendBitStr { get; set; } = "";

        public string SendBitName { get; set; } = "";

        public string SendPinName { get; set; } = "";

        public string DigSrcSignalName { get; set; } = "";

        public string Xpins { get; set; } = "";

        public string UnusedIoPins { get; set; } = "";

        #endregion

        public List<string> MeasIPowerPinList = new List<string>();
        public List<string> MeasVPowerPinList = new List<string>();
        public List<string> MeasIioPinList = new List<string>();
        public List<string> MeasVioPinList = new List<string>();
        public List<string> MeasFPinList = new List<string>();
        public List<string> MeasVdiffPinList = new List<string>();
        public List<string> MeasIdiffPinList = new List<string>();
        public List<string> MeasVdiff2PinList = new List<string>();
        public List<string> MeasFdiffPinList = new List<string>();
        public List<string> MeasVocmPinList = new List<string>();
        public List<string> MeasR1PinList = new List<string>();
        public List<string> MeasR2PinList = new List<string>();
        public List<HardIpSeqInfo> SeqInfo = new List<HardIpSeqInfo>();
        public List<string> SpecialLimits = new List<string>();
        public List<string> ForcePrePatRaw = new List<string>();
        public List<string> DisconnectPrePat = new List<string>();
        public List<string> MiscInfo
        {
            get { return _miscinfo; }
            set
            {
                _miscinfo.AddRange(value.SelectMany(p => p.Split(';')));
            }
        }
        public string DigSrcAssignment = "";
        public ForceCondition ForceCondition
        {
            get { return ExtractForcePrePat(); }
        }

        public bool PatInfoExist { get; set; }
        public bool IllegalChar = false;
        public bool IsIgnoreComment { get; set; }
        public List<string> ForceVPinList = new List<string>();
        public List<string> ForceIPinList = new List<string>();
        public string MeasName = "";
        public string CapStoreName = "";
        public PowerOnFly PowerOnFlyItem = null;
        public string IsSendLsbFirstInfo = "";
        public string IsCapLsbFirstInfo = "";
        public Dictionary<string, List<Dictionary<string, List<string>>>> MeaSet = new Dictionary<string, List<Dictionary<string, List<string>>>>();
        public Dictionary<string, List<string>> OriMeaSet = new Dictionary<string, List<string>>();
        public HardIpInfoNew NewInfo;

        public void MergeInfo(HardIpInfo otherInfo)
        {
            CapBit += otherInfo.CapBit;
            CapBitStr = string.Join("|", new List<string> { CapBitStr, otherInfo.CapBitStr }).Trim('|');
            CapBitName = string.Join("|", new List<string> { CapBitName, otherInfo.CapBitName }).Trim('|');
            CapPinName = string.Join("|", new List<string> { CapPinName, otherInfo.CapPinName }).Trim('|');
            MultiDsscOut += otherInfo.DsscOut + "|";
            //Merge DsscOut = 
            if (string.IsNullOrEmpty(DsscOut))
            {
                DsscOut = otherInfo.DsscOut;
            }
            else if (!string.IsNullOrEmpty(DsscOut) && !string.IsNullOrEmpty(otherInfo.DsscOut))
            {
                DsscOut += otherInfo.DsscOut.Replace("DSSC_OUT", "");
            }
            SendBit = string.Join("|", new List<string> { SendBit, otherInfo.SendBit }).Trim('|');
            SendBitStr = string.Join("|", new List<string> { SendBitStr, otherInfo.SendBitStr }).Trim('|');
            SendBitName = string.Join("|", new List<string> { SendBitName, otherInfo.SendBitName }).Trim('|');
            SendPinName = string.Join("|", new List<string> { SendPinName, otherInfo.SendPinName }).Trim('|');
            PatInfoExist = otherInfo.PatInfoExist;
            DigSrcSignalName = otherInfo.DigSrcSignalName;
            VmVector = otherInfo.VmVector;
            CallSubrs = otherInfo.CallSubrs;
            CallSubrsCnt += otherInfo.CallSubrsCnt;
            CalcStoreName = otherInfo.CalcStoreName;
            MeasIPowerPinList.AddRange(otherInfo.MeasIPowerPinList);
            MeasVPowerPinList.AddRange(otherInfo.MeasVPowerPinList);
            MeasIioPinList.AddRange(otherInfo.MeasIioPinList);
            MeasVioPinList.AddRange(otherInfo.MeasVioPinList);
            MeasFPinList.AddRange(otherInfo.MeasFPinList);
            MeasVdiffPinList.AddRange(otherInfo.MeasVdiffPinList);
            MeasIdiffPinList.AddRange(otherInfo.MeasIdiffPinList);
            MeasVdiff2PinList.AddRange(otherInfo.MeasVdiff2PinList);
            MeasFdiffPinList.AddRange(otherInfo.MeasFdiffPinList);
            MeasVocmPinList.AddRange(otherInfo.MeasVocmPinList);
            MeasR1PinList.AddRange(otherInfo.MeasR1PinList);
            MeasR2PinList.AddRange(otherInfo.MeasR2PinList);
            TrimRegName.AddRange(otherInfo.TrimRegName);
            if (otherInfo.SeqInfo.Count > 0)
            {
                int offset = SeqInfo.Count;
                foreach (HardIpSeqInfo seq in otherInfo.SeqInfo)
                {
                    HardIpSeqInfo newSeq = new HardIpSeqInfo().Copy(seq);
                    newSeq.MeasPins.ForEach(p => p.SequenceIndex += offset);
                    SeqInfo.Add(newSeq);
                }
            }
            MeasSeqStr = string.Join(",", SeqInfo.Select(p => p.SeqName));
        }

        public string IsMsbFirst()
        {
            if (string.IsNullOrEmpty(IsSendLsbFirstInfo))
            {
                return "";
            }

            foreach (string isLsb in IsSendLsbFirstInfo.Split('+'))
            {
                if (isLsb.Equals("false", StringComparison.OrdinalIgnoreCase))
                {
                    return "TRUE";
                }
            }
            foreach (string isLsb in IsCapLsbFirstInfo.Split('+'))
            {
                if (isLsb.Equals("false", StringComparison.OrdinalIgnoreCase))
                {
                    return "TRUE";
                }
            }
            return "";
        }

        public bool IsRegisterReverse(string register)
        {
            bool result = false;
            List<string> regList = SendBitName.Split('+').ToList();
            List<string> sendRevList = IsSendLsbFirstInfo.Split('+').ToList();
            if (regList.Contains(register, StringComparer.OrdinalIgnoreCase))
            {
                int index = regList.FindIndex(p => p.Equals(register, StringComparison.OrdinalIgnoreCase));
                if (index < sendRevList.Count)
                {
                    if (sendRevList[index].Equals("false", StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
            }
            return result;
        }

        public ForceCondition ExtractForcePrePat()
        {
            string regAc = @"(?<AC>AC:[\w,&]+:\w+)";
            var forcecondition = new ForceCondition();
            foreach (string forceStr in ForcePrePatRaw)
            {
                if (Regex.IsMatch(forceStr.Replace(" ", ""), regAc, RegexOptions.IgnoreCase))
                {
                    MatchCollection forceACs = Regex.Matches(forceStr.Replace(" ", ""), regAc, RegexOptions.IgnoreCase);
                    foreach (object ac in forceACs)
                    {
                        var acpin = new ForcePin();
                        for (int acindex = 0; acindex < ac.ToString().Split(':').Length; acindex++)
                        {
                            switch (acindex)
                            {
                                case 0:
                                    acpin.ForceType = ac.ToString().Split(':')[acindex];
                                    break;
                                case 1:
                                    acpin.PinName = ac.ToString().Split(':')[acindex].Replace("&", "::");
                                    break;
                                case 2:
                                    acpin.ForceValue = ac.ToString().Split(':')[acindex];
                                    break;
                            }
                        }
                        forcecondition.ForcePins.Add(acpin);
                    }
                }
                else
                {
                    string regForce = @"(?<pin>[\w\W]+):(?<type>[\w\W]+):(?<value>[\w\W]+);*";
                    string type = "";
                    var pins = new List<string>();
                    string value = "";
                    if (Regex.IsMatch(forceStr, regForce, RegexOptions.IgnoreCase))
                    {
                        type = Regex.Match(forceStr, regForce, RegexOptions.IgnoreCase).Groups["type"].ToString();
                        pins = Regex.Match(forceStr, regForce, RegexOptions.IgnoreCase).Groups["pin"].ToString().Split(',').ToList();
                        value = Regex.Match(forceStr, regForce, RegexOptions.IgnoreCase).Groups["value"].ToString();
                        foreach (string pin in pins)
                        {

                            var forcepin = new ForcePin
                            {
                                PinName = pin.Replace("&", "::"),
                                ForceType = type,
                                ForceValue = value
                            };
                            forcecondition.ForcePins.Add(forcepin);

                        }
                    }

                }
            }
            return forcecondition;
        }
    }
}
