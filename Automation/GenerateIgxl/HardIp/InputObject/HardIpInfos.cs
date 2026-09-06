using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;

namespace Automation.GenerateIgxl.HardIp.InputObject
{
    public class HardIpInfos : Dictionary<string, List<HardIpInfo>>
    {
        public HardIpInfos(IEnumerable<HardIpInfo> collection)
            : base(StringComparer.OrdinalIgnoreCase)
        {
            if (collection == null)
            {
                return;
            }

            foreach (HardIpInfo info in collection)
            {
                if (!this.ContainsKey(info.Payload))
                {
                    this[info.Payload] = new List<HardIpInfo>();
                }
                this[info.Payload].Add(info);
            }
        }

        public HardIpInfos(params HardIpInfo[] items)
            : base(StringComparer.OrdinalIgnoreCase)
        {
            if (items == null)
            {
                return;
            }

            foreach (HardIpInfo info in items)
            {
                if (!this.ContainsKey(info.Payload))
                {
                    this[info.Payload] = new List<HardIpInfo>();
                }
                this[info.Payload].Add(info);
            }
        }

        public HardIpInfo GetHardIpInfo(string patternName, bool useThisVersion = true)
        {
            var hardIpInfo = new HardIpInfo();
            if (this.TryGetValue(patternName, out List<HardIpInfo> results))
            {
                if (useThisVersion && results.Any(x => x.UseThisVersion))
                {
                    return results.FirstOrDefault(x => x.UseThisVersion);
                }
                else if (results.Any())
                {
                    return results.FirstOrDefault();
                }
            }
            return hardIpInfo;
        }

        public HardIpInfo GetHardIpInfo(HardIpPattern pattern)
        {
            var hardIpInfo = new HardIpInfo();
            var payloadName = new List<string>();
            foreach (string patternName in pattern.Pattern.PatternSetList.SelectMany(p => p))
            {
                payloadName.Add(patternName);
                hardIpInfo.MergeInfo(GetHardIpInfo(patternName, useThisVersion: true) ?? new HardIpInfo());
                hardIpInfo.Payload = string.Join("+", payloadName);
            }
            hardIpInfo.MultiDsscOut = hardIpInfo.MultiDsscOut.Substring(0, hardIpInfo.MultiDsscOut.Length - 1);

            if (Regex.IsMatch(pattern.MiscInfo, HardIpConstData.IgnorePatInfo, RegexOptions.IgnoreCase) &&
                hardIpInfo.SeqInfo.Count > 0)
            {
                var newHardIpInfo = new HardIpInfo
                {
                    PatInfoExist = hardIpInfo.PatInfoExist,
                    CapBit = hardIpInfo.CapBit,
                    CapBitName = hardIpInfo.CapBitName,
                    CapBitStr = hardIpInfo.CapBitStr,
                    CapPinName = hardIpInfo.CapPinName,
                    CallSubrs = hardIpInfo.CallSubrs,
                    CallSubrsCnt = hardIpInfo.CallSubrsCnt,
                    DigSrcAssign = hardIpInfo.DigSrcAssign,
                    DigSrcDataWidth = hardIpInfo.DigSrcDataWidth,
                    DsscOut = hardIpInfo.DsscOut,
                    MultiDsscOut = hardIpInfo.MultiDsscOut,
                    Payload = hardIpInfo.Payload,
                    SendBit = hardIpInfo.SendBit,
                    SendBitName = hardIpInfo.SendBitName,
                    SendBitStr = hardIpInfo.SendBitStr,
                    SendPinName = hardIpInfo.SendPinName,
                    Subroutine = hardIpInfo.Subroutine,
                    TimeSet = hardIpInfo.TimeSet,
                    IsIgnoreComment = true
                };
                return newHardIpInfo;
            }
            if (Regex.IsMatch(pattern.MiscInfo, HardIpConstData.IgnorePatDigSrc, RegexOptions.IgnoreCase))
            {
                hardIpInfo.SendBit = "0";
                hardIpInfo.SendBitName = "";
                hardIpInfo.SendBitStr = "";
                hardIpInfo.SendPinName = "";
            }
            return hardIpInfo;
        }
    }
}
