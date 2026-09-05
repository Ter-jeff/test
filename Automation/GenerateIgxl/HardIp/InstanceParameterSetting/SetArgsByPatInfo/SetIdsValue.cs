using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.Static;
using Automation.Utility.HardIP;

using CommonLib.Enums;
using CommonLib.Extension;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo
{
    public class SetIdsValue : SetValueBase
    {
        private Dictionary<string, string> _miscInfoDict = new Dictionary<string, string>();
        private readonly Regex _regexIdsHighLimitSpecialHandler = new Regex(@"^Binning\((?<str>((.+[+|-].+([+|-].+)?)))\)\*(?<ratio>(\d+\.\d+))", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public SetIdsValue(HardIpInputData hardIpInputData, HardIpSheet hardIpSheet) : base(hardIpInputData, hardIpSheet)
        {
            ReservedMiscInfoKeys = new List<string>
            {
                "DisableClockPortName",
                "DisableFRCPinName",
                "FRC_RelayPin",
                "Fuse_Enable",
                "dictName"
            };
        }

        public override void SetArgsListValue(HardIpPattern pattern, ref Function function, string voltage)
        {
            _miscInfoDict = pattern.MiscInfoDict;
            if (function.Type == ".NET")
            {
                CsProcess(pattern, function, voltage);
            }
            else
            {
                VbtProcess(pattern, function, voltage);
            }
        }

        private void CsProcess(HardIpPattern pattern, Function function, string voltage)
        {
            HardIpInfo info = HardIpService.GetHardIpInfo(pattern);
            function.SetParamValue("pattern", pattern.Pattern.GetInstancePatternName());
            string job = "";
            string ttrJob = SearchInfo.GetTtrEnable(pattern.TtrStr, "NV", LocalSpecs.AllJobsIds);
            if (ttrJob.ContainsIgnoreCase("CP"))
            {
                job = "CP";
            }
            else if (ttrJob.ContainsIgnoreCase("FT"))
            {
                job = "FT";
            }
            IEnumerable<string> measPins = pattern.MeasPins.Where(x => !x.PinName.Contains("fakepin", StringComparison.OrdinalIgnoreCase)).Select(x => x.PinName);
            string sortCpFtPin = DataConvertor.SortCpFtPin(string.Join(",", measPins));

            string measuredPinsArg = GetMeasPinsByJob(job, sortCpFtPin);
            if
                ("VDD_AVE,VDD_CPU_SRAM,VDD_DCS_DDR,VDD_DISP,VDD_ECPU,VDD_FIXED,VDD_GPU,VDD_LOW,VDD_PCPU,VDD_SOC,VDD_SRAM_GPU,VDD_SRAM_SOC,VDD_AMPH_DDR,VDD_FIXED_AMUX,VDD_FIXED_CIO,VDD_FIXED_CPU,VDD_FIXED_ECPU_MTR,VDD_FIXED_LPDP_RX_DCVI,VDD_FIXED_LPDP_TX_INT_DCVI,VDD_FIXED_LPDP_TX_SEC_DCVI,VDD_FIXED_MTR_CPM_PCPU,VDD_FIXED_PCIE_DCVI,VDD_FIXED_PCIE_REFBUF,VDD_FIXED_PLL,VDD_FIXED_PLL_SOC,VDD_FIXED_USB,VDD_FIXED_XTAL,VDD_SRAM_ULPPLL_FLPPLL_SLC,VDD_SRAM_USB_DEBUG,VDD_SRAM_VID_PLL,VDD_SRAM_VIDSEC_PLL,VDD10_ADC_SOC,VDD10_AMUX_FMON,VDD10_CIO,VDD10_CPLL_SOC,VDD10_HSCDFT0,VDD10_HSCDFT1,VDD10_LPDP_RX_DCVI,VDD10_LPDP_TX_INT,VDD10_LPDP_TX_SEC,VDD10_PCIE_DCVI,VDD10_PCIE_REFBUF,VDD10_PLL,VDD10_PLL_DDR,VDD10_SLC_PLL,VDD10_ULPPLL_FLPPLL,VDD10_VID_PLL,VDD10_VIDSEC_PLL,VDD10_XTAL,VDD12_USB,VDD12_USB_DEBUG,VDDIO12_AOP,VDDIO12_AOP_2,VDDIO12_GRP,VDDQL_DDR" ==
                 measuredPinsArg)
            {

            }
            function.SetParamValue("measuredPins", measuredPinsArg);
            function.SetParamValue("autoRangePins", pattern.MiscInfoDict.ContainsKey("iRangeRatio") && !string.IsNullOrEmpty(pattern.MiscInfoDict["iRangeRatio"]) ? "" : measuredPinsArg);

            //DisableClk
            _miscInfoDict.TryGetValue("DisableClockPortName", out string disableClockPortName);

            _miscInfoDict.TryGetValue("DisableFRCPinName", out string disableFrcPinName);

            _miscInfoDict.TryGetValue("FRC_RelayPin", out string frcRelayPin);

            _miscInfoDict.TryGetValue("Fuse_Enable", out string fuseEnable);

            _miscInfoDict.TryGetValue("dictName", out string dictName);

            string disableClock = "";
            if (!string.IsNullOrEmpty(disableClockPortName))
            {
                disableClock += $",{disableClockPortName}";
            }

            if (!string.IsNullOrEmpty(disableFrcPinName))
            {
                disableClock += $",{disableFrcPinName}";
            }

            if (!string.IsNullOrEmpty(frcRelayPin))
            {
                disableClock += $":{frcRelayPin}";
            }

            disableClock = disableClock.Trim(',').Trim(':');
            function.SetParamValue("disableClock", disableClock);

            //Calc_Eqn and storeName
            function.SetParamValue("calcFunc", GetCalcFunc(pattern));

            function.SetParamValue("initialCRFromLimit", "TRUE");

            function.SetParamValue("searchSteps", "7");

            function.SetParamValue("flagWait", SearchInfo.GetCpuflag(info, pattern));

            if (dictName != null)
            {
                function.SetParamValue("dictName", dictName.Trim('\"'));
            }

            function.SetParamValue("fuseEnable", GetFuseEnable(pattern.SubBlock, fuseEnable));

            function.SetParamValue("digSrcAssignment", pattern.RegisterAssignment.Replace("[", ":").Replace("]", ""));
            //DigSrc_Equation: From patInfo file "Send Bit Name"
            function.SetParamValue("digSrcEquation", pattern.DigSrcEquation);
            //DigSrc_Pin
            function.SetParamValue("digSrcPin", SearchInfo.GetSrcPin(info));

            function.SetParamValue("interposePreMeasure", SearchInfo.GetPrePat(pattern, voltage));

            function.SetParamValue("interposePostMeasure", SearchInfo.GetPostMeas(pattern));

            function.SetParamValue("idsHighLimitSpecialHandler", GetBvIdsHighLimitRatio(pattern, job));
        }

        private void VbtProcess(HardIpPattern pattern, Function function, string voltage)
        {
            function.SetParamValue("patt", pattern.Pattern.GetInstancePatternName());
            string measPins = "";

            foreach (MeasPin pin in pattern.MeasPins)
            {
                if (pin.PinName.Contains("fakepin", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                measPins += pin.PinName + ",";
            }

            if (LocalSpecs.Options.Device != EnumDevice.RF)
            {
                function.SetParamValue("DCVS_Power_Pin", DataConvertor.SortCpFtPin(measPins.TrimEnd(',')));
            }
            else
            {
                function.SetParamValue("DCVS_Power_Pin", CommonConst.DcvsPower);
                function.SetParamValue("DCVI_Power_Pin", CommonConst.DcviPower);
            }

            //repeat_count
            string repeatStr = HardIpService.GetRepeatMapping(pattern.MiscInfo);
            int repeatCount = 1;
            if (repeatStr != "")
            {
                repeatCount = repeatStr.Split(',').Length;
            }

            function.SetParamValue("repeat_count", repeatCount.ToString("G15"));

            //DisableClk
            function.SetParamValue("DisableClock", "1");
            //Calc_Eqn and storeName
            function.SetParamValue("Calc_Eqn", DataConvertor.ConvertValueSpec(pattern.CalcEqn));

            function.SetParamValue("FlowLimitForInitIRange", "-1");

            function.SetParamValue("DigSrc_Assignment", pattern.RegisterAssignment.Replace("[", ":").Replace("]", ""));
            //DigSrc_Equation: From patInfo file "Send Bit Name"
            function.SetParamValue("DigSrc_Equation", pattern.DigSrcEquation);

            HardIpInfo info = HardIpService.GetHardIpInfo(pattern);
            function.SetParamValue("DigSrc_Sample_Size", info.SendBit);
            function.SetParamValue("DigSrc_DataWidth", SearchInfo.GetDigDataWidth(info.SendBitStr, "0"));
            //DigSrc_Pin
            function.SetParamValue("DigSrc_Pin", SearchInfo.GetSrcPin(info));
            //Interpose_Meas_before
            function.SetParamValue("Interpose_Meas_before", SearchInfo.GetPrePat(pattern, voltage));
            function.SetParamValue("FlagWait", SearchInfo.GetCpuflag(info, pattern));
            function.SetParamValue("BV_SpecialHandle", GetBvSpecialValue(pattern));
        }

        private string GetBvIdsHighLimitRatio(HardIpPattern pattern, string job)
        {
            var result = new List<string>();

            if (job != "CP" && job != "FT")
            {
                return "";
            }

            foreach (MeasPin measpin in pattern.MeasPins)
            {
                if (measpin.MeasLimitsN.Any())
                {
                    var jobGroup = measpin.MeasLimitsN.GroupBy(x => x.JobName).ToDictionary(x => x.Key.ToUpper(), x => x.ToList());
                    KeyValuePair<string, List<MeasLimit>> jobLimit = jobGroup.FirstOrDefault(x => x.Key.StartsWith(job));
                    if (jobLimit.Value != null)
                    {
                        string hiLimit = jobLimit.Value.First().HiLimit;
                        if (_regexIdsHighLimitSpecialHandler.IsMatch(hiLimit))
                        {
                            string content = _regexIdsHighLimitSpecialHandler.Match(hiLimit).Groups["str"].ToString();
                            string ratio = _regexIdsHighLimitSpecialHandler.Match(hiLimit).Groups["ratio"].ToString();
                            var replaceDic = new Dictionary<string, string>();
                            foreach (string str in content.Split('-', '+'))
                            {
                                if (!str.Contains('.'))
                                {
                                    continue;
                                }

                                string replaceStr = str.Split('.').First();
                                replaceDic.Add(str, replaceStr);
                            }
                            foreach (KeyValuePair<string, string> str in replaceDic)
                            {
                                string oriStr = str.Key;
                                string replaceStr = str.Value;
                                content = content.Replace(oriStr, replaceStr);
                            }
                            result.Add($"{measpin.PinName.ToUpper()}=({content})*{ratio}");
                        }
                        else if (!string.IsNullOrEmpty(jobLimit.Key) && hiLimit.Split('*').Length > 1)
                        {
                            string target = jobLimit.Value.First().HiLimit.Split('*')[1].Trim();
                            if (double.TryParse(target, out double outValue))
                            {
                                if (outValue != 1.0)
                                {
                                    result.Add($"{measpin.PinName.ToUpper()}*{target}");
                                }
                            }
                        }
                    }
                }
            }
            return string.Join(",", result);
        }

        /**
         * job:pin* ratio, pin2 * ratio | job2: pin * ratio ,pin2 * ratio
         * */
        private string GetBvSpecialValue(HardIpPattern pattern)
        {
            var ret = new List<string>();
            //get nv limit since ids is only nv
            var dictionaryLimitByJob = new Dictionary<string, List<string>>();
            foreach (MeasPin measpin in pattern.MeasPins)
            {
                if (measpin.MeasLimitsN.Any())
                {
                    var jobGroup = measpin.MeasLimitsN.GroupBy(x => x.JobName).ToDictionary(x => x.Key, x => x.ToList());
                    foreach (KeyValuePair<string, List<MeasLimit>> job in jobGroup)
                    {
                        if (!dictionaryLimitByJob.ContainsKey(job.Key))
                        {
                            dictionaryLimitByJob.Add(job.Key, new List<string>());
                            dictionaryLimitByJob[job.Key].Add(measpin.PinName + ":" + job.Value.First().HiLimit);
                        }
                        else
                        {
                            dictionaryLimitByJob[job.Key].Add(measpin.PinName + ":" + job.Value.First().HiLimit);
                        }
                    }
                }
            }

            foreach (string job in dictionaryLimitByJob.Keys)
            {
                var allLimitWithRatioByJob = new List<string>();
                foreach (string limit in dictionaryLimitByJob[job])
                {
                    if (limit.Contains("*") && limit.Contains(":"))
                    {
                        string pin = limit.Split(':')[0];
                        string ratio = limit.Split('*')[1];
                        if (limit.Contains("="))
                        {
                            string pinJob = limit.Split('=')[0];
                            if (!job.StartsWith(pinJob))
                            {
                                continue;
                            }

                            allLimitWithRatioByJob.Add(pin.Split('=')[1] + "*" + ratio);
                        }
                        else
                        {
                            allLimitWithRatioByJob.Add(pin + "*" + ratio);
                        }
                    }
                }
                if (allLimitWithRatioByJob.Any())
                {
                    ret.Add(job + ":" + string.Join(",", allLimitWithRatioByJob));
                }
            }

            return string.Join("|", ret);
        }

        private string GetMeasPinsByJob(string job, string measPins)
        {
            List<string> measPinsByJob = measPins.ToUpper().Split(';').ToList();
            if (!measPinsByJob.Any(x => x.StartsWith("CP=") || x.StartsWith("FT=")))
            {
                return measPins;
            }

            if (job == "CP")
            {
                string pinsByJob = measPinsByJob.FirstOrDefault(x => x.ToUpper().StartsWith("CP="));
                if (string.IsNullOrEmpty(pinsByJob))
                {
                    return "";
                }

                return pinsByJob.Replace("CP=", "");
            }
            if (job == "FT")
            {
                string pinsByJob = measPinsByJob.FirstOrDefault(x => x.ToUpper().StartsWith("FT="));
                if (string.IsNullOrEmpty(pinsByJob))
                {
                    return "";
                }

                return pinsByJob.Replace("FT=", "");
            }
            return measPins;
        }

        private string GetFuseEnable(string subBlock, string fuseEnable)
        {
            if (string.IsNullOrEmpty(fuseEnable))
            {
                return "";
            }

            if (fuseEnable.Trim().ToUpper() == "TRUE" || fuseEnable.Trim().ToUpper() == "FALSE")
            {
                return fuseEnable.Trim().ToUpper();
            }

            if (subBlock.ContainsIgnoreCase("OFF"))
            {
                return "";
            }

            fuseEnable = fuseEnable.Trim(';').Trim('\"');
            string[] fuseEnableList = fuseEnable.Split(';');

            if (subBlock.ToUpper().EndsWith("CP"))
            {
                string cpSetting = fuseEnableList.FirstOrDefault(x => x.ToUpper().StartsWith("CP="));
                string setting = cpSetting.Replace("CP=", "");
                return setting.ToUpper();
            }

            if (subBlock.ToUpper().EndsWith("FT"))
            {
                string cpSetting = fuseEnableList.FirstOrDefault(x => x.ToUpper().StartsWith("FT="));
                string setting = cpSetting.Replace("FT=", "");
                return setting.ToUpper();
            }

            return "";
        }

        private string GetCalcFunc(HardIpPattern pattern)
        {
            var allCalcFunc = new List<string>();
            pattern.MiscInfoDict.TryGetValue("calcFunc", out string getValueFromMiscInfo);
            if (!string.IsNullOrEmpty(getValueFromMiscInfo))
            {
                allCalcFunc.Add(getValueFromMiscInfo);
            }
            pattern.MeasPins.ForEach
            (
                x =>
                {
                    if (x.MiscInfoDict.TryGetValue("calcFunc", out getValueFromMiscInfo))
                    {
                        if (!string.IsNullOrEmpty(getValueFromMiscInfo))
                        {
                            allCalcFunc.Add(getValueFromMiscInfo);
                        }
                    }
                }
            );
            return string.Join(";", allCalcFunc);
        }
    }
}
