using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.Static;
using Automation.Utility.HardIP;

using CommonLib.Extension;

using LogLib.Utility;

using TestPlanLib.Static;

namespace Automation.GenerateIgxl.HardIp.DividerManager.FlowDividerManager
{
    public class FlowLimitDivider
    {
        /// <summary>
        /// Separate jobs by different use-limit
        /// </summary>
        /// <param name="patternList"></param>
        /// <param name="hardIpInputData"></param>
        /// <returns></returns>
        public List<HardIpPattern> DivideUseLimit(List<HardIpPattern> patternList, HardIpInputData hardIpInputData)
        {
            var resuList = new List<HardIpPattern>();
            foreach (HardIpPattern pattern in patternList)
            {
                resuList.Add(pattern);
                if (pattern.MeasPins.Count == 0)//If no MeasPin, do not need divide limit 
                {
                    continue;
                }

                pattern.UseLimitsH = GenerateLimitByJob(pattern, "H");
                pattern.UseLimitsL = GenerateLimitByJob(pattern, "L");
                pattern.UseLimitsN = GenerateLimitByJob(pattern, "N");

            }
            return resuList;
        }

        internal List<MeasPin> GenerateLimitByJob(HardIpPattern pattern, string voltage)
        {
            HardIpInfo info = HardIpService.GetHardIpInfo(pattern);
            int seqCount = info.SeqInfo.Count == 0 ? pattern.TestPlanSequences.Count : info.SeqInfo.Count;
            var measPins = pattern.MeasPins.Where(x => !x.MiscInfo.ContainsIgnoreCase(HardIpConstData.IgnoreFlowLimit)).ToList();
            List<MeasPin> otherLimit = pattern.FunctionName == VbtFunctionLibShared.VifName || pattern.FunctionName == VbtFunctionLibShared.HardIpmtdTest ?
                measPins.Where(x =>
                    Regex.IsMatch(x.MeasType,
                        $"{MeasType.MeasCalc}|{MeasType.MeasLimit}|{MeasType.MeasCalcLimit}|{MeasType.MeasC}")).ToList() :
                    measPins.Where(x => x.SequenceIndex == 0).ToList();
            var newList = new List<MeasPin>();
            try
            {
                for (int sequenceIndex = 1; sequenceIndex <= seqCount; sequenceIndex++)
                {
                    int loopCnt = 0;
                    var measPin = measPins.Where(x => x.SequenceIndex == sequenceIndex).ToList();
                    foreach (MeasPin pin in measPin)
                    {
                        var pinList =
                            pin.ForceConditions.SelectMany(x => x.ForcePins)
                                .Where(y => y.PinName == pin.PinName)
                                .ToList();
                        loopCnt += pinList.Any() ? pinList.Max(y => y.ForceCnt) : 0;
                    }
                    //loopCnt = loopCnt / measPin.Count;
                    if (measPin.Any())
                    {
                        loopCnt = measPin[0].MeasType == MeasType.MeasR2 ? loopCnt / measPin.Count / 2 : loopCnt / measPin.Count;
                    }

                    loopCnt = loopCnt == 0 ? 1 : loopCnt;
                    for (int i = 0; i < loopCnt; i++)
                    {
                        var tempList = new List<MeasPin>();
                        foreach (MeasPin pin in measPin)
                        {
                            Dictionary<string, List<string>> groupLimits = GroupLimits(pin, voltage);
                            foreach (string limit in groupLimits.Keys)
                            {
                                List<string> jobList = groupLimits[limit];
                                List<string> values = limit.Split(',').ToList();
                                int index = 0;

                                List<string> newjobs = jobList;
                                if (pin.PinName.Contains('='))
                                {
                                    newjobs =
                                        newjobs.Where(
                                            p =>
                                                Regex.IsMatch(p, pin.PinName.Split('=')[0], RegexOptions.IgnoreCase))
                                            .ToList();
                                    if (newjobs.Count == 0)
                                    {
                                        index++;
                                        continue;
                                    }
                                }

                                var newPin = new MeasPin();
                                newPin.Copy(pin);
                                newPin.PinName = pin.PinName;
                                newPin.LowLimit = values[index].Split('$')[0];
                                newPin.HighLimit = values[index].Split('$')[1];
                                newPin.Job = string.Join(",", newjobs);
                                newPin.SequenceIndex = pin.SequenceIndex;
                                if (CheckMergePower(newPin))
                                //"true" means the pin is power merge pin, and mismatch with its job-enable, skip this pin
                                {
                                    index++;
                                    continue;
                                }

                                tempList.Add(newPin);
                                index++;
                            }
                        }

                        newList.AddRange(tempList);
                    }
                }

                foreach (MeasPin pin in otherLimit)
                {
                    Dictionary<string, List<string>> groupLimits = GroupLimits(pin, voltage);
                    foreach (string limit in groupLimits.Keys)
                    {
                        List<string> jobList = groupLimits[limit];
                        List<string> values = limit.Split(',').ToList();
                        int index = 0;

                        var newPin = new MeasPin();
                        newPin.Copy(pin);
                        newPin.PinName = pin.PinName;
                        List<string> newjobs = jobList;
                        if (pin.PinName.Contains('='))
                        {
                            newjobs =
                                newjobs.Where(
                                    p => Regex.IsMatch(p, pin.PinName.Split('=')[0], RegexOptions.IgnoreCase))
                                    .ToList();
                            if (newjobs.Count == 0)
                            {
                                index++;
                                continue;
                            }

                        }
                        newPin.LowLimit = values[index].Split('$')[0];
                        newPin.HighLimit = values[index].Split('$')[1];
                        newPin.Job = string.Join(",", newjobs);
                        newList.Add(newPin);
                        index++;
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }
            return newList;
        }

        internal Dictionary<string, List<string>> GroupLimits(MeasPin pin, string voltage = "")
        {
            var groupLimits = new Dictionary<string, List<string>>();
            //bool allMeasC = true;
            var valueList = new Dictionary<string, string>();
            List<MeasLimit> limits;
            switch (voltage)
            {
                case "":
                    limits = pin.MeasLimitsN;
                    break;
                case "H":
                    limits = pin.MeasLimitsH;
                    break;
                case "L":
                    limits = pin.MeasLimitsL;
                    break;
                case "N":
                    limits = pin.MeasLimitsN;
                    break;
                default:
                    limits = pin.MeasLimitsN;
                    break;
            }


            #region Initial MeasLimits for the pin which is not specified in TestPlan
            if (limits.Count == 0)
            {
                foreach (string job in LocalSpecs.AllJobs)
                {
                    var newLimit = new MeasLimit(job);
                    limits.Add(newLimit);
                }
            }
            #endregion
            foreach (MeasLimit limit in limits)
            {
                if (valueList.ContainsKey(limit.JobName))
                {
                    valueList[limit.JobName] += limit.LoLimit + "$" + limit.HiLimit + ",";
                }
                else
                {
                    valueList.Add(limit.JobName, limit.LoLimit + "$" + limit.HiLimit + ",");
                }
            }

            foreach (KeyValuePair<string, string> group in valueList)
            {
                GenGroups(group.Value, group.Key, groupLimits);
            }
            return groupLimits;
        }

        private static void GenGroups(string limitStr, string jobName, Dictionary<string, List<string>> groupLimits)
        {
            if (limitStr != "")
            {
                limitStr = limitStr.Remove(limitStr.Length - 1, 1);
                if (groupLimits.ContainsKey(limitStr))
                {
                    groupLimits[limitStr].Add(jobName);
                }
                else
                {
                    var jobList = new List<string> { jobName };
                    groupLimits.Add(limitStr, jobList);
                }
            }
        }

        /// <summary>
        /// Check whether the PowerMerge pin matches with its job enable
        /// </summary>
        /// <param name="pin"></param>
        /// <returns></returns>
        internal bool CheckMergePower(MeasPin pin)
        {
            if (!pin.PinName.Contains("VDD", StringComparison.OrdinalIgnoreCase) ||
                TestProgram.IgxlWorkBk.PinMapPair.Value.IsGroupExist(pin.PinName))//PowerMerge pins are Power pin(VDD_)
            {
                return false;
            }

            if (!TestPlanStatic.PowerMergeSheet.PowerMerge.CpPowers.ContainsValue(RemoveJobName(pin.PinName)) &&
                !TestPlanStatic.PowerMergeSheet.PowerMerge.FtPowers.ContainsValue(RemoveJobName(pin.PinName)))
            {
                return false;
            }

            string jobList = pin.Job;
            if (jobList.Contains("CP") && !jobList.Contains("FT") &&
                (!TestPlanStatic.PowerMergeSheet.PowerMerge.CpPowers.ContainsValue(RemoveJobName(pin.PinName)) || pin.PinName.Contains("FT=")))
            {
                return true;
            }

            if (jobList.Contains("FT") && !jobList.Contains("CP") && (!TestPlanStatic.PowerMergeSheet.PowerMerge.FtPowers.ContainsValue(RemoveJobName(pin.PinName)) || pin.PinName.Contains("CP=")))
            {
                return true;
            }

            if (jobList.Contains("FT") && jobList.Contains("CP"))
            {
                MergeCpFtJobs(pin, jobList);
            }
            return false;
        }

        private void MergeCpFtJobs(MeasPin pin, string jobList)
        {
            var newJobList = new List<string>();
            //Remove ft jobs for only-CP pin
            if (!TestPlanStatic.PowerMergeSheet.PowerMerge.FtPowers.ContainsValue(RemoveJobName(pin.PinName)) || pin.PinName.Contains("CP="))
            {
                foreach (string job in jobList.Split(','))
                {
                    if (job.Contains("CP"))
                    {
                        newJobList.Add(job);
                    }
                }
            }
            //Remove cp jobs for only-FT pin
            if (!TestPlanStatic.PowerMergeSheet.PowerMerge.CpPowers.ContainsValue(RemoveJobName(pin.PinName)) || pin.PinName.Contains("FT="))
            {
                foreach (string job in jobList.Split(','))
                {
                    if (job.Contains("FT"))
                    {
                        newJobList.Add(job);
                    }
                }
            }
            if (newJobList.Count > 0)
            {
                pin.Job = string.Join(",", newJobList);
            }
        }

        internal string RemoveJobName(string pinName)
        {
            if (pinName.Contains("="))
            {
                return pinName.Split('=')[1];
            }

            return pinName;
        }
    }
}
