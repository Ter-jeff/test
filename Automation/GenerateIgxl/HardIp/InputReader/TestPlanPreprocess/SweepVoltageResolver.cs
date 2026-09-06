using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

namespace Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess
{
    internal class SweepVoltageResolver
    {
        //sweep(PinA:V:0.1) or sweepY(PinA:V:0.1)
        private static readonly Regex _regSweepVoltage =
            new Regex(@"sweep(?<label>\w*)\s*\((?<SweepStr>.*)\)(?<CustomSetting>:.+)?", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        //sweep[0.001,0.002] or sweepY[0.001,0.002]
        private static readonly Regex _regSweepVoltageList =
            new Regex(@"sweep(?<label>\w*)\s*\[(?<SweepStr>.*)\](?<CustomSetting>:.+)?", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        //NestSweep(PinA:V:0.1) or sweepY(PinA:V:0.1)
        private static readonly Regex _regNestSweepVoltage =
            new Regex(@"nestsweep(?<label>\D*)(?<order>\d*)\s*\((?<SweepStr>.*)\)(?<CustomSetting>:.+)?", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        //NestSweep[0.001,0.002] or sweepY[0.001,0.002]
        private static readonly Regex _regNestSweepVoltageList =
            new Regex(@"nestsweep(?<label>\D*)(?<order>\d*)\s*\[(?<SweepStr>[^\]]+)\](?<CustomSetting>:.+)?", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private readonly TestPlanSheet _planSheet;

        public SweepVoltageResolver(TestPlanSheet planSheet)
        {
            _planSheet = planSheet;
        }

        // this function would derive SweepVoltage from Miscinfo, the return type would be dictionary for X and Y
        // the axis Y should be outside of X
        internal Dictionary<string, List<SweepVData>> GetSweepVoltage(PatternRow pattern)
        {
            bool errorFlagOfMiscInfo = false;
            int miscInfoIndex = _planSheet.PlanHeaderIdx["miscInfoIndex"];
            var result = new Dictionary<string, List<SweepVData>>();
            foreach (string miscinfo in pattern.MiscInfo.Split(';'))
            {
                ParseSweepVoltageFromMiscInfo(miscinfo, result, ref errorFlagOfMiscInfo);
            }
            if (errorFlagOfMiscInfo)
            {
                ErrorReportManager.AddError(
                    HardIpErrorType.E_WrongSweepStep_01,
                    pattern.SheetName,
                    pattern.RowNum,
                    miscInfoIndex,
                    ["Misc Info"]
                );
            }

            bool errorFlagOfForceCondition = false;
            int forceConditionIndex = _planSheet.PlanHeaderIdx["forceIndex"];

            foreach (string forceCondition in pattern.ForceCondition.GetPrePatForceCondition().Split(';'))
            {
                ParseSweepVoltageFromForceCondition(forceCondition, result, ref errorFlagOfForceCondition);
            }

            if (errorFlagOfForceCondition)
            {
                ErrorReportManager.AddError(
                    HardIpErrorType.E_WrongSweepStep_01,
                    pattern.SheetName,
                    pattern.RowNum,
                    forceConditionIndex,
                    ["Force Condition"]
                );
            }
            return result;
        }

        private void ParseSweepVoltageFromMiscInfo(string miscinfo, Dictionary<string, List<SweepVData>> result, ref bool errorFlagOfMiscInfo)
        {
            if (Regex.IsMatch(miscinfo.Split(':')[0], "forloop", RegexOptions.IgnoreCase))
            {
                string loopInfo = miscinfo.Split(':')[1];
                if (loopInfo.Split(',').Length == 4)
                {
                    string label = loopInfo.Split(',')[0];
                    if (!result.ContainsKey(label))
                    {
                        result.Add(label, new List<SweepVData>());
                    }

                    result[label].Add(new SweepVData(loopInfo, false));
                    if (!new SweepVData(loopInfo, false).CheckStep)
                    {
                        errorFlagOfMiscInfo = true;
                    }
                }
            }

            Match matchSweepVoltage = _regSweepVoltage.Match(miscinfo);
            Match matchSweepVoltageList = _regSweepVoltageList.Match(miscinfo);
            Match matchNestSweepVoltage = _regNestSweepVoltage.Match(miscinfo);
            Match matchNestSweepVoltageList = _regNestSweepVoltageList.Match(miscinfo);

            if (matchSweepVoltage.Success || matchNestSweepVoltage.Success)
            {
                string sweepStr = matchSweepVoltage.Groups["SweepStr"].ToString();
                Group labelGroup = matchSweepVoltage.Groups["label"];
                string label = labelGroup.Success ? "V" + labelGroup.Value : "V";

                foreach (string sweepItem in sweepStr.Split('|'))
                {
                    if (sweepItem.Split(':').Length == 2)
                    {
                        if (!result.ContainsKey(label))
                        {
                            result.Add(label, new List<SweepVData> { new SweepVData(sweepItem, label) });
                        }
                        else
                        {
                            result[label].Add(new SweepVData(sweepItem, label));
                        }

                        if (!new SweepVData(sweepItem, label).CheckStep)
                        {
                            errorFlagOfMiscInfo = true;
                        }
                    }
                }
            }
            else if (matchSweepVoltageList.Success || matchNestSweepVoltageList.Success)
            {
                string sweepStr = matchSweepVoltageList.Groups["SweepStr"].ToString();
                Group labelGroup = matchSweepVoltage.Groups["label"];
                string label = labelGroup.Success ? "V" + labelGroup.Value : "V";

                foreach (string sweepItem in sweepStr.Split('|'))
                {
                    if (sweepItem.Split(':').Length == 2)
                    {
                        if (!result.ContainsKey(label))
                        {
                            result.Add(label, new List<SweepVData> { new SweepVData(sweepItem, label) });
                        }
                        else
                        {
                            result[label].Add(new SweepVData(sweepItem, label));
                        }

                        if (!new SweepVData(sweepItem, label).CheckStep)
                        {
                            errorFlagOfMiscInfo = true;
                        }
                    }
                }
            }
        }

        private void ParseSweepVoltageFromForceCondition(string forceCondition, Dictionary<string, List<SweepVData>> result, ref bool errorFlagOfForceCondition)
        {
            string instanceVoltage = GetInstanceVoltageInCondition(forceCondition);

            Match matchSweepVoltage = _regSweepVoltage.Match(forceCondition);
            Match matchNestSweepVoltage = _regNestSweepVoltage.Match(forceCondition);
            Match matchNestSweepVoltageList = _regNestSweepVoltageList.Match(forceCondition);

            if (matchSweepVoltage.Success || matchNestSweepVoltage.Success)
            {
                string sweepStr = matchSweepVoltage.Success ? matchSweepVoltage.Groups["SweepStr"].ToString() : matchNestSweepVoltage.Groups["SweepStr"].ToString();
                string label = sweepStr.Split(':').Length == 3 ? sweepStr.Split(':')[1] : "V";
                string order = "";
                string forceType = label;
                string customSetting = matchSweepVoltage.Success ? matchSweepVoltage.Groups["CustomSetting"].ToString() : matchNestSweepVoltage.Groups["CustomSetting"].ToString();

                if (matchNestSweepVoltage.Success)
                {
                    order = matchNestSweepVoltage.Groups["order"].Value;
                    label = "NestSweep";
                }
                else if (matchSweepVoltage.Success)
                {
                    label += matchSweepVoltage.Groups["label"].Value;
                }

                foreach (string sweepItem in sweepStr.Split('|'))
                {
                    if (sweepItem.Split(':').Length == 2 || sweepItem.Split(':').Length == 3)
                    {
                        if (!result.ContainsKey(label))
                        {
                            result.Add(label, new List<SweepVData> {
                                new SweepVData(sweepItem, label, false, true, order, instanceVoltage: instanceVoltage, forceType: forceType, customSetting: customSetting) });
                        }
                        else
                        {
                            result[label].Add(
                                new SweepVData(sweepItem, label, false, true, order, instanceVoltage: instanceVoltage, forceType: forceType, customSetting: customSetting));
                        }

                        if (!new SweepVData(sweepItem, label, false, true, instanceVoltage: instanceVoltage, forceType: forceType, customSetting: customSetting).CheckStep)
                        {
                            errorFlagOfForceCondition = true;
                        }
                    }
                }
            }
            else if (matchNestSweepVoltageList.Success)
            {
                string str = matchNestSweepVoltageList.Groups["SweepStr"].ToString();
                string label = "NestSweep"; //default to V since it's sweep voltage
                string order = matchNestSweepVoltageList.Groups["order"].Value;

                string[] forceconArr = forceCondition.Split(':');
                string forceType = forceconArr.Length > 1 ? forceconArr[1] : "V";

                string customSetting = matchNestSweepVoltageList.Groups["CustomSetting"].ToString();

                if (!result.ContainsKey(label))
                {
                    result.Add(label, new List<SweepVData> {
                        new SweepVData(forceconArr[0], str, label, order, instanceVoltage: instanceVoltage, forceType: forceType, customSetting: customSetting)});
                }
                else
                {
                    result[label].Add(
                        new SweepVData(forceconArr[0], str, label, order, instanceVoltage: instanceVoltage, forceType: forceType, customSetting: customSetting));
                }
            }
        }

        internal string GetInstanceVoltageInCondition(string forceCondition)
        {
            if (forceCondition.StartsWith("NV@", StringComparison.OrdinalIgnoreCase))
            {
                return "NV";
            }

            if (forceCondition.StartsWith("HV@", StringComparison.OrdinalIgnoreCase))
            {
                return "HV";
            }

            if (forceCondition.StartsWith("LV@", StringComparison.OrdinalIgnoreCase))
            {
                return "LV";
            }

            return "";
        }
    }
}
