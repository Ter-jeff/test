using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.Utility.HardIP;

using CommonLib.Utility;

namespace Automation.GenerateIgxl.HardIp.InputReader
{
    internal class PatternBurstResolver
    {
        private static readonly Regex _regex = new Regex("ref_subblock", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        internal void SetPatternBurst(ConcurrentDictionary<string, HardIpSheet> planDic)
        {
            #region process pattern burst

            IEnumerable<HardIpPattern> rows = planDic.Select(p => p.Value).SelectMany(x => x.Rows).ToList()
                .Where(x => !string.IsNullOrEmpty(x.SubBlockCopy)).ToList();
            var refDic = rows.GroupBy(x => x.SubBlockCopy).ToDictionary(x => x.Key, x => x.First());
            var refDicGlobal = rows.GroupBy(x => x.SheetSubBlockName).ToDictionary(x => x.Key, x => x.First());

            foreach (HardIpPattern pattern in planDic.Select(p => p.Value).SelectMany(x => x.Rows).ToList())
            {
                if (pattern.Pattern.IsMultiple() || pattern.MiscInfo.IndexOf("ref_subblock", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    ProcessBurstPattern(pattern, planDic, refDic, refDicGlobal);
                }
            }
            #endregion
        }

        private void ProcessBurstPattern(HardIpPattern pattern, ConcurrentDictionary<string, HardIpSheet> planDic,
            Dictionary<string, HardIpPattern> refDic, Dictionary<string, HardIpPattern> refDicGlobal)
        {
            var assignments = new List<string>();
            var multiMiscInfo = new List<string>();
            var equation = new List<string>();
            var calc = new List<string>();
            var regAssignList = new List<string>();
            string blockName = CommonGenerator.GetBlockNameFromSheetName(pattern.SheetName);
            if (!(pattern.MiscInfo.IndexOf("ref_subblock", StringComparison.OrdinalIgnoreCase) >= 0))
            {
                if (pattern.BurstPatterns.Count == 0)
                {
                    BuildSubBlockBurstPatterns(pattern, assignments, multiMiscInfo, equation, calc, regAssignList);
                }
            }
            if (pattern.MeasPins.Count > 0)
            {
                return;
            }

            ProcessRefSubBlockMisc(pattern, planDic, refDic, refDicGlobal, blockName, assignments, multiMiscInfo, equation, calc, regAssignList);
        }

        private void BuildSubBlockBurstPatterns(HardIpPattern pattern, List<string> assignments, List<string> multiMiscInfo,
            List<string> equation, List<string> calc, List<string> regAssignList)
        {
            List<string> assignList = pattern.RegisterAssignment.Trim().Split('|').ToList();
            if (assignList.Count == 1)
            {
                assignList = Enumerable.Repeat(assignList.First(), pattern.Pattern.PatternSetList.SelectMany(p => p).Count()).ToList();
            }

            int srcIndex = 0;
            foreach (string subPatternName in pattern.Pattern.PatternSetList.SelectMany(p => p))
            {
                HardIpInfo info = HardIpService.GetHardIpInfo(subPatternName);
                HardIpPattern subPattern;
                if (!subPatternName.Equals(pattern.Pattern.GetLastPayload()))
                {
                    subPattern = new HardIpPattern
                    {
                        SheetName = pattern.SheetName,
                        Pattern = new PatternClass(subPatternName),
                        DigSrcEquation = info.SendBitName
                    };
                }
                else
                {
                    subPattern = pattern.Copy();
                    subPattern.Pattern = new PatternClass(subPatternName);
                    subPattern.BurstPatterns.Clear();
                    subPattern.DigSrcEquation = info.SendBitName;
                }

                if (assignList.Count > srcIndex)
                {
                    if (string.IsNullOrEmpty(info.DigSrcAssignment) && string.IsNullOrEmpty(info.SendBitName))
                    {
                        equation.Add("");
                        assignments.Add("");
                        subPattern.RegisterAssignment = "";
                    }
                    else
                    {
                        equation.Add(info.SendBitName);
                        assignments.Add(assignList[srcIndex]);
                        subPattern.RegisterAssignment = assignList[srcIndex];
                    }
                }
                else
                {
                    equation.Add("");
                    assignments.Add("");
                    subPattern.RegisterAssignment = "";
                }
                pattern.BurstPatterns.Add(subPattern);
                srcIndex++;
            }

            if (pattern.SweepCodes.Count == 0 && pattern.BurstPatterns.Any())
            {
                pattern.SweepCodes = pattern.BurstPatterns.Last().SweepCodes;
            }

            if (pattern.SweepVoltage.Count == 0 && pattern.BurstPatterns.Any())
            {
                pattern.SweepVoltage = pattern.BurstPatterns.Last().SweepVoltage;
            }

            if (pattern.Shmoo.CharSteps.Count == 0 && pattern.BurstPatterns.Any())
            {
                pattern.Shmoo = pattern.BurstPatterns.Last().Shmoo;
            }

            if (string.IsNullOrEmpty(pattern.RegAssignName))
            {
                pattern.RegAssignName = string.Join("|", regAssignList);
            }

            if (string.IsNullOrEmpty(pattern.CalcEqn))
            {
                pattern.CalcEqn = string.Join(";", calc).Trim(';');
            }

            pattern.MiscInfo += ";" + string.Join(";", multiMiscInfo);
            if (string.IsNullOrEmpty(pattern.DigSrcEquation))
            {
                pattern.DigSrcEquation = string.Join("|", equation);
            }
        }

        private void ProcessRefSubBlockMisc(HardIpPattern pattern, ConcurrentDictionary<string, HardIpSheet> planDic,
            Dictionary<string, HardIpPattern> refDic, Dictionary<string, HardIpPattern> refDicGlobal, string blockName,
            List<string> assignments, List<string> multiMiscInfo, List<string> equation, List<string> calc, List<string> regAssignList)
        {
            foreach (string misc in pattern.MiscInfo.Split(';'))
            {
                if (_regex.IsMatch(misc))
                {
                    List<string> refPatterns = misc.Split(':')[1].Split('#', '+').ToList();
                    int patIndex = 0;
                    int offset = 0;
                    foreach (string refPattern in refPatterns)
                    {
                        if (refDic.ContainsKey(refPattern.Replace("_", "")) && (refDic[refPattern.Replace("_", "")].SheetName.Equals(pattern.SheetName, StringComparison.OrdinalIgnoreCase) || pattern.Pattern.IsMultiTimeDomain()))
                        {
                            HardIpPattern refPatsData = refDic[refPattern.Replace("_", "")];
                            offset = AppendRefPatternData(pattern, planDic, refPatsData, refPattern, blockName, patIndex, offset, assignments, equation, calc, multiMiscInfo, regAssignList);
                        }
                        else if (refDicGlobal.ContainsKey(Combination.CombineByUnderLine(pattern.SheetName, refPattern.Replace("_", ""))))
                        {
                            HardIpPattern hardIpPattern = refDicGlobal[Combination.CombineByUnderLine(pattern.SheetName, refPattern.Replace("_", ""))];
                            offset = AppendGlobalRefPatternData(pattern, planDic, hardIpPattern, refPattern, blockName, patIndex, offset, assignments, equation, calc, multiMiscInfo, regAssignList);
                        }
                        patIndex++;
                    }

                    if (pattern.ForceConditionList.Count == 0 && pattern.BurstPatterns.Any())
                    {
                        pattern.ForceConditionList = pattern.ForceConditionList.Any() ? pattern.ForceConditionList :
                            pattern.BurstPatterns.Last().ForceConditionList;
                        pattern.ForceCondition = !string.IsNullOrEmpty(pattern.ForceCondition.ForceCondition) ? pattern.ForceCondition :
                            pattern.BurstPatterns.Last().ForceCondition;
                    }
                    if (pattern.SweepCodes.Count == 0 && pattern.BurstPatterns.Any())
                    {
                        pattern.SweepCodes = pattern.BurstPatterns.Last().SweepCodes;
                    }

                    if (pattern.SweepVoltage.Count == 0 && pattern.BurstPatterns.Any())
                    {
                        pattern.SweepVoltage = pattern.BurstPatterns.Last().SweepVoltage;
                    }

                    if (pattern.Shmoo.CharSteps.Count == 0 && pattern.BurstPatterns.Any())
                    {
                        pattern.Shmoo = pattern.BurstPatterns.Last().Shmoo;
                    }

                    if (string.IsNullOrEmpty(pattern.RegAssignName))
                    {
                        pattern.RegAssignName = string.Join("|", regAssignList);
                    }

                    if (string.IsNullOrEmpty(pattern.RegisterAssignment))
                    {
                        pattern.RegisterAssignment = string.Join("|", assignments);
                    }

                    if (string.IsNullOrEmpty(pattern.CalcEqn))
                    {
                        pattern.CalcEqn = string.Join(";", calc).Trim(';');
                    }

                    pattern.MiscInfo += ";" + string.Join(";", multiMiscInfo);
                    if (string.IsNullOrEmpty(pattern.DigSrcEquation))
                    {
                        pattern.DigSrcEquation = string.Join("|", equation);
                    }
                }
            }
        }

        private int AppendRefPatternData(HardIpPattern pattern, ConcurrentDictionary<string, HardIpSheet> planDic,
            HardIpPattern refPatsData, string refPattern, string blockName, int patIndex, int offset,
            List<string> assignments, List<string> equation, List<string> calc, List<string> multiMiscInfo, List<string> regAssignList)
        {
            assignments.Add(refPatsData.RegisterAssignment);
            equation.Add(refPatsData.DigSrcEquation);
            bool isUpdateOffset = false;
            foreach (MeasPin measPin in refPatsData.MeasPins)
            {
                MeasPin newPin = measPin.Copy();
                if (newPin.SequenceIndex != 0)
                {
                    newPin.SequenceIndex += offset;
                    isUpdateOffset = true;
                }
                newPin.PatternIndex = patIndex;
                if (pattern.Pattern.IsMultiTimeDomain())
                {
                    newPin.TestName += "_" + refPattern;
                }

                pattern.MeasPins.Add(newPin);
            }
            if (isUpdateOffset)
            {
                offset += refPatsData.MeasPins.Max(pin => pin.SequenceIndex);
            }

            pattern.TestPlanSequences.AddRange(refPatsData.TestPlanSequences);
            calc.Add(refPatsData.CalcEqn);
            List<HardIpPattern> hardIpPatterns = pattern.SheetName.Equals("HARDIP_MTD") ? planDic.Select(sheet => sheet.Value).SelectMany(x => x.Rows).ToList() : planDic[pattern.SheetName].Rows;
            multiMiscInfo.Add(hardIpPatterns.Where(x =>
                CommonGenerator.GetSubBlockName(x.Pattern.TestPlanPatternName, x.MiscInfo, blockName).Equals(refPattern)).Select(x => x.MiscInfo).ToList().FirstOrDefault());
            regAssignList.Add(refPatsData.RegAssignName);
            pattern.BurstPatterns.Add(refPatsData);

            return offset;
        }

        private int AppendGlobalRefPatternData(HardIpPattern pattern, ConcurrentDictionary<string, HardIpSheet> planDic,
            HardIpPattern hardIpPattern, string refPattern, string blockName, int patIndex, int offset,
            List<string> assignments, List<string> equation, List<string> calc, List<string> multiMiscInfo, List<string> regAssignList)
        {
            assignments.Add(hardIpPattern.RegisterAssignment);
            equation.Add(hardIpPattern.DigSrcEquation);
            bool isUpdateOffset = false;
            foreach (MeasPin measPin in hardIpPattern.MeasPins)
            {
                MeasPin newPin = measPin.Copy();
                if (newPin.SequenceIndex != 0)
                {
                    newPin.SequenceIndex += offset;
                    isUpdateOffset = true;
                }
                newPin.PatternIndex = patIndex;
                pattern.MeasPins.Add(newPin);
            }
            if (isUpdateOffset)
            {
                offset += hardIpPattern.MeasPins.Max(pin => pin.SequenceIndex);
            }

            pattern.TestPlanSequences.AddRange(hardIpPattern.TestPlanSequences);
            calc.Add(hardIpPattern.CalcEqn);
            IEnumerable<HardIpPattern> hardIpPatterns = pattern.SheetName.Equals("HARDIP_MTD") ? planDic.Select(sheet => sheet.Value).SelectMany(x => x.Rows) : planDic[pattern.SheetName].Rows;
            multiMiscInfo.Add(hardIpPatterns.Where(x =>
                CommonGenerator.GetSubBlockName(x.Pattern.TestPlanPatternName, x.MiscInfo, blockName).Equals(refPattern)).Select(x => x.MiscInfo).ToList().FirstOrDefault());
            regAssignList.Add(hardIpPattern.RegAssignName);
            pattern.BurstPatterns.Add(hardIpPattern);

            return offset;
        }
    }
}
