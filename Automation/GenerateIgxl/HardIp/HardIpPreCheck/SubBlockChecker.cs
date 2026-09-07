using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;
using CommonLib.Utility;

namespace Automation.GenerateIgxl.HardIp.HardIpPreCheck
{
    public class SubBlockChecker
    {
        private readonly List<string> _currentSubBlockList = new List<string>();
        private readonly List<string> _localSubBlockList = new List<string>();
        private readonly Dictionary<string, HardIpPattern> _refPatternsGlobal;

        public SubBlockChecker(Dictionary<string, HardIpSheet> planDic)
        {
            var list = planDic.SelectMany(x => x.Value.Rows).ToList();
            _refPatternsGlobal = list.GroupBy(x => x.SheetSubBlockName).ToDictionary(g => g.Key, g => g.First());
        }

        public void ResetSubBlockList()
        {
            _localSubBlockList.Clear();
        }

        public void Check(HardIpSheet hardIpSheet, HardIpPattern pattern)
        {
            int miscInfoIndex = hardIpSheet.PlanHeaderIdx["miscInfoIndex"];

            foreach (string misc in pattern.MiscInfo.Split(';'))
            {
                if (misc.ContainsIgnoreCase("ref_subblock"))
                {
                    string subblock = CommonGenerator.GetSubBlockName(pattern.Pattern.GetLastPayload(), pattern.MiscInfo, "");
                    if (_refPatternsGlobal.ContainsKey(Combination.CombineByUnderLine(pattern.SheetName, subblock)))
                    {
                        _refPatternsGlobal[Combination.CombineByUnderLine(pattern.SheetName, subblock)].IsBurst = true;
                    }

                    List<string> refPatterns = misc.Split(':')[1].Split('+', '#').ToList();
                    foreach (string refPattern in refPatterns)
                    {
                        if (_refPatternsGlobal.ContainsKey(Combination.CombineByUnderLine(pattern.SheetName, refPattern.Replace("_", ""))))
                        {
                            _refPatternsGlobal[Combination.CombineByUnderLine(pattern.SheetName, refPattern.Replace("_", ""))].IsBurst = true;
                        }
                        else
                        {
                            //not found in global and local
                            ErrorReportManager.AddError(
                                HardIpErrorType.E_SubBlockNotFound_01,
                                pattern.SheetName,
                                pattern.RowNum,
                                miscInfoIndex,
                                [refPattern, pattern.SheetName]
                            );
                        }
                    }
                }
            }

            foreach (string item in pattern.MiscInfo.Split(';'))
            {
                if (!item.Contains(":"))
                {
                    continue;
                }

                string paramName = item.Split(':')[0].Trim();
                string paramValue = item.Split(':')[1].Trim();

                if (paramName.Equals(HardIpConstData.SubBlockName, StringComparison.OrdinalIgnoreCase))
                {
                    List<string> arr = item.Split(':').ToList();
                    if (arr.Count != 2) //=> exceed 1 or lower than 1 is not valid
                    {
                        ErrorReportManager.AddError(
                            HardIpErrorType.E_SubBlockUsage_01,
                            pattern.SheetName,
                            pattern.RowNum,
                            miscInfoIndex,
                            [string.Join(":", arr)]
                        );
                    }
                    else // check is repeated
                    {
                        string subBlockName = string.Join("_", arr);
                        if (_localSubBlockList.Contains(subBlockName))
                        {
                            ErrorReportManager.AddError(
                                HardIpErrorType.E_RepeatSubBlock_01,
                                pattern.SheetName,
                                pattern.RowNum,
                                miscInfoIndex,
                                [subBlockName]
                            );
                        }
                        else
                        {
                            _localSubBlockList.Add(subBlockName);
                            if (!_currentSubBlockList.Contains(subBlockName))
                            {
                                _currentSubBlockList.Add(subBlockName);
                            }
                            else
                            {
                                ErrorReportManager.AddError(
                                    HardIpErrorType.W_ExistedSubBlock_01,
                                    pattern.SheetName,
                                    pattern.RowNum,
                                    miscInfoIndex,
                                    [subBlockName]
                                );
                            }
                        }

                    }
                    if (Regex.IsMatch("-", paramValue))
                    {
                        ErrorReportManager.AddError(
                            HardIpErrorType.W_RepeatSubBlock_01,
                            pattern.SheetName,
                            pattern.RowNum,
                            miscInfoIndex,
                            [paramValue]
                        );
                    }
                    break;
                }
            }
        }

        public void CheckNoBurstItem(KeyValuePair<string, HardIpSheet> dic, bool needBurst, Dictionary<string, HardIpPattern> refPatterns, Dictionary<string, HardIpPattern> refPatternsGlobal)
        {
            bool isFoundInLocal = refPatterns.All(subblock => !needBurst || subblock.Value.IsBurst);
            if (isFoundInLocal)
            {
                return;
            }

            foreach (KeyValuePair<string, HardIpPattern> globalSubblock in refPatternsGlobal)
            {
                if (globalSubblock.Value.SheetName.Equals(dic.Key) && !globalSubblock.Value.IsBurst)
                {
                    string subBlock = CommonGenerator.GetSubBlockName(
                        globalSubblock.Value.Pattern.GetLastPayload(),
                        globalSubblock.Value.MiscInfo,
                        ""
                    );
                    ErrorReportManager.AddError(
                        HardIpErrorType.W_NoBurstSubBlock_01,
                        globalSubblock.Value.SheetName,
                        globalSubblock.Value.RowNum,
                        dic.Value.PlanHeaderIdx["miscInfoIndex"],
                        [subBlock]
                    );
                }
            }
        }
    }
}
