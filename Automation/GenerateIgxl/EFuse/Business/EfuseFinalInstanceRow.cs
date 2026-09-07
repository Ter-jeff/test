using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.EFuse.Enums;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Utility;

using IgxlLib.IgxlBase;

using TestPlanLib.Basic;
using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.Efuse.Input;

namespace Automation.GenerateIgxl.EFuse.Business
{
    public class EfuseFinalInstanceRow
    {
        public EfusePatternRow EfusePatternRow = new EfusePatternRow();
        public BinCutInstanceRow EfuseInstanceRow = new BinCutInstanceRow();
        public EfuseFlowSiteVar SiteVar = new EfuseFlowSiteVar();
        public EfuseExtraType ExtraType = EfuseExtraType.Normal;
        public bool IsFinalInit = false;
        public string PayloadType = "";
        public string BankName = "";
        public string TestName { get; set; } = "";
        public string TestNameTemp { get; set; } = "";
        public string InitPatName { get; set; } = "";
        public string PatSetName { get; set; } = ""; //with RowNum
        public string PatSetNameTemp = "";          //Without RowNum
        public EfuseReadRow EfuseReadRow;
        public string Voltage = "";

        public bool NopByEnableWord
        {
            get
            {
                return EfuseInstanceRow.EnableFlow.Equals("NOP", StringComparison.CurrentCultureIgnoreCase);
            }
        }
        private PatSet _patSet;
        public PatSet PatSet
        {
            get
            {
                if (_patSet == null)
                {
                    _patSet = new PatSet();
                    string commentPatSet = EfusePatternRow.SheetName + ": RowNum" + EfusePatternRow.RowNum;
                    bool isNeedMergeInitPats = LocalSpecs.Options.MergeInitPatterns;
                    if (EfusePatternRow.InitList.Count > 1 && isNeedMergeInitPats && PatSetName.Contains("Init_ALL"))
                    {
                        _patSet.PatSetName = PatSetName;
                        foreach (string pattern in EfusePatternRow.InitList.Distinct())
                        {
                            if (!string.IsNullOrEmpty(pattern))
                            {
                                _patSet.AddRow(FillPatSetRow(pattern, commentPatSet));
                            }
                        }
                    }

                    if (EfusePatternRow.InitList.Count > 1 && EfuseInstanceRow.InitList.Any(p => !string.IsNullOrEmpty(p)))
                    {
                        _patSet.PatSetName = PatSetName;
                        foreach (string pattern in EfusePatternRow.InitList.Distinct())
                        {
                            if (!string.IsNullOrEmpty(pattern))
                            {
                                _patSet.AddRow(FillPatSetRow(pattern, commentPatSet));
                            }
                        }
                    }

                    if (EfusePatternRow.PayloadList.Count > 1)
                    {
                        _patSet.PatSetName = PatSetName;
                        foreach (string pattern in EfusePatternRow.PayloadList.Distinct())
                        {
                            if (!string.IsNullOrEmpty(pattern))
                            {
                                _patSet.AddRow(FillPatSetRow(pattern, commentPatSet));
                            }
                        }
                    }
                }
                return _patSet;
            }
        }

        private PatSetRow FillPatSetRow(string pattern, string comment = "")
        {
            var resultRow = new PatSetRow { File = pattern, Burst = "No", Comment = comment };
            return resultRow;
        }

        protected internal string GetPatSetNameForArgument()
        {
            return EfusePatternRow.PayloadList.Count == 1 ? EfusePatternRow.PayloadList[0] : PatSetName;
        }

        protected internal string GetInitPatSetNameForArgument(bool isInit = false)
        {
            bool isNeedMergeInitPats = LocalSpecs.Options.MergeInitPatterns;
            string initPatName = Combination.CombineByUnderLine(new List<string>
                        {
                            EFuseConst.ConvertToFullBankName(EfusePatternRow.BankName),
                            "Init",
                            "ALL"
                        });
            int idx = 0;
            if (Regex.IsMatch(TestName, @"_init(?<idx>\d+)", RegexOptions.IgnoreCase))
            {
                idx = int.Parse(Regex.Match(TestName, @"_init(?<idx>\d+)", RegexOptions.IgnoreCase).Groups["idx"].ToString());
            }
            return isInit
                ? EfusePatternRow.InitList.Any() ? EfusePatternRow.InitList[idx] : ""
                : isNeedMergeInitPats
                    ? EfusePatternRow.InitList.Count > 1
                        ? initPatName
                        : EfusePatternRow.InitList.Any() ? EfusePatternRow.InitList[idx] : ""
                    : "";
        }

        protected internal string GetAcCategory(string timeSet)
        {
            string acCategory = AcTSetCategoryMapSingleton.Instance().GetCategory(timeSet, BlockType.Efuse);
            if (acCategory.Equals("TBD", StringComparison.OrdinalIgnoreCase))
            {
                acCategory = AcTSetCategoryMapSingleton.Instance().GetCategory(timeSet);
            }

            return acCategory;
        }

        protected internal string GetTimeSet(string pattern)
        {
            if (AcTSetCategoryMapSingleton.Instance().PatternList.TryGetValue(pattern.ToLower(), out PatternData patternData))
            {
                return patternData.TimeSetVersion;
            }

            return "";
        }

        protected internal string GetTimeSet()
        {
            var timeSets = new List<string>();

            foreach (string patternName in EfusePatternRow.InitList)
            {
                if (AcTSetCategoryMapSingleton.Instance().PatternList.TryGetValue(patternName.ToLower(), out PatternData patternData))
                {
                    timeSets.Add(patternData.TimeSetVersion);
                }
            }

            foreach (string patternName in EfusePatternRow.PayloadList)
            {
                if (AcTSetCategoryMapSingleton.Instance().PatternList.TryGetValue(patternName.ToLower(), out PatternData patternData))
                {
                    timeSets.Add(patternData.TimeSetVersion);
                }
            }

            return string.Join(",", timeSets.Distinct().ToList());
        }

        protected internal string GenerateLevel()
        {
            return "Levels_Efuse";
        }
    }
}
