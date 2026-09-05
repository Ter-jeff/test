using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.BinCut.Flow;
using TestPlanLib.BinCut.FlowNew;

namespace Automation.GenerateIgxl.BinCut.Base
{
    public class BinCutSourceItem
    {
        private const string TempSensor = "TEMP SENSOR";
        private const string StrCpu = "CPU";
        private const string StrGfx = "GFX";
        private const string StrGpu = "GPU";
        private const string StrSoc = "SOC";
        private const string PerformanceModeMatchPattern = @"^\s*(?<str>[a-zA-Z]+\d+)";

        public string Enable;
        public int RowNum;
        public bool Nop;
        public string Job;
        public string PerformanceMode;
        public readonly string ColumnContent;
        public string PerformanceModeFromColumnContent
        {
            get
            {
                foreach (string match in ColumnContent.Split(new[] { ' ', ':' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    if (Regex.IsMatch(match, PerformanceModeSingleton.RegContainPerformanceModeWithGroup, RegexOptions.IgnoreCase))
                    {
                        return match;
                    }
                }
                return "";
            }
        }

        public string BinningDomain { get; }

        public string TargetPerformanceMode
        {
            get
            {
                return PerformanceMode.Split('_').First();
            }
        }
        public string AllOther { get; }

        public EnumColumnName ColumnName { get; }

        public bool InstOrCallFlowByBms
        {
            get
            {
                return ColumnName.Equals(EnumColumnName.E1Voltage)
                    || ColumnName.Equals(EnumColumnName.RelayOn)
                    || ColumnName.Equals(EnumColumnName.RelayOff)
                    || ColumnName.Equals(EnumColumnName.CallNwireEnable)
                    || ColumnName.Equals(EnumColumnName.CallNwireDisable)
                    || ColumnName.Equals(EnumColumnName.CallTMPS);
            }
        }
        public string Level
        {
            get
            {
                if (ColumnName.Equals(EnumColumnName.TD))
                {
                    return "TD Binning";
                }

                if (ColumnName.Equals(EnumColumnName.Mbist))
                {
                    return "Mbist Binning";
                }

                return "SPI Binning";
            }
        }
        public int JobCount;
        public List<PinInfo> BinValues { set; get; }

        public EnumBinCutTableType TableType { get; }

        public EnumBinCutTableBinType TableBinType { get; }

        public bool CanBeMerged { get; set; }

        public BinCutSourceItem(BinCutFlowSheetRow binCutFlowSheetRow, NewBinCutFlowSheetRow newBinCutFlowSheetRow, EnumColumnName columnName, string columnContent)
        {
            RowNum = newBinCutFlowSheetRow.RowNum;
            BinningDomain = !string.IsNullOrEmpty(newBinCutFlowSheetRow.BinningDomain) ? newBinCutFlowSheetRow.BinningDomain : binCutFlowSheetRow.BinningDomain;
            PerformanceMode = binCutFlowSheetRow.PerformanceMode;
            AllOther = binCutFlowSheetRow.AllOther;
            ColumnName = columnName;
            ColumnContent = columnContent;
            BinValues = binCutFlowSheetRow.PinInfos;
            TableType = newBinCutFlowSheetRow.TableType;
            TableBinType = newBinCutFlowSheetRow.TableBinType;
        }

        public BinCutSourceItem(BinCutFlowSheetRow binCutFlowSheetRow, EnumColumnName columnName, string columnContent, List<PinInfo> pinInfos)
        {
            RowNum = binCutFlowSheetRow.RowNum;
            BinningDomain = binCutFlowSheetRow.BinningDomain;
            PerformanceMode = binCutFlowSheetRow.PerformanceMode;
            AllOther = binCutFlowSheetRow.AllOther;
            ColumnName = columnName;
            ColumnContent = columnContent;
            BinValues = pinInfos;
            TableType = binCutFlowSheetRow.TableType;
            TableBinType = binCutFlowSheetRow.TableBinType;
            Nop = binCutFlowSheetRow.Nop;
        }

        #region Public method
        public virtual BinCutFinalInstanceRow FillBlankRow()
        {
            var row = new BinCutFinalInstanceRow { Nop = !InstOrCallFlowByBms, PerformanceMode = PerformanceModeFromColumnContent };
            string module = "";
            var paraList = new List<string>();
            string domain = GetDomainOfContent();
            if (domain.Equals("DDR", StringComparison.CurrentCultureIgnoreCase))
            {
                module = domain;
            }
            else if (ColumnName == EnumColumnName.TD)
            {
                module = domain + "Td";
            }
            else if (ColumnName == EnumColumnName.Mbist)
            {
                module = domain + "Mbist";
            }
            else if (ColumnName == EnumColumnName.FUNC)
            {
                module = BinCutConstant.ConSpi;
            }

            if (!InstOrCallFlowByBms)
            {
                paraList.Add(module);
                paraList.Add(TargetPerformanceMode);
                paraList.Add(
                    ColumnContent.Replace(" ", "_")
                        .Replace(@"\", "_")
                        .Replace("/", "_")
                        .Replace(";", "_")
                        .Replace("#", "_")
                        .Replace(":", "_"));
            }
            else if (ColumnName.Equals(EnumColumnName.E1Voltage))
            {
                paraList.Add("Set_E1_Voltage");
            }
            else if (ColumnName.Equals(EnumColumnName.RelayOn)
                || ColumnName.Equals(EnumColumnName.RelayOff)
                || ColumnName.Equals(EnumColumnName.CallNwireEnable)
                || ColumnName.Equals(EnumColumnName.CallNwireDisable))
            {
                paraList.Add("Call_" + ColumnContent);
            }
            row.PatSetName = string.Join("_", paraList);
            if (ColumnContent.StartsWith("Flow_TMPS", StringComparison.CurrentCultureIgnoreCase))
            {
                row.PatSetName = ColumnContent;
            }

            row.VbtFunction = "";
            row.PayloadList.Add("");
            row.PatternList.Add("");
            row.FinalJobs.Add(Job);
            return row;
        }

        public string GetBinningDomain()
        {
            if (BinningDomain.Contains(","))
            {
                List<string> domains = BinningDomain.Split(',').ToList();
                string domainPin = "VDD_" + domains.First() + "_" + PerformanceMode;
                domainPin = domainPin + "," + "VDD_" + string.Join("_", domains);
                var pinList = domains.Select(p => "VDD_" + p).ToList();
                AddPinGroup("VDD_" + string.Join("_", domains), pinList);
                return domainPin;
            }
            return "VDD_" + BinningDomain + "_" + PerformanceMode;
        }

        public void AddPinGroup(string groupName, List<string> pinList)
        {

            PinMapSheet pinMap = TestProgram.IgxlWorkBk.PinMapPair.Value;
            if (pinMap.IsGroupExist(groupName))
            {
                return;
            }

            var group = new PinGroup(groupName, "power");
            foreach (string pin in pinList)
            {
                if (pinMap.IsPinExist(pin))
                {
                    var newPin = new Pin(pin, "power");
                    group.AddPin(newPin);
                }
            }
            pinMap.AddRow(group);
        }

        public bool JudgeIsTargetFlow(BinCutFinalInstanceRow binCutInstDataRow)
        {
            string flowName = binCutInstDataRow.BinCutInstanceRow.FlowName;
            string jobTestStage = binCutInstDataRow.BinCutInstanceRow.JobTestStage;
            if (ColumnContent.Contains(";"))
            {
                foreach (string text in ColumnContent.Split(';'))
                {
                    string blockName = Regex.Replace(text, "/s{2,}", " ");
                    if (blockName.Trim().Equals(flowName, StringComparison.CurrentCultureIgnoreCase))
                    {
                        if (string.IsNullOrEmpty(jobTestStage) ||
                            jobTestStage.Equals("All", StringComparison.CurrentCultureIgnoreCase))
                        {
                            return true;
                        }

                        var testingStageArr = jobTestStage.Replace(" ", "").Split(',').Select(x => x.Trim()).ToList();
                        return GetJobFlag(testingStageArr);
                    }
                }
            }

            string blockName2 = Regex.Replace(ColumnContent, "/s{2,}", " ");
            if (ColumnContent.Contains("#"))
            {
                string blockNameRegex = blockName2.Replace("#", "[ |_]");
                if (Regex.IsMatch(flowName, blockNameRegex, RegexOptions.IgnoreCase))
                {
                    return true;
                }
            }
            else if (ColumnContent.Contains(":"))
            {
                string newBlockName = blockName2.Split(new[] { ':' }, StringSplitOptions.RemoveEmptyEntries).Last();
                if (newBlockName.Trim().Equals(flowName, StringComparison.CurrentCultureIgnoreCase))
                {
                    return true;
                }
            }
            else
            {
                if (blockName2.Trim().Equals(flowName, StringComparison.CurrentCultureIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private bool GetJobFlag(List<string> testingStageArr)
        {
            var flags = new List<bool>();
            if (testingStageArr.Exists(x => x.StartsWith("!")))
            {
                foreach (string stage in testingStageArr)
                {
                    if (stage.StartsWith("!", StringComparison.CurrentCultureIgnoreCase))
                    {
                        flags.Add(!stage.ToUpper().TrimStart('!').Contains(Job.ToUpper()));
                    }
                    else
                    {
                        return stage.ContainsIgnoreCase(Job.ToUpper());
                    }
                }
            }
            else
            {
                return testingStageArr.Exists(x => x.ContainsIgnoreCase(Job.ToUpper()));
            }
            return flags.TrueForAll(x => x);
        }

        public string GetBinType()
        {
            if (TableType == EnumBinCutTableType.Hv)
            {
                return "HBV";
            }

            return "BV";
        }

        public string GetDcCategory()
        {
            string dcCategory = Regex.Replace(AllOther, "_LV$", "", RegexOptions.IgnoreCase);
            dcCategory = Regex.Replace(dcCategory, "_NV$", "", RegexOptions.IgnoreCase);
            dcCategory = Regex.Replace(dcCategory, "_HV$", "", RegexOptions.IgnoreCase);
            return dcCategory;
        }

        public string GetDomainOfMode()
        {
            return GetDomain(PerformanceMode);
        }

        public string GetDomainOfContent()
        {
            return GetDomain(ColumnContent);
        }

        public string GetEnableWord(List<string> finalJobs)
        {
            if (finalJobs.Count != JobCount)
            {
                string enable = "";
                foreach (string job in finalJobs)
                {
                    enable = enable + "||" + job;
                }
                enable = enable.Substring(2);
                return enable;
            }
            return "";
        }

        public string GetVbtFunction()
        {
            if (TableType == EnumBinCutTableType.Hv)
            {
                if (Regex.IsMatch(PerformanceMode, "TMPS", RegexOptions.IgnoreCase) || Regex.IsMatch(ColumnContent, TempSensor, RegexOptions.IgnoreCase))
                {
                    return "GradeSearch_HVCC_TMPS_VT";
                }

                if (Regex.IsMatch(ColumnContent, "CPM", RegexOptions.IgnoreCase))
                {
                    return "GradeSearch_postBinCut_VT";
                }

                if (ColumnName == EnumColumnName.FUNC && Regex.IsMatch(ColumnContent, "ilb", RegexOptions.IgnoreCase))
                {
                    return "GradeSearch_HVCC_CallInstance_VT";
                }

                if (ColumnName == EnumColumnName.FUNC && Regex.IsMatch(ColumnContent, "elb", RegexOptions.IgnoreCase))
                {
                    return "GradeSearch_HVCC_CallInstance_VT";
                }

                if (ColumnName == EnumColumnName.FUNC)
                {
                    return "GradeSearch_HVCC_CallInstance_VT";
                }

                return "GradeSearch_HVCC_VT";
            }

            if (Regex.IsMatch(PerformanceMode, "TMPS", RegexOptions.IgnoreCase) || Regex.IsMatch(ColumnContent, TempSensor, RegexOptions.IgnoreCase))
            {
                return "GradeSearch_TMPS_VT";
            }

            if (Regex.IsMatch(ColumnContent, "CPM", RegexOptions.IgnoreCase))
            {
                return "GradeSearch_postBinCut_VT_HIP_CPM";
            }

            if (ColumnName == EnumColumnName.FUNC && Regex.IsMatch(ColumnContent, "ilb", RegexOptions.IgnoreCase))
            {
                return "GradeSearch_CallInstance_VT";
            }

            if (ColumnName == EnumColumnName.FUNC && Regex.IsMatch(ColumnContent, "elb", RegexOptions.IgnoreCase))
            {
                return "GradeSearch_CallInstance_VT";
            }

            if (ColumnName == EnumColumnName.FUNC)
            {
                return "GradeSearch_CallInstance_VT";
            }

            if (ColumnName == EnumColumnName.E1Voltage)
            {
                return "Set_E1_Voltage_ForPmode";
            }

            return "GradeSearch_VT";
        }

        public string GetVbtFunctionCs()
        {
            if (ColumnName == EnumColumnName.E1Voltage)
            {
                return "SetVoltageWithoutTest";
            }

            return "BinCutTest";
        }

        public bool IsTargetInstanceName(string instanceName)
        {
            instanceName = instanceName.Replace("<", "").Replace(">", "");

            foreach ((Func<ValidateData, bool> shouldApply, Func<ValidateData, bool?> evaluate) in _targetInstanceRules)
            {
                ValidateData vd = new ValidateData()
                {
                    InstanceName = instanceName,
                    Item = this,
                };
                if (!shouldApply(vd))
                {
                    continue;
                }
                bool? result = evaluate(vd);
                if (result.HasValue)
                {
                    return result.Value;
                }
            }

            return false;
        }

        public bool IsTargetSpiPattern(string pTaget, string pPattern)
        {
            bool lBResult = true;
            string[] subNames = pPattern.ToUpper().Split('_');
            var subNameList = new List<string>(subNames);

            const string reDelimiter = @"\s|_";
            List<string> subWords = Regex.Split(pTaget, reDelimiter).ToList();

            foreach (string word in subWords)
            {
                if (!subNameList.Exists(p => p.Contains(word)))
                {
                    lBResult = false;
                    break;
                }
            }
            return lBResult;
        }

        public string GetDomain(string keyword)
        {
            List<string> arr = keyword.Split(' ').ToList();
            if (arr.Exists(x => x.Equals(StrCpu, StringComparison.CurrentCultureIgnoreCase)))
            {
                return BinCutConstant.ConCpu;
            }

            if (arr.Exists(x => x.Equals(StrGfx, StringComparison.CurrentCultureIgnoreCase)) ||
                arr.Exists(x => x.Equals(StrGpu, StringComparison.CurrentCultureIgnoreCase)))
            {
                return BinCutConstant.ConGpu;
            }

            if (arr.Exists(x => x.Equals(StrSoc, StringComparison.CurrentCultureIgnoreCase)))
            {
                return BinCutConstant.ConSoc;
            }

            if (arr.Exists(x => x.Equals("DDR", StringComparison.CurrentCultureIgnoreCase)))
            {
                return BinCutConstant.ConDdr;
            }

            if (arr.Exists(x => x.Equals("SC", StringComparison.CurrentCultureIgnoreCase)))
            {
                return BinCutConstant.ConCpu;
            }

            if (arr.Exists(x => Regex.IsMatch(x, PerformanceModeMatchPattern, RegexOptions.IgnoreCase)))
            {
                string mode = Regex.Match(keyword, PerformanceModeMatchPattern, RegexOptions.IgnoreCase).Groups["str"].ToString();
                if (mode.Substring(1, 1).Equals("C", StringComparison.CurrentCultureIgnoreCase))
                {
                    return BinCutConstant.ConCpu;
                }

                if (mode.Substring(1, 1).Equals("G", StringComparison.CurrentCultureIgnoreCase))
                {
                    return BinCutConstant.ConGpu;
                }

                if (mode.Substring(1, 1).Equals("S", StringComparison.CurrentCultureIgnoreCase))
                {
                    return BinCutConstant.ConSoc;
                }

                if (mode.Substring(1, 1).Equals("D", StringComparison.CurrentCultureIgnoreCase))
                {
                    return BinCutConstant.ConDdr;
                }
            }
            return "";
        }

        private string GetBinCutKey(string[] spt)
        {
            // MSX001 SOC TD: SocTd    MSX002 SOC BIST: SocMbist
            // MSX001 GPU TD: GfxTd    MSX001 GPU BIST: GfxMbist
            // MSX001 PCPU TD, MSX001 ECPU TD: CpuTd
            // SOC TD: SocTd    SOC BIST: SocMbist
            // GPU TD: GfxTd    GPU BIST: GpuMbist
            string keyword = "";
            if (spt[0].Trim().ToUpper().StartsWith("M"))
            {
                if (spt.Length == 2) //no domain information case
                {
                    keyword = spt[1].Trim().ContainsIgnoreCase("BIST") ? "MBIST" : spt[1].Trim();
                }
                else
                {
                    keyword = spt[1].Trim().ContainsIgnoreCase("CPU") ? "CPU" : spt[1].Trim().ContainsIgnoreCase("GPU") ? "GFX" : spt[1].Trim();
                    keyword += spt[2].Trim().ContainsIgnoreCase("BIST") ? "MBIST" : spt[2].Trim();
                }
            }
            else if (spt.Length == 2)
            {
                keyword = spt[0].Trim().ContainsIgnoreCase("CPU") ? "CPU" : spt[0].Trim().ContainsIgnoreCase("GPU") ? "GFX" : spt[0].Trim();
                keyword += spt[1].Trim().ContainsIgnoreCase("BIST") ? "MBIST" : spt[1].Trim();
            }
            return keyword;
        }
        #endregion

        private static bool IsNonLvBvInstance(ValidateData vd)
        {
            return vd.InstanceName.EndsWith("_BV") && !vd.Item.AllOther.Contains("LV");
        }

        private static bool IsLvHbvInstance(ValidateData vd)
        {
            return vd.InstanceName.EndsWith("_HBV") && vd.Item.AllOther.Contains("LV");
        }

        private static bool IsTempRelated(ValidateData vd)
        {
            return Regex.IsMatch(vd.Item.PerformanceMode, "TMPS", RegexOptions.IgnoreCase)
                || Regex.IsMatch(vd.Item.ColumnContent, TempSensor, RegexOptions.IgnoreCase);
        }

        private static bool IsElbRelated(ValidateData vd)
        {
            return vd.Item.ColumnContent.EndsWith("ELB", StringComparison.OrdinalIgnoreCase)
                || vd.Item.ColumnContent.EndsWith("ILB", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsPerformanceModeInstance(ValidateData vd)
        {
            return vd.InstanceName.StartsWith(vd.Item.PerformanceMode, StringComparison.OrdinalIgnoreCase)
                && (
                    (vd.Item.ColumnName != EnumColumnName.Mbist && vd.Item.ColumnName != EnumColumnName.FUNC)
                    || (vd.Item.ColumnName == EnumColumnName.Mbist && vd.InstanceName.ContainsIgnoreCase("mbist"))
                );
        }

        private static bool IsFuncColumn(ValidateData vd)
        {
            return vd.Item.ColumnName == EnumColumnName.FUNC;
        }

        private static bool? EvaluateTempRelated(ValidateData vd)
        {
            string type = vd.Item.GetBinType();
            string tmpsInstanceName = vd.Item.TargetPerformanceMode + "_TMPS_";
            if (vd.InstanceName.StartsWith(tmpsInstanceName, StringComparison.OrdinalIgnoreCase) &&
                vd.InstanceName.EndsWith(type, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return false;
        }

        private static bool? EvaluateElbRelated(ValidateData vd)
        {
            if (vd.InstanceName.ContainsIgnoreCase(vd.Item.TargetPerformanceMode.ToLower()) &&
                        (vd.InstanceName.ContainsIgnoreCase("_ELB_")
                         || vd.InstanceName.ContainsIgnoreCase("_ILB_")))
            {
                return true;
            }
            return null;
        }

        private static bool? EvaluatePerformanceModeInstance(ValidateData vd)
        {
            string instanceName = vd.InstanceName;
            const string regexPattern = "[(](?<str>.+)[)]";
            string[] arr = vd.Item.ColumnContent.Split(';');
            foreach (string value in arr)
            {
                string mode = value.Split(' ').First();
                string[] spt = value.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                string block = vd.Item.GetBinCutKey(spt).ToLower();

                if (!instanceName.ContainsIgnoreCase(mode.ToLower()))
                {
                    continue;
                }

                if (Regex.IsMatch(value, regexPattern, RegexOptions.IgnoreCase))
                {
                    List<string> wordList = Regex.Match(value, regexPattern, RegexOptions.IgnoreCase)
                        .Groups["str"]
                        .ToString()
                        .ToUpper()
                        .Split(',')
                        .ToList();

                    foreach (string word in wordList)
                    {
                        string trimWord = word.Trim();
                        string[] keywordArray = trimWord.Split('_');
                        if (trimWord.ContainsIgnoreCase("except") || trimWord.ContainsIgnoreCase("exclude"))
                        {
                            if (keywordArray.All(p => !instanceName.ContainsIgnoreCase(p)))
                            {
                                return true;
                            }
                        }

                        if (keywordArray.All(instanceName.ContainsIgnoreCase))
                        {
                            return true;
                        }
                    }
                    return false;
                }

                if (block != "" && instanceName.ContainsIgnoreCase(block))
                {
                    return true;
                }

                return false;
            }

            return null;
        }

        private static bool? EvaluateFuncColumn(ValidateData vd)
        {
            if (vd.InstanceName.StartsWith(vd.Item.PerformanceMode, StringComparison.OrdinalIgnoreCase))
            {
                return vd.Item.IsTargetSpiPattern(vd.Item.ColumnContent, vd.InstanceName);
            }
            return null;
        }

        private static readonly (
            Func<ValidateData, bool> ShouldApply,
            Func<ValidateData, bool?> Evaluate
        )[] _targetInstanceRules =
        {
            (IsNonLvBvInstance, vd => false),
            (IsLvHbvInstance, vd => false),
            (IsTempRelated, EvaluateTempRelated),
            (IsElbRelated, EvaluateElbRelated),
            (IsPerformanceModeInstance, EvaluatePerformanceModeInstance),
            (IsFuncColumn, EvaluateFuncColumn)
        };

        private class ValidateData
        {
            public string InstanceName { get; set; }

            public BinCutSourceItem Item { get; set; }
        }
    }
}
