using System;
using System.Collections.Generic;
using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.Static;
using Automation.Utility.HardIP;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

namespace Automation.GenerateIgxl.HardIp.InputObject
{
    public class TestPlanSheet
    {
        public string SheetName { get; set; }
        public string SubBlock { get; set; }
        public List<PatternRow> PatternRows { get; set; } = new List<PatternRow>();
        public List<HardIpPattern> PatternItems { get; set; } = new List<HardIpPattern>();

        public string ForceStr { get; set; }
        public int ForceIndex { get; set; }
        public Dictionary<string, int> PlanHeaderIdx { get; set; } = new Dictionary<string, int>();

        public void ConvertRealPatternName(ScghData scghData)
        {
            //const string errorMsgWrongPat = "Patterns that not start with “dd_”, ”cz_” or “pp_” will be ignored.";
            foreach (PatternRow row in PatternRows)
            {
                if (HardIpConstData.RegInsInPatt.IsMatch(row.Pattern.RealPatternName) ||
                    HardIpConstData.RegOpcodeInPatt.IsMatch(row.Pattern.RealPatternName))
                {
                    continue;
                }
                row.Pattern.ConvertRealPatternName(scghData);
                foreach (string pat in row.Pattern.PatternSetList.SelectMany(x => x))
                {
                    if (!SearchInfo.IsValidPatName(pat))
                    {
                        ErrorReportManager.AddError(
                            HardIpErrorType.W_WrongPatternName_01,
                            SheetName,
                            row.RowNum,
                            0,
                            []
                        );
                    }
                }
            }
        }

        public void DividePatternRow()
        {
            for (int i = 0; i < PatternRows.Count; i++)
            {
                PatternRow patRow = PatternRows[i];

                //Divide pattern row by MultiplePayload
                if (patRow != null)
                {
                    if (patRow.Pattern.IsMultiTimeDomain())
                    {
                        string oriPatName = patRow.Pattern.RealPatternName;
                        foreach (string pats in oriPatName.Split('#'))
                        {
                            PatternRow splitPatRow = patRow.Copy();
                            splitPatRow.Pattern = new PatternClass(pats) { RealPatternName = oriPatName };
                            PatternRows.Insert(i, splitPatRow);
                            i++;
                        }
                    }
                    else
                    {
                        string oriPatName = patRow.Pattern.RealPatternName;
                        string[] arr = oriPatName.Split(',');
                        for (int a = 1; a < arr.Length; a++)
                        {
                            PatternRow secondPatRow = patRow.Copy();
                            secondPatRow.Pattern = new PatternClass(arr[a]);
                            i++;
                            PatternRows.Insert(i, secondPatRow);
                        }

                        patRow.Pattern = new PatternClass(arr[0]);
                    }
                }

                // Set for CZ2only
                bool cz2Only = HardIpService.IsCz2Only(patRow.MiscInfo);
                if (cz2Only)
                {
                    patRow.ForceCondition.IsShmooInProdInst = false;
                    patRow.ForceCondition.IsShmooInProdFlow = false;
                    patRow.ForceCondition.IsShmooInCharInst = true;
                    patRow.ForceCondition.IsShmooInCharFlow = true;
                }

                // Divide pattern row by shmoo
                patRow.ForceCondition.IsShmooInForce = IsShmooInForce(patRow);
                // VT shmoo for another flow sheet
                patRow.ForceCondition.IsVtShmoo = IsVtShmooInForce(patRow);
                if (patRow.ForceCondition.IsShmooInForce)
                {
                    if (cz2Only)
                    {
                        patRow.ForceCondition.IsShmooInProdInst = false;
                        patRow.ForceCondition.IsShmooInProdFlow = false;
                        patRow.ForceCondition.IsShmooInCharInst = true;
                        patRow.ForceCondition.IsShmooInCharFlow = true;
                    }
                    else
                    {
                        patRow.ForceCondition.IsShmooInProdInst = true;
                        patRow.ForceCondition.IsShmooInProdFlow = true;
                        patRow.ForceCondition.IsShmooInCharInst = false;
                        patRow.ForceCondition.IsShmooInCharFlow = true;
                    }
                }
            }
        }

        private bool IsShmooInForce(PatternRow patRow)
        {
            foreach (string condition in patRow.ForceCondition.ForceCondition.Split(';'))
            {
                if (HardIpConstData.RegShmoo.IsMatch(condition))
                {
                    return true;
                }
            }
            return false;
        }

        private bool IsVtShmooInForce(PatternRow patRow)
        {
            foreach (string condition in patRow.ForceCondition.ForceCondition.Split(';'))
            {
                if (HardIpConstData.RegVtShmoo.IsMatch(condition))
                {
                    return true;
                }
            }
            return false;
        }

        public void UpdateRelay()
        {
            List<HardIpPattern> patterns = PatternItems;

            int? lastCzIdx = null;
            int? lastNonCzIdx = null;

            List<string> typeForCz =
                LocalSpecs.Options.CharacterizationType.Split(',').ToList();

            for (int i = 0; i < patterns.Count; i++)
            {
                var lastPatRelaySetting = new Dictionary<string, string>();

                bool isCz =
                    typeForCz.Exists(x => patterns[i].Pattern.GetLastPayload().StartsWith(x, StringComparison.CurrentCultureIgnoreCase))
                    || patterns[i].Pattern.InstancePatternName.All(name => typeForCz.Any(prefix => name.StartsWith(prefix, StringComparison.CurrentCultureIgnoreCase)))
                    || patterns[i].MiscInfoDict.ContainsKey(CommonConst.Cz2Only);

                if (isCz)
                {
                    if (lastCzIdx.HasValue)
                    {
                        lastPatRelaySetting =
                            patterns[lastCzIdx.Value].NewRelaySetting;
                    }

                    lastCzIdx = i;
                }
                else
                {
                    if (lastNonCzIdx.HasValue)
                    {
                        lastPatRelaySetting =
                            patterns[lastNonCzIdx.Value].NewRelaySetting;
                    }

                    lastNonCzIdx = i;
                }

                GetRelaySetting(patterns[i], lastPatRelaySetting);
            }
        }

        public Dictionary<string, string> GetRelaySetting(HardIpPattern pattern, Dictionary<string, string> exSetting)
        {
            var oriRelaySetting = new Dictionary<string, string>();
            var newRelaySetting = new Dictionary<string, string>();
            var actualRelaySetting = new Dictionary<string, string>();
            List<HardIpRelay> settingInMisc = GetsRelaysInMisc(pattern.MiscInfo);
            IEnumerable<IGrouping<string, HardIpRelay>> relayGroups = from item in settingInMisc group item by item.Job into g select g;
            foreach (IGrouping<string, HardIpRelay> relayGroup in relayGroups)
            {
                string job = relayGroup.Key;
                string relayOnSetting = string.Join("&", relayGroup.ToList().Where(s => s.Status == RelayStatus.On).Select(s => s.Name.Replace(",", "&")).ToList());
                string relayOffSetting = string.Join("&", relayGroup.ToList().Where(s => s.Status == RelayStatus.Off).Select(s => s.Name.Replace(",", "&")).ToList());
                if (!string.IsNullOrEmpty(relayOnSetting))
                {
                    relayOnSetting = "RelayOn:" + relayOnSetting;
                }

                if (!string.IsNullOrEmpty(relayOffSetting))
                {
                    relayOffSetting = "RelayOff:" + relayOffSetting;
                }

                oriRelaySetting.Add(job, (relayOnSetting + "_" + relayOffSetting).Trim('_'));
                newRelaySetting.Add(job, (relayOnSetting + "_" + relayOffSetting).Trim('_'));
            }
            pattern.NewRelaySetting = newRelaySetting;
            var jobs = newRelaySetting.Keys.ToList();
            jobs.AddRange(exSetting.Keys);
            jobs = jobs.Distinct().ToList();
            foreach (string job in jobs)
            {
                exSetting.TryGetValue(job, out string exSettingInJob);
                oriRelaySetting.TryGetValue(job, out string newSettingInJob);
                if (string.IsNullOrEmpty(exSettingInJob))
                {
                    actualRelaySetting.Add(job, newSettingInJob);
                    continue;
                }
                if (exSettingInJob == newSettingInJob)
                {
                    continue;
                }
                actualRelaySetting.Add(job, (HardIpService.ReverseRelaySetting(exSettingInJob) + ";" + newSettingInJob).Trim(';'));
            }
            pattern.RelaySetting = actualRelaySetting;
            pattern.NewRelaySetting = newRelaySetting;
            return newRelaySetting;
        }

        public List<HardIpRelay> GetsRelaysInMisc(string miscInfo)
        {
            var settingInMisc = new List<HardIpRelay>();
            foreach (string param in miscInfo.Split(';'))
            {
                string text = param.Trim();
                if (!(text.Replace("_", "").StartsWith($"{HardIpConstData.RelayOn.Replace("_", "")}:", StringComparison.CurrentCultureIgnoreCase)
                    || text.Replace("_", "").StartsWith($"{HardIpConstData.RelayOff.Replace("_", "")}:", StringComparison.CurrentCultureIgnoreCase)))
                {
                    continue;
                }

                string[] relayArr = param.Split(':');
                RelayStatus status = relayArr[0].Replace("_", "").Equals("relayon", StringComparison.CurrentCultureIgnoreCase) ? RelayStatus.On : RelayStatus.Off;
                string job = param.Split(':').Length == 3 ? param.Split(':')[2].ToUpper() : "";
                var relay = new HardIpRelay(job, relayArr[1], status);
                settingInMisc.Add(relay);
            }
            return settingInMisc;
        }
    }
}
