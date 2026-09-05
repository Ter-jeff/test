using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Singleton;
using Automation.Static;

using CommonLib.Extension;
using CommonLib.Utility;

using IgxlLib.IgxlBase;

using TestPlanLib.Basic;
using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.Utility;

namespace Automation.GenerateIgxl.BinCut.Base
{
    public class BinCutFinalInstanceRow
    {
        public BinCutInstanceRow BinCutInstanceRow = new BinCutInstanceRow();
        public bool IsDuplicateName;
        public string FinalInstName;
        public bool JobNotMap;
        public bool IsUsed;
        public string PerformanceMode = "";
        public List<string> InitList = new List<string>();
        public List<string> PayloadList = new List<string>();
        public List<string> PatternList = new List<string>();
        public string VbtFunction = "";
        public string Enable = "";
        public string Module = "";
        public string VddPin = "";
        public bool Nop;

        public bool NopByEnableWord
        {
            get
            {
                return BinCutInstanceRow.EnableFlow.Equals("NOP", StringComparison.CurrentCultureIgnoreCase);
            }
        }

        public bool CanBeBurst
        {
            get
            {
                JudgeBurst();
                return BinCutInstanceRowUtility.IsBist(BinCutInstanceRow.FlowName) ? _plBurst && _inBurst : _allBurst;
            }
        }

        private bool _plBurst;
        private bool _inBurst;
        private bool _allBurst;
        public string InitPatSetNameNew { get; set; }
        public string InitPatSetNameNewTemp;
        public string PatSetName { get; set; } = ""; //with RowNum
        public string PatSetNameTemp = "";          //Without RowNum

        public void JudgeBurst()
        {
            if (BinCutInstanceRow.Burst.ToUpper().Equals("YES"))
            {
                _plBurst = true;
                _inBurst = true;
                _allBurst = true;
            }
            else if (BinCutInstanceRow.Burst.ToUpper().Equals("NO"))
            {
                _plBurst = false;
                _inBurst = false;
                _allBurst = false;
            }
        }

        public PatSet PatSet
        {
            get
            {
                var patSet = new PatSet();
                string commentPatSet = BinCutInstanceRow.FlowName + " , " + BinCutInstanceRow.SheetName + ": RowNum" + BinCutInstanceRow.RowNum;
                if (BinCutInstanceRowUtility.IsBist(BinCutInstanceRow.FlowName) || !string.IsNullOrEmpty(BinCutInstanceRow.Char))
                {
                    if (PayloadList.Any())
                    {
                        patSet.PatSetName = PatSetName;
                        foreach (string pattern in PayloadList)
                        {
                            if (!string.IsNullOrEmpty(pattern))
                            {
                                patSet.AddRow(FillPatSetRow(pattern, _plBurst, commentPatSet));
                            }
                        }
                    }
                    else
                    {
                        return null;
                    }
                }
                else
                {
                    patSet.PatSetName = PatSetName;
                    foreach (string pattern in PatternList)
                    {
                        if (!string.IsNullOrEmpty(pattern))
                        {
                            patSet.AddRow(FillPatSetRow(pattern, _allBurst, commentPatSet));
                        }
                    }
                }
                return patSet;
            }
        }

        public string InitPatSetName
        {
            get
            {
                if (InitPatSet != null)
                {
                    return InitPatSet.PatSetName;
                }

                return "";
            }
        }
        public PatSet InitPatSet
        {
            get
            {
                if (BinCutInstanceRowUtility.IsBist(BinCutInstanceRow.FlowName) || !string.IsNullOrEmpty(BinCutInstanceRow.Char))
                {
                    var prePatSet = new PatSet();
                    string commentPatSet = BinCutInstanceRow.FlowName + " , " + BinCutInstanceRow.SheetName + ": RowNum" + BinCutInstanceRow.RowNum;
                    if (PatSetName.Contains("RowNum"))
                    {
                        List<string> arr = PatSetName.Split('_').ToList();
                        prePatSet.PatSetName = string.Join("_", arr);
                    }
                    else
                    {
                        prePatSet.PatSetName = PatSetName; //+ "_INIT";
                    }

                    if (!string.IsNullOrEmpty(InitPatSetNameNew))
                    {
                        prePatSet.PatSetName = InitPatSetNameNew;
                    }

                    foreach (string init in InitList)
                    {
                        prePatSet.AddRow(FillPatSetRow(init, _inBurst, commentPatSet)); //commentPatSet+ "_INIT"
                    }

                    return prePatSet;
                }
                return null;
            }
        }

        public List<string> FinalJobs = new List<string>();

        public string Domain { get; set; } = string.Empty;
        public string Block { get; set; }
        public string ModeByFlowName { get; set; }
        public string InitName { get; set; }
        public string PayloadName { get; set; }

        public bool IsEvs2()
        {
            return !string.IsNullOrEmpty(BinCutInstanceRow.SiteVar) &&
                   !BinCutInstanceRow.BinOutStage.Equals("X", StringComparison.CurrentCultureIgnoreCase);
        }

        public string GetEnableWord()
        {
            foreach (string pattern in PayloadList)
            {
                if (pattern.ContainsIgnoreCase("_BIR_"))
                {
                    return "BIR";
                }
            }
            return "";
        }

        public object Clone(string job)
        {
            var row = new BinCutFinalInstanceRow
            {
                Enable = Enable,
                PerformanceMode = PerformanceMode,
                VbtFunction = VbtFunction,
                Module = Module,
                VddPin = VddPin,
                Nop = Nop,
                PatSetName = PatSetName,
                PatternList = PatternList,
                PayloadList = PayloadList,
                InitList = InitList,
                PatSetNameTemp = PatSetNameTemp,
                InitPatSetNameNew = InitPatSetNameNew,
                InitPatSetNameNewTemp = InitPatSetNameNewTemp,
                Domain = Domain,
                BinCutInstanceRow = BinCutInstanceRow,
                IsDuplicateName = IsDuplicateName,
                FinalInstName = FinalInstName,
                JobNotMap = JobNotMap,
                IsUsed = IsUsed
            };

            row.FinalJobs.Add(job);
            return row;
        }

        public string GetSubFlowName()
        {
            string sheetName = BinCutInstanceRow.FlowName.ToUpper();
            if (!sheetName.StartsWith("Flow_", StringComparison.CurrentCultureIgnoreCase))
            {
                sheetName = "Flow_" + sheetName;
            }
            return sheetName;
        }

        public string GetBlockByFlowName()
        {
            string block;
            if (BinCutInstanceRow.FlowName.ContainsIgnoreCase("TD"))
            {
                block = "Td";
            }
            else if (BinCutInstanceRow.FlowName.ContainsIgnoreCase("BIST") || BinCutInstanceRow.FlowName.ContainsIgnoreCase("BIRA"))
            {
                block = "Mbist";
            }
            else if (Regex.IsMatch(BinCutInstanceRow.FlowName, "ELB|ILB", RegexOptions.IgnoreCase))
            {
                block = "DDR";
            }
            else if (Regex.IsMatch(BinCutInstanceRow.FlowName, "CPM", RegexOptions.IgnoreCase))
            {
                block = "Td";
            }
            else if (Regex.IsMatch(BinCutInstanceRow.FlowName, "SCAN", RegexOptions.IgnoreCase))
            {
                block = "Scan";
            }
            else
            {
                block = "Spi";
            }
            return block;
        }

        public string GetVoltageType()
        {
            string typeByDc = BinCutInstanceRowUtility.GetTypeByFlowNameOrDcCategory(BinCutInstanceRow.DCcategory);
            if (!typeByDc.Equals("UnknowType"))
            {
                return typeByDc;
            }

            string typeByFlow = BinCutInstanceRowUtility.GetTypeByFlowNameOrDcCategory(BinCutInstanceRow.FlowName);
            if (!typeByFlow.Equals("UnknowType"))
            {
                return typeByFlow;
            }

            return "UnknowType";
        }

        public string GetTestInstanceFailFlag(string flowParameter)
        {
            return string.IsNullOrEmpty(flowParameter) ? $"F_{GetParameter()}" : $"F_{flowParameter}";
        }

        public string GetParameter()
        {
            var texts = new List<string> { PatSetName };
            if (IsDuplicateName)
            {
                texts.Add("RowNum" + BinCutInstanceRow.RowNum);
            }
            texts.Add(GetVoltageType());
            return string.Join("_", texts.Where(x => !string.IsNullOrEmpty(x)));
        }

        public string GetEvsParameter()
        {
            var texts = new List<string>();
            if (!PatSetName.StartsWith("EVS"))
            {
                texts.Add("EVS");
            }

            texts.Add(PatSetName);
            if (!string.IsNullOrEmpty(BinCutInstanceRow.Instance))
            {
                texts.Add(BinCutInstanceRow.Instance);
            }

            if (IsDuplicateName)
            {
                texts.Add(BinCutInstanceRow.SheetName);
                texts.Add("RowNum" + BinCutInstanceRow.RowNum);
            }

            texts.Add(GetVoltageType());
            return string.Join("_", texts.Where(x => !string.IsNullOrEmpty(x)));
        }

        public string GetEvsRampPowerName()
        {
            var texts = new List<string>();
            string testName = BinCutInstanceRow.FlowName.Replace("Flow_EVS_", "").Replace("_HV", "");
            testName = testName.StartsWith("Flow") ? testName.Replace("Flow_", "") : testName;
            texts.Add("EVS_Static_Power_Ramp");
            if (Domain.ToLower().Equals("cpu") && string.IsNullOrEmpty(BinCutInstanceRow.PatSetNameOrange))
            {
                texts.Add(testName.Split('_')[0]);
            }
            texts.Add("Multi");
            texts.Add(!string.IsNullOrEmpty(BinCutInstanceRow.PatSetNameOrange) ? BinCutInstanceRow.PatSetNameOrange : testName);
            if (!string.IsNullOrEmpty(BinCutInstanceRow.Instance))
            {
                texts.Add(BinCutInstanceRow.Instance);
            }

            if (IsDuplicateName)
            {
                texts.Add(BinCutInstanceRow.SheetName);
                texts.Add("RowNum" + BinCutInstanceRow.RowNum);
            }
            return string.Join("_", texts.Where(x => !string.IsNullOrEmpty(x)));
        }

        public string GetJob()
        {
            if (BinCutInstanceRow != null)
            {
                return string.Join(",", BinCutInstanceRowUtility.GetJobs(BinCutInstanceRow.JobTestStage));
            }

            return "";
        }

        public string GetEnable()
        {
            if (BinCutInstanceRow != null)
            {
                return BinCutInstanceRow.EnableFlow;
            }

            return "";
        }

        public bool IsSsn()
        {
            if (TestPlanStatic.MappingCoreTable == null)
            {
                return false;
            }

            if (!TestPlanStatic.MappingCoreTable.Rows.Exists(x => x.IsSsn()))
            {
                return false;
            }

            List<string> ssnPatList = PatternList.FindAll(x => x.Split('_').ToList().Exists(y => Regex.IsMatch(y, "^(SSC|SSU)$", RegexOptions.IgnoreCase)));
            bool foundSsn = false;
            foreach (string pat in ssnPatList)
            {
                if (!Regex.IsMatch(pat, @"_PL\w{2}_", RegexOptions.IgnoreCase) || Regex.IsMatch(pat, "_(LPB|CON)", RegexOptions.IgnoreCase))
                {
                    continue;
                }
                if (TestPlanStatic.MappingCoreTable.RegexPlList.Exists(x => x.IsMatch(pat.ToLower())))
                {
                    foundSsn = true;
                }
            }
            return foundSsn;
        }

        public List<string> GetInstanceFailFlags()
        {
            var flags = new List<string>();

            if (!string.IsNullOrEmpty(BinCutInstanceRow.FailFlag) && !BinCutInstanceRow.FailFlag.Equals("X", StringComparison.OrdinalIgnoreCase))
            {
                IEnumerable<string> failFlags = BinCutInstanceRow.FailFlag.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.TrimStart('!'));
                flags.AddRange(failFlags);
            }

            return flags;
        }

        public List<string> GetPinGroupFailFlags()
        {
            var flags = new List<string>();
            if (BinCutInstanceRow.IsHarvesting.Equals("TRUE"))
            {
                if (!string.IsNullOrEmpty(BinCutInstanceRow.PinGroupBinoutFlag))
                {
                    string[] binOutFlags = BinCutInstanceRow.PinGroupBinoutFlag.Split(new[] { ',', ' ', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    flags.AddRange(binOutFlags);
                }
            }

            return flags;
        }

        public List<string> GetDefaultFailFlags(string payloadType, string mode)
        {
            var flags = new List<string>();

            string defaultFlag = GetDefaultBinTable(payloadType, mode);
            if (!string.IsNullOrEmpty(defaultFlag))
            {
                flags.Add(defaultFlag);
            }

            return flags;
        }

        public List<string> GetBinOutFlags(string payloadType, string mode)
        {
            List<string> binTableItemList = GetDefaultFailFlags(payloadType, mode);
            if (GetInstanceFailFlags().Any() && GetPinGroupFailFlags().Any())
            {
                binTableItemList = GetInstanceFailFlags();
            }
            else if (GetInstanceFailFlags().Any())
            {
                binTableItemList = GetInstanceFailFlags();
            }
            else if (GetPinGroupFailFlags().Any())
            {
                binTableItemList = GetPinGroupFailFlags();
            }
            return binTableItemList;
        }

        public List<List<string>> GetFlagWithPinFail(string instanceFailFlag)
        {
            var flagSelection = new List<List<string>>
            {
                new List<string> { instanceFailFlag }
            };

            if (JudgePinAndInstBinOut()) //combination
            {
                flagSelection = new List<List<string>>();
                GetPinGroupFailFlags().ForEach(x => flagSelection.Add(new List<string> { x, instanceFailFlag }));
            }

            return flagSelection;
        }

        public bool JudgePinAndInstBinOut()
        {
            return GetInstanceFailFlags().Any() && GetPinGroupFailFlags().Any();
        }

        public IEnumerable<string> GetFlowFailAction(string payloadType, string mode, string flowParameter = "")
        {
            var flags = new List<string>();

            foreach (string item in GetInstanceFailFlags())
            {
                flags.AddRange(item.Replace("&&", ",").Replace("||", ",").Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries));
            }

            if (!flags.Any())
            {
                flags.AddRange(GetDefaultFailFlags(payloadType, mode));
            }
            return flags.Distinct().Append(GetTestInstanceFailFlag(flowParameter));
        }

        public bool IsBurstNonBinCutInstance()
        {
            return CanBeBurst;
        }

        private string GetCFandCsTypeByPayload()
        {
            string gfxType = "";

            foreach (string pattern in PayloadList)
            {
                List<string> arr = pattern.Split('_').ToList();
                if (arr.Count > 5)
                {
                    if (Regex.IsMatch(arr[5], "CF", RegexOptions.IgnoreCase))
                    {
                        gfxType = "CFXX";
                    }
                    else if (Regex.IsMatch(arr[5], "CS", RegexOptions.IgnoreCase))
                    {
                        gfxType = "CSXX";
                    }
                }
                if (!string.IsNullOrEmpty(gfxType))
                {
                    return gfxType;
                }
            }
            return "";
        }

        public string GetDefaultBinTable(string payloadType, string mode)
        {
            if (BinCutInstanceRow != null && (BinCutInstanceRow.IsHarvesting == "TRUE" || !string.IsNullOrEmpty(BinCutInstanceRow.FailFlag)))
            {
                return null;
            }

            return GenFlag(payloadType, mode);
        }

        public string GenFlag(string payloadType, string mode)
        {
            string flowType = GetVoltageType();
            var flagList = new List<string> { "F", Domain, // Cpu
                payloadType // SaChain
            };
            if (BinCutInstanceRow != null &&
                Domain.Equals("GFX", StringComparison.OrdinalIgnoreCase) &&
                Regex.IsMatch(BinCutInstanceRow.FlowName, "_SA_", RegexOptions.IgnoreCase))
            {
                string type = GetCFandCsTypeByPayload();
                if (!string.IsNullOrEmpty(type))
                {
                    flagList.Add(type);
                }
            }
            flagList.Add(mode); //MC701
            flagList.Add(flowType);
            return Combination.CombineByUnderLine(flagList);
        }

        public string GetDcCategory()
        {
            string dcCategory = Regex.Replace(BinCutInstanceRow.DCcategory, @"\s*_*LV$", "", RegexOptions.IgnoreCase);
            dcCategory = Regex.Replace(dcCategory, @"\s*_*NV$", "", RegexOptions.IgnoreCase);
            dcCategory = Regex.Replace(dcCategory, @"\s*_*HV$", "", RegexOptions.IgnoreCase);
            return dcCategory;
        }

        public string GetTimeSetVersion(List<string> patternNames)
        {
            if (BinCutInstanceRow != null)
            {
                if (!string.IsNullOrEmpty(BinCutInstanceRow.TimeSet))
                {
                    string timeSet = BinCutInstanceRow.TimeSet.Contains(":")
                        ? BinCutInstanceRow.TimeSet.Split(':').First()
                        : BinCutInstanceRow.TimeSet;
                    string specificTimeSet = Path.GetFileNameWithoutExtension(timeSet);
                    string timeSetVersion = AcTSetCategoryMapSingleton.Instance().GetTimeSetVersion(specificTimeSet);
                    if (string.IsNullOrEmpty(timeSetVersion))
                    {
                        return specificTimeSet;
                    }

                    return timeSetVersion;
                }
            }

            var timeSets = new List<string>();
            foreach (string patternName in patternNames)
            {
                if (AcTSetCategoryMapSingleton.Instance().PatternList.TryGetValue(patternName.ToLower(), out PatternData patternData))
                {
                    timeSets.Add(patternData.TimeSetVersion);
                }
            }

            return string.Join(",", timeSets.Distinct().ToList());
        }

        public string GetBinOutStage()
        {
            if (BinCutInstanceRow != null)
            {
                return BinCutInstanceRow.BinOutStage;
            }

            return "";
        }

        private PatSetRow FillPatSetRow(string pattern, bool isBurst, string comment = "")
        {
            var resultRow = new PatSetRow { File = pattern, Burst = isBurst ? "Yes" : "No", Comment = comment };
            return resultRow;
        }

        public string GetPatSetNameForArgument()
        {
            return !string.IsNullOrEmpty(PatSetName) ? PatSetName : "";
        }
    }
}
