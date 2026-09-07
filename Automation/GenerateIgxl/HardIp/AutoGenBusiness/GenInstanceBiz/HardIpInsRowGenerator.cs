using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.Singleton;
using Automation.Static;
using Automation.Utility.HardIP;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using IgxlLib.IgxlBase;

using LogLib.Static;

using TestPlanLib.Singleton;
using TestPlanLib.Static;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenInstanceBiz
{
    public class HardIpInsRowGenerator : InsRowGenerator
    {
        public HardIpInsRowGenerator(HardIpInputData hardIpInputData, HardIpSheet hardIpSheet, string sheetName) : base(hardIpInputData, hardIpSheet, sheetName)
        {
        }

        /// <summary>
        /// Generate instance rows for one pattern, if contains ReTest, two instances will be generated, else only one instance will be generated
        /// </summary>
        /// <returns></returns>
        public override List<InstanceRow> GenInsRows()
        {
            var insRowlist = new List<InstanceRow>();

            #region insert fail flag

            if (Regex.IsMatch(Pattern.MiscInfo, HardIpConstData.MappingTtrBranch, RegexOptions.IgnoreCase))
            {
                string mapTtrKey = HardIpService.GetMappingTtrBranchItem(Pattern.MiscInfo);
                if (InitOriginalItems.TryGetValue(mapTtrKey, out string item))
                {
                    Pattern.HipPreWriteFlag = item;
                }
            }
            else
            {
                Pattern.HipPreWriteFlag =
                VbtFunctionLibShared.EfusePrewriteFunctionList.Exists(f => f.Equals(Pattern.FunctionName, StringComparison.CurrentCultureIgnoreCase))
                || VbtFunctionLibShared.EfuseReadFunctionList.Exists(f => f.Equals(Pattern.FunctionName, StringComparison.CurrentCultureIgnoreCase))
                ? LastPatFailAction
                : CreateTestFailAction();
            }

            Pattern.HipPreWriteFlag = Regex.Replace(Pattern.HipPreWriteFlag, "_Multiple", "", RegexOptions.IgnoreCase);

            if (Regex.IsMatch(Pattern.MiscInfo, HardIpConstData.OriginalTtrBranch, RegexOptions.IgnoreCase))
            {
                string oriTtrKey = HardIpService.GetOriginalTtrBranchItem(Pattern.MiscInfo);
                if (!InitOriginalItems.ContainsKey(oriTtrKey))
                {
                    InitOriginalItems.Add(oriTtrKey, Pattern.HipPreWriteFlag);
                }
            }

            if (Regex.IsMatch(Pattern.MiscInfo, HardIpConstData.FuseStage + @"\:"))
            {
                List<string> miscList = Pattern.MiscInfo.Split(';').ToList();
                foreach (string cmd in miscList)
                {
                    if (cmd.StartsWith(HardIpConstData.FuseStage))
                    {
                        string voltageLabel = cmd.Split(':')[1].ToUpper()[0].ToString();
                        List<string> flag = LastPatFailAction.Split('_').ToList();
                        flag[flag.Count - 2] = voltageLabel;
                        Pattern.HipPreWriteFlag = string.Join("_", flag);
                        break;
                    }
                }
            }
            else
            {
                LastPatFailAction = Pattern.HipPreWriteFlag;
            }

            #endregion

            #region basic row

            InstanceRow insRow = GenHardIpRow();
            insRowlist.Add(insRow);

            if (Pattern.ForceCondition.IsVtShmoo)
            {
                insRowlist.Add(GenHardIpRow(true));
            }
            #endregion

            #region reTest row

            if (Pat.MiscInfoDict.ContainsKey(HardIpConstData.ReTest))
            {
                InstanceRow reTestRow = GenReTestRow(insRow);
                insRowlist.Add(reTestRow);
            }

            #endregion
            #region adaptive cooling 
            if (LabelVoltage.Equals(HardIpConstData.LabelNv, StringComparison.CurrentCultureIgnoreCase) &&
                HardIpInputData.ConfigData.AdaptiveCoolingItem.Contains($"{BlockName}_{SubBlockName}"))
            {
                SubBlockName += "MON";
                insRow = GenHardIpRow();
                insRowlist.Add(insRow);
            }
            #endregion

            insRowlist.ForEach(ins => ins.RowNum = Pat.RowNum);
            return insRowlist;
        }

        /// <summary>
        /// Genrator Hardip instance row(Except IDS)
        /// </summary>
        /// <returns></returns>
        protected InstanceRow GenHardIpRow(bool isVtShmoo = false)
        {
            if (isVtShmoo)
            {
                Pattern.IsVtShmooUsed = true;
            }
            var insRow = new InstanceRow
            {
                SheetName = SheetName,
                TestName = CreateHardIpTestName(),
                VbtType = CreateType(),
                VbtName = CreateVbtName(),
                ArgList = CreateArgList(),
                Args = CreateArgs(),
                TimeSets = CreateHardIpTimeSets()
            };
            //Check the sequence exist DutyCycle(RF) or D(AP)
            if (insRow.GetArgument("TestSequence").ContainsIgnoreCase("DUTYCYCLE") || (insRow.GetArgument("TestSequence").Split(',').ToList().Exists(x => x.Equals("D", StringComparison.CurrentCultureIgnoreCase)) && !string.IsNullOrEmpty(insRow.TimeSets)))
            {
                insRow.TimeSets += LocalSpecs.Options.Device == EnumDevice.AP ? ",JitterSheet" : ",CheckDuty_Jitter";
            }
            insRow.DcCategory = CreateHardIpDcCategory();
            insRow.DcSelector = CreateDcSelector();
            if (insRow.TimeSets != null)
            {
                insRow.AcCategory = CreateHardIpAcCategory(insRow.TimeSets.Split(',')[0]);
            }

            insRow.AcSelector = string.IsNullOrEmpty(insRow.AcCategory) ? "" : CreateAcSelector();
            if (TestPlanStatic.ConcurrentFlowSheet != null
                    && TestPlanStatic.ConcurrentFlowSheet.Rows.Exists(x => x.Subflows.Exists(y => y.ToUpper().Replace("FLOW_HARDIP_", "").Equals(SheetName.Replace("HARDIP_", "")))))
            {
                insRow.PinLevels = CreateHardIpLevelConcurrent(insRow.TimeSets);
            }
            else
            {
                insRow.PinLevels = insRow.PinLevels = !string.IsNullOrEmpty(Pat.LevelUsed) ? "Levels_" + Pat.LevelUsed : CreateHardIpPinLevel();
            }

            insRow.MixedSignalTiming = CreateHardIpMixSignalSheet();
            return insRow;
        }

        protected override void SetBasicInfoByPattern(HardIpPattern pattern)
        {
            BlockName = CommonGenerator.GetBlockNameFromSheetName(pattern.SheetName);
            SubBlockName = CommonGenerator.GetSubBlockName(pattern.Pattern.GetLastPayload(), pattern.MiscInfo, BlockName, pattern.ForceCondition.IsCz2InstName);
            TimingAc = CommonGenerator.GetTimingAc(pattern);
            InstNameSubStr = HardIpService.GetInstNameSubStr(pattern.MiscInfo);
            VbtFunction = CommonGenerator.GetVbtFunctionBase(pattern.FunctionName);
            NoPattern = HardIpService.IsNoPattern(pattern.Pattern.GetLastPayload());
            HardIpDcSetting = CommonGenerator.GetHardIpDcSetting(pattern.LevelUsed);
        }

        #region Generate each columns Methods
        /// <summary>
        /// generate TestName
        /// </summary>
        /// <returns></returns>
        protected string CreateHardIpTestName()
        {
            if (!string.IsNullOrEmpty(Pattern.TestName))
            {
                return Pattern.TestName + "_" + LabelVoltage;
            }

            string patternName = Pattern.Pattern.GetPatternName();
            return CommonGenerator.GenHardIpInsTestName(BlockName, Pattern.IsVtShmooUsed ? SubBlockName + "VTSHMOO" : SubBlockName, patternName, Pattern.PatternIndexFlag, TimingAc,
                     Pattern.ForceVoltageFlag, InstNameSubStr, LabelVoltage, NoPattern, Pat.WirelessData.IsNeedPostBurn, false, Pat.WirelessData.IsDoMeasure);
        }
        /// <summary>
        /// Create DcCategory:
        /// if exist in PatternSummary, use what in summary
        /// if specified Hardip DC, use what in HardIpDc
        /// if patternName contains MD[XX], use "HARDIP_MD[XX]
        /// else use default: HardIp
        /// </summary>
        /// <returns></returns>
        protected virtual string CreateHardIpDcCategory()
        {
            string dcCategory = ExceptionListSingleton.Instance().GetDcCategoryByInstance(CreateHardIpTestName());
            if (dcCategory != string.Empty)
            {
                return dcCategory;
            }

            if (Pattern.DcCategory != "")
            {
                string dcUsed = GetSpecifyInfo(Pattern.DcCategory, "DC");
                if (dcUsed != "")
                {
                    dcCategory = dcUsed;
                }
            }

            if (!string.IsNullOrEmpty(dcCategory))
            {
                if (!MultiTestSettingSheetsSingleton.Instance().DcCategoryInfos.Exists(s => s.CategoryName.Equals(dcCategory, StringComparison.OrdinalIgnoreCase)))
                {
                    Response.Report("Can not find HardIp category: " + dcCategory + " from TestSettting sheet", EnumMessageLevel.Error, 45);
                    ErrorReportManager.AddError(
                        HardIpErrorType.E_MissingHardIpCategory_01,
                        Pattern.SheetName,
                        Pattern.RowNum,
                        0,
                        [dcCategory]
                    );
                }
                return dcCategory;
            }

            string performanceMode = "";
            string hardipDcName = HardIpDcSetting != null && !string.IsNullOrEmpty(HardIpDcSetting.DcCategory) ? HardIpDcSetting.CategoryName : "";
            string[] patternSplitArr = Pattern.Pattern.GetLastPayload().Split('_');
            if (patternSplitArr.Length >= 10 && ModuleSingleton.GetModuleByPerformanceMode(patternSplitArr[9]) == ModuleSingleton.Instance().ModuleDdr)
            {
                performanceMode = patternSplitArr[9];
            }

            List<string> patterns = new List<string> { Pattern.Pattern.GetLastPayload() };

            dcCategory = MultiTestSettingSheetsSingleton.Instance().DcCategoryInfos.FindHardIpCatgeoryName(TestPlanStatic.PowerInfoSheet, performanceMode, hardipDcName, BlockName, patterns, Pattern.Chiplet, out EnumMessageLevel msgLevel, out string errorMsg);
            if (!string.IsNullOrEmpty(errorMsg))
            {
                Response.Report(errorMsg, msgLevel, 45);
                ErrorReportManager.AddError(
                    HardIpErrorType.E_MissingHardIpCategory_01,
                    Pattern.SheetName,
                    Pattern.RowNum,
                    0,
                    [dcCategory]
                );
            }
            return dcCategory;
        }

        /// <summary>
        /// Create PinLevel
        /// if exist in PatternSummary, use what in summary
        /// if specified Hardip DC, use what in HardIpDc
        /// </summary>
        /// <returns></returns>
        internal virtual string CreateHardIpPinLevel()
        {
            if (HardIpDcSetting != null)
            {
                return HardIpDcSetting.LevelSheet;
            }

            return HardIpConstData.LevelDefault;
        }

        internal string CreateHardIpLevelConcurrent(string timeSet)
        {
            if (AcTSetCategoryMapSingleton.Instance().MultiTimeSetSheets.Find(x => x.Name.Equals(timeSet)) == null)
            {
                return "TBD(ConcurrentLevelError)";
            }

            string timeDomain =
                AcTSetCategoryMapSingleton.Instance()
                    .MultiTimeSetSheets.Find(x => x.Name.Equals(timeSet))
                    .TimeDomain;
            return "Levels_Con_H_" + timeDomain;
        }

        /// <summary>
        /// Create TimeSets
        /// if exist in PatternSummary, use what in summary
        /// if specified MCG, use MCD timeset
        /// if exists in patternList csv file, use what in patternList
        /// </summary>
        /// <returns></returns>
        protected string CreateHardIpTimeSets()
        {
            string patternName = Pattern.Pattern.GetLastPayload();
            if (VbtFunctionLibShared.EfusePrewriteFunctionList.Exists(f => f.Equals(Pattern.FunctionName, StringComparison.CurrentCultureIgnoreCase))
                || VbtFunctionLibShared.EfuseReadFunctionList.Exists(f => f.Equals(Pattern.FunctionName, StringComparison.CurrentCultureIgnoreCase)))
            {
                patternName = Lastpatname;
                if (!string.IsNullOrEmpty(patternName))
                {
                    return AcTSetCategoryMapSingleton.Instance().PatternList[patternName].TimeSetVersion;
                }
            }
            if (!string.IsNullOrEmpty(Pattern.TimeSetUsed.McgSetting))
            {
                if (!string.IsNullOrEmpty(Pattern.TimeSetUsed.TimeSetMcg))
                {
                    Lastpatname = patternName;
                }
                return Pattern.TimeSetUsed.TimeSetMcg;
            }

            if (Pattern.Pattern.IsMultiTimeDomain())
            {
                var mtdTimeSets = new List<string>();
                foreach (string lastpat in Pattern.Pattern.GetMtdEachLastPayLoad())
                {
                    if (AcTSetCategoryMapSingleton.Instance().PatternList.ContainsKey(lastpat))
                    {
                        if (AcTSetCategoryMapSingleton.Instance().PatternList[lastpat].IsExist)
                        {
                            Lastpatname = lastpat;
                        }
                        mtdTimeSets.Add(AcTSetCategoryMapSingleton.Instance().PatternList[lastpat].TimeSetVersion);
                    }
                }
                return string.Join(",", mtdTimeSets);
            }

            if (AcTSetCategoryMapSingleton.Instance().PatternList.ContainsKey(patternName))
            {
                if (AcTSetCategoryMapSingleton.Instance().PatternList[patternName].IsExist)
                {
                    Lastpatname = patternName;
                }
                return AcTSetCategoryMapSingleton.Instance().PatternList[patternName].TimeSetVersion;
            }
            return "";
        }

        protected string CreateHardIpMixSignalSheet()
        {
            return Pattern.MixSigUsed;
        }

        protected virtual string CreateTestFailAction(string forceVoltageLabel = "")
        {
            if (Pattern.IsIgnorePatBinOut())
            {
                return "";
            }

            string patternName = Pattern.Pattern.GetPatternName();
            return string.IsNullOrEmpty(forceVoltageLabel) ? CommonGenerator.GenHardIpFlowFailAction(SheetName, BlockName, SubBlockName, patternName, TimingAc, InstNameSubStr,
                LabelVoltage, Pat.MiscInfo, NoPattern) : CommonGenerator.GenHardIpFlowFailAction(SheetName, BlockName, SubBlockName, patternName, TimingAc, InstNameSubStr,
                forceVoltageLabel, Pat.MiscInfo, NoPattern);
        }
        #endregion
    }
}
