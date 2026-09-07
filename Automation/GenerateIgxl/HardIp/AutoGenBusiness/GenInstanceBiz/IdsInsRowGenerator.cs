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

using IgxlLib.IgxlBase;

using LogLib.Static;

using TestPlanLib.Singleton;
using TestPlanLib.Static;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenInstanceBiz
{
    public class IdsInsRowGenerator : HardIpInsRowGenerator
    {
        #region Const properties

        protected bool NoFuse
        {
            get { return HardIpService.IdsNoFuse(Pattern.MiscInfo); }
        }

        protected bool IsNandPattern
        {
            get { return HardIpService.IsNandPattern(Pattern.Pattern.GetLastPayload()); }
        }

        protected bool IsSpiPattern
        {
            get { return HardIpService.IsSpiPattern(Pattern.Pattern.GetLastPayload()); }
        }
        #endregion

        #region Constructor
        public IdsInsRowGenerator(HardIpInputData hardIpInputData, HardIpSheet hardIpSheet, string sheetName) : base(hardIpInputData, hardIpSheet, sheetName)
        {
        }
        #endregion

        #region Main Methods
        /// <summary>
        /// Generate test instance rows for IDS sheet
        /// if pattern specified as IDS_NoFuse, use naming rule for normal hardip pattern
        /// else use IDS naming rule
        /// </summary>
        /// <returns></returns>
        public override List<InstanceRow> GenInsRows()
        {
            var insRowlist = new List<InstanceRow>();
            if (!NoFuse && !IsNandPattern && !IsSpiPattern)
            {
                return insRowlist;
            }

            Pattern.HipPreWriteFlag =
                VbtFunctionLibShared.EfusePrewriteFunctionList.Exists(f => f.Equals(Pattern.FunctionName, StringComparison.CurrentCultureIgnoreCase))
                ? LastPatFailAction
                : CreateTestFailAction();
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

            var insRow = new InstanceRow
            {
                SheetName = SheetName,
                VbtType = CreateType(),
                VbtName = CreateVbtName(),
                ArgList = CreateArgList(),
                Args = CreateArgs(),
                TestName = CreateIdsTestName(),
                TimeSets = string.IsNullOrEmpty(Pattern.RtosIdsTimeSetUsed) ? CreateIdsTimeSets() : Pattern.RtosIdsTimeSetUsed
            };
            insRow.DcCategory = insRow.VbtName.Contains("IedaSetting") ? "" : CreateIdsDcCategory();
            insRow.DcSelector = insRow.VbtName.Contains("IedaSetting") ? "" : CreateDcSelector();
            insRow.AcCategory = CreateHardIpAcCategory(insRow.TimeSets);
            insRow.AcSelector = string.IsNullOrEmpty(insRow.AcCategory) ? "" : CreateAcSelector();
            insRow.PinLevels = insRow.VbtName.Contains("IedaSetting") ? "" : TestPlanStatic.ConcurrentFlowSheet != null
                && TestPlanStatic.ConcurrentFlowSheet.Rows.Exists(x => x.Subflows.Exists(y => y.ToUpper().Replace("FLOW_HARDIP_", "").Equals(SheetName.Replace("HARDIP_", ""))))
                ? CreateHardIpLevelConcurrent(insRow.TimeSets) : CreateIdsPinLevel();
            UpdateIdsMappingTable(Pattern.SubBlock, insRow.TestName);
            insRowlist.Add(insRow);

            if (HardIpInputData.ConfigData.AdaptiveCoolingItem.Contains($"{BlockName}_{SubBlockName}"))
            {
                VbtFunction = CommonGenerator.GetVbtFunctionBase(VbtFunctionLibShared.VifName);
                int seqIndex = VbtFunction.Parameters.Split(',').ToList().FindIndex(s => s.Equals("TestSequence", StringComparison.OrdinalIgnoreCase));
                int measPinIdx = VbtFunction.Parameters.Split(',').ToList().FindIndex(s => s.Equals("MeasI_PinS", StringComparison.OrdinalIgnoreCase));
                insRow = new InstanceRow
                {
                    SheetName = SheetName,
                    VbtType = CreateType(),
                    VbtName = CreateVbtName(),
                    ArgList = CreateArgList(),
                    Args = CreateArgs()
                };
                if (seqIndex != -1)
                {
                    insRow.Args[seqIndex] = "N";
                }

                if (measPinIdx != -1)
                {
                    insRow.Args[measPinIdx] = "";
                }

                insRow.TestName = CommonGenerator.GenHardIpInsTestName(BlockName, SubBlockName + "_TMPSMON", Pattern.Pattern.GetLastPayload(),
                    Pat.PatternIndexFlag,
                    TimingAc, Pat.ForceVoltageFlag, InstNameSubStr, "NV", NoPattern, Pat.WirelessData.IsNeedPostBurn, true, Pat.WirelessData.IsDoMeasure);
                insRow.TimeSets = string.IsNullOrEmpty(Pattern.RtosIdsTimeSetUsed) ? CreateIdsTimeSets() : Pattern.RtosIdsTimeSetUsed;
                insRow.DcCategory = CreateIdsDcCategory();
                insRow.DcSelector = CreateDcSelector();
                insRow.AcCategory = CreateHardIpAcCategory(insRow.TimeSets);
                insRow.AcSelector = string.IsNullOrEmpty(insRow.AcCategory) ? "" : CreateAcSelector();
                insRow.PinLevels = CreateIdsPinLevel();

                insRowlist.Add(insRow);
            }

            return insRowlist;
        }

        #endregion

        #region Generate each columns Methods
        /// <summary>
        /// Create Ids TestName: IDS_[NandTestName/SpiTestName]_[NV/LV/HV]
        /// </summary>
        /// <returns></returns>
        protected virtual string CreateIdsTestName()
        {
            string patternName = Pattern.Pattern.GetPatternName();
            LabelVoltage = HardIpConstData.LabelNv;

            if (!string.IsNullOrEmpty(Pat.TestName))
            {
                return Pat.TestName + "_" + LabelVoltage;
            }

            return CommonGenerator.GenHardIpInsTestName(BlockName, SubBlockName, patternName,
                Pat.PatternIndexFlag,
                TimingAc, Pat.ForceVoltageFlag, InstNameSubStr, "NV", NoPattern, Pat.WirelessData.IsNeedPostBurn, true, Pat.WirelessData.IsDoMeasure);
        }

        /// <summary>
        /// Create Ids DcCategory
        /// if exist in PatternSummary, use what in summary
        /// if patternName contains MD[XX], use "HARDIP_MD[XX]
        /// else use default: IDS
        /// </summary>
        /// <returns></returns>
        protected string CreateIdsDcCategory()
        {
            string dcCategory = ExceptionListSingleton.Instance().GetDcCategoryByInstance(CreateIdsTestName());
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
                    Response.Report("Can not find HardIp category: " + dcCategory, EnumMessageLevel.Error, 45);
                }
                return dcCategory;
            }

            string performanceMode = "";
            string hardipDcName = HardIpDcSetting != null ? HardIpDcSetting.CategoryName : "";
            string[] patternSplitArr = Pattern.Pattern.GetLastPayload().Split('_');
            if (patternSplitArr.Length >= 10 && ModuleSingleton.GetModuleByPerformanceMode(patternSplitArr[9]) == ModuleSingleton.Instance().ModuleDdr)
            {
                performanceMode = patternSplitArr[9];
            }
            List<string> patterns = new List<string> { Pattern.Pattern.GetLastPayload() };
            dcCategory = MultiTestSettingSheetsSingleton.Instance().DcCategoryInfos.FindIdsCatgeoryName(performanceMode, hardipDcName, patterns, out EnumMessageLevel msgLevel, out string errorMsg);
            if (!string.IsNullOrEmpty(errorMsg))
            {
                Response.Report(errorMsg, msgLevel, 45);
            }
            return dcCategory;

        }

        /// <summary>
        /// Create Ids PinLevel
        /// if exist in PatternSummary, use what in summary
        /// else use default: Levels_IDS
        /// </summary>
        /// <returns></returns>
        protected string CreateIdsPinLevel()
        {
            if (HardIpDcSetting != null)
            {
                return HardIpDcSetting.LevelSheet;
            }

            return HardIpConstData.IdsLevelDefault;
        }

        /// <summary>
        /// Create Ids PinLevel
        /// if exist in PatternSummary, use what in summary
        /// if exists in patternList csv file, use what in patternList
        /// </summary>
        /// <returns></returns>
        protected string CreateIdsTimeSets()
        {
            string spectimeset = GetSpecifyInfo(Pattern.ForceCondition.ForceCondition, "TimeSets");
            if (!string.IsNullOrEmpty(spectimeset))
            {
                return spectimeset;
            }

            if (AcTSetCategoryMapSingleton.Instance().PatternList.ContainsKey(Pattern.Pattern.GetLastPayload()))
            {
                return AcTSetCategoryMapSingleton.Instance().PatternList[Pattern.Pattern.GetLastPayload()].TimeSetVersion;
            }
            return string.Empty;
        }

        private void UpdateIdsMappingTable(string subBlock, string instanceName)
        {
            if (!instanceName.ToUpper().EndsWith("NV"))
            {
                return;
            }

            if (TestPlanStatic.IdsMappingSheet == null)
            {
                return;
            }

            TestPlanStatic.IdsMappingSheet.Rows
                .Where(x => x.SubBlock.Trim().Replace("_", "")
                            .Equals(subBlock.Trim().Replace("_", ""), StringComparison.OrdinalIgnoreCase))
                .ToList()
                .ForEach(x => x.InstanceName = instanceName);
        }

        #endregion
    }
}
