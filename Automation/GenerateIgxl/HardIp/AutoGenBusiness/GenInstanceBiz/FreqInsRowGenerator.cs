using System;
using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.Singleton;
using Automation.Static;
using Automation.Utility.HardIP;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using IgxlLib.IgxlBase;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenInstanceBiz
{
    public class FreqInsRowGenerator : InsRowGenerator
    {
        public string BlockScbi;
        public string Module;

        public FreqInsRowGenerator(HardIpInputData hardIpInputData, HardIpSheet hardIpSheet, string sheetName) : base(hardIpInputData, hardIpSheet, sheetName)
        {
        }

        private BlockType _block = BlockType.HardIp;
        /// <summary>
        /// Generate instance rows for one pattern, if contains ReTest, two instances will be generated, else only one instance will be generated
        /// </summary>
        /// <returns></returns>
        public override List<InstanceRow> GenInsRows()
        {
            var insRowlist = new List<InstanceRow>();

            #region basic row

            InstanceRow insRow = GenHardIpRow();
            insRowlist.Add(insRow);

            #endregion

            #region reTest row
            if (Pat.MiscInfoDict.ContainsKey(HardIpConstData.ReTest))
            {
                InstanceRow reTestRow = GenReTestRow(insRow);
                insRowlist.Add(reTestRow);
            }
            #endregion

            return insRowlist;
        }

        /// <summary>
        /// Genrator Hardip instance row(Except IDS)
        /// </summary>
        /// <returns></returns>
        protected InstanceRow GenHardIpRow()
        {
            var insRow = new InstanceRow
            {
                SheetName = SheetName,
                VbtType = CreateType(),
                VbtName = CreateVbtName(),
                ArgList = CreateArgList(),
                Args = CreateArgs(),
                TestName = CreateHardIpTestName(),
                TimeSets = CreateHardIpTimeSets(),
                DcCategory = CreateHardIpDcCategory(),
                DcSelector = CreateDcSelector()
            };
            insRow.AcCategory = CreateHardIpAcCategory(insRow.TimeSets);
            insRow.AcSelector = string.IsNullOrEmpty(insRow.AcCategory) ? "" : CreateAcSelector();
            insRow.PinLevels = TestPlanStatic.ConcurrentFlowSheet != null
                && TestPlanStatic.ConcurrentFlowSheet.Rows.Exists(x => x.Subflows.Exists(y => y.ToUpper().Replace("FLOW_HARDIP_", "").Equals(SheetName.Replace("HARDIP_", ""))))
                ? CreateHardIpLevelConcurrent(insRow.TimeSets) : CreateHardIpPinLevel();
            insRow.MixedSignalTiming = "";
            return insRow;
        }

        protected override void SetBasicInfoByPattern(HardIpPattern pattern)
        {
            BlockName = CommonGenerator.GetBlockNameFromSheetName(pattern.SheetName);
            SubBlockName = CommonGenerator.GetSubBlockName(pattern.Pattern.GetLastPayload(), pattern.MiscInfo, BlockName, pattern.ForceCondition.IsCz2InstName);
            TimingAc = CommonGenerator.GetTimingAc(pattern);
            InstNameSubStr = HardIpService.GetInstNameSubStr(pattern.MiscInfo);
            VbtFunction = CommonGenerator.GetVbtFunctionBase(pattern.FunctionName);
            if (VbtFunction.FunctionName == "" && pattern.MiscInfo != "")
            {
                ErrorReportManager.AddError(ClockCheckErrorType.E_MissingLibrary_01, pattern.SheetName, pattern.RowNum, 0, [pattern.MiscInfo]);
            }

            NoPattern = HardIpService.IsNoPattern(pattern.Pattern.GetLastPayload());
            HardIpDcSetting = CommonGenerator.GetHardIpDcSetting(pattern.LevelUsed);
            _block = pattern.BlockType.Equals("scan", StringComparison.OrdinalIgnoreCase) ? BlockType.Scan : BlockType.Mbist;
        }

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
            return CommonGenerator.GenHardIpInsTestName(BlockName, SubBlockName, patternName, Pattern.PatternIndexFlag, TimingAc,
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
            string dc = GetSpecifyInfo(Pattern.DcCategory, "DC");
            return dc;

        }

        /// <summary>
        /// Create PinLevel
        /// if exist in PatternSummary, use what in summary
        /// if specified Hardip DC, use what in HardIpDc
        /// </summary>
        /// <returns></returns>
        protected virtual string CreateHardIpPinLevel()
        {
            if (!string.IsNullOrEmpty(Pattern.LevelUsed))
            {
                return GetSpecifyInfo(Pattern.LevelUsed, "Level");
            }
            BlockType type = _block;
            if (type == BlockType.Scan)
            {
                return "Levels_Scan";
            }

            if (type == BlockType.Mbist)
            {
                return "Levels_Mbist";
            }

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

            string timeDomain = AcTSetCategoryMapSingleton.Instance().MultiTimeSetSheets.Find(x => x.Name.Equals(timeSet)).TimeDomain;
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

            if (AcTSetCategoryMapSingleton.Instance().PatternList.ContainsKey(patternName))
            {
                return AcTSetCategoryMapSingleton.Instance().PatternList[patternName].TimeSetVersion;
            }

            if (AcTSetCategoryMapSingleton.Instance().PatternList.ContainsKey(patternName.ToUpper()))
            {
                return AcTSetCategoryMapSingleton.Instance().PatternList[patternName.ToUpper()].TimeSetVersion;


            }

            return "";
        }
    }
}
