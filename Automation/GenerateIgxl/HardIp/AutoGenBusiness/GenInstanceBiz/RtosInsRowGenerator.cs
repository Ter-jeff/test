using System;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using IgxlLib.IgxlBase;

using LogLib.Static;

using TestPlanLib.Static;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenInstanceBiz
{
    public class RtosInsRowGenerator : HardIpInsRowGenerator
    {

        #region Constructor
        public RtosInsRowGenerator(HardIpInputData hardIpInputData, HardIpSheet hardIpSheet, string sheetName) : base(hardIpInputData, hardIpSheet, sheetName)
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
            if (VbtFunction.FunctionName.Equals(FuncNameConst.CSharpFuncNameRtosRunScenario, StringComparison.OrdinalIgnoreCase))
            {
                var voltages = new List<string> { "NV", "HV", "LV" };
                foreach (string vol in voltages)
                {
                    var insRow = new InstanceRow
                    {
                        SheetName = SheetName,
                        VbtType = CreateType(),
                        VbtName = CreateVbtName(),
                        ArgList = CreateArgList(),
                        Args = CreateArgs(),
                        TimeSets = string.IsNullOrEmpty(Pattern.RtosIdsTimeSetUsed) ? CreateHardIpTimeSets() : Pattern.RtosIdsTimeSetUsed,
                        DcCategory = CreateRtosDcCategory(),
                        DcSelector = CreateDcSelector(),
                        TestName = CreateRtosTestName() + "_" + vol
                    };
                    insRow.AcCategory = CreateHardIpAcCategory(insRow.TimeSets);
                    insRow.AcSelector = string.IsNullOrEmpty(insRow.AcCategory) ? "" : CreateAcSelector();
                    insRow.PinLevels = CreateRtosPinLevel();
                    insRowlist.Add(insRow);
                }
            }
            else
            {
                var insRow = new InstanceRow
                {
                    SheetName = SheetName,
                    VbtType = CreateType(),
                    VbtName = CreateVbtName(),
                    ArgList = CreateArgList(),
                    Args = CreateArgs(),
                    TimeSets = string.IsNullOrEmpty(Pattern.RtosIdsTimeSetUsed) ? CreateHardIpTimeSets() : Pattern.RtosIdsTimeSetUsed,
                    DcCategory = CreateRtosDcCategory(),
                    DcSelector = CreateDcSelector(),
                    TestName = CreateRtosTestName()
                };
                insRow.AcCategory = CreateHardIpAcCategory(insRow.TimeSets);
                insRow.AcSelector = string.IsNullOrEmpty(insRow.AcCategory) ? "" : CreateAcSelector();
                insRow.PinLevels = CreateRtosPinLevel();
                UpdateIdsMappingTable(Pattern.SubBlock, insRow.TestName);

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
        protected virtual string CreateRtosTestName()
        {
            if (Pattern.FunctionName.Equals(VbtFunctionLibShared.RtosBootUp, StringComparison.OrdinalIgnoreCase) || Pattern.FunctionName.Equals(FuncNameConst.CSharpFuncNameRtosRunScenario, StringComparison.OrdinalIgnoreCase))
            {
                string rtosTestName = CommonGenerator.GenRtosInsTestName(Pattern.FunctionName, BlockName, SubBlockName, Pattern.Pattern.GetPatternName());
                if (rtosTestName != "__CONTINUE__" && !string.IsNullOrEmpty(rtosTestName))
                {
                    return rtosTestName;
                }

            }


            string patternName = Pattern.Pattern.GetPatternName();
            LabelVoltage = HardIpConstData.LabelNv;

            if (!string.IsNullOrEmpty(Pat.TestName))
            {
                return Pat.TestName + "_" + LabelVoltage;
            }

            return CommonGenerator.GenHardIpInsTestName(BlockName, SubBlockName, patternName,
                Pat.PatternIndexFlag,
                TimingAc, Pat.ForceVoltageFlag, InstNameSubStr, "", NoPattern, Pat.WirelessData.IsNeedPostBurn, true, Pat.WirelessData.IsDoMeasure);
        }

        /// <summary>
        /// Create Ids PinLevel
        /// if exist in PatternSummary, use what in summary
        /// else use default: Levels_IDS
        /// </summary>
        /// <returns></returns>
        protected string CreateRtosPinLevel()
        {

            if (!string.IsNullOrEmpty(Pattern.LevelUsed))
            {
                return Pattern.LevelUsed;
            }

            return HardIpConstData.RtosLevelDefault;
        }
        #endregion
        public void UpdateIdsMappingTable(string subBlock, string instanceName)
        {
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

        protected virtual string CreateRtosDcCategory()
        {
            string dcCategory = MultiTestSettingSheetsSingleton.Instance().FindRtosCategory("", out EnumMessageLevel _, out string _);
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
                    Response.Report("Can not find Rtos category: " + dcCategory + " from TestSettting sheet", EnumMessageLevel.Error, 45);
                    ErrorReportManager.AddError(
                        RtosErrorType.E_MissingParameter_01,
                        Pattern.SheetName,
                        Pattern.RowNum,
                        0,
                        [dcCategory]
                    );
                }
                return dcCategory;
            }

            return dcCategory;
        }


    }
}
