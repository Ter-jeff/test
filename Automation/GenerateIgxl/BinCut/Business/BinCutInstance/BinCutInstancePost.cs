using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.BinCut.Base;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.Singleton;
using Automation.Static;
using Automation.Utility.Atpg;

using IgxlLib.IgxlBase;

using TestPlanLib.Basic;
using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.BinCut.Business.BinCutInstance
{
    public class BinCutInstancePost : BinCutInstanceBase
    {
        public BinCutInstancePost(BinCutFinalInstanceRow binCutFinalInstanceRow, BinCutSourceItem sourceRow, BinCutInputData binCutInputManager)
            : base(binCutFinalInstanceRow, sourceRow, binCutInputManager)
        {

        }

        public override HashSet<string> AllJobs
        {
            get
            {
                if (BinCutInputManager.NewPostBinCutFlowTables != null && BinCutInputManager.NewPostBinCutFlowTables.Any())
                {
                    return BinCutInputManager.NewPostBinCutFlowTables.Select(x => x.JobName).ToHashSet(StringComparer.OrdinalIgnoreCase);
                }

                return BinCutInputManager.BinCutPostFlowSheets.SelectMany(x => x).Select(x => x.JobName).ToHashSet(StringComparer.OrdinalIgnoreCase);
            }
        }

        internal override string GetInstanceName()
        {
            string modePatSetName = GetModePatSetName();
            string parameter = BinCutFinalInstanceRow.BinCutInstanceRow.Type == BincutInstanceType.Hardip || BinCutFinalInstanceRow.BinCutInstanceRow.Type == BincutInstanceType.Rtos ? BinCutFinalInstanceRow.BinCutInstanceRow.Type + "_" + modePatSetName : modePatSetName;
            parameter += $"_{BinCutFinalInstanceRow.BinCutInstanceRow.Instance}";
            parameter += "_outsidebincut";
            parameter = DeleteModeInInstanceName(parameter);
            parameter = AdditionInfoInstanceName(parameter);
            return parameter + "_" + SourceRow.GetBinType();
        }

        protected override string GetVbtName()
        {
            return "GradeSearch_postBinCut_VT";
        }

        protected override string GetFlag()
        {
            string block;
            if (SourceRow.ColumnName == EnumColumnName.TD)
            {
                block = "TD";
            }
            else if (SourceRow.ColumnName == EnumColumnName.Mbist)
            {
                block = "Mbist";
            }
            else
            {
                block = "SPI";
            }

            string section = block + "_";
            return "F_" + section + SourceRow.PerformanceMode + "_outsidebincut_BV";
        }

        protected override void GenerateArgsAndArgList(InstanceRow row)
        {
            ExecuteCoreLogic(row,
                (n, r, i) => GenerateCSharpInstanceRow(n, r),
                (n, r, i) => GenerateVbtInstanceRow(n, r));
        }

        private void GenerateCSharpInstanceRow(Function function, InstanceRow row)
        {
            function.SetParamValue("patterns", BinCutFinalInstanceRow.GetPatSetNameForArgument());
            if (!string.IsNullOrEmpty(BinCutFinalInstanceRow.InitPatSetNameNew))
            {
                function.SetParamValue("MbistInitPatterns", BinCutFinalInstanceRow.InitPatSetNameNew);
            }

            function.SetParamValue("instanceFailFlag", GetFlag());
            if (BinCutFinalInstanceRow.CanBeBurst)
            {
                function.SetParamValue("resultmode", "1");
            }

            if (BinCutFinalInstanceRow.BinCutInstanceRow.FunctionName.Equals("BinCutRetentionTest", StringComparison.CurrentCultureIgnoreCase))
            {
                function.SetParamValue("retentionWaitMillis", BinCutFinalInstanceRow.BinCutInstanceRow.RetentionWaitTime);
                List<string> payloadList = new List<string>();
                var retentionWaitOccurrencesList = new List<string>();
                payloadList.AddRange(BinCutFinalInstanceRow.PayloadList);
                foreach (int retention in BinCutFinalInstanceRow.BinCutInstanceRow.RetentionWaitIdx)
                {
                    payloadList.Insert(retention - BinCutFinalInstanceRow.InitList.Count, "Retention_Pause");
                }
                foreach (string payload in payloadList)
                {
                    retentionWaitOccurrencesList.Add(payload.Equals("Retention_Pause") ? "wait" : "");
                }
                function.SetParamValue("retentionWaitOccurrences", string.Join(",", retentionWaitOccurrencesList));
            }
            if (!string.IsNullOrEmpty(BinCutFinalInstanceRow.BinCutInstanceRow.BinOutStage))
            {
                HashSet<string> binoOutJobs = BinCutFinalInstanceRow.BinCutInstanceRow.BinOutStage.Split(',').ToHashSet(StringComparer.OrdinalIgnoreCase);
                string disableBinoOutJobs = string.Join(",", AllJobs.Where(x => !binoOutJobs.Contains(x)));
                function.SetParamValue("disableBinOut", binoOutJobs.Contains("x") ? "ALL" : disableBinoOutJobs);
            }
            function.SetParamValue("performanceMode", SourceRow.PerformanceMode);

            UserFunctionTableRow ufFuncSetting = null;
            if (!string.IsNullOrEmpty(BinCutFinalInstanceRow.BinCutInstanceRow.UserFunction) && TestPlanStatic.UserFunctionSheet != null)
            {
                ufFuncSetting = TestPlanStatic.UserFunctionSheet.Rows
                .FirstOrDefault(x => x.UserFunction.Equals(BinCutFinalInstanceRow.BinCutInstanceRow.UserFunction, StringComparison.OrdinalIgnoreCase));
            }
            List<string> ufDigSrcPats = TestPlanStatic.UfDigSrcSheets
                .SelectMany(x => x.Rows)
                .Where(y => !string.IsNullOrEmpty(y.PatternName)).Select(z => z.PatternName).ToList();
            AtpgService.SetDigSrc(
                BinCutFinalInstanceRow,
                ufDigSrcPats,
                ufFuncSetting,
                LocalSpecs.HardIpInfos, $":{SourceRow.PerformanceMode}",
                BinCutFinalInstanceRow.InitList.Any() ? BinCutFinalInstanceRow.InitList : BinCutFinalInstanceRow.PatternList,
                ref function
            );
            row.Args = function.ArgList;
            row.ArgList = function.Parameters;
        }

        private void GenerateVbtInstanceRow(Function function, InstanceRow row)
        {
            string dsscPat = "";
            if (SourceRow.ColumnName == EnumColumnName.Mbist)
            {
                function.ArgList[0] = BinCutFinalInstanceRow.GetPatSetNameForArgument();
            }
            else
            {
                function.ArgList[0] = BinCutFinalInstanceRow.PatSetName;
            }

            if (!string.IsNullOrEmpty(BinCutFinalInstanceRow.InitPatSetName))
            {
                function.SetParamValue("PrePatt", BinCutFinalInstanceRow.InitPatSetName);
            }

            if (BinCutFinalInstanceRow.CanBeBurst)
            {
                function.SetParamValue("DecomposePatt", "No");
            }

            function.SetParamValue("EnableBinOut", "TRUE");
            function.SetParamValue("Performance_mode", GetBinningDomain());
            function.SetParamValue("HarvPinGrp_Enable", BinCutFinalInstanceRow.BinCutInstanceRow.HarvPinGrpEnable);
            if (BinCutFinalInstanceRow.PatternList.Any())
            {
                foreach (string pat in BinCutFinalInstanceRow.PatternList)
                {
                    if (Regex.IsMatch(pat, @"\w*DSSC\w*", RegexOptions.IgnoreCase))
                    {
                        dsscPat = pat;
                    }
                }
            }
            if (!string.IsNullOrEmpty(dsscPat))
            {
                string sendPinName = "JTAG_TDI";
                if (LocalSpecs.HardIpInfos != null)
                {
                    HardIpInfo target = LocalSpecs.HardIpInfos.GetHardIpInfo(dsscPat);
                    sendPinName = string.IsNullOrEmpty(target.SendPinName) ? sendPinName : target.SendPinName;
                }
                function.SetParamValue("DigSrc_Pin", sendPinName);
            }
            row.Args = function.ArgList;
            row.ArgList = function.Parameters;
        }

        internal override string GenerateAcCategory(InstanceRow pRow)
        {
            BlockType type = BlockType.Scan;
            if (SourceRow.ColumnName == EnumColumnName.TD ||
                SourceRow.ColumnName == EnumColumnName.FUNC)
            {
                type = BlockType.Scan;
            }
            else if (SourceRow.ColumnName == EnumColumnName.Mbist)
            {
                type = BlockType.Mbist;
            }
            else if (SourceRow.ColumnName == EnumColumnName.ELB || SourceRow.ColumnName == EnumColumnName.ILB)
            {
                type = BlockType.HardIp;
            }
            else if (Regex.IsMatch(SourceRow.PerformanceMode, "TMPS", RegexOptions.IgnoreCase) ||
                     Regex.IsMatch(SourceRow.ColumnContent, "TEMP SENSOR", RegexOptions.IgnoreCase) ||
                     Regex.IsMatch(BinCutFinalInstanceRow.BinCutInstanceRow.FlowName, "ILB", RegexOptions.IgnoreCase) ||
                     Regex.IsMatch(BinCutFinalInstanceRow.BinCutInstanceRow.FlowName, "ELB", RegexOptions.IgnoreCase) ||
                     (SourceRow.ColumnName == EnumColumnName.FUNC && SourceRow.GetDomainOfMode() == "DDR"))
            {
                type = BlockType.HardIp;
            }
            else if (Regex.IsMatch(BinCutFinalInstanceRow.BinCutInstanceRow.FlowName, "CPM", RegexOptions.IgnoreCase))
            {
                type = BlockType.Scan;
            }

            string timeSet = pRow.TimeSets;
            string acCategory;
            if (BinCutFinalInstanceRow.BinCutInstanceRow != null && !string.IsNullOrEmpty(BinCutFinalInstanceRow.BinCutInstanceRow.ShiftSpeed))
            {
                acCategory = GetAcCategory(timeSet, type) + "_" + BinCutFinalInstanceRow.BinCutInstanceRow.ShiftSpeed;
            }
            else
            {
                acCategory = GetAcCategory(timeSet, type);
            }

            return acCategory;
        }

        private string GenerateTd_CpmLevel()
        {
            // Rtos DV level don't need add "SCAN" JN
            if (GetModuleFromInstanceName().Contains("Rtos"))
            {
                return "Levels_" + GetModuleFromInstanceName();
            }
            if (MultiTestSettingSheetsSingleton.Instance().DcCategoryInfos.IsSplitDcSpecs(LocalSpecs.Options.IsSplitDcSpecs))
            {
                return "Levels_BinCut";
            }

            return "Levels_Scan";
        }

        private string GenerateMbistLevel()
        {
            if (MultiTestSettingSheetsSingleton.Instance().DcCategoryInfos.IsSplitDcSpecs(LocalSpecs.Options.IsSplitDcSpecs))
            {
                return "Levels_BinCut";
            }

            return "Levels_Mbist";
        }

        protected override string GenerateLevel()
        {
            if (!string.IsNullOrEmpty(BinCutFinalInstanceRow?.BinCutInstanceRow?.Levels))
            {
                return BinCutFinalInstanceRow?.BinCutInstanceRow?.Levels;
            }

            if (SourceRow.ColumnName == EnumColumnName.TD ||
                SourceRow.ColumnName == EnumColumnName.FUNC ||
                Regex.IsMatch(BinCutFinalInstanceRow.BinCutInstanceRow.FlowName, "CPM", RegexOptions.IgnoreCase))
            {
                return GenerateTd_CpmLevel();
            }

            if (SourceRow.ColumnName == EnumColumnName.Mbist)
            {
                return GenerateMbistLevel();
            }

            if (SourceRow.ColumnName == EnumColumnName.ELB || SourceRow.ColumnName == EnumColumnName.ILB)
            {
                return "Levels_HardIP";
            }

            if (Regex.IsMatch(SourceRow.PerformanceMode, "TMPS", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(SourceRow.ColumnContent, "TEMP SENSOR", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(BinCutFinalInstanceRow.BinCutInstanceRow.FlowName, "ILB", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(BinCutFinalInstanceRow.BinCutInstanceRow.FlowName, "ELB", RegexOptions.IgnoreCase) ||
                (SourceRow.ColumnName == EnumColumnName.FUNC && SourceRow.GetDomainOfMode() == "DDR"))
            {
                return "Levels_HardIP";
            }
            return GenerateTd_CpmLevel();
        }
    }
}
