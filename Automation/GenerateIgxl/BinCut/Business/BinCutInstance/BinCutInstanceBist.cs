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
using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.BinCut.Business.BinCutInstance
{
    public class BinCutInstanceBist : BinCutInstanceBase
    {
        public BinCutInstanceBist(BinCutFinalInstanceRow binCutFinalInstanceRow, BinCutSourceItem sourceRow, BinCutInputData binCutInputManager)
            : base(binCutFinalInstanceRow, sourceRow, binCutInputManager)
        {
            Block = "Mbist";
        }

        protected override void GenerateArgsAndArgList(InstanceRow row)
        {
            ExecuteCoreLogic(row,
                (n, r, i) => GenerateCSharpInstanceRow(n, r),
                (n, r, i) => GenerateVbtInstanceRow(n, r));
        }

        private void GenerateCSharpInstanceRow(Function function, InstanceRow row)
        {
            bool isHbv = SourceRow.TableType.ToString() == "Hv";
            function.SetParamValue("patterns", BinCutFinalInstanceRow.GetPatSetNameForArgument());
            if (!string.IsNullOrEmpty(BinCutFinalInstanceRow.InitPatSetNameNew))
            {
                function.SetParamValue("MbistInitPatterns", BinCutFinalInstanceRow.InitPatSetNameNew);
            }

            function.SetParamValue("instanceFailFlag", isHbv ? GetFlag() : GetFlagBv());
            if (BinCutFinalInstanceRow.CanBeBurst)
            {
                function.SetParamValue("resultMode", "1");
            }
            else
            {
                function.SetParamValue("resultMode", "0");
            }

            function.SetParamValue("performanceMode", (isHbv ? "HBV_" : "") + SourceRow.PerformanceMode);
            function.SetParamValue("isHarvesting", BinCutFinalInstanceRow.BinCutInstanceRow.IsHarvesting);
            if (!string.IsNullOrEmpty(BinCutFinalInstanceRow.BinCutInstanceRow.BinOutStage))
            {
                HashSet<string> binOutJobs = BinCutFinalInstanceRow.BinCutInstanceRow.BinOutStage.Split(',').ToHashSet(StringComparer.OrdinalIgnoreCase);
                string disableBinOutJobs = string.Join(",", AllJobs.Where(x => !binOutJobs.Contains(x)));
                function.SetParamValue("disableBinOut", binOutJobs.Contains("x") ? "ALL" : disableBinOutJobs);
            }

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
                LocalSpecs.HardIpInfos, $":{(isHbv ? "HBV_" : "") + SourceRow.PerformanceMode}",
                BinCutFinalInstanceRow.InitList,
                ref function
            );
            row.Args = function.ArgList;
            row.ArgList = function.Parameters;
        }

        private void GenerateVbtInstanceRow(Function function, InstanceRow row)
        {
            string dsscPat = "";
            function.ArgList[0] = BinCutFinalInstanceRow.GetPatSetNameForArgument();
            if (!string.IsNullOrEmpty(BinCutFinalInstanceRow.InitPatSetName))
            {
                function.SetParamValue("PrePatt", BinCutFinalInstanceRow.InitPatSetName);
            }

            if (BinCutFinalInstanceRow.InitList.Any())
            {
                foreach (string pat in BinCutFinalInstanceRow.InitList)
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
            if (BinCutFinalInstanceRow.CanBeBurst)
            {
                function.SetParamValue("DecomposePatt", "No");
                function.SetParamValue("result_mode", "1");
            }
            function.SetParamValue("Offset_testType", SourceRow.ColumnName.ToString());
            function.SetParamValue("Performance_mode", GetBinningDomain());
            function.SetParamValue("HarvPinGrp_Enable", BinCutFinalInstanceRow.BinCutInstanceRow.HarvPinGrpEnable);
            function.SetParamValue("HarvestBinningFlag", BinCutFinalInstanceRow.BinCutInstanceRow.HarvestBinningFlag);
            row.Args = function.ArgList;
            row.ArgList = function.Parameters;
        }

        internal override string GenerateAcCategory(InstanceRow pRow)
        {
            string timeSet = pRow.TimeSets;
            string acCategory;
            if (BinCutFinalInstanceRow.BinCutInstanceRow != null && !string.IsNullOrEmpty(BinCutFinalInstanceRow.BinCutInstanceRow.ShiftSpeed))
            {
                acCategory = GetAcCategory(timeSet, BlockType.Mbist) + "_" + BinCutFinalInstanceRow.BinCutInstanceRow.ShiftSpeed;
            }
            else
            {
                acCategory = GetAcCategory(timeSet, BlockType.Mbist);
            }

            return acCategory;
        }

        protected override string GenerateLevel()
        {
            if (!string.IsNullOrEmpty(BinCutFinalInstanceRow?.BinCutInstanceRow?.Levels))
            {
                return BinCutFinalInstanceRow?.BinCutInstanceRow?.Levels;
            }

            if (MultiTestSettingSheetsSingleton.Instance().DcCategoryInfos.IsSplitDcSpecs(LocalSpecs.Options.IsSplitDcSpecs))
            {
                return "Levels_BinCut";
            }

            return "Levels_" + Block;
        }
    }
}
