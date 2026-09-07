using System.Collections.Generic;
using System.Linq;

using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Static;

using CommonLib.Datalog;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using TestPlanLib.BinCut.Binning;

namespace BinCutScriptLib.Algorithm.GradeSearch
{
    public class OneGradeSearch
    {
        public string Name = "";
        public List<LimitLine> JudgeResultLines = [];
        public List<LimitLine> PatternResultLines = [];
        public Dictionary<int, BinCutLineBase> FailSite = [];
        public bool IsSiteMismatch;

        public List<VBinResultWoTestLine1> VBinResultLines = [];
        public List<SearchResultWoTestLine1> SearchResultWoTestLine = [];
        public BinCutLineBase InstanceLine = new();
        public InstanceBinCut? InstanceBinCut;
        public string HarvestBinningFlag = "";
        public List<OneStep> Steps = [];
        public List<InterpolationRow> EnRows = [];
        public bool IsCallInstance;
        public EnumPreVddStatus PreVddStatus = EnumPreVddStatus.None;
        public EnumFstpStatus FstpStatus = EnumFstpStatus.None;
        public bool IsReJudge;
        public bool IsHarvAssignedTrue;

        public string BvMode { get; internal set; } = "";
        public string Rail { get; internal set; } = "";
        public string Pmode { get; internal set; } = "";

        public void ClearAll()
        {
            Name = "";
            JudgeResultLines.Clear();
            VBinResultLines.Clear();
            EnRows.Clear();
        }

        public BvName GetFirstBvName(List<string> powerNames)
        {
            foreach (OneStep step in Steps)
            {
                foreach (BvLineInfo oneStepBvLineInfo in step.OneStepBvLineInfo)
                {
                    var bvName = new BvName
                    {
                        Name = oneStepBvLineInfo.BvName,
                        NameRemoveBv = Reg.RegexBv.Replace(oneStepBvLineInfo.BvName, "")
                    };
                    bvName.Mode = BinCutAlgorithmService.GetModeByName(bvName.Name);
                    bvName.Index = GetPowerIndexByBv(powerNames, oneStepBvLineInfo.BvName);
                    bvName.Sites = [.. Steps.SelectMany(x => x.OneStepBvLineInfo.Select(y => y.Site)).Distinct()];
                    return bvName;
                }
            }
            return null!;
        }

        public BvName GetFirstBvNameCs(List<PinInfo> pinInfos)
        {
            if (!string.IsNullOrEmpty(BvMode))
            {
                var bvName = new BvName
                {
                    Name = BvMode.Replace("'", ""),
                    NameRemoveBv = Reg.RegexBv.Replace(BvMode, "").Trim('\''),
                    Mode = BinCutAlgorithmService.GetModeByName(BvMode).Trim('\''),
                    Index = GetPowerIndexByBv([.. pinInfos.Select(x => x.Mode)], BvMode),
                    Sites = [.. Steps.SelectMany(x => x.OneStepBvLineInfo.Select(y => y.Site)).Distinct()]
                };
                return bvName;
            }
            return null!;
        }

        private static int GetPowerIndexByBv(List<string> powerNames, string bvName)
        {
            //eq. BV_VDD_SOC_MS001 => Index of List
            for (int index = 0; index < powerNames.Count; index++)
            {
                if (bvName.Contains(powerNames[index]))
                {
                    return index;
                }
            }
            string errorMessage = $"{bvName} can't be found in powerName list of <Judge_stored_IDS>!!!";
            if (!ErrorReportManager.GetErrorList().Select(x => x.Message).Contains(errorMessage))
            {
                ErrorReportManager.AddError(BasicErrorType.E_MissingPin_11, "Datalog", 0, 0, $"{bvName} can't be found in powerName list of <Judge_stored_IDS>!!!", [bvName]);
            }

            return -1;
        }
    }
}
