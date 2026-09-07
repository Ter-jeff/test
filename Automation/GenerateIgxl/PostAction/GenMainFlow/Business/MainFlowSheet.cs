using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.PostAction.GenMainFlow.Base;
using Automation.Static;

namespace Automation.GenerateIgxl.PostAction.GenMainFlow.Business
{
    public class MainFlowSheet
    {
        public const string FlowMainConst = "Flow_M_";
        public const string SubProgramFlowEnableWdConst = "Enable_";
        private static readonly Regex _regexFlowCoreOverFailing = new Regex("^Flow_CoreOverFailing(?:_(?<blockName>\\w+))?$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public MainFlowSheet(List<MainFlowBase> mainFlowBases, List<string> jobs)
        {
            Rows = mainFlowBases;
            Jobs = jobs;

            SetEvsDeferredBinOut();
            SetCoreOverFailingFlows();
        }
        public MainFlowSheet(List<MainFlowBase> mainFlowBases, int sourceCol, int sheetNameCol, int subFlowCol, int subprogramCol, int enableWdCol, int bintableEnableWdCol, int siteFlagPerSiteCol, int failFlagCol, int commentCol, int moduleCol, int includeCol, int groupCol, int optionCol, List<string> jobs, List<string> enableModules)
        {
            Rows = mainFlowBases;
            SourceCol = sourceCol;
            SheetNameCol = sheetNameCol;
            SubFlowCol = subFlowCol;
            SubprogramCol = subprogramCol;
            EnableWdCol = enableWdCol;
            BintableEnableWdCol = bintableEnableWdCol;
            SiteFlagPerSiteCol = siteFlagPerSiteCol;
            FailFlagCol = failFlagCol;
            CommentCol = commentCol;
            ModuleCol = moduleCol;
            IncludeCol = includeCol;
            GroupCol = groupCol;
            OptionCol = optionCol;
            Jobs = jobs;
            EnableModules = enableModules;

            SetEvsDeferredBinOut();
            SetCoreOverFailingFlows();
        }

        public List<MainFlowBase> Rows { get; }
        public int SourceCol { get; } = -1;
        public int SheetNameCol { get; } = -1;
        public int SubFlowCol { get; } = -1;
        public int SubprogramCol { get; } = -1;
        public int EnableWdCol { get; } = -1;
        public int BintableEnableWdCol { get; } = -1;
        public int SiteFlagPerSiteCol { get; } = -1;
        public int FailFlagCol { get; } = -1;
        public int CommentCol { get; } = -1;
        public int ModuleCol { get; } = -1;
        public int IncludeCol { get; } = -1;
        public int GroupCol { get; } = -1;
        public int OptionCol { get; } = -1;
        public List<string> Jobs { get; }
        public List<string> EnableModules { get; }
        public HashSet<string> EvsDeferSubFlowSheets { get; private set; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, string> CoreOverFailingFlows { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        private void SetEvsDeferredBinOut()
        {
            if (EnableModules?.Contains(BlockStatus.Evs) != true)
            {
                return;
            }

            HashSet<string> evsDeferSubFlowSheets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            HashSet<string> deferredBinoutSequences = Rows.First().SequencesNew
                .Where(x => x.Module.Equals("Defer", StringComparison.OrdinalIgnoreCase)
                    && x.IsEvsDeferredBinout)
                .Select(x => x.SheetName).ToHashSet(StringComparer.OrdinalIgnoreCase);
            IEnumerable<FlowSequenceNew> needEvsDeferredSequences = Rows.First().SequencesNew
                .Where(x => x.OptionDict.ContainsKey("Defer"));
            foreach (FlowSequenceNew sequence in needEvsDeferredSequences)
            {
                if (deferredBinoutSequences.Contains(sequence.OptionDict["Defer"]))
                {
                    string sourceSheet = string.IsNullOrEmpty(sequence.SubFlowName) ?
                        $"{sequence.SheetName}" : $"{sequence.SheetName}:{sequence.SubFlowName}";
                    evsDeferSubFlowSheets.Add(sourceSheet);
                }
            }
            EvsDeferSubFlowSheets = evsDeferSubFlowSheets;
        }

        private void SetCoreOverFailingFlows()
        {
            if (Rows == null || !Rows.Any())
            {
                return;
            }

            Dictionary<string, string> coreOverFailingFlows = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string sheetName in Rows.First().SequencesNew.Select(x => x.SheetName))
            {
                Match match = _regexFlowCoreOverFailing.Match(sheetName);
                if (match.Success)
                {
                    if (!coreOverFailingFlows.ContainsKey(sheetName))
                    {
                        coreOverFailingFlows[sheetName] = match.Groups["blockName"].Value.ToUpper();
                    }
                }
            }
            CoreOverFailingFlows = coreOverFailingFlows;
        }
    }
}
