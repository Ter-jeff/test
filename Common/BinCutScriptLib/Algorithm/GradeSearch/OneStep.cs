using System.Collections.Generic;
using System.Linq;

using BinCutScriptLib.Base.Line;

namespace BinCutScriptLib.Algorithm.GradeSearch
{
    public class OneStep
    {
        public bool OneStepNoBinOut;
        public List<int> OneStepMfstpNoBinOut = [];
        public List<BinCutLineBase> OneStepInitLines = [];

        public List<BinCutLineBase> OneStepLimits = [];
        public List<BinCutLineBase> OneStepMeasures = [];
        public List<PatternRow> OneStepPatternRows = [];
        public List<BinCutLineBase> OneStepTmps = [];
        public List<BinCutLineBase> OneStepIlbs = [];
        public List<BinCutLineBase> OneStepElbs = [];

        public List<BvLineInfo> OneStepBvLineInfo = [];
        public List<BinCutLineBase> OneStepDssc = [];
        public List<OffsetLine1> OneStepOffset = [];
        public List<BinCutLineBase> OneRunPattOnly = [];
        public List<HarvestSourceCodetRow> OneHarvestSourceCodetRows = [];
        public List<BinCutLineBase> OneStepDisabledPinGroupLine = [];
        public List<BinCutLineBase> OneStepSyncUpLine = [];
        public List<BinCutLineBase> OneStepDebugErrorLine = [];
        public List<BinCutLineBase> OneStepHarvFailLine = [];

        public List<BinCutLineBase> EqLineCs = [];
        public List<BvLineInfo> SafeVoltageCs = [];

        public OneStep()
        {
        }

        public OneStep(OneStep oneStep)
        {
            if (oneStep == null)
            {
                return;
            }

            OneStepNoBinOut = oneStep.OneStepNoBinOut;
            CopyLineListsPart1(oneStep);
            CopyLineListsPart2(oneStep);
        }

        private void CopyLineListsPart1(OneStep oneStep)
        {
            OneStepMfstpNoBinOut = oneStep.OneStepMfstpNoBinOut?.ToList() ?? [];
            OneStepInitLines = oneStep.OneStepInitLines?.ToList() ?? [];
            OneStepLimits = oneStep.OneStepLimits?.ToList() ?? [];
            OneStepMeasures = oneStep.OneStepMeasures?.ToList() ?? [];
            OneStepPatternRows = oneStep.OneStepPatternRows?.ToList() ?? [];
            OneStepTmps = oneStep.OneStepTmps?.ToList() ?? [];
            OneStepIlbs = oneStep.OneStepIlbs?.ToList() ?? [];
            OneStepElbs = oneStep.OneStepElbs?.ToList() ?? [];
            OneStepBvLineInfo = oneStep.OneStepBvLineInfo?.ToList() ?? [];
            OneStepDssc = oneStep.OneStepDssc?.ToList() ?? [];
        }

        private void CopyLineListsPart2(OneStep oneStep)
        {
            OneStepOffset = oneStep.OneStepOffset?.ToList() ?? [];
            OneRunPattOnly = oneStep.OneRunPattOnly?.ToList() ?? [];
            OneHarvestSourceCodetRows = oneStep.OneHarvestSourceCodetRows?.ToList() ?? [];
            OneStepDisabledPinGroupLine = oneStep.OneStepDisabledPinGroupLine?.ToList() ?? [];
            OneStepSyncUpLine = oneStep.OneStepSyncUpLine?.ToList() ?? [];
            OneStepDebugErrorLine = oneStep.OneStepDebugErrorLine?.ToList() ?? [];
            OneStepHarvFailLine = oneStep.OneStepHarvFailLine?.ToList() ?? [];
            EqLineCs = oneStep.EqLineCs?.ToList() ?? [];
            SafeVoltageCs = oneStep.SafeVoltageCs?.ToList() ?? [];
        }

        public OneStep Copy()
        {
            return new OneStep(this);
        }
    }
}
