using Automation.GenerateIgxl.BinCut.Base;
using Automation.InputManager.Data;
using Automation.Singleton;
using Automation.Static;

using IgxlLib.IgxlBase;

using TestPlanLib.BinCut.BinCutInstance;

namespace Automation.GenerateIgxl.BinCut.Business.BinCutInstance
{
    public class BinCutInstanceTdDdr : BinCutInstanceBase
    {
        public BinCutInstanceTdDdr(BinCutFinalInstanceRow binCutFinalInstanceRow, BinCutSourceItem sourceRow, BinCutInputData binCutInputManager)
            : base(binCutFinalInstanceRow, sourceRow, binCutInputManager)
        {
            Block = "Scan";
        }

        internal override string GetInstanceName()
        {
            string modePatSetName = GetModePatSetName();
            string parameter = BinCutFinalInstanceRow.BinCutInstanceRow.Type == BincutInstanceType.Hardip || BinCutFinalInstanceRow.BinCutInstanceRow.Type == BincutInstanceType.Rtos ? BinCutFinalInstanceRow.BinCutInstanceRow.Type + "_" + modePatSetName : modePatSetName;
            parameter += $"_{BinCutFinalInstanceRow.BinCutInstanceRow.Instance}";
            parameter = DeleteModeInInstanceName(parameter);
            parameter = AdditionInfoInstanceName(parameter);
            return parameter + "_" + SourceRow.GetBinType();
        }

        internal override string GenerateAcCategory(InstanceRow pRow)
        {
            string timeSet = pRow.TimeSets;
            string acCategory;
            if (BinCutFinalInstanceRow.BinCutInstanceRow != null && !string.IsNullOrEmpty(BinCutFinalInstanceRow.BinCutInstanceRow.ShiftSpeed))
            {
                acCategory = GetAcCategory(timeSet, BlockType.Scan) + "_" + BinCutFinalInstanceRow.BinCutInstanceRow.ShiftSpeed;
            }
            else
            {
                acCategory = GetAcCategory(timeSet, BlockType.Scan);
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
