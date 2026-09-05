using Automation.GenerateIgxl.BinCut.Base;
using Automation.InputManager.Data;
using Automation.Singleton;

using IgxlLib.IgxlBase;

using TestPlanLib.BinCut.Flow;

namespace Automation.GenerateIgxl.BinCut.Business.BinCutInstance
{
    public class BinCutInstanceIlb : BinCutInstanceBase
    {
        public BinCutInstanceIlb(BinCutFinalInstanceRow binCutFinalInstanceRow, BinCutSourceItem sourceRow, BinCutInputData binCutInputManager)
            : base(binCutFinalInstanceRow, sourceRow, binCutInputManager)
        {
        }

        internal override string GetInstanceName()
        {
            return GetIlbElbInstanceName();
        }

        internal override string GenerateAcCategory(InstanceRow pRow)
        {
            string timeSet = pRow.TimeSets;
            string acCategory;
            if (BinCutFinalInstanceRow.BinCutInstanceRow != null && !string.IsNullOrEmpty(BinCutFinalInstanceRow.BinCutInstanceRow.ShiftSpeed))
            {
                acCategory = GetAcCategory(timeSet, BlockType.HardIp) + "_" + BinCutFinalInstanceRow.BinCutInstanceRow.ShiftSpeed;
            }
            else
            {
                acCategory = GetAcCategory(timeSet, BlockType.HardIp);
            }

            return acCategory;
        }

        protected override string GenerateLevel()
        {
            if (!string.IsNullOrEmpty(BinCutFinalInstanceRow?.BinCutInstanceRow?.Levels))
            {
                return BinCutFinalInstanceRow?.BinCutInstanceRow?.Levels;
            }

            return "Levels_HardIP";
        }

        protected override string GetFlag()
        {
            if (SourceRow.TableType == EnumBinCutTableType.Post)
            {
                return $"F_{SourceRow.ColumnName}_{SourceRow.PerformanceMode}_outsidebincut_BV";
            }

            return "F_ILB_" + SourceRow.TargetPerformanceMode + "_HBV";
        }
    }
}
