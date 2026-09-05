using Automation.GenerateIgxl.BinCut.Base;
using Automation.InputManager.Data;
using Automation.Singleton;

using IgxlLib.IgxlBase;

namespace Automation.GenerateIgxl.BinCut.Business.BinCutInstance
{
    public class BinCutInstanceDdrLb : BinCutInstanceBase
    {
        public BinCutInstanceDdrLb(BinCutFinalInstanceRow binCutFinalInstanceRow, BinCutSourceItem sourceRow, BinCutInputData binCutInput)
            : base(binCutFinalInstanceRow, sourceRow, binCutInput)
        {
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
            return "F_DDR_" + SourceRow.TargetPerformanceMode + "_HBV";
        }
    }
}
