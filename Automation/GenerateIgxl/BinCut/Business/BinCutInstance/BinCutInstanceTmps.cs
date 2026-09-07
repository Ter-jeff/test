using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.BinCut.Base;
using Automation.InputManager.Data;
using Automation.Singleton;

using IgxlLib.IgxlBase;


namespace Automation.GenerateIgxl.BinCut.Business.BinCutInstance
{
    public class BinCutInstanceTmps : BinCutInstanceBase
    {
        public BinCutInstanceTmps(BinCutFinalInstanceRow binCutFinalInstanceRow, BinCutSourceItem sourceRow, BinCutInputData binCutInputManager)
            : base(binCutFinalInstanceRow, sourceRow, binCutInputManager)
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

        protected override string GetBinningDomain()
        {
            if (SourceRow.BinningDomain.Contains(","))
            {
                List<string> domains = SourceRow.BinningDomain.Split(',').ToList();
                string domainPin = "VDD_" + domains.First() + "_" + SourceRow.TargetPerformanceMode;
                domainPin = domainPin + "," + "VDD_" + string.Join("_", domains);
                var pinList = domains.Select(p => "VDD_" + p).ToList();
                SourceRow.AddPinGroup("VDD_" + string.Join("_", domains), pinList);
                return domainPin;
            }
            return "VDD_" + SourceRow.BinningDomain + "_" + SourceRow.TargetPerformanceMode;
        }

        protected override string GetFlag()
        {
            return "F_TMPS_" + SourceRow.TargetPerformanceMode + "_HBV";
        }
    }
}
