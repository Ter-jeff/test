using System.Collections.Generic;
using System.Linq;

namespace TestPlanLib.BinCut.PowerBinning
{
    public class PowerBinningSheet
    {
        public string SheetName { get; set; } = string.Empty;
        public List<PowerBinningSheetRow> BinnedModeList { get; set; } = [];
        public List<PowerBinningSheetRow> OtherModeList { get; set; } = [];
        public double Offset { get; set; }
        public double Spec { get; set; }
        public bool NewMethod { get; set; }

        public PowerBinningSheet()
        {
        }

        protected PowerBinningSheet(PowerBinningSheet powerBinningSheet)
        {
            if (powerBinningSheet == null)
            {
                return;
            }

            SheetName = powerBinningSheet.SheetName;
            BinnedModeList = powerBinningSheet.BinnedModeList?.Select(x => x.Copy()).ToList() ?? [];
            OtherModeList = powerBinningSheet.OtherModeList?.Select(x => x.Copy()).ToList() ?? [];
            Offset = powerBinningSheet.Offset;
            Spec = powerBinningSheet.Spec;
            NewMethod = powerBinningSheet.NewMethod;
        }

        public PowerBinningSheet Copy()
        {
            return new PowerBinningSheet(this);
        }
    }
}
