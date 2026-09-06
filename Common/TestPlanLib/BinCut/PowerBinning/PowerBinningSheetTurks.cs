namespace TestPlanLib.BinCut.PowerBinning
{
    public class PowerBinningSheetTurks : PowerBinningSheet
    {
        public PowerBinningSheetTurks(string name)
        {
            SheetName = name;
            BinnedModeList = [];
            OtherModeList = [];
        }

        public PowerBinningSheetTurks()
        {
            BinnedModeList = [];
            OtherModeList = [];
        }
    }
}
