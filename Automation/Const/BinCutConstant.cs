using System.Collections.Generic;

namespace Automation.Const
{
    public class BinCutConstant
    {
        public const string ConCpu = "Cpu";
        public const string ConGpu = "Gfx";
        public const string ConSoc = "Soc";
        public const string ConSpi = "Spi";
        public const string ConScan = "Scan";
        public const string ConDdr = "DDR";

        public const string VddBinningFailStop = "Vddbinning_Fail_Stop";
        public const string VddBinningFailStopFlag = "F_Vddbinning_Fail_Stop";
        public const string VddBinningInterpolationFail = "Other_Vddbinning_Interpolation_fail";
        public const string VddBinningInterpolationFailCs = "Vddbinning_Interpolation_fail";
        public const string VddBinningInterpolationFailFlag = "F_Other_Vddbinning_Interpolation_fail";
        public const string VddBinningInterpolationFailFlagCs = "F_Vddbinning_Interpolation_fail";
        public const string PowerBinningFail = "Vddbinning_Power_Binning_Fail_Stop";
        public const string PowerBinningFailFlag = "F_Power_Binning_Fail";

        public static readonly Dictionary<string, HashSet<string>> GradeJobMap = new Dictionary<string, HashSet<string>>
        {
            { "FT1", new HashSet<string> { "FT1", "FT2_25C", "WLFT", "WLFT1", "FT_ROOM", "RMA_ROOM"} },
            { "FT2", new HashSet<string> { "FT2", "FT2_85C", "WLFT2", "FT_HOT", "RMA_HOT" } }
        };
    }
}
