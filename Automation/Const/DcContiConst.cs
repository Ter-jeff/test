using System.Text.RegularExpressions;

namespace Automation.Const
{
    public class DcContiConst
    {
        public const string VbtFuncNameOpenShort = "ppmu_continuity";
        public const string VbtFuncNamePowerShort = "p2p_short_Power_FVMI";
        public const string P2PShortPowerFvmiParallel = "p2p_short_Power_FVMI_Parallel";
        public const string CSharpFuncNamePowerShort = "PowerShort";
        public const string VbtFuncNamePowerSense = "PowerSensePins_continuity";
        public const string VbtFuncNameSenseImpedance = "PPMU_Impedance_Function";
        public const string CSharpFuncNameSenseImpedance = "PPMUImpedance";
        public const string VbtFuncNameGndSense = "GndSensePins_continuity";
        public const string CSharpFuncNameSensePinConti = "SensePinContinuity";
        public const string VbtFuncNameRelayControl = "Relay_Control";
        public const string CsFuncNameRelayControl = "ControlRelay";
        public const string VbtFuncNameFunctionalT = "Functional_T_updated";
        public const string CSharpFuncNameFunctionalT = "FuncTestMain";
        public const string CSharpFuncNameWalkingZContinuity = "WalkingZContinuity";
        public const string CSharpFuncNameFuncTestCharMain = "FuncTestCharMain";
        public const string CSharFuncNamePrintShmooInfoMain = "PrintShmooInfoMain";
        public const string VbtFuncNameEvsRampPowerPa = "EVS_Static_Power_Ramp_PAEVS";
        public const string VbtFuncNameEvsRampPower = "EVS_Static_Power_Ramp";
        public const string CSharpFuncNameEvsRampPower = "EVSStaticPowerRamp";
        public const string VbtFuncNameAutoZ = "AutoZ_Continuity";
        public const string CSharpFuncNameAutoZ = "AutoZContinuity";
        public const string VbtCreateOverlayCpm = "Create_Overlay_CPM";
        public const string VbtCpmEFuseRead = "CPM_eFuse_Read";
        public const string CSharpCpmEFuseRead = "Copy_From_Fuse_To_Flag";
        public const string VbtCpmEFuseWrite = "CPM_eFuse_Write";
        public const string CSharpCpmEFuseWrite = "Set_CPM";
        public const string VbtFuncNameSetPpmuClamp = "Set_PPMU_Clamp";
        public const string CsFuncNameSetPpmuClamp = "ConfigurePPMUClamp";
        public const string VbtContiWalkingZ = "Conti_WalkingZ";

        public const string PinGroupAllDcvi = "All_DCVI";
        public const string PinGroupAllDigitalPowerUp = "All_Digital_PowerUp";
        public const string FlagNamePowerShort = "F_powershort";
        public const string FlagNamePowerOpen = "F_poweropen";
        public const string FlagNamePowerSense = "F_powersense";
        public const string FlagNameGroundSense = "F_groundsense";
        public const string FlagNameOpen = "F_open";
        public const string FlagNameShort = "F_short";
        public const string FlagNameAutoZ = "F_AutoZ_check";
        public const string FlagNameCres = "F_CRES";

        public const string BinNameOpenShort = "Bin_DC_open_short";
        public const string BinNameOpen = "Bin_DC_open";
        public const string BinNameShort = "Bin_DC_short";
        public const string BinNamePowerShort = "Bin_DC_powershort";
        public const string BinNamePowerOpen = "Bin_DC_poweropen";
        public const string BinNamePowerSense = "Bin_DC_powersense";
        public const string BinNameCres = "Bin_DC_powerResistance";

        public const string BinNameGndSense = "Bin_DC_GND_sense";
        public const string BinNameAutoZCheck = "Bin_AutoZ_check";

        public const string BinPowerUpAlarm = "Bin_PowerUp_Alarm";
        public const string FPowerUpAlarm = "F_PowerUp_Alarm";

        public const string RelayWaitTime = "0.003";

        public static readonly Regex FailFlagRegex = new Regex("^F_", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        public static readonly Regex BinDcRegex = new Regex("^Bin_DC_", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    }
}
