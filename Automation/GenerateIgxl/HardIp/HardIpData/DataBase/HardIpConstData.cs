using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Automation.GenerateIgxl.HardIp.HardIpData.DataBase
{
    public class HardIpConstData
    {
        public const string ReTestFlag = "F_Retest";
        public const string FreeRunClkDisable = "FreeRunClk_Disable";
        public const string FreeRunClkEnable = "FreeRunClk_Enable";

        public const string DigCapName = "MeasCapName";
        public const string CusStrDigCapData = "CUS_Str_DigCapData";
        public const string DictStoreCodeName = "Dict_Store_Code_Name";
        public const string TrimStoreName = "TrimStoreName"; //Changed from "TrimName" to "TrimStoreName" on 2017/3/8
        public const string TrimCodeStoreName = "TrimCodeStoreName";
        public const string FakePin = "FakePin";

        //Additional syntax for C#
        public const string TrimDictionaryStoreName = "TrimDictionaryStoreName";
        public const string DigCapDataCustomString = "digCapDataCustomString";

        # region default DC/AC/Timeset/Level
        public const string TimesetNwire = "TIMESET_nWire";
        public const string LevelNwire = "Levels_nWire";
        public const string HardIp = "HardIP";
        public const string AcCommonDefault = "Common";
        public const string LevelDefault = "Levels_HardIP";
        public const string IdsLevelDefault = "Levels_IDS";
        public const string RtosLevelDefault = "Levels_RTOS";
        # endregion

        #region Misc info Standard name
        public const string IdsNoFuse = "IDS_NoFuse";
        public const string Limit = "Limit";
        public const string VbtKey = "VBT";
        public const string Timing = "Timing";
        public const string Vbt = "Func|Trim|Meas|Trim"; //Change from VBT to Func on 2016/6/27
        //Func:XXX
        public const string Opcode = "opcode";
        public static readonly Regex RegOpcodeInMisc =
            new Regex(@"opcode\s*:", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        //Add on 2016/6/29
        public const string HighTemp = "High_Temp_Only";
        public const string RoomTemp = "Room_Temp_Only";
        public const string RelayOn = "Relay_On";
        public const string RelayOff = "Relay_Off";

        public const string NvOnly = "NV_Only_For_HLN_Flow|NvOnly"; //Change from NvOnly to NV_Only_For_HLN_Flow on 2016/6/27
        public const string HvOnly = "HV_Only_For_HLN_Flow|HvOnly"; //Added to support HV_Only_For_HLN_Flow on 2016/6/27
        public const string LvOnly = "LV_Only_For_HLN_Flow|LvOnly"; //Added to support LV_Only_For_HLN_Flow on 2016/6/27
        public const string RunNv = "Run_NV_Flow_Only";
        public const string RunLv = "Run_LV_Flow_Only";
        public const string RunHv = "Run_HV_Flow_Only";
        public const string RemoveNv = "RemoveNv"; //Change to Run_NV_Flow_Only on 2016/6/28
        public const string RemoveLv = "RemoveLv"; //Change to Run_LV_Flow_Only on 2016/6/28
        public const string RemoveHv = "RemoveHv"; //Change to Run_HV_Flow_Only on 2016/6/28
        public const string ReTest = "Fail_Retest";
        public const string PreNwireEnaOrDis = "FreerunClk";
        public const string FreerunClkEnableWord = PreNwireEnaOrDis + "Enable";
        public const string FreeRunClkDisableWord = PreNwireEnaOrDis + "Disable";
        public const string NoBin = "No_Fail_Flag_All|NoBin"; //Change from NoBin to No_Fail_Flag_All on 2016/6/24
        public const string FuseStage = "FuseStage";
        public const string OriginalTtrBranch = "OriginalTTRBranch";
        public const string MappingTtrBranch = "MappingTTRBranch";

        public const string NoBinUseLimit = "No_Fail_Flag_UseLimit|NoBinUseLimit";
        //Change from NoBinUseLimit to No_Fail_Flag_UseLimit on 2016/6/24

        public const string RemovePattern = "Do_Not_Generate|RemovePattern";
        //Change from RemovePattern to Fail_Retest on 2016/6/23

        public const string Manual = "Generate_But_Manually_Modify|Manual";
        public const string SkipCheck = "Skip_Pre_Check";
        public const string InstNameSubStr = "InstNameSubStr";
        public const string Calc = "Calc";
        public const string NvCalc = "NV@Calc";
        public const string LvCalc = "LV@Calc";
        public const string HvCalc = "HV@Calc";
        public const string CalcParameter = "CalcArg";
        public const string NvCalcParameter = "NV@CalcArg";
        public const string LvCalcParameter = "LV@CalcArg";
        public const string HvCalcParameter = "HV@CalcArg";

        public const string IgnorePatInfo = "Ignore_Patt_Comment";
        public const string IgnorePatMeasC = "Ignore_Patt_MeasC";
        public const string IgnorePatDigSrc = "Ignore_Patt_DigSrc";
        public const string IgnorePatBinOut = "Ignore_Patt_BinOut";
        public const string RealtimePatBinOut = "Realtime_Patt_BinOut";

        public const string IgnoreFlowLimit = "Ignore_Flow_Limit";
        public const string IgnoreSplitPOWERPowerMerge = "Ignore_Split_POWERPowerMerge";

        public const string EnablePattBinout = "Enable_Patt_Binout";
        public static readonly string[] IgrPatBinOutPrefix = { "MN_", "CZ_", "DD_", "HT_", "FA_", "DP_" };

        public const string StoreDigAll = "StoreDigAll"; // For HardIP MRR
        public const string RepeatLimit = "^Repeat_Limit";
        public const string NoPattern = "No_patt";

        public static readonly Regex RegInsInPatt =
            new Regex(@"^Instance[\s]*:[\s]*(?<InsName>[\w]+)",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static readonly Regex RegOpcodeInPatt =
            new Regex(@"^Opcode[\s]*:[\s]*(?<Opcode>[\w]+)(?:[\s]*:[\s]*(?<Parameter>.+))?",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public const string Block = "Block";
        public const string SubBlockName = "SubBlock";
        public const string SubBlockCzName = "SubBlockCZ";
        public const string InstSpecialSetup = "InstSpecialSetting";
        public const string MtrLoop = "MTRLoop";
        //Added on 2016/11/24 by Jackie
        public const string ExecCond = "ExecCond";

        public const string CallExtraFlow = "Call";
        // To add extra flow before pattern test for Flow_Nwire_default on 2017/08/31

        public const string KeepDsscOut = "Disable_MeasC_Split";
        public const string Cz2Only = "CZ2_Only";

        public const string Slope = "Slope";
        public const string SweepRange = "SweepRange";
        public const string LoadFile = "LoadFile";
        public const string TransitionSlope = "TransitionSlope";
        public const string RefSubBlock = "Ref_SubBlock";
        public const string ForLoop = "ForLoop";
        public const string Level = "Level";
        public const string EnableWord = "Enableword";
        public const string CheckSelsram = "CheckSelsram";

        public const string PrefixAtgRelay = "AtgRelay_";

        #endregion

        #region Naming rules

        public const string RegNand = "(_nan)|(_ids)";
        public const string RegSpi = "_spi";
        public const string RegCzPattern = "cz_";

        public const string HardipBinEnable = "HardIPBin";
        public const string EnvTtr = "TTR";
        public const string PrefixHardIp = "HARDIP_";
        public const string PrefixDcTest = "DCTEST_";
        public const string PrefixWireless = "WIRELESS_";
        public const string PrefixHardIpFailAction = "F";
        public const string SuffixHardIpFailAction = "_Flag";
        public const string SubfixHipEfuseReadBinTableName = "Fuse_Read_Non_Zero_Check";
        public const string GpioBlockName = "GPIO";
        public const string TmpsBlockName = "TMPS";
        public const string BinFlowFlag = "Bin";
        public const string PrefixInsSheetByVoltage = "TestInst_HARDIP_";
        public const string LabelAll = "All";
        public const string LabelNv = "NV";
        public const string LabelLv = "LV";
        public const string LabelHv = "HV";
        public static readonly List<string> LabelVolList = new List<string> { LabelNv, LabelLv, LabelHv };
        public const string LabelChar = "Char";
        public const string SelectMax = "Max";
        public const string SelectMin = "Min";
        public const string SelectTyp = "Typ";
        public const string PrefixReTest = "retest_";
        public const string RegTestSequence = "^TestSequence:";
        #endregion

        public static readonly List<string> IgnoredHardIpSheetList =
            new List<string> { "HardIP_ErrorReport", "HardIP_DC" };

        #region Shmoo

        public const string PrefixShmooSetupName = "CZ2_";
        public static readonly Regex RegShmoo = new Regex(@"(xshmoo|yshmoo)\s*\((?<ShmooStr>[^)]+)\)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        public static readonly Regex RegVtShmoo = new Regex(@"(xVtshmoo|YVtshmoo)\s*\((?<ShmooStr>[^)]+)\)*", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        #endregion

        public static readonly Regex RegMergeIndex = new Regex(@"Merge_\d+", RegexOptions.Compiled);

        public static readonly List<string> MiscKeyList =
            new List<string> { Limit, Timing, Vbt, VbtKey, HighTemp,
                RoomTemp, RelayOn, RelayOff, RelayOn.Replace("_",""), RelayOff.Replace("_",""),
                NvOnly, HvOnly, LvOnly, RunNv, RunLv, RunHv, RemoveHv, RemoveLv,
                RemoveNv, ReTest, FreeRunClkDisableWord,
                FreerunClkEnableWord, NoBin, NoBinUseLimit,
                RemovePattern, Manual, SkipCheck, Opcode, InstNameSubStr,
                Calc, NvCalc, HvCalc, LvCalc, CalcParameter, NvCalcParameter, HvCalcParameter, LvCalcParameter,
                DigCapName, ExecCond, SubBlockName, SubBlockCzName, CallExtraFlow, RepeatLimit ,
                Slope, SweepRange, LoadFile, TransitionSlope, IgnorePatMeasC, IgnorePatInfo, IgnorePatBinOut, EnablePattBinout, RealtimePatBinOut,
                InstSpecialSetup, MtrLoop, StoreDigAll, KeepDsscOut, RefSubBlock, ForLoop, Cz2Only, Level, EnableWord,
                CheckSelsram, OriginalTtrBranch, MappingTtrBranch };
    }
}
