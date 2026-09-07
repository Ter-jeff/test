using System;
using System.Collections.Generic;

using BinCutScriptLib.ByProject;
using BinCutScriptLib.SetFunction.SetStartStep;

using CommonLib.Extension;

using TestPlanLib.BinCut.BinCutConfig;

namespace BinCutScriptLib.Static
{
    public static class BinCutConfig
    {
        public const int MaxSitForRetest = 10;

        public static List<string> ActiveEnableWords { get; set; } = [];
        public static bool MultiFstpEnable { get; set; }
        public static Dictionary<string, EnumPowerType> PowerType { get; set; } = new(StringExtensions.IgnoreCase);
        public static List<Tuple<string, string>> IdsNames { get; set; } = [];
        public static Dictionary<string, string> DomainInOtherRail2Power { get; set; } = new(StringExtensions.IgnoreCase);
        public static Dictionary<string, bool> PowerBinningConfig { get; set; } = new(StringExtensions.IgnoreCase);
        public static Dictionary<string, string> PowerBiningMappingTable { get; set; } = new(StringExtensions.IgnoreCase);
        public static List<InterpolatioNode> InterpolationNodes { get; set; } = [];

        public static bool IsSimulationMode { get; set; }
        public static bool IsDebugPrint { get; set; }

        public static bool EfuseFloor { get; set; }

        //BinCut config
        public static bool IsDoAll { get; set; }
        public static bool IsOverrideFailstop { get; set; }
        public static string VddbinningFailStopFlag { get; set; } = "";
        public static string VddbinningPowerBinningFailStopFlag { get; set; } = "";
        public static string VddbinningIdsFailFlag { get; set; } = "";
        public static string VddbinningInterpolationFailFlag { get; set; } = "";
        public static bool IsEnablePowerBinningHarvest { get; set; } = true;
        public static string GradeSearchMethodSelected { get; set; } = "";
        public static bool FlagIdsDistributionEnable { get; set; }
        public static bool StartEqn { get; set; }
        public static bool IsCompareByProductValueOnly { get; set; }
        public static bool FlagUseNewInterpolationMonotonicity { get; set; }
        public static bool FlagUseCofInstance { get; set; }
        public static bool FlagUseNewInterpolationMonoton { get; set; }
        public static bool VddbinCofStepInheritance { get; set; }
        public static bool FlagCp1IdsNotUseHarvResult { get; set; }
        public static bool FlagSkipPrintingSafeVoltage { get; set; }
        public static bool FlagSkipPrintingSelSrmDsscInfo { get; set; }
        public static bool FlagEnablePreAdjustVddbinning { get; set; }
        public static bool? IsBinCutJobForStepSearch { get; set; } = false;
        //For COF
        public static bool VddbinCofStepInheritanceNewLogic { get; set; }
        public static bool DebugBinCutCofE1 { get; set; }
        public static bool DebugBinCutCofStored { get; set; }
        //For GPU domain cross domain inherit
        public static bool FlagCrossDomainCheckEnable { get; set; }
        //For check havrestbin name from pwrbinning sheet
        public static bool FlagPowerBinningHarvestBinFieldName { get; set; }
        public static string MappingTestJobNameToBincutJobName { get; set; } = string.Empty;

        public static string NonBinningRail { get; set; } = "Non_Binning_Rail";
        public static string VddBinningDefAppA1 { get; set; } = "Vdd_Binning_Def_appA_1";
        public static string VddBinningDefAppA2 { get; set; } = "Vdd_Binning_Def_appA_2";
        public static string VddBinningDefAppA3 { get; set; } = "Vdd_Binning_Def_appA_3";
        public static string PowerBinningSeqSheet { get; set; } = "PwrBinning_V";
        //vbt name is Non_Binning_Rail
        //For new Csharp log
        public static string BincutAteConditionNon { get; set; } = "bincut_ate_condition_non";
        //vbt name is Vdd_Binning_Def_appA_1
        public static string BincutEqnAppA { get; set; } = "bincut_eqn_appA";

        public static string SafeVoltageFollowPayload { get; set; } = string.Empty;
        public static ProjectConfig ProjectConfig { get; set; } = new();
        //For set initail step to bin1 eq1 when BIN4_CAND flag turn on
        public static bool FlagVddbinHarvestBin4RunBin1Eqn1 { get; set; }
        public static bool FlagNewSelsrmVthEnable { get; set; }
        //For totx determination 
        public static bool FlagT0TxHotFormat { get; set; }
        public static bool FlagT0TxRoomFormat { get; set; }
        public static bool FlagSyncUpDcvsOutputEnable { get; set; }
        public static Dictionary<string, List<string>> HarvFlags { get; set; } = [];

        public static bool EnableWordLogLevelDebug { get; set; }
        public static bool EnableWordLogLevelNpi { get; set; }
        public static bool EnableWordLogLevelMp { get; set; }
        public static bool DebugPrintFlag { get; set; }

        public static void Initialize()
        {
            NonBinningRail = "Non_Binning_Rail";
            VddBinningDefAppA1 = "Vdd_Binning_Def_appA_1";
            VddBinningDefAppA2 = "Vdd_Binning_Def_appA_2";
            VddBinningDefAppA3 = "Vdd_Binning_Def_appA_3";
            PowerBinningSeqSheet = "PwrBinning_V";

            BincutAteConditionNon = "bincut_ate_condition_non";
            BincutEqnAppA = "bincut_eqn_appA";

            ActiveEnableWords = [];
            MultiFstpEnable = false;
            PowerType = new Dictionary<string, EnumPowerType>(StringExtensions.IgnoreCase);
            IdsNames = [];
            DomainInOtherRail2Power = new Dictionary<string, string>(StringExtensions.IgnoreCase);
            PowerBinningConfig = new Dictionary<string, bool>(StringExtensions.IgnoreCase);
            PowerBiningMappingTable = new Dictionary<string, string>(StringExtensions.IgnoreCase);
            InterpolationNodes = [];

            IsSimulationMode = false;
            IsDebugPrint = false;
            EfuseFloor = false;
            IsDoAll = false;
            IsOverrideFailstop = false;
            VddbinningFailStopFlag = "";
            VddbinningPowerBinningFailStopFlag = "";
            VddbinningIdsFailFlag = "";
            VddbinningInterpolationFailFlag = "";
            IsEnablePowerBinningHarvest = true;
            GradeSearchMethodSelected = "";
            FlagIdsDistributionEnable = false;
            StartEqn = false;
            IsCompareByProductValueOnly = false;
            FlagUseNewInterpolationMonotonicity = false;
            FlagUseCofInstance = false;
            VddbinCofStepInheritance = false;
            FlagCp1IdsNotUseHarvResult = false;
            FlagSkipPrintingSafeVoltage = false;
            FlagSkipPrintingSelSrmDsscInfo = false;
            FlagEnablePreAdjustVddbinning = false;
            IsBinCutJobForStepSearch = false;
            VddbinCofStepInheritanceNewLogic = false;
            DebugBinCutCofE1 = false;
            DebugBinCutCofStored = false;
            FlagCrossDomainCheckEnable = false;
            FlagPowerBinningHarvestBinFieldName = false;
            MappingTestJobNameToBincutJobName = "";

            SafeVoltageFollowPayload = "";
            ProjectConfig = new ProjectConfig();
            FlagVddbinHarvestBin4RunBin1Eqn1 = false;
            FlagNewSelsrmVthEnable = false;
            FlagT0TxHotFormat = false;
            FlagT0TxRoomFormat = false;
            FlagSyncUpDcvsOutputEnable = false;
            HarvFlags = [];

            EnableWordLogLevelDebug = false;
            EnableWordLogLevelNpi = false;
            EnableWordLogLevelMp = false;
            DebugPrintFlag = false;
        }

        public static void GetProjectConfig(string projectName)
        {
            switch (projectName)
            {
                case "Ellis":
                    {
                        ProjectConfig = new Ellis();
                        break;
                    }
                case "JadeCdie":
                    {
                        ProjectConfig = new JadeCdie();
                        break;
                    }
                case "Staten":
                    {
                        ProjectConfig = new Staten();
                        break;
                    }
                case "Crete":
                    {
                        ProjectConfig = new Crete();
                        break;
                    }
                case "Bora":
                    {
                        ProjectConfig = new Bora();
                        break;
                    }
                case "JC-Chop":
                    {
                        ProjectConfig = new JcChop1();
                        break;
                    }
                case "Rhodes":
                    {
                        ProjectConfig = new Rhodes();
                        break;
                    }
                case "Caicos":
                    {
                        ProjectConfig = new Caicos();
                        break;
                    }
                case "Brava_C":
                    {
                        ProjectConfig = new BravaC();
                        break;
                    }
                case "Brava_S":
                    {
                        ProjectConfig = new BravaC();
                        break;
                    }
                case "Donan":
                    {
                        ProjectConfig = new Donan();
                        break;
                    }
                default:
                    {
                        ProjectConfig = new ProjectConfig();
                        break;
                    }
            }
        }
    }
}
