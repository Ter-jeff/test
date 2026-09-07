using System.Collections.Generic;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

namespace TestPlanLib.Static
{
    public static class VbtFunctionLibShared
    {
        public const string RtosBootUp = "bootup";
        public const string RtosRunScenarioT = "RTOS_RunScenario_T";

        //HardIP
        public const string VirName = "meas_vir_io_universal_func";
        public const string VirGpioTtrName = "meas_vir_io_universal_func_gpio_ttr";

        // RF
        public const string DvdcTrim3D = "wi_3dtrimuniversalfunc";
        public const string DvdcTrim = "wi_trimuniversalfunc";
        public const string RfTrim = "rftrim_doall_genericpgm";
        public const string RfTrim2D = "rf2dtrim_doall_genericpgm";
        public const string RfFunc = "rffunc_trx_universal";
        public const string RfHtolFunc = "rffunc_trx_htol_universal";
        public const string OtpCheckDefaultReal = "OTP_CHECK_DefaulReal";
        public const string OtpBlankBitCheck = "OTP_Blank_Bit_Check";
        public const string OtpBurn = "OTP_Burn";
        public const string OtpRealAllSetOtpData = "OTP_Read_ALL_SetOTPdata";

        // LCD
        public const string LcdTrim = "lcdtrimuniversalfunc";
        public const string LcdMeas = "measuniversalfunc";

        public static string FunctionalName { get; set; } = "functional_t_updated";
        public static string PowerUp { get; set; } = "PowerUp_Parallel";
        public static string PowerDown { get; set; } = "PowerDown_Parallel";
        public static string Ids { get; set; } = "ids_main_current";
        public static string IdsMathFunc { get; set; } = "IDSMathFunc";

        #region HardIP
        public static string VifName { get; set; } = "meas_freqvoltcurr_universal_func";
        public static string HardIpmtdTest { get; set; } = "hardip_mtd_test";
        public static List<string> EfusePrewriteFunctionList { get; set; } = ["hip_efuse_write", "hip_efuse_write_multi", "ids_efuse_write", "HardIPFuseWrite"];
        public static List<string> EfuseReadFunctionList { get; set; } = ["hip_efuse_read", "hip_efuse_read_multi", "HardIPFuseRead"];
        #endregion

        public static List<List<string>> ParamMappingList { get; set; } = [];
        public static Dictionary<string, int> GeneratedVbtFunctionDic { get; set; } = [];

        // CA1868: Changed to HashSet for performance and direct .Add guarding
        private static readonly HashSet<string> _missingParams = new(StringExtensions.IgnoreCase);

        private static HashSet<string>? _usedPatternList;
        public static HashSet<string> UsedPatternList
        {
            get => _usedPatternList ??= new HashSet<string>(StringExtensions.IgnoreCase);
            set => _usedPatternList = value;
        }

        public static void Clear()
        {
            ParamMappingList = [];
            GeneratedVbtFunctionDic = [];
            _missingParams.Clear();
            _usedPatternList = new HashSet<string>(StringExtensions.IgnoreCase);
        }

        public static void CheckMissingParameter(string functionName, string paramName, string block, string type)
        {
            string compositeKey = functionName + "&" + paramName;

            // CA1868 Fix: .Add returns false if item already existed, bypassing .Contains lookup rules
            if (_missingParams.Add(compositeKey))
            {
                string message = $"Missing Parameter in {functionName}({type}) : {paramName}";
                string[] args = [functionName, type, paramName];

                switch (block.ToLower())
                {
                    case "conti":
                        ErrorReportManager.AddError(BasicErrorType.E_MissingParameter_01, "", 0, 0, message, args);
                        break;
                    case "scan":
                        ErrorReportManager.AddError(ScanErrorType.E_MissingParameter_01, "", 0, 0, message, args);
                        break;
                    case "mbist":
                        ErrorReportManager.AddError(MbistErrorType.E_MissingParameter_01, "", 0, 0, message, args);
                        break;
                    case "rtos":
                        ErrorReportManager.AddError(RtosErrorType.E_MissingParameter_01, "", 0, 0, message, args);
                        break;
                    case "hardip":
                        ErrorReportManager.AddError(HardIpErrorType.E_MissingParameter_03, "", 0, 0, message, args);
                        break;
                    case "efuse":
                        ErrorReportManager.AddError(EFuseErrorType.E_MissingParameter_01, "", 0, 0, args);
                        break;
                    case "bincut":
                        ErrorReportManager.AddError(BinCutErrorType.E_MissingParameter_01, "", 0, 0, message, args);
                        break;
                    case "evs":
                        ErrorReportManager.AddError(EvsErrorType.E_MissingParameter_01, "", 0, 0, message, args);
                        break;
                    default:
                        ErrorReportManager.AddError(PreActionErrorType.E_MissingParameter_01, "", 0, 0, args);
                        break;
                }
            }
        }

        public static void ReportMissingLibraryModuleError(string block, string functionName, bool isCsharp)
        {
            string libraryType = isCsharp ? "C#" : "VBT";
            string message = $"The function: {functionName} can not find in {libraryType} library!";
            string[] args = [functionName, libraryType];

            switch (block.ToLower())
            {
                case "conti":
                    ErrorReportManager.AddError(BasicErrorType.E_MissVbtModule_01, "", 0, 0, message, args);
                    break;
                case "scan":
                    ErrorReportManager.AddError(ScanErrorType.E_MissVbtModule_01, "", 0, 0, message, args);
                    break;
                case "mbist":
                    ErrorReportManager.AddError(MbistErrorType.E_MissVbtModule_01, "", 0, 0, message, args);
                    break;
                case "rtos":
                    ErrorReportManager.AddError(RtosErrorType.E_MissingLibrary_01, "", 0, 0, message, args);
                    break;
                case "hardip":
                    break;
                case "efuse":
                    ErrorReportManager.AddError(EFuseErrorType.E_MissVbtModule_01, "", 0, 0, message, args);
                    break;
                case "bincut":
                    ErrorReportManager.AddError(BinCutErrorType.E_MissVbtModule_01, "", 0, 0, message, args);
                    break;
                case "evs":
                    ErrorReportManager.AddError(EvsErrorType.E_MissVbtModule_01, "", 0, 0, message, args);
                    break;
                case "basic":
                    ErrorReportManager.AddError(EvsErrorType.E_MissVbtModule_02, "", 0, 0, message, args);
                    break;
                case "htol":
                    ErrorReportManager.AddError(HtolErrorType.E_MissingLibrary_01, "", 0, 0, args);
                    break;
                default:
                    ErrorReportManager.AddError(PreActionErrorType.E_MissingLibrary_01, "", 0, 0, args);
                    break;
            }
        }
    }
}
