using CommonLib.ErrorReport.Base;

namespace CommonLib.ErrorReport.ErrorCodes
{
    public static class AutoAiErrorType
    {
        public static readonly ErrorCode E_MissingTimesetFile_01 = new(
            enumErrorCategory: EnumErrorCategory.AutoAi,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Timing,
            code: 1,
            template: "[TimeSet Latest] The timeset \"{0}\" in pattern dash board is not existed in pattern folder",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "The TimeSet specified in the Pattern Dashboard must exist in the pattern folder. " +
                    "Please verify the TimeSet name and ensure the corresponding file is available.");

        public static readonly ErrorCode E_MissingPatternFile_01 = new(
            enumErrorCategory: EnumErrorCategory.AutoAi,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 1,
            template: "[File Versions] The pattern \"{0}\" in pattern dash board is not existed in pattern folder",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "The pattern specified in the Pattern Dashboard must exist in the pattern folder. " +
                    "Please verify the pattern name and ensure the corresponding pattern file is available.");

        public static readonly ErrorCode E_MissingHeader_01 = new(
            enumErrorCategory: EnumErrorCategory.AutoAi,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Header,
            code: 1,
            template: "[Header] Missing the header: \" {0} \"",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Wrong header in input sheet, please check!");

        public static readonly ErrorCode W_MissingHeader_02 = new(
            enumErrorCategory: EnumErrorCategory.AutoAi,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Header,
            code: 2,
            template: "[Header] Missing the header: \" {0} \"",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Wrong header in input sheet, please check!");

        public static readonly ErrorCode E_MissingSheet_01 = new(
            enumErrorCategory: EnumErrorCategory.AutoAi,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.File,
            code: 1,
            template: "Can not find \" {0} \" sheet in input file !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "TBD");

        public static readonly ErrorCode W_FormatError_01 = new(
            enumErrorCategory: EnumErrorCategory.AutoAi,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 1,
            template: "[UseNotUse] The syntax \"{0}\" can not be identified, and this row will be ignored !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "The UseNotUse field must be set to either 'Use' or 'Not Use'. " +
                    "Please correct the value to ensure the row is processed.");

        public static readonly ErrorCode E_FormatError_02 = new(
            enumErrorCategory: EnumErrorCategory.AutoAi,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 2,
            template: "[AIType] The syntax \"{0}\" can not be identified, and treat this as Data log !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "The AIType field must be set to 'Data log', '1D', or '2D'. " +
                    "Unrecognized values will be treated as 'Data log'.");

        public static readonly ErrorCode E_FormatError_03 = new(
            enumErrorCategory: EnumErrorCategory.AutoAi,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 3,
            template: "[Data Logging Setting] The syntax \"{0}\" can not be identified, and treat this as NA !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "The Data Logging Setting field must be set to 'NA', 'DFCList...', or 'DFCStep...'. " +
                    "Unrecognized values will be treated as 'NA'.");

        public static readonly ErrorCode E_FormatError_04 = new(
            enumErrorCategory: EnumErrorCategory.AutoAi,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 4,
            template: "[Data Logging Setting] The DFC syntax \"{0}\" can not be identified, and treat this as NA !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "The DFCList setting must follow the format DFCList(value1,value2,...). " +
                    "Please verify the syntax and ensure the values are enclosed in parentheses.");

        public static readonly ErrorCode E_FormatError_05 = new(
            enumErrorCategory: EnumErrorCategory.AutoAi,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 5,
            template: "[Data Logging Setting] The DFC arg syntax \"{0}\" can not be identified, and treat this as NA !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "All values specified in DFCList must be valid integer numbers. " +
                    "Please verify the DFC values and remove any unsupported characters or formats.");

        public static readonly ErrorCode E_FormatError_06 = new(
            enumErrorCategory: EnumErrorCategory.AutoAi,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 6,
            template: "[Data Logging Setting] The DFC syntax \"{0}\" can not be identified, and treat this as NA !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "The DFCStep setting must follow the format DFCStep(start,step). " +
                    "Please verify the syntax and ensure the values are enclosed in parentheses.");

        public static readonly ErrorCode E_FormatError_07 = new(
            enumErrorCategory: EnumErrorCategory.AutoAi,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 7,
            template: "[Data Logging Setting] The DFC arg syntax \"{0}\" can not be identified, and treat this as NA !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "The DFCStep setting must contain exactly two integer values in the format DFCStep(start,step). " +
                    "Please verify the number and format of the arguments.");

        public static readonly ErrorCode E_FormatError_08 = new(
            enumErrorCategory: EnumErrorCategory.AutoAi,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 8,
            template: "[SELSRM] The syntax \"{0}\" can not be identified, and this will be ignored !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "The SELSRAM DSSC setting must start with 'SELSRM' or 'DSELSRM'. " +
                    "Please verify the number and format of the arguments.");

        public static readonly ErrorCode W_CanNotDetermineWhichSpecToUse_01 = new(
            enumErrorCategory: EnumErrorCategory.AutoAi,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Timing,
            code: 1,
            template: "\"[Mapping] Multi timeset for payload({0}) in base program: {1}",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Each payload in the base program must be associated with only one timeset. " +
                    "Please verify the payload configuration and remove any unexpected timeset assignments.");

        public static readonly ErrorCode E_CanNotDetermineWhichSpecToUse_02 = new(
            enumErrorCategory: EnumErrorCategory.AutoAi,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Timing,
            code: 2,
            template: "[Mapping] None of timeset for payload({0}) in base program",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Each payload in the base program must have a corresponding timeset. " +
                    "Please verify the payload configuration and ensure an timeset is defined.");

        public static readonly ErrorCode W_CanNotDetermineWhichSpecToUse_03 = new(
            enumErrorCategory: EnumErrorCategory.AutoAi,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.DcCategory,
            code: 1,
            template: "[Mapping] Multi dc category for payload({0}) in base program: {1}",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Each payload in the base program must be associated with only one DC category. " +
                    "Please verify the payload configuration and remove any unexpected DC category assignments.");

        public static readonly ErrorCode E_CanNotDetermineWhichSpecToUse_04 = new(
            enumErrorCategory: EnumErrorCategory.AutoAi,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.DcCategory,
            code: 1,
            template: "[Mapping] None of dc category for payload({0}) in base program",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Each payload in the base program must have a corresponding DC category. " +
                    "Please verify the payload configuration and ensure an DC category is defined.");
    }
}
