using CommonLib.ErrorReport.Base;

namespace CommonLib.ErrorReport.ErrorCodes
{
    public static class HarvestErrorType
    {
        /// <summary>Template: "Missing job name in harvesting truth table sheet name: "{0}"."</summary>
        /// <remarks>{0} = sheet name</remarks>
        public static readonly ErrorCode E_InvalidFormat_01 = new(
            enumErrorCategory: EnumErrorCategory.Harvest,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 1,
            template: "Missing job name in harvesting truth table sheet name: \"{0}\".",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please ensure the sheet name contains the job name.");

        /// <summary>Template: "Invalid content "{0}". Only X, numeric values, and ranges are allowed."</summary>
        /// <remarks>{0} = content</remarks>
        public static readonly ErrorCode E_InvalidFormat_02 = new(
            enumErrorCategory: EnumErrorCategory.Harvest,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 2,
            template: "Invalid content \"{0}\"."
                    + " Only X, numeric values, and ranges are allowed.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please enter only X, a number, or a valid range.");

        /// <summary>Template: "Cannot fill range {0} for single flag : {1}."</summary>
        /// <remarks>{0} = range value and flag name context, {1} = flag name</remarks>
        public static readonly ErrorCode E_InvalidFormat_03 = new(
            enumErrorCategory: EnumErrorCategory.Harvest,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 3,
            template: "Cannot fill range {0} for single flag : {1}.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Unable to define the range. Ensure the flag defined in the top row is defined with a valid range.");

        /// <summary>Template: "Header order is incorrect. '{0}' must appear after '{1}'."</summary>
        /// <remarks>{0} = column or field name, {1} = column or field name</remarks>
        public static readonly ErrorCode E_InvalidOrder_01 = new(
            enumErrorCategory: EnumErrorCategory.Harvest,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Order,
            code: 1,
            template: "Header order is incorrect. '{0}' must appear after '{1}'.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify that headers in HarvestingTruthTable table follow the required order.");

        /// <summary>Template: "Field {0} with \"Real\" type was not found in EFUSE_BitDef_Table."</summary>
        /// <remarks>{0} = mismatch context description</remarks>
        public static readonly ErrorCode E_MissingField_01 = new(
            enumErrorCategory: EnumErrorCategory.Harvest,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Field,
            code: 1,
            template: "Field {0} with \"Real\" type was not found in EFUSE_BitDef_Table.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please verify that the field exists in EFUSE_BitDef_Table and is defined with type \"Real\".");

        /// <summary>Template: "Read Fuse flag {0} is defined in CP2 but is not defined in this truth table."</summary>
        /// <remarks>{0} = flag name</remarks>
        public static readonly ErrorCode E_MissingFlag_01 = new(
            enumErrorCategory: EnumErrorCategory.Harvest,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Flag,
            code: 1,
            template: "Read Fuse flag \"{0}\" is defined in CP2 but is not defined in this truth table.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please ensure that flag definitions are consistent across the both harvest truth tables.");

        /// <summary>Template: "Cannot found header: {0}, please check."</summary>
        /// <remarks>{0} = header name</remarks>
        public static readonly ErrorCode E_MissingHeader_01 = new(
            enumErrorCategory: EnumErrorCategory.Harvest,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Header,
            code: 1,
            template: "Cannot found header: {0}, please check.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open HarvestingTruthTable and ensure the missing column header(s) are present and spelled correctly.");

        /// <summary>Template: "Read Fuse Flag "{0}" is defined in this truth table but is missing in CP2."</summary>
        /// <remarks>{0} = flag name</remarks>
        public static readonly ErrorCode E_RedundantFlag_01 = new(
            enumErrorCategory: EnumErrorCategory.Harvest,
            enumErrorBehavior: EnumErrorBehavior.Redundant,
            enumErrorTarget: EnumErrorTarget.Flag,
            code: 1,
            template: "Read Fuse Flag \"{0}\" is defined in this truth table but is missing in CP2.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please ensure that flag definitions are consistent across the both harvest truth tables.");

        /// <summary>Template: "Special character(s) are not allowed: {0}."</summary>
        /// <remarks>{0} = condition value</remarks>
        public static readonly ErrorCode W_InvalidFormat_01 = new(
            enumErrorCategory: EnumErrorCategory.Harvest,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 1,
            template: "Special character(s) are not allowed: {0}.",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Use only letters, numbers, and underscores (_).");

        /// <summary>Template: "Cannot found any matching pattern in Mapping_DigitalCores table. Pattern : \"{0}\""</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode W_MismatchPattern_01 = new(
            enumErrorCategory: EnumErrorCategory.Harvest,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 1,
            template: "Cannot found any matching pattern in Mapping_DigitalCores table. Pattern : \"{0}\"",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Because PatternPinGroup is specified in this instance,"
                    + " verify that the corresponding Pattern Name or Pattern Name keyword exists in Mapping_DigitalCores table.");

        /// <summary>Template: "Cannot found header: {0}, please check."</summary>
        /// <remarks>{0} = header name</remarks>
        public static readonly ErrorCode W_MissingHeader_01 = new(
            enumErrorCategory: EnumErrorCategory.Harvest,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Header,
            code: 1,
            template: "Cannot found header: {0}, please check.",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open HarvestingTruthTable and ensure the missing column header(s) are present and spelled correctly.");

        /// <summary>Template: "Cannot found header \"FUSING({0})\", please check."</summary>
        /// <remarks>{0} = job name</remarks>
        public static readonly ErrorCode W_MissingHeader_02 = new(
            enumErrorCategory: EnumErrorCategory.Harvest,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Header,
            code: 2,
            template: "Cannot found header \"FUSING({0})\", please check.",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open HarvestingTruthTable and add the FUSING column for the flagged job. "
                    + "Ensure the column name matches the expected format 'FUSING(<job>)'.");
    }
}
