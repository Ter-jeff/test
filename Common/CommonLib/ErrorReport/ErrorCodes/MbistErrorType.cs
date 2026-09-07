using CommonLib.ErrorReport.Base;

namespace CommonLib.ErrorReport.ErrorCodes
{
    public static class MbistErrorType
    {
        /// <summary>Template: "Column[Voltage]: Same label with previous row, but Voltage mismatch."</summary>
        public static readonly ErrorCode E_BurstInfoMismatch_01 = new(
            enumErrorCategory: EnumErrorCategory.Mbist,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Voltage,
            code: 1,
            template: "Column[Voltage]: Same label with previous row, but Voltage mismatch.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the MBIST sheet and check the Voltage column for the flagged row. "
                    + "Rows sharing the same label must have identical Voltage values.");

        /// <summary>Template: "Column[FailBranch]: Same label with previous row, but FailBranch mismatch."</summary>
        public static readonly ErrorCode E_BurstInfoMismatch_02 = new(
            enumErrorCategory: EnumErrorCategory.Mbist,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Logic,
            code: 2,
            template: "Column[FailBranch]: Same label with previous row, but FailBranch mismatch.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the MBIST sheet and check the FailBranch column for the flagged row. "
                    + "Rows sharing the same label must have identical FailBranch values.");

        /// <summary>Template: "Column[TimeSet]: Same label with previous row, but TimeSet mismatch."</summary>
        public static readonly ErrorCode E_BurstInfoMismatch_03 = new(
            enumErrorCategory: EnumErrorCategory.Mbist,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Timing,
            code: 3,
            template: "Column[TimeSet]: Same label with previous row, but TimeSet mismatch.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the MBIST sheet and check the TimeSet column for the flagged row. "
                    + "Rows sharing the same label must have identical TimeSet values.");

        /// <summary>Template: "Argument \"patternBeforeWait\" for retention will be empty."</summary>
        public static readonly ErrorCode E_BurstInfoMismatch_04 = new(
            enumErrorCategory: EnumErrorCategory.Mbist,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Parameter,
            code: 4,
            template: "Argument \"patternBeforeWait\" for retention will be empty.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the retention pattern configuration and ensure 'patternBeforeWait' is assigned a valid pattern. "
                    + "Verify the burst info definition includes this argument.");

        /// <summary>Template: "{0} should be last row in {1}!!!"</summary>
        /// <remarks>{0} = rule or condition description, {1} = ?</remarks>
        public static readonly ErrorCode E_Business_01 = new(
            enumErrorCategory: EnumErrorCategory.Mbist,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Logic,
            code: 5,
            template: "{0} should be last row in {1}!!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "The action column contest for the last row of Mbist sheet need to be \"PASS\" . ");

        /// <summary>Template: "Cannot find any payload from pass branch : {0} , selsram will use init pattern category."</summary>
        /// <remarks>{0} = context description</remarks>
        public static readonly ErrorCode E_Business_02 = new(
            enumErrorCategory: EnumErrorCategory.Mbist,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 6,
            template: "Cannot find any payload from pass branch : {0} , selsram will use init pattern category.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Review the MBIST business logic constraint that was violated. "
                    + "Check the non-logical data definition and correct the configuration.");

        /// <summary>Template: "Error! Can not found DC Spec by {0} in {1}"</summary>
        /// <remarks>{0} = context description, {1} = ?</remarks>
        public static readonly ErrorCode E_Business_03 = new(
            enumErrorCategory: EnumErrorCategory.Mbist,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.DcCategory,
            code: 7,
            template: "Error! Can not found DC Spec by {0} in {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "DC category defined in flagged row can not be find in voltage table, add DC category into voltage table");

        /// <summary>Template: "Missing Parameter in {0}({1}) : {2}"</summary>
        /// <remarks>{0} = parameter name, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_MissingParameter_01 = new(
            enumErrorCategory: EnumErrorCategory.Mbist,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Parameter,
            code: 8,
            template: "Missing Parameter in {0}({1}) : {2}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Missing argument in C# library function, check C# library version. ");

        /// <summary>Template: "The VBT function: {0} can not find in {1} library!"</summary>
        /// <remarks>{0} = module name, {1} = ?</remarks>
        public static readonly ErrorCode E_MissVbtModule_01 = new(
            enumErrorCategory: EnumErrorCategory.Mbist,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Module,
            code: 9,
            template: "The VBT function: {0} can not find in {1} library!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Missing C# library function, check C# library version. ");
    }
}
