using CommonLib.ErrorReport.Base;

namespace CommonLib.ErrorReport.ErrorCodes
{
    public static class ClockCheckErrorType
    {
        /// <summary>Template: "Type/INIT sequence: "{0}" does not have a matching entry in \"Sub Flow\" column of Instance_Clock_Check sheet."</summary>
        /// <remarks>{0} = clock check type</remarks>
        public static readonly ErrorCode E_MissingFlow_01 = new(
            enumErrorCategory: EnumErrorCategory.ClockCheck,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Flow,
            code: 1,
            template: "Type/INIT sequence: \"{0}\" does not have a matching entry in \"Sub Flow\" column of Instance_Clock_Check sheet.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the Type/INIT sequence value in Clock_Check sheet has a corresponding entry in the \"Sub Flow\" column of of Instance_Clock_Check sheet.");

        /// <summary>Template: "No library setting was found based on MiscInfo column value "{0}"."</summary>
        /// <remarks>{0} = misc info content</remarks>
        public static readonly ErrorCode E_MissingLibrary_01 = new(
            enumErrorCategory: EnumErrorCategory.ClockCheck,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Library,
            code: 1,
            template: "No library setting was found based on MiscInfo column value \"{0}\".",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify that the MiscInfo column value is correct and that a corresponding library setting is defined.");

        /// <summary>Template: "Instance not include from clock check"</summary>
        public static readonly ErrorCode E_MissingSetting_01 = new(
            enumErrorCategory: EnumErrorCategory.ClockCheck,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Setting,
            code: 1,
            template: "This instance does not have a matching entry in \"{0}\" sheet.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify that each Sub Flow in the Instance_Clock_Check sheet exists as a Type/INIT sequence in Clock_Check sheet.");

        /// <summary>Template: "Burst pattern instance uses fail flag \"{1}\"."</summary>
        /// <remarks>{0} = fail flag</remarks>
        public static readonly ErrorCode I_BurstPatJudgeFlag_01 = new(
            enumErrorCategory: EnumErrorCategory.ClockCheck,
            enumErrorBehavior: EnumErrorBehavior.Info,
            enumErrorTarget: EnumErrorTarget.Flag,
            code: 1,
            template: "Burst pattern instance uses fail flag \"{0}\".",
            enumErrorLevel: EnumErrorLevel.Info,
            guidance: "Verify the burst pattern judge flag definition. "
                    + "Ensure the flag name and condition are valid and match the burst pattern flow.");
    }
}
