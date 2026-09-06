using CommonLib.ErrorReport.Base;

namespace CommonLib.ErrorReport.ErrorCodes
{
    public static class EvsErrorType
    {
        /// <summary>Template: "Missing Parameter in {0}({1}) : {2}"</summary>
        /// <remarks>{0} = parameter name, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_MissingParameter_01 = new(
            enumErrorCategory: EnumErrorCategory.Evs,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 1,
            template: "Missing Parameter in {0}({1}) : {2}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the EVS sheet and add the required parameter. "
                    + "Check the EVS specification to confirm which parameters are mandatory.");

        /// <summary>Template: "The VBT function: {0} can not find in {1} library!"</summary>
        /// <remarks>{0} = module name, {1} = ?</remarks>
        public static readonly ErrorCode E_MissVbtModule_01 = new(
            enumErrorCategory: EnumErrorCategory.Evs,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Module,
            code: 1,
            template: "The VBT function: {0} can not find in {1} library!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the VBT module name against the expected module list. "
                    + "Add the missing module or correct the reference.");

        /// <summary>Template: "The VBT function: {0} can not find in {1} library!"</summary>
        /// <remarks>{0} = module name, {1} = ?</remarks>
        public static readonly ErrorCode E_MissVbtModule_02 = new(
            enumErrorCategory: EnumErrorCategory.Evs,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Module,
            code: 2,
            template: "The VBT function: {0} can not find in {1} library!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the VBT module name against the expected module list for the basic block. "
                    + "Add the missing module definition or correct the reference.");

        /// <summary>Template: "The pattern : {0} format illegal"</summary>
        /// <remarks>{0} = pattern name</remarks>
        public static readonly ErrorCode E_Pattern_01 = new(
            enumErrorCategory: EnumErrorCategory.Evs,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 1,
            template: "The pattern : {0} format illegal",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the pattern name and its definition in the EVS sheet. "
                    + "Check that the pattern exists and is correctly referenced.");
    }
}
