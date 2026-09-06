using CommonLib.ErrorReport.Base;

namespace CommonLib.ErrorReport.ErrorCodes
{
    public static class PatternMissing
    {
        /// <summary>Template: "{0} doesn't exist in pattern folder."</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingPatternFile_01 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 1,
            template: "{0} doesn't exist in pattern folder.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the pattern file path and name against the test-plan entry. "
                      + "Run the pattern compilation step if the file is generated. "
                      + "Check source-control to confirm the file was committed and not accidentally excluded.");

        /// <summary>Template: "Pattern's version changes from {0} to {1}"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_VersionChange_01 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 1,
            template: "Pattern's version changes from {0} to {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "The pattern was regenerated with a different version than the test plan expects. "
                      + "Re-run the full pattern generation flow and update the test plan to reference the new version, "
                      + "or roll back the pattern source to the expected version.");
    }
}
