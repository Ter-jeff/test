using CommonLib.ErrorReport.Base;

namespace CommonLib.ErrorReport.ErrorCodes
{
    public static class FlowMainErrorType
    {
        /// <summary>Template: "No data found in range: A1:J10."</summary>
        public static readonly ErrorCode E_InvalidFormat_01 = new(
            enumErrorCategory: EnumErrorCategory.FlowMain,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 1,
            template: "No data found in range: A1:J10.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify that the correct input document was provided and that the Flow Main sheet is not empty.");

        /// <summary>Template: "Referenced Flow Main Sheet '{0}' does not exist in the test plan."</summary>
        /// <remarks>{0} = Flow Main Sheet name</remarks>
        public static readonly ErrorCode E_MissingDocument_01 = new(
            enumErrorCategory: EnumErrorCategory.FlowMain,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Document,
            code: 1,
            template: "Referenced Flow Main Sheet '{0}' does not exist in the test plan.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify that the specified Flow Main Sheet exists in the test plan."
                    + "Check for any sheet name mismatches or missing sheets.");

        /// <summary>Template: "The Sheet Name: "{0}"{1} defined for Job "{2}" does not match any existing sheet or source."</summary>
        /// <remarks>{0} = target sheet name, {1} = target sub flow msg, {2} = job</remarks>
        public static readonly ErrorCode W_MismatchFlow_01 = new(
            enumErrorCategory: EnumErrorCategory.FlowMain,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Flow,
            code: 1,
            template: "The Sheet Name: \"{0}\"{1} defined for Job \"{2}\" does not match any existing sheet or source.",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Verify that the specified mapping matches an existing sheet or source in the test plan.");
    }
}
