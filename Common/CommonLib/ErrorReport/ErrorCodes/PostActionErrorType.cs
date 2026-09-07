using CommonLib.ErrorReport.Base;

namespace CommonLib.ErrorReport.ErrorCodes
{
    public static class PostActionErrorType
    {
        /// <summary>Template: "Duplicate sheet name detected during IGXL composition: {0}. The previously loaded sheet was removed."</summary>
        /// <remarks>{0} = file name</remarks>
        public static readonly ErrorCode W_DuplicateFile_01 = new(
            enumErrorCategory: EnumErrorCategory.PostAction,
            enumErrorBehavior: EnumErrorBehavior.Duplicate,
            enumErrorTarget: EnumErrorTarget.File,
            code: 1,
            template: "Duplicate sheet name detected during IGXL composition: {0}. The previously loaded sheet was removed.",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "A duplicate sheet was found and has been removed from the output. "
                    + "Verify that the intended source file is the one retained.");

        /// <summary>Template: "Duplicate instance '{0}' was detected {1} times in the following sheet(s): {2}."</summary>
        /// <remarks>{0} = instance name, {1} = duplicate times, {2} = instance sheet name(s)</remarks>
        public static readonly ErrorCode W_DuplicateInstance_01 = new(
            enumErrorCategory: EnumErrorCategory.PostAction,
            enumErrorBehavior: EnumErrorBehavior.Duplicate,
            enumErrorTarget: EnumErrorTarget.Instance,
            code: 1,
            template: "Duplicate instance '{0}' was detected {1} times in the following sheet(s): {2}.",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Please check whether the same instance has been defined multiple times in the input document,"
                    + " as duplicate definitions may lead to generation issues.");

        /// <summary>Template: "Test Number {0} duplicate, Please check {1} in {2}"</summary>
        /// <remarks>{0} = test number, {1} = first flow sheet name, {2} = first item name, {3} = second flow sheet name, {4} = second item name</remarks>
        public static readonly ErrorCode W_DuplicateTestNumber_01 = new(
            enumErrorCategory: EnumErrorCategory.PostAction,
            enumErrorBehavior: EnumErrorBehavior.Duplicate,
            enumErrorTarget: EnumErrorTarget.TestNumber,
            code: 1,
            template: "Duplicate TNum {0} detected between Flow Sheet '{1}' Item '{2}' and Flow Sheet '{3}' Item '{4}'.",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Check the test number assignments for repeated values. "
                    + "Each test must have a unique number; assign the next available number to the duplicate.");
    }
}
