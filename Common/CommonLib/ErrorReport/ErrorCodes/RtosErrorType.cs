using CommonLib.ErrorReport.Base;

namespace CommonLib.ErrorReport.ErrorCodes
{
    public static class RtosErrorType
    {
        /// <summary>Template: "The function: {0} can not find in {1} library!"</summary>
        /// <remarks>{0} = function name, {1} = C#/VBT</remarks>
        public static readonly ErrorCode E_MissingLibrary_01 = new(
            enumErrorCategory: EnumErrorCategory.Rtos,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Library,
            code: 1,
            template: "The function: {0} can not find in {1} library!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please verify that the function exists in the specified library and that the library version and function reference are correct.");

        /// <summary>Template: "Missing Parameter in {0}({1}) : {2}"</summary>
        /// <remarks>{0} = function name, {1} = VBT/.NET, {2} = parameter name</remarks>
        public static readonly ErrorCode E_MissingParameter_01 = new(
            enumErrorCategory: EnumErrorCategory.Rtos,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Parameter,
            code: 1,
            template: "Missing parameter \"{2}\" in function \"{0}\" of library \"({1})\".",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please ensure the parameter exists in the library and that the argument name matches the documented definition.");

        /// <summary>Template: "Can not find Rtos category: {0} from TestSettting sheet"</summary>
        /// <remarks>{0} = category name</remarks>
        public static readonly ErrorCode E_MissingRtosCategory_01 = new(
            enumErrorCategory: EnumErrorCategory.Rtos,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.DcCategory,
            code: 1,
            template: "Can not find Rtos category: {0} from TestSettting sheet",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Add the missing Rtos category definition. "
                    + "Verify the category name against the expected Rtos configuration.");
    }
}
