using CommonLib.ErrorReport.Base;

namespace CommonLib.ErrorReport.ErrorCodes
{
    public static class HtolErrorType
    {
        /// <summary>Template: "The function: {0} can not find in {1} library!"</summary>
        /// <remarks>{0} = function name, {1} = C#/VBT</remarks>
        public static readonly ErrorCode E_MissingLibrary_01 = new(
            enumErrorCategory: EnumErrorCategory.Htol,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Library,
            code: 1,
            template: "The function: {0} can not find in {1} library!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please verify that the function exists in the specified library and that the library version and function reference are correct.");
    }
}
