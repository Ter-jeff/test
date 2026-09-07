using CommonLib.ErrorReport.Base;

namespace CommonLib.ErrorReport.ErrorCodes
{
    public static class PatInfoType
    {
        /// <summary>Template: "Please Check the pattern \"{0}\" , the vm vector and generic name is not match."</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MismatchGenericName_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.PatternInfo,
            code: 2,
            template: "Please Check the pattern \"{0}\" , the vm vector and generic name is not match.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the generic name used in PatInfo matches the name defined in the source definition. "
                    + "Check for case differences or renamed entries.");

        /// <summary>Template: "Please Check the pattern \"{0}\" , the vm vector and call subrs is not match."</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MismatchVmandSubr_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.PatternInfo,
            code: 3,
            template: "Please Check the pattern \"{0}\" , the vm vector and call subrs is not match.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Compare the VM definition against the call subroutine entries for the pattern. "
                    + "Ensure they reference the same signals and measurement targets.");

        /// <summary>Template: "Please Check the pattern \"{0}\" , the vm vector is missing"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingVmvector_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.PatternInfo,
            code: 3,
            template: "Please Check the pattern \"{0}\" , the vm vector is missing",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the PatInfo sheet and add the required VM vector definition for the pattern. "
                    + "Verify the VM vector name and ensure it is referenced correctly.");

        /// <summary>Template: "Pattern : \"{0}\" , Sen bit is Empty while pattern info has DigSrc Signal Name"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_WrongDigSrcSignalName_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.PatternInfo,
            code: 4,
            template: "Pattern : \"{0}\" , Sen bit is Empty while pattern info has DigSrc Signal Name",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the digital source signal name against the defined signal list. "
                    + "Check for typos or renamed signals in the source definition.");
    }
}
