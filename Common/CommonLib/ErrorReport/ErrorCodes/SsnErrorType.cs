using CommonLib.ErrorReport.Base;

namespace CommonLib.ErrorReport.ErrorCodes
{
    public static class SsnErrorType
    {
        /// <summary>Template: "Cannot found SsnCoreName \"{0}\" of pattern (HardipInfo) in the \"{1}\"(HarvestPinFlag_Table). Pattern : \"{2}\""</summary>
        /// <remarks>{0} = ?, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_MismatchCore_01 = new(
            enumErrorCategory: EnumErrorCategory.Ssn,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Core,
            code: 1,
            template: "Cannot found SsnCoreName \"{0}\" of pattern (HardipInfo) in the \"{1}\"(HarvestPinFlag_Table). Pattern : \"{2}\"",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Compare the SSN core name against the defined core list. "
                    + "Update the SSN definition or the core reference to ensure they match.");

        /// <summary>Template: "Cannot found SsnCoreName \"{0}\" from \"{1}\" (HarvestPinFlag_Table) in the pattern (HardipInfo). Pattern : \"{2}\""</summary>
        /// <remarks>{0} = ?, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_MismatchCore_02 = new(
            enumErrorCategory: EnumErrorCategory.Ssn,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Core,
            code: 2,
            template: "Cannot found SsnCoreName \"{0}\" from \"{1}\" (HarvestPinFlag_Table) in the pattern (HardipInfo). Pattern : \"{2}\"",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Compare the core names between the HarvestPinFlag_Table entry and the HardipInfo pattern definition. "
                    + "Add the missing core to HardipInfo or remove the extra entry from HarvestPinFlag_Table.");

        /// <summary>Template: "Duplicate Pattern : {0}"</summary>
        /// <remarks>{0} = pattern name</remarks>
        public static readonly ErrorCode E_DuplicatePattern_01 = new(
            enumErrorCategory: EnumErrorCategory.Ssn,
            enumErrorBehavior: EnumErrorBehavior.Duplicate,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 1,
            template: "Duplicate Pattern : {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the SSN pattern list for duplicate entries. "
                    + "Remove the duplicate or rename one entry so each pattern name is unique.");

        /// <summary>Template: "Duplicate Pin Group : {0}"</summary>
        /// <remarks>{0} = pin group name</remarks>
        public static readonly ErrorCode E_DuplicatePinGroup_01 = new(
            enumErrorCategory: EnumErrorCategory.Ssn,
            enumErrorBehavior: EnumErrorBehavior.Duplicate,
            enumErrorTarget: EnumErrorTarget.PinGroup,
            code: 1,
            template: "Duplicate Pin Group : {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the SSN pin group definitions for duplicate entries. "
                    + "Remove or rename the duplicate group so each name is unique.");

        /// <summary>Template: "Illegal Flag Name : {0}"</summary>
        /// <remarks>{0} = flag name</remarks>
        public static readonly ErrorCode E_InvalidFlag_01 = new(
            enumErrorCategory: EnumErrorCategory.Ssn,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Flag,
            code: 1,
            template: "Illegal Flag Name : {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Rename the SSN flag to avoid reserved keywords or illegal characters. "
                    + "Check the SSN naming rules for the list of disallowed names.");

        /// <summary>Template: "Pattern : \"{0}\" , Missing SSN info in HardipInfo"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingPatternInfo_01 = new(
            enumErrorCategory: EnumErrorCategory.Ssn,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.PatternInfo,
            code: 1,
            template: "Pattern : \"{0}\" , Missing SSN info in HardipInfo",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the SSN definition and add the required SSN information. "
                    + "Verify all mandatory SSN fields are populated for the specified instance.");

        /// <summary>Template: "PinGroup Count : {0} , PinFlag Count : {1} Mismatch"</summary>
        /// <remarks>{0} = pin group or flag name, {1} = ?</remarks>
        public static readonly ErrorCode E_MismatchFlag_01 = new(
            enumErrorCategory: EnumErrorCategory.Ssn,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Flag,
            code: 1,
            template: "PinGroup Count : {0} , PinFlag Count : {1} Mismatch",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify that the SSN pin group and the associated flag definition are aligned. "
                    + "Ensure the pin count and flag assignments are consistent between the two definitions.");

        /// <summary>Template: "Corresponding patterns &gt; 1 in the HarvestPinFlag_Table. Pattern : \"{0}\""</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode W_MismatchPattern_01 = new(
            enumErrorCategory: EnumErrorCategory.Ssn,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 1,
            template: "Corresponding patterns > 1 in the HarvestPinFlag_Table. Pattern : \"{0}\"",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the HarvestPinFlag_Table and verify that each pattern has a unique match. "
                    + "Remove or disambiguate the duplicate entries so only one row matches the pattern.");

        /// <summary>Template: "Cannot found any matching pattern in the HarvestPinFlag_Table. Pattern : \"{0}\""</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode W_MismatchPattern_02 = new(
            enumErrorCategory: EnumErrorCategory.Ssn,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 2,
            template: "Cannot found any matching pattern in the HarvestPinFlag_Table. Pattern : \"{0}\"",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Check the SSN field value against the expected specification. "
                    + "Ensure the value is within allowed bounds and uses the correct format.");

        /// <summary>Template: "Pattern : \"{0}\" , SsnCoreName : \"{1}\" SsnFlag : \"{2}\" mismatch"</summary>
        /// <remarks>{0} = ?, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode W_MismatchFlag_01 = new(
            enumErrorCategory: EnumErrorCategory.Ssn,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Flag,
            code: 1,
            template: "Pattern : \"{0}\" , SsnCoreName : \"{1}\" SsnFlag : \"{2}\" mismatch",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the HarvestPinFlag_Table and check the flag name for the flagged SSN core. "
                    + "Ensure the flag follows the expected naming convention (F_<CoreName>).");
    }
}
