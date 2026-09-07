using CommonLib.ErrorReport.Base;

namespace CommonLib.ErrorReport.ErrorCodes
{
    public static class BinOutReportErrorType
    {
        /// <summary>Template: "Count mismatch detected for job "{0}": binout report count ({1}) does not match the count in program sheet "{2}" ({3}). (The updated value of Binout status and limit value is not exactly the same in each row)"</summary>
        /// <remarks>{0} = JobPart, {0} = expected count, {1} = flow sheet name, {2} = actual count</remarks>
        public static readonly ErrorCode E_MismatchCount_01 = new(
            enumErrorCategory: EnumErrorCategory.BinOutReport,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Count,
            code: 1,
            template: "Count mismatch detected for job \"{0}\": binout report count ({1}) does not match the count in program sheet \"{2}\" ({3}). (The updated value of Binout status and limit value is not exactly the same in each row)",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please check for any missing or extra rows in the flow sheet that may cause the count mismatch.");

        /// <summary>Template: "Count mismatch detected for job "{0}": binout report count ({1}) does not match the count in program sheet "{2}" ({3}). (The updated value of Binout status is not exactly the same in each row)"</summary>
        /// <remarks>{0} = JobPart, {0} = expected count, {1} = flow sheet name, {2} = actual count</remarks>
        public static readonly ErrorCode E_MismatchCount_02 = new(
            enumErrorCategory: EnumErrorCategory.BinOutReport,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Count,
            code: 2,
            template: "Count mismatch detected for job \"{0}\": binout report count ({1}) does not match the count in program sheet \"{2}\" ({3}). (The updated value of Binout status is not exactly the same in each row)",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please check for any missing or extra rows in the flow sheet that may cause the count mismatch.");

        /// <summary>Template: "Count mismatch detected for job "{0}": binout report count ({1}) does not match the count in program sheet "{2}" ({3}). (The updated value of limit value is not exactly the same in each row)"</summary>
        /// <remarks>{0} = JobPart, {0} = expected count, {1} = flow sheet name, {2} = actual count</remarks>
        public static readonly ErrorCode E_MismatchCount_03 = new(
            enumErrorCategory: EnumErrorCategory.BinOutReport,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Count,
            code: 3,
            template: "Count mismatch detected for job \"{0}\": binout report count ({1}) does not match the count in program sheet \"{2}\" ({3}). (The updated value of limit value is not exactly the same in each row)",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please check for any missing or extra rows in the flow sheet that may cause the count mismatch.");

        /// <summary>Template: "Count mismatch for functional item "{0}" between TP and datalog. The value cannot be updated in "{1}"."</summary>
        /// <remarks>{0} = functional item name, {1} = JobPart</remarks>
        public static readonly ErrorCode E_MismatchCount_04 = new(
            enumErrorCategory: EnumErrorCategory.BinOutReport,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Count,
            code: 4,
            template: "Count mismatch for functional item \"{0}\" between TP and datalog. The value cannot be updated in \"{1}\".",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please check for any missing or extra rows in the flow sheet that may cause the count mismatch.");

        /// <summary>Template: "Functional item "{0}" with TName "{1}" was not found in the binout report for "{2}", so it cannot be updated."</summary>
        /// <remarks>{0} = item name, {1} = TName, {2} = JobPart</remarks>
        public static readonly ErrorCode E_MismatchCount_05 = new(
            enumErrorCategory: EnumErrorCategory.BinOutReport,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Count,
            code: 5,
            template: "\"{0}\" with TName \"{1}\" was not found in the binout report for \"{2}\", so it cannot be updated.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please verify that the item and TName exist in the report for the specified job and part.");

        /// <summary>Template: "TP/binout report count mismatch or inconsistent status/limit values found for "{0}" with TName "{1}" in job "{2}". It cannot be updated."</summary>
        /// <remarks>{0} = item name, {1} = TName, {2} = JobPart</remarks>
        public static readonly ErrorCode E_MismatchCount_06 = new(
            enumErrorCategory: EnumErrorCategory.BinOutReport,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Count,
            code: 6,
            template: "TP/binout report count mismatch or inconsistent status/limit values found for \"{0}\" with TName \"{1}\" in \"{2}\". It cannot be updated.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please verify that the TP and binout report counts match and that the status and limit values are consistent for the same item and TName within the job and part.");

        /// <summary>Template: "Fail flag(s) "{0}" are not referenced by any bintable row in Bin_Table_HardIP."</summary>
        /// <remarks>{0} = flag(s)</remarks>
        public static readonly ErrorCode E_MissingBinTable_01 = new(
            enumErrorCategory: EnumErrorCategory.BinOutReport,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.BinTable,
            code: 1,
            template: "Fail flag(s) \"{0}\" are not referenced by any bintable row in Bin_Table_HardIP.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please ensure the fail flag(s) are referenced in at least one row of the Bin_Table_HardIP sheet.");

        /// <summary>Template: "Bin table "{0}" is not referenced in flow sheet "{1}"."</summary>
        /// <remarks>{0} = bintable name, {1} = flow sheet name</remarks>
        public static readonly ErrorCode E_MissingBinTable_02 = new(
            enumErrorCategory: EnumErrorCategory.BinOutReport,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.BinTable,
            code: 2,
            template: "Bin table \"{0}\" is not referenced in flow sheet \"{1}\".",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please ensure the bin table is referenced in the flow sheet.");

        /// <summary>Template: "No fail flags found for instance "{0}"."</summary>
        /// <remarks>{0} = instance name</remarks>
        public static readonly ErrorCode E_MissingFlag_01 = new(
            enumErrorCategory: EnumErrorCategory.BinOutReport,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Flag,
            code: 1,
            template: "No fail flags found for instance \"{0}\".",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Binout update requires at least one fail flag. "
                    + "Please verify that the correct flags are assigned to the instance.");

        /// <summary>Template: "Instance {0} not be found in the flow sheet {1}."</summary>
        /// <remarks>{0} = instance name, {1} = flow sheet name</remarks>
        public static readonly ErrorCode E_MissingInstance_01 = new(
            enumErrorCategory: EnumErrorCategory.BinOutReport,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Instance,
            code: 1,
            template: "Instance {0} not be found in the flow sheet {1}.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please ensure the instance exists in the flow sheet before performing the binout status update.");

        /// <summary>Template: "Instance "{0}" is listed in the report, but no corresponding flow row was found in any flow sheet of the test program."</summary>
        /// <remarks>{0} = instance name</remarks>
        public static readonly ErrorCode E_MissingInstance_02 = new(
            enumErrorCategory: EnumErrorCategory.BinOutReport,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Instance,
            code: 2,
            template: "Instance \"{0}\" is listed in the report, but no corresponding flow row was found in any flow sheet of the test program.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please verify that the instance exists and is referenced in the correct flow sheet");

        /// <summary>Template: "The {0} limit "{1}" in flow sheet "{2}" is not a numeric value."</summary>
        /// <remarks>{0} = high/low, {1} = value, {2} = flow sheet name</remarks>
        public static readonly ErrorCode W_InvalidLimit_01 = new(
            enumErrorCategory: EnumErrorCategory.BinOutReport,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Limit,
            code: 1,
            template: "The {0} limit \"{1}\" in flow sheet \"{2}\" is not a numeric value.",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Since the value is not numeric, please verify that it is the intended value.");

        /// <summary>Template: "Instance "{0}" is not tested in job "{1}", so the tool will not update the value in the test program."</summary>
        /// <remarks>{0} = item name, {1} = details</remarks>
        public static readonly ErrorCode W_MismatchJob_01 = new(
            enumErrorCategory: EnumErrorCategory.BinOutReport,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Job,
            code: 1,
            template: "Instance \"{0}\" is not tested in job \"{1}\", so the tool will not update the value in the test program.",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Please ensure the instance is included in the test flow for the specified job before updating the test program.");

        /// <summary>Template: "Original flag for instance "{0}" was not found. The tool will generate a new flag "{1}"."</summary>
        /// <remarks>{0} = instance name, {1} = new flag</remarks>
        public static readonly ErrorCode W_MissingFlag_01 = new(
            enumErrorCategory: EnumErrorCategory.BinOutReport,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Flag,
            code: 1,
            template: "Original flag for instance \"{0}\" was not found. The tool will generate a new flag \"{1}\".",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Please verify the original flag assignment in the flow sheet and confirm that the generated flag is correct.");
    }
}
