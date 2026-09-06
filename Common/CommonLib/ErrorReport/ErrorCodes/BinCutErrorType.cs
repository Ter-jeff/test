using CommonLib.ErrorReport.Base;

namespace CommonLib.ErrorReport.ErrorCodes
{
    public static class BinCutErrorType
    {
        /// <summary>Template: "BinCut AllowEqual setting '{0}' is invalid."</summary>
        /// <remarks>{0} = equal condition or context</remarks>
        public static readonly ErrorCode E_AllowEqual_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.PerformanceMode,
            code: 1,
            template: "The format of Allow Equal \" {0} \" was worng (Should be {1})!!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the AllowEqual mode column in the binning table. "
                    + "AllowEqual mode can only be the previous performance mode");

        /// <summary>Template: "The Allow Equal should not be empty !!!"</summary>
        public static readonly ErrorCode E_AllowEqual_02 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.PerformanceMode,
            code: 2,
            template: "The Allow Equal should not be empty !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Fill in the Allow Equal column for the flagged mode row in the binning table. "
                    + "Allow Equal column should have same content for same performance mode");

        /// <summary>Template: "The Comment should not be empty !!!"</summary>
        public static readonly ErrorCode E_AllowEqual_03 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.PerformanceMode,
            code: 3,
            template: "The Comment should not be empty !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Fill in the Comment column for the flagged mode row in the binning table. "
                    + "Comment column should have same content for same performance mode if it is start with \"Max PV\"");

        /// <summary>Template: "The Comment should contain mode {0} !!!"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_AllowEqual_04 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.PerformanceMode,
            code: 4,
            template: "The Comment should contain mode {0} !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Update the Comment value to include all modes from the allow-equal group.");

        /// <summary>Template: "Row {0} and row {1} of the init patterns are the same!!!"</summary>
        /// <remarks>{0} = pattern name, {1} = ?</remarks>
        public static readonly ErrorCode E_BistInitPattern_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 5,
            template: "Row {0} and row {1} of the init patterns are the same!!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check that the BIST initialization pattern, they need to be different between two neighbor rows ");

        /// <summary>Template: "The Ids {0} of {1} was larger than the spec {2} in {0} @ Bin {1}"</summary>
        /// <remarks>{0} = rule or condition description, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_Business_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Ids,
            code: 6,
            template: "The Ids {0} of {1} was larger than the spec {2} in {3} @ Bin {4}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Ids value is larger the maximum Ids value. ");

        /// <summary>Template: "The domain : {0} in Other Rails can't be found."</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_Business_02 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Domain,
            code: 7,
            template: "The domain : {0} in Other Rails can't be found.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check if the domain for other rail is defined in Binning table");

        /// <summary>Template: "C value {0} is different from {1} in {2} when setting allow equal mode!!!"</summary>
        /// <remarks>{0} = C value or context, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_C_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.C,
            code: 8,
            template: "C value {0} is different from {1} in {2} when setting allow equal mode!!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the BinCut C (constant) value definition in Binning table, should be the same when setting to allow equal mode ");

        /// <summary>Template: "CPGB value {0} is different from {1} in {2} when setting allow equal mode!!!"</summary>
        /// <remarks>{0} = Cp2Gb value or context, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_Cp2Gb_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.GB,
            code: 9,
            template: "CP2GB value {0} is different from {1} in {2} when setting allow equal mode!!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the Cp2Gb (CP2 guardband) value in the Binning table, should be the same with its allow equal mode when setting allow equal mode. ");

        /// <summary>Template: "CPGB value {0} is different from {1} in {2} when setting allow equal mode!!!"</summary>
        /// <remarks>{0} = Cpgb value or context, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_Cpgb_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.GB,
            code: 10,
            template: "CPGB value {0} is different from {1} in {2} when setting allow equal mode!!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the Cpgb (CP guardband) value in the Binning table, should be the same with its allow equal mode when setting allow equal mode. ");

        /// <summary>Template: "The cpVMax {0} + {1} - {2} is larger than the limit in efuseBitDef {3} - ( 2 ^ {4} - 1 ) * {5} !!!"</summary>
        /// <remarks>{0} = CP max value or context, {1} = ?, {2} = ?, {3} = ?, {4} = ?, {5} = ?</remarks>
        public static readonly ErrorCode E_CpMax_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Cpvmax,
            code: 11,
            template: "The cpVMax {0} + {1} - {2} is larger than the limit in efuseBitDef {3} - ( 2 ^ {4} - 1 ) * {5} !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the CP maximum value against the specification. "
                    + "Ensure the value does not exceed the allowed limit for this bin.");

        /// <summary>Template: "The domain {0} is not existed !!!"</summary>
        /// <remarks>{0} = domain name or value</remarks>
        public static readonly ErrorCode E_Domain_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Domain,
            code: 12,
            template: "The domain {0} is not existed !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Check the domain name in the Binning table against the supported domain list. "
                    + "Update the configuration to use a valid domain name.");

        /// <summary>Template: "The domain {0} of {1} is correct (Should be {2}) !!!"</summary>
        /// <remarks>{0} = domain name and mode context, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_Domain_02 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Domain,
            code: 13,
            template: "The domain {0} of {1} is incorrect (Should be {2}) !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the BinCut binning table and correct the domain assignment for the flagged row. "
                    + "All rows with the same performance mode must reference the same domain.");

        /// <summary>Template: "The IDS_VDD_{0} can not be found in efuseBitDef, and only found relative name {1} !!!"</summary>
        /// <remarks>{0} = eFuse bit definition name, {1} = ?</remarks>
        public static readonly ErrorCode E_efuseBitDef_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Field,
            code: 14,
            template: "The IDS_VDD_{0} can not be found in efuseBitDef, and only found relative name {1} !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Verify the eFuse bit definition name against IDS name. "
                    + "Check for typos or missing bit definitions.");

        /// <summary>Template: "The IDS_VDD_{0} can not be found in efuseBitDef !!!"</summary>
        /// <remarks>{0} = domain name</remarks>
        public static readonly ErrorCode E_efuseBitDef_02 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Field,
            code: 15,
            template: "The IDS_VDD_{0} can not be found in efuseBitDef !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the eFuse bit definition name against IDS name. "
                    + "Check for typos or missing bit definitions.");

        /// <summary>Template: "{0} and {1} have different number of equation"</summary>
        /// <remarks>{0} = equation text or name, {1} = ?</remarks>
        public static readonly ErrorCode E_Equation_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Equation,
            code: 16,
            template: "{0} and {1} have different number of equation",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the equation count between flagged mode and its allow equal mode in Binning table. "
                    + "Ensure all referenced variables are defined and the expression is well-formed.");

        /// <summary>Template: "Equation {0} not be found in {1} when setting allow equal mode!!!"</summary>
        /// <remarks>{0} = equation and allow-equal mode context, {1} = ?</remarks>
        public static readonly ErrorCode E_Equation_02 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Equation,
            code: 17,
            template: "Equation {0} not be found in {1} when setting allow equal mode!!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the BinCut binning table and verify that the flagged equation exists in the allow-equal mode rows. ");

        /// <summary>Template: "BinCut flow core power pin mismatch between jobs: {0}"</summary>
        /// <remarks>{0} = job and pin name context</remarks>
        public static readonly ErrorCode E_Flow_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 18,
            template: "Please check core power pin in different jobs {0} {1} vs {2} !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Compare the core power pin definitions across all jobs in the BinCut flow sheet. "
                    + "Ensure the pin names and order are consistent between all job definitions.");

        /// <summary>Template: "BinCut flow core power pin count in job '{0}' does not match the reference job."</summary>
        /// <remarks>{0} = job name context</remarks>
        public static readonly ErrorCode E_Flow_02 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 19,
            template: "The core power count in {0} does not match with {1} !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the BinCut flow sheet and verify that all jobs define the same number of core power pins. "
                    + "Align the pin counts across all job definitions.");

        /// <summary>Template: "The mode {0} in flow can not be found in the column performance mode !!!"</summary>
        /// <remarks>{0} = mode name</remarks>
        public static readonly ErrorCode E_Flow_03 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.PerformanceMode,
            code: 20,
            template: "The mode {0} in flow can not be found in the column performance mode !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the BinCut flow sheet and verify that the flagged mode appears in the performance mode column. "
                    + "Add the missing mode or correct the reference.");

        /// <summary>Template: "The flowName \"{0}\" is missing in the all BinCut_Instance sheet !!!"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_FlowName_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Flow,
            code: 21,
            template: "The flowName \"{0}\" is missing in the all BinCut_Instance sheet !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the flow name against the defined flow list. "
                    + "Check for typos or missing flow definitions.");

        /// <summary>Template: "The first init can not be empty in one TTR instance !!!"</summary>
        public static readonly ErrorCode E_FlowName_02 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 22,
            template: "The first init can not be empty in one TTR instance !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the BinCut instance sheet and assign a valid init pattern to the first row of the TTR instance group. "
                    + "The first init pattern is required for TTR instance initialization.");

        /// <summary>Template: "The flowName \"{0}\" with extra spaces !!!"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_FlowName_03 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 23,
            template: "The flowName \"{0}\" with extra spaces !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the BinCut instance sheet and remove the extra spaces from the flagged flow name. "
                    + "Flow names must not contain leading, trailing, or embedded spaces.");

        /// <summary>Template: "The subflow {0} had not start with Performance Mode !!!"</summary>
        /// <remarks>{0} = field or row description</remarks>
        public static readonly ErrorCode E_FormatError_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 24,
            template: "The subflow {0} had not start with Performance Mode !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the BinCut sheet and inspect the flagged field for formatting issues. "
                    + "Ensure all values conform to the expected format and data type.");

        /// <summary>Template: "Can not find voltage condition in flow sheet (Job: {0}, TableType: {1} ,TableBinType: {2}, mode: {3})!!!"</summary>
        /// <remarks>{0} = flow or condition description, {1} = ?, {2} = ?, {3} = ?</remarks>
        public static readonly ErrorCode E_FormatError_02 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 25,
            template: "Can not find voltage condition in flow sheet (Job: {0}, TableType: {1} ,TableBinType: {2}, mode: {3})!!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the BinCut flow sheet and verify the voltage condition entry exists for the specified job, table type, and performance mode. "
                    + "Add the missing row or correct the reference.");

        /// <summary>Template: "BinCut format error in order configuration '{0}'."</summary>
        /// <remarks>{0} = field or row description</remarks>
        public static readonly ErrorCode E_FormatError_03 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.PerformanceMode,
            code: 26,
            template: "The performance mode {0} can not be found in flow sheet!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the BinCut order sheet and inspect the flagged row for formatting issues. "
                    + "Ensure the performance mode and related fields conform to the expected format.");

        /// <summary>Template: "BinCut order row format error: '{0}'."</summary>
        /// <remarks>{0} = field or row description</remarks>
        public static readonly ErrorCode E_FormatError_04 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 27,
            template: "The columns TD, BIST or FUNC can not be null or empty!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the BinCut order sheet and inspect the flagged row for formatting issues. "
                    + "Verify all required columns are populated with valid values.");

        /// <summary>Template: "The missing column {0} in sheet {1} !!!"</summary>
        /// <remarks>{0} = sheet or context description, {1} = ?</remarks>
        public static readonly ErrorCode E_FormatError_05 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 28,
            template: "The missing column {0} in sheet {1} !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Compare the header and row definitions across all binning table sheets. "
                    + "Ensure all sheets use consistent column names and value formats.");

        /// <summary>Template: "CPGB value {0} is different from {1} in {2} when setting allow equal mode!!!"</summary>
        /// <remarks>{0} = FtRoomGb value or context, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_FtRoomGb_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.GB,
            code: 29,
            template: "FT1GB value {0} is different from {1} in {2} when setting allow equal mode!!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the BinCut binning table and compare the FT1GB value for the flagged row against the allow-equal mode row. "
                    + "The FT1 GB values must be identical when allow-equal mode is set.");

        /// <summary>Template: "CPGB value {0} is different from {1} in {2} when setting allow equal mode!!!"</summary>
        /// <remarks>{0} = CPGB value and allow-equal mode context, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_FtHotGb_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.GB,
            code: 30,
            template: "FT2GB value {0} is different from {1} in {2} when setting allow equal mode!!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the BinCut binning table and compare the FT2GB value for the flagged row against the allow-equal mode row. "
                    + "The FT2 GB values must be identical when allow-equal mode is set.");

        /// <summary>Template: "BinCut HTOLGb value '{0}' is invalid."</summary>
        /// <remarks>{0} = HTOLGb value or context</remarks>
        public static readonly ErrorCode E_HTOLGb_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.GB,
            code: 31,
            template: "The GB of HTOL {0} is bigger than FT {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the HTOLGb (HTOL guardband) value in the BinCut definition. "
                    + "Ensure it is positive and within the specification limits for HTOL testing.");

        /// <summary>Template: "The ID column does not exist in sheet {0}, please add it!!!"</summary>
        /// <remarks>{0} = ID value or name</remarks>
        public static readonly ErrorCode E_ID_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 32,
            template: "The ID column does not exist in sheet {0}, please add it!!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check if ID column is existed in Bincut binning table.");

        /// <summary>Template: "The ID value incorrect"</summary>
        public static readonly ErrorCode E_ID_02 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 33,
            template: "The ID value invalid",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the BinCut binning table and verify the ID value for the first row of the flagged performance mode group. "
                    + "The ID must be a valid numeric value.");

        /// <summary>Template: "The larger equation {0} ({1}) needs to set a larger ID value than {2} ({3})"</summary>
        /// <remarks>{0} = equation and ID ordering context, {1} = ?, {2} = ?, {3} = ?</remarks>
        public static readonly ErrorCode E_ID_04 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 34,
            template: "The larger equation {0} ({1}) needs to set a larger ID value than {2} ({3})",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the BinCut binning table and correct the ID values so that larger equations have larger IDs. "
                    + "The ID values must be monotonically increasing with equation number.");

        /// <summary>Template: "The smaller equation {0} ({1}) needs to set a smaller ID value than {2} ({3})"</summary>
        /// <remarks>{0} = equation and ID ordering context, {1} = ?, {2} = ?, {3} = ?</remarks>
        public static readonly ErrorCode E_ID_05 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 35,
            template: "The smaller equation {0} ({1}) needs to set a smaller ID value than {2} ({3})",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the BinCut binning table and correct the ID values so that smaller equations have smaller IDs. "
                    + "The ID values must be monotonically decreasing with equation number.");

        /// <summary>Template: "The {0}(ID:{1}) of the performance mode {2} needs to be set to be greater than the {3}(ID:{4}) of the {5}"</summary>
        /// <remarks>{0} = equation, ID, and performance mode context, {1} = ?, {2} = ?, {3} = ?, {4} = ?, {5} = ?</remarks>
        public static readonly ErrorCode E_ID_06 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 36,
            template: "The {0}(ID:{1}) of the performance mode {2} needs to be set to be greater than the {3}(ID:{4}) of the {5}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the BinCut binning table and ensure that the last equation's ID in each performance mode is greater than in the preceding mode. "
                    + "Reorder or reassign ID values to satisfy this constraint.");

        /// <summary>Template: "The ID value is duplicated in the {0} domain, please change!!!"</summary>
        /// <remarks>{0} = domain name</remarks>
        public static readonly ErrorCode E_ID_07 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Duplicate,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 37,
            template: "The ID value is duplicated in the {0} domain, please change!!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the BinCut binning table and assign unique ID values within the flagged domain. "
                    + "No two rows within the same domain may share the same ID value.");

        /// <summary>Template: "The IdsMax {0} is larger than the limit in efuseBitDef {1} - ( 2 ^ {2} - 1 )  !!!"</summary>
        /// <remarks>{0} = IDS max value or context, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_IdsMax_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Ids,
            code: 38,
            template: "The IdsMax {0} is larger than the limit in efuseBitDef {1} - ( 2 ^ {2} - 1 )  !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the IDS maximum value against the device specification. "
                    + "Ensure the value is within the allowable range for this bin category.");

        /// <summary>Template: "Instance name too long(over 150), please modify instance name by PatSetName(Orange)!!!"</summary>
        public static readonly ErrorCode E_InstanceName_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 39,
            template: "Instance name too long(over 150), please modify instance name by PatSetName(Orange)!!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Instance name generated by autogen wil be too long and cause error in run time, use PatSetName(Orange) in test plan to shorten it. ");

        /// <summary>Template: "Interpolation ModeH values are not consistent across equations for mode '{0}'."</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_Interpolation_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Interpolation,
            code: 40,
            template: "The value of performance {0} is different with other equations!!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Ensure all equations for the same mode have the same Interpolation ModeH, ModeL, Offset and SkipTest value.");

        /// <summary>Template: "The value of Int_Mode_L/Int_Mode_H {0} is not a legal performance mode for {1}!!!</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_Interpolation_03 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Interpolation,
            code: 41,
            template: "The value of Int_Mode_L/Int_Mode_H {0} is not a legal performance mode for {1}!!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Ensure Int_Mode_L/Int_Mode_H values are performance modes.");

        /// <summary>Template: "Jobs in the flow sheet (Performance Mode / FlowName : {0} / {1}) is {2} , but jobs in BinCut_instance is {3} !!!"</summary>
        /// <remarks>{0} = job name or context, {1} = ?, {2} = ?, {3} = ?</remarks>
        public static readonly ErrorCode E_JobMisMatch_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Job,
            code: 42,
            template: "Jobs in the flow sheet (Performance Mode / FlowName : {0} / {1}) is {2} , but jobs in BinCut_instance is {3} !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Compare the job assignment in the BinCut sheet against the expected job definition. "
                    + "Ensure the job name is consistent between all references.");

        /// <summary>Template: "M value {0} is different from {1} in {2} when setting allow equal mode!!!"</summary>
        /// <remarks>{0} = M value or context, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_M_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.M,
            code: 43,
            template: "M value {0} is different from {1} in {2} when setting allow equal mode!!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the BinCut M (multiplier) value definition. "
                    + "Ensure the value is within the allowed range and correctly specified.");

        /// <summary>Template: "The instance row {0} not be used in the flow sheet ({1})!!!"</summary>
        /// <remarks>{0} = item or field description, {1} = ?</remarks>
        public static readonly ErrorCode E_Missing_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Redundant,
            enumErrorTarget: EnumErrorTarget.Flow,
            code: 44,
            template: "The instance row {0} not be used in the flow sheet ({1})!!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Identify the missing item in the BinCut configuration and add it. "
                    + "Check the BinCut specification to confirm which items are required.");
        /// <summary>Template: "The instance row {0} not be used in the flow sheet ({1})!!!"</summary>
        /// <remarks>{0} = item or field description, {1} = ?</remarks>
        public static readonly ErrorCode E_Missing_02 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Flow,
            code: 45,
            template: "The IPL/EPL instance {0} does not have the corresponding sub flow {1} in the flow_vddbinning sheet!!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Identify the missing sub flows in the flow_vddbinning sheet. "
                    + "Check the Binning sheet to confirm the performance items are not required to generate IPL instances.");

        /// <summary>Template: "Missing Parameter in {0}({1}) : {2}"</summary>
        /// <remarks>{0} = parameter name, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_MissingParameter_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Parameter,
            code: 46,
            template: "Missing Parameter in {0}({1}) : {2}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Missing argument in C# library function, check C# library version.");

        /// <summary>Template: "The VBT function: {0} can not find in {1} library!"</summary>
        /// <remarks>{0} = module name, {1} = ?</remarks>
        public static readonly ErrorCode E_MissVbtModule_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Module,
            code: 47,
            template: "The VBT function: {0} can not find in {1} library!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Missing function in C# library, check C# library version. ");

        /// <summary>Template: "The pattern : {0} format illegal"</summary>
        /// <remarks>{0} = pattern name</remarks>
        public static readonly ErrorCode E_Pattern_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 48,
            template: "The pattern : {0} format illegal",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the pattern name follows the expected naming convention (minimum 11 underscore-separated segments). "
                    + "Check that the pattern exists and is correctly referenced in the BinCut sheet.");

        /// <summary>Template: "The eqn start {0} can not exceed the max eqn {1} !!!"</summary>
        /// <remarks>{0} = equation start value and max equation context, {1} = ?</remarks>
        public static readonly ErrorCode E_Pattern_02 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Equation,
            code: 49,
            template: "The eqn start {0} can not exceed the max eqn {1} !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the BinCut IDS distribution configuration and reduce the equation start value. "
                    + "The start value must not exceed the maximum equation number defined for this pattern.");

        /// <summary>Template: "The type of patten {0} is unknown !!!"</summary>
        /// <remarks>{0} = pattern name</remarks>
        public static readonly ErrorCode E_Pattern_03 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 50,
            template: "The type of patten {0} is unknown !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the pattern type classification in the instance sheet reader. "
                    + "Ensure the pattern name follows the expected naming convention for RTOS pattern types.");

        /// <summary>Template: "The type of patten {0} is unknown !!!"</summary>
        /// <remarks>{0} = pattern name</remarks>
        public static readonly ErrorCode E_Pattern_04 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 51,
            template: "The type of patten {0} is unknown !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the pattern type classification in the secondary parsing path. "
                    + "Ensure the pattern name follows the expected naming convention for all supported pattern types.");

        /// <summary>Template: "An empty string exists in the pattern: {0} !!!"</summary>
        /// <remarks>{0} = pattern cell value containing the empty string</remarks>
        public static readonly ErrorCode E_Pattern_05 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 52,
            template: "An empty string exists in the pattern: {0} !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the BinCut instance sheet and remove the empty string from the pattern cell. "
                    + "Pattern cells must not contain whitespace-only values.");

        /// <summary>Template: "The fisrt pattern can not be empty !!!"</summary>
        public static readonly ErrorCode E_Pattern_06 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 53,
            template: "The fisrt pattern can not be empty !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the BinCut instance sheet and provide a valid pattern name in the first pattern column. "
                    + "The first pattern in every row is required and must not be empty.");

        /// <summary>Template: "This test can not find any time set by pattern dashboard !!!"</summary>
        public static readonly ErrorCode E_PatternTimeSet_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 54,
            template: "This test can not find any time set by pattern dashboard !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the time set assigned to the BinCut pattern against the defined time set list. "
                    + "Ensure the time set is compatible with the pattern's requirements.");

        /// <summary>Template: "There are multi time set {0} by patterns !!!"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_PatternTimeSet_02 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 55,
            template: "There are multi time set {0} by patterns !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Multiple patterns reference different time sets and no override is provided. "
                    + "Set the TimeSet column in the BinCut instance sheet to explicitly select the intended time set.");

        /// <summary>Template: "There are multi time set {0} by patterns !!!"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_PatternTimeSet_03 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 56,
            template: "There are multi time set {0} by patterns !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Multiple patterns reference different time sets in the secondary check path and no override is provided. "
                    + "Set the TimeSet column in the BinCut instance sheet to explicitly select the intended time set.");

        /// <summary>Template: "BinCut performance mode '{0}' is invalid."</summary>
        /// <remarks>{0} = performance mode name</remarks>
        public static readonly ErrorCode E_PerformanceMode_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Duplicate,
            enumErrorTarget: EnumErrorTarget.PerformanceMode,
            code: 57,
            template: "The performance {0} is duplicated in other post flow sheets !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the performance mode name against the supported mode list. "
                    + "Update the configuration to use a valid performance mode.");

        /// <summary>Template: "The efuse {0} resolution {1} is different with StepSize {2} from the Notes sheet !!!"</summary>
        /// <remarks>{0} = resolution value or context, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_Resolution_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Resolution,
            code: 58,
            template: "The efuse {0} resolution {1} is different with StepSize {2} from the Notes sheet !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the BinCut resolution value against the specification. "
                    + "Ensure the resolution step size is valid for the bin type.");

        /// <summary>Template: "BinCut SRAM threshold '{0}' is invalid."</summary>
        /// <remarks>{0} = SRAM threshold value or context</remarks>
        public static readonly ErrorCode E_SRAMthresh_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.SRAM,
            code: 59,
            template: "The SRAMthresh_Product value is not allowed <= 0 !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the SRAM threshold value in the BinCut definition. "
                    + "Ensure the threshold is within the valid range for the specified SRAM type.");

        /// <summary>Template: "The SRAMthresh_binSearch (CP1) value is not allowed to be &lt;= 0."</summary>
        public static readonly ErrorCode E_SRAMthresh_02 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.SRAM,
            code: 60,
            template: "The SRAMthresh_binSearch (CP1) value is not allowed to be <= 0.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Set the SRAMthresh_binSearch column value to a positive number greater than zero.");

        /// <summary>Template: "The SRAMthresh_binSearch (CP2) value is not allowed to be &lt;= 0."</summary>
        public static readonly ErrorCode E_SRAMthresh_03 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.SRAM,
            code: 61,
            template: "The SRAMthresh_binSearch (CP2) value is not allowed to be <= 0.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Set the SRAMthresh_binSearch CP2 column value to a positive number greater than zero.");

        /// <summary>Template: "The job name of Testing Stage \"{0}\" is unknown !!!"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_TestingStage_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Job,
            code: 62,
            template: "The job name of Testing Stage \"{0}\" is unknown !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the testing stage name against the defined stage list. "
                    + "Check for typos or missing stage definitions in the configuration.");

        /// <summary>Template: "TimeSet {0} does not exist is K folder"</summary>
        /// <remarks>{0} = time set name</remarks>
        public static readonly ErrorCode E_TimeSet_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Timing,
            code: 63,
            template: "TimeSet {0} does not exist is K folder",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the time set name against the defined time set list. "
                    + "Add the missing time set or correct the reference.");

        /// <summary>Template: "Can't get any SCAN Tset in all payload"</summary>
        public static readonly ErrorCode E_TimeSet_02 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Timing,
            code: 64,
            template: "Can't get any SCAN Tset in all payload",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the pattern dashboard for all payload patterns in this BinCut instance row. "
                    + "Ensure at least one payload pattern has a valid SCAN time set (ScanTset) defined.");

        /// <summary>Template: "Can not use multi enable {0} in one TTR instance !!!"</summary>
        /// <remarks>{0} = TTR value or context</remarks>
        public static readonly ErrorCode E_TTR_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 65,
            template: "Can not use multi enable {0} in one TTR instance !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Check the Enable value for TTR group in the BinCut_Instance. "
                    + "Enable value should be the same in one TTR group.");

        /// <summary>Template: "DigSrc {0} in UserFunction is not defined in UF_DigSrc sheet"</summary>
        /// <remarks>{0} = function name</remarks>
        public static readonly ErrorCode E_UserFunction_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Userfunction,
            code: 66,
            template: "DigSrc {0} in UserFunction is not defined in UF_DigSrc sheet",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check if the Userfunction key is defined in Userfunction sheet.");

        /// <summary>Template: "Need to assign digital source group for digital source pattern {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_UserFunction_02 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Userfunction,
            code: 67,
            template: "Need to assign digital source group for digital source pattern {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Assign a DigSrc group in the UserFunction column for the row containing the digital source pattern.");

        /// <summary>Template: "UserFunction format error, should be \"DigSrc:[DigSrcGroup]\""</summary>
        public static readonly ErrorCode E_UserFunction_03 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Userfunction,
            code: 68,
            template: "UserFunction format error, should be \"DigSrc:[DigSrcGroup]\"",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Correct the UserFunction value to follow the format 'DigSrc:[DigSrcGroup]'.");

        /// <summary>Template: "Voltage type is different in the Sub Flow column and Voltage Category column !!!"</summary>
        public static readonly ErrorCode E_VoltageType_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Voltage,
            code: 69,
            template: "Voltage type is different in the Sub Flow column and Voltage Category column !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the voltage type against the list of supported BinCut voltage types. "
                    + "Update the configuration to use a valid voltage type.");

        /// <summary>Template: "Unable to get domain from flow name and pattern !!!"</summary>
        public static readonly ErrorCode W_Domain_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Domain,
            code: 70,
            template: "Unable to get domain from flow name and pattern !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the BinCut instance sheet and verify that the flow name or at least one pattern can be resolved to a known domain. "
                    + "Ensure the flow name follows the naming convention that encodes the domain.");

        /// <summary>Template: "Unable to get domain from flow name !!!"</summary>
        public static readonly ErrorCode W_Domain_02 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Domain,
            code: 71,
            template: "Unable to get domain from flow name !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the BinCut instance sheet and verify the flow name contains a recognizable domain syntax. "
                    + "The pattern list domain will be used as a fallback.");

        /// <summary>Template: "Unable to get domain from pattern !!!"</summary>
        public static readonly ErrorCode W_Domain_03 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Domain,
            code: 72,
            template: "Unable to get domain from pattern !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the BinCut instance sheet and verify the pattern names contain a recognizable domain syntax. "
                    + "The flow name domain will be used as a fallback.");

        /// <summary>Template: "This domain of flowName {0} is different from the pattern {1}!!!"</summary>
        /// <remarks>{0} = flow name and pattern domain context, {1} = ?</remarks>
        public static readonly ErrorCode W_Domain_04 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Domain,
            code: 73,
            template: "This domain of flowName {0} is different from the pattern {1}!!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the BinCut instance sheet and verify the domain consistency between the flow name and the assigned patterns. "
                    + "All patterns in a row should belong to the same domain as the flow name.");

        /// <summary>Template: "This domain is different from other patterns!!!"</summary>
        public static readonly ErrorCode W_Domain_05 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Domain,
            code: 74,
            template: "This domain is different from other patterns!!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the BinCut instance sheet and verify that all patterns in the row belong to the same domain. "
                    + "Mixed-domain patterns within a single row are not supported.");

        /// <summary>Template: "The VDD_{0}_{1} can not be found in efuseBitDef, and only found relative name {2} !!!"</summary>
        /// <remarks>{0} = domain and mode context, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode W_efuseBitDef_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Field,
            code: 75,
            template: "The VDD_{0}_{1} can not be found in efuseBitDef, and only found relative name {2} !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the eFuse bit definition table and add an exact entry for the flagged VDD domain+mode combination. "
                    + "The relative match may not produce correct voltage mapping.");

        /// <summary>Template: "The ids name of {0} in flow can not be found in efuseBitDef, and only found relative name {1} !!!"</summary>
        /// <remarks>{0} = pin name and relative name context, {1} = ?</remarks>
        public static readonly ErrorCode W_efuseBitDef_02 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Field,
            code: 76,
            template: "The ids name of {0} in flow can not be found in efuseBitDef, and only found relative name {1} !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the eFuse bit definition table and add an exact entry for the flagged IDS pin. "
                    + "The relative match may not produce correct IDS value mapping in the flow.");

        /// <summary>Template: "The ids name of {0} in flow can not be found in efuseBitDef !!!"</summary>
        /// <remarks>{0} = pin name</remarks>
        public static readonly ErrorCode W_efuseBitDef_03 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Field,
            code: 77,
            template: "The ids name of {0} in flow can not be found in efuseBitDef !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the eFuse bit definition table and add an entry for the flagged IDS pin. "
                    + "Verify the IDS pin name follows the naming convention used in the efuseBitDef configuration.");

        /// <summary>Template: "The mode {0} has multiple types in FUNC block !!!"</summary>
        /// <remarks>{0} = flow name or reference</remarks>
        public static readonly ErrorCode W_Flow_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Flow,
            code: 78,
            template: "The mode {0} has multiple types in FUNC block !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Verify the flow reference in the BinCut sheet against the defined flow list. "
                    + "Check for typos or missing flow definitions.");

        /// <summary>Template: "BinCut flow performance mode '{0}' syntax does not match any entry in the binning sheet."</summary>
        /// <remarks>{0} = performance mode name</remarks>
        public static readonly ErrorCode W_Flow_02 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Flow,
            code: 79,
            template: "Please check the {0} syntax of performance, that should be existed in binning sheeet !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the BinCut flow sheet and verify the performance mode prefix against the binning sheet. "
                    + "Ensure the performance mode value begins with a recognized mode name.");

        /// <summary>Template: "BinCut flow AllOther voltage value in row {0} is empty."</summary>
        /// <remarks>{0} = row number</remarks>
        public static readonly ErrorCode W_Flow_03 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Voltage,
            code: 80,
            template: "AllOther in row {0} is empty !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the BinCut flow sheet and fill in the AllOther voltage value for the flagged row. "
                    + "Every row must have a non-empty AllOther voltage assignment.");

        /// <summary>Template: "BinCut flow voltage value for pin '{0}' is empty."</summary>
        /// <remarks>{0} = pin name and row number context</remarks>
        public static readonly ErrorCode W_Flow_04 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Voltage,
            code: 81,
            template: $"{0} in row {1} is empty !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the BinCut flow sheet and fill in the missing voltage value for the flagged pin column. "
                    + "Every power pin must have a valid voltage assignment in each row.");

        /// <summary>Template: "The mode {0} in flow can not be found in binning, binning_binX ro binning_binY!!!"</summary>
        /// <remarks>{0} = mode name</remarks>
        public static readonly ErrorCode W_Flow_05 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.PerformanceMode,
            code: 82,
            template: "The mode {0} in flow can not be found in binning, binning_binX ro binning_binY!!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the BinCut binning sheets and verify the performance mode definition. "
                    + "Add the missing mode to the appropriate binning sheet.");

        /// <summary>Template: "BinCut new flow tables: performance mode '{0}' syntax does not match any entry in the binning sheet."</summary>
        /// <remarks>{0} = performance mode name</remarks>
        public static readonly ErrorCode W_Flow_06 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.PerformanceMode,
            code: 83,
            template: "Please check the {0} syntax of performance, that should be existed in binning sheeet !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the new BinCut flow sheet and verify the performance mode prefix against the binning sheet. "
                    + "Ensure the performance mode value begins with a recognized mode name.");

        /// <summary>Template: "Can not use multi flow name {0} in one TTR instance !!!"</summary>
        /// <remarks>{0} = list of flow names</remarks>
        public static readonly ErrorCode W_FlowName_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 84,
            template: "Can not use multi flow name {0} in one TTR instance !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the BinCut instance sheet and assign a single flow name per TTR instance row. "
                    + "Merge or split TTR rows as needed to satisfy the one-flow-name-per-instance constraint.");

        /// <summary>Template: "The flowName \"{0}\" not found in the flow sheet !!!"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode W_FlowName_02 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Flow,
            code: 85,
            template: "The flowName \"{0}\" not found in the flow sheet !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the BinCut flow sheet and verify the flow name is defined. "
                    + "Add the missing flow definition or correct the flow name reference in the instance sheet.");

        /// <summary>Template: "BinCut HTOL guardband (RO) is larger than the FT hot-temperature guardband: {0}"</summary>
        /// <remarks>{0} = HTOL and FT guardband values</remarks>
        public static readonly ErrorCode W_HTOLGb_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.GB,
            code: 86,
            template: "The GB of HTOL {0} is bigger than FT {1}",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the BinCut binning table and review the HTOL and FT hot guardband values. "
                    + "The HTOL RO guardband should not exceed the FT hot guardband.");

        /// <summary>Template: "BinCut HTOL guardband (room) is larger than the FT room-temperature guardband: {0}"</summary>
        /// <remarks>{0} = HTOL and FT guardband values</remarks>
        public static readonly ErrorCode W_HTOLGb_02 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.GB,
            code: 87,
            template: "The GB of HTOL {0} is bigger than FT {1}",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the BinCut binning table and review the HTOL room and FT room guardband values. "
                    + "The HTOL room guardband should not exceed the FT room guardband.");

        /// <summary>Template: "BinCut HTOL guardband (hot) is larger than the FT room-temperature guardband: {0}"</summary>
        /// <remarks>{0} = HTOL and FT guardband values</remarks>
        public static readonly ErrorCode W_HTOLGb_03 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.GB,
            code: 88,
            template: "The GB of HTOL {0} is bigger than FT {1}",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the BinCut binning table and review the HTOL hot and FT room guardband values. "
                    + "The HTOL hot guardband should not exceed the FT room guardband.");

        /// <summary>Template: "{0} can't be found the pattern in CSV !!!"</summary>
        /// <remarks>{0} = pattern name</remarks>
        public static readonly ErrorCode W_Pattern_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 89,
            template: "{0} can't be found the pattern in CSV !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Verify the pattern name against the pattern CSV dashboard. "
                    + "Ensure the pattern is listed and the name spelling is correct.");

        /// <summary>Template: "This pattern {0} is \"Dont_useInCsv\" !!!"</summary>
        /// <remarks>{0} = pattern name</remarks>
        public static readonly ErrorCode W_Pattern_02 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Info,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 90,
            template: "This pattern {0} is \"Dont_useInCsv\" !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "The pattern is intentionally excluded via the 'dont_use' flag in the CSV. "
                    + "If the pattern is needed, update its use flag in the pattern CSV dashboard.");

        /// <summary>Template: "There are no FileVersion for {0} !!!"</summary>
        /// <remarks>{0} = pattern name</remarks>
        public static readonly ErrorCode W_Pattern_03 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 91,
            template: "There are no FileVersion for {0} !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Check the pattern CSV dashboard for the flagged pattern. "
                    + "Ensure a valid FileVersion value is assigned to the pattern entry.");

        /// <summary>Template: "{0} can't be found the pattern in CSV !!!"</summary>
        /// <remarks>{0} = pattern name</remarks>
        public static readonly ErrorCode W_Pattern_04 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 92,
            template: "{0} can't be found the pattern in CSV !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Verify the pattern name against the pattern CSV dashboard in the secondary check path. "
                    + "Ensure the pattern is listed and the name spelling is correct.");

        /// <summary>Template: "This pattern {0} is \"Dont_useInCsv\" !!!"</summary>
        /// <remarks>{0} = pattern name</remarks>
        public static readonly ErrorCode W_Pattern_05 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Info,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 93,
            template: "This pattern {0} is \"Dont_useInCsv\" !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "The pattern is intentionally excluded via the 'dont_use' flag in the CSV. "
                    + "If the pattern is needed, update its use flag in the pattern CSV dashboard.");

        /// <summary>Template: "The pattern can not be empty !!!"</summary>
        public static readonly ErrorCode W_Pattern_06 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 94,
            template: "The pattern can not be empty !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the BinCut instance sheet and fill in the empty pattern cell, or remove the extra column. "
                    + "Non-first pattern cells should either contain a valid pattern name or be omitted.");

        /// <summary>Template: "There are multi time set {0} by patterns overwrite to {1}!!!"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode W_PatternTimeSet_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Duplicate,
            enumErrorTarget: EnumErrorTarget.Timing,
            code: 95,
            template: "There are multi time set {0} by patterns overwrite to {1}!!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Multiple patterns reference different time sets. The specified time set value was used as an override. "
                    + "Verify that the override time set is correct for all patterns in this row.");

        /// <summary>Template: "This test can not find any time set by pattern dashboard !!!"</summary>
        public static readonly ErrorCode W_PatternTimeSet_02 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 96,
            template: "This test can not find any time set by pattern dashboard !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Check the pattern dashboard for all payload patterns in this row. "
                    + "Ensure at least one pattern has a valid time set (TimeSetVersion) defined in the CSV.");

        /// <summary>Template: "There are multi time set {0} by patterns overwrite to {1}!!!"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode W_PatternTimeSet_03 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Duplicate,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 97,
            template: "There are multi time set {0} by patterns overwrite to {1}!!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Multiple patterns reference different time sets in the secondary check path. The specified time set value was used as an override. "
                    + "Verify that the override time set is correct for all patterns in this row.");

        /// <summary>Template: "Can not use multi enable {0} in one TTR instance !!!"</summary>
        /// <remarks>{0} = list of enable values</remarks>
        public static readonly ErrorCode W_TTR_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 98,
            template: "Can not use multi enable {0} in one TTR instance !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the BinCut instance sheet and reduce the TTR enable values to a single entry per instance. "
                    + "Merge or split rows as needed to satisfy the one-enable-per-instance constraint.");

        /// <summary>Template: "Can not use multi JobTestStage {0} in one TTR instance !!!"</summary>
        /// <remarks>{0} = list of JobTestStage values</remarks>
        public static readonly ErrorCode W_TTR_02 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 99,
            template: "Can not use multi JobTestStage {0} in one TTR instance !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the BinCut instance sheet and reduce the TTR JobTestStage values to a single entry per instance. "
                    + "Each TTR instance must specify exactly one JobTestStage.");

        /// <summary>Template: "Can not use multi SiteVar {0} in one TTR instance !!!"</summary>
        /// <remarks>{0} = list of SiteVar values</remarks>
        public static readonly ErrorCode W_TTR_03 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 100,
            template: "Can not use multi SiteVar {0} in one TTR instance !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the BinCut instance sheet and reduce the TTR SiteVar values to a single entry per instance. "
                    + "Each TTR instance must reference exactly one SiteVar.");

        /// <summary>Template: "Can not use multi ShiftSpeed {0} in one TTR instance !!!"</summary>
        /// <remarks>{0} = list of ShiftSpeed values</remarks>
        public static readonly ErrorCode W_TTR_04 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 101,
            template: "Can not use multi ShiftSpeed {0} in one TTR instance !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the BinCut instance sheet and reduce the TTR ShiftSpeed values to a single entry per instance. "
                    + "Each TTR instance must have exactly one ShiftSpeed assignment.");

        /// <summary>Template: "Can not use multi Voltage Category {0} in one TTR instance !!!"</summary>
        /// <remarks>{0} = list of Voltage Category values</remarks>
        public static readonly ErrorCode W_TTR_05 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 102,
            template: "Can not use multi Voltage Category {0} in one TTR instance !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the BinCut instance sheet and reduce the TTR Voltage Category values to a single entry per instance. "
                    + "Each TTR instance must reference exactly one voltage category.");

        /// <summary>Template: "Can not use multi timeSet {0} in one TTR instance !!!"</summary>
        /// <remarks>{0} = list of timeSet values</remarks>
        public static readonly ErrorCode W_TTR_06 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 103,
            template: "Can not use multi timeSet {0} in one TTR instance !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the BinCut instance sheet and reduce the TTR timeSet values to a single entry per instance. "
                    + "Each TTR instance must specify exactly one time set.");

        /// <summary>Template: "Can not use multi FailFlag {0} in one TTR instance !!!"</summary>
        /// <remarks>{0} = list of FailFlag values</remarks>
        public static readonly ErrorCode W_TTR_07 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 104,
            template: "Can not use multi FailFlag {0} in one TTR instance !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the BinCut instance sheet and reduce the TTR FailFlag values to a single entry per instance. "
                    + "Each TTR instance must reference exactly one FailFlag.");

        /// <summary>Template: "Can not use multi PatSetName(Orange) {0} in one TTR instance !!!"</summary>
        /// <remarks>{0} = list of PatSetName values</remarks>
        public static readonly ErrorCode W_TTR_08 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 105,
            template: "Can not use multi PatSetName(Orange) {0} in one TTR instance !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the BinCut instance sheet and reduce the TTR PatSetName (Orange) values to a single entry per instance. "
                    + "Each TTR instance must reference exactly one pattern set name.");

        /// <summary>Template: "Format error in BinCut flow table voltage entry."</summary>
        public static readonly ErrorCode E_FormatError_06 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 106,
            template: "Please check syntax for\"{0}\" ,the count of \"(\" and \")\" are mismatch !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the voltage pin entry in the BinCut flow table for format issues.");

        /// <summary>Template: "Format error in BinCut flow table voltage entry (second check)."</summary>
        public static readonly ErrorCode E_FormatError_07 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 107,
            template: "The syntax \"{0}\"  of {1} was unknown !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the voltage pin entry in the BinCut flow table for format issues.");

        /// <summary>Template: "Format error in BinCut flow table performance mode entry."</summary>
        public static readonly ErrorCode E_FormatError_08 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 108,
            template: "{0} has not evaluate bin before Bin Result !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the performance mode column entry in the BinCut flow table for format issues.");

        /// <summary>Template: "Format error in BinCut flow table AllOther entry."</summary>
        public static readonly ErrorCode E_FormatError_09 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 109,
            template: $"The {0} is incorrect DC category in the Bincut !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the AllOther column entry in the BinCut flow table for format issues.");

        /// <summary>Template: "Format error in BinCut flow table ATPG entry."</summary>
        public static readonly ErrorCode E_FormatError_10 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 110,
            template: "There are extra spaces !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the ATPG column entry in the BinCut flow table for format issues.");

        /// <summary>Template: "Format error in BinCut flow table MBIST entry."</summary>
        public static readonly ErrorCode E_FormatError_11 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 111,
            template: "There are extra spaces !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the MBIST column entry in the BinCut flow table for format issues.");

        /// <summary>Template: "Format error in BinCut flow table SPI/RTOS entry."</summary>
        public static readonly ErrorCode E_FormatError_12 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 112,
            template: "There are extra spaces !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the SPI/RTOS column entry in the BinCut flow table for format issues.");

        /// <summary>Template: "Duplicate sheet from {0} in PowerBinning Workbook!"</summary>
        /// <remarks>{0} = sheet name</remarks>
        public static readonly ErrorCode W_DuplicateProgramSheet_01 = new(
            enumErrorCategory: EnumErrorCategory.BinCut,
            enumErrorBehavior: EnumErrorBehavior.Duplicate,
            enumErrorTarget: EnumErrorTarget.Source,
            code: 113,
            template: "Duplicate sheet from {0} in PowerBinning Workbook!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Check for duplicate sheet names in the program workbook. "
                    + "Remove or rename the duplicate sheet so each sheet has a unique name.");
    }
}
