using CommonLib.ErrorReport.Base;

namespace CommonLib.ErrorReport.ErrorCodes
{
    public static class PreActionErrorType
    {
        /// <summary>Template: "Duplicate DC category: {0}"</summary>
        /// <remarks>{0} = category name or pair</remarks>
        public static readonly ErrorCode E_DuplicateDcCategory_01 = new(
            enumErrorCategory: EnumErrorCategory.PreAction,
            enumErrorBehavior: EnumErrorBehavior.Duplicate,
            enumErrorTarget: EnumErrorTarget.DcCategory,
            code: 1,
            template: "Duplicate DC category: {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Duplicate DC category definition found in the voltage table. "
                    + "Please ensure each DC category name is unique.");

        /// <summary>Template: "Duplicate library: \"{0}\" in file {1}!={2}"</summary>
        /// <remarks>{0} = ?, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_DuplicateLibrary_01 = new(
            enumErrorCategory: EnumErrorCategory.PreAction,
            enumErrorBehavior: EnumErrorBehavior.Duplicate,
            enumErrorTarget: EnumErrorTarget.Library,
            code: 1,
            template: "Duplicate library: \"{0}\" in file {1}!={2}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the library list for duplicate entries. "
                    + "Remove the duplicate module or rename one entry so each module name is unique.");

        /// <summary>Template: "Worksheet does not contain any data."</summary>
        /// <remarks>{0} = field or context description</remarks>
        public static readonly ErrorCode E_InvalidDocument_01 = new(
            enumErrorCategory: EnumErrorCategory.PreAction,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Document,
            code: 1,
            template: "Worksheet does not contain any data.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify that the document is not empty and contains the expected content.");

        /// <summary>Template: "PinGroup {0} from IO_PinMap/IO_PinGroup has inconsistent pin definitions between {1}. Please check these pins in IO_PinMap/IO_PinGroup : {2}"</summary>
        /// <remarks>{0} = pin group name, {1} = source msg, {2} = different pins</remarks>
        public static readonly ErrorCode E_MismatchPinGroup_01 = new(
            enumErrorCategory: EnumErrorCategory.PreAction,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.PinGroup,
            code: 1,
            template: "Pin group {0} from IO_PinMap/IO_PinGroup has inconsistent pin definitions between {1}. "
                    + "Please check these pins in IO_PinMap/IO_PinGroup : {2}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please compare the pin definitions and resolve any inconsistencies.");

        /// <summary>Template: "Pin group {0} type mismatch: {1} uses type \"{2}\", while the corresponding pin group in IO_PinMap/IO_PinGroup uses type \"{3}\","</summary>
        /// <remarks>{0} = pin group name, {1} = source msg?, {2} = source pin type, {3} = IO_PinMap/IO_PinGroup pin type</remarks>
        public static readonly ErrorCode E_MismatchPinGroup_02 = new(
            enumErrorCategory: EnumErrorCategory.PreAction,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.PinGroup,
            code: 2,
            template: "Pin group {0} type mismatch: {1} uses type \"{2}\", "
                    + "while the corresponding pin group in IO_PinMap/IO_PinGroup uses type \"{3}\",",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please verify the pin group type definitions in both sources and ensure they match.");

        /// <summary>Template: "Missing DC category name."</summary>
        public static readonly ErrorCode E_MissingDcCategory_01 = new(
            enumErrorCategory: EnumErrorCategory.PreAction,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.DcCategory,
            code: 1,
            template: "Missing DC category name.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "A category name is missing in a flagged column of the voltage table. "
                    + "Every category column must have a non-empty name in the header row.");

        /// <summary>Template: "Required sheet IO_PinMap or IO_Continuity was not found in the test plan."</summary>
        public static readonly ErrorCode E_MissingDocument_01 = new(
            enumErrorCategory: EnumErrorCategory.PreAction,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Document,
            code: 1,
            template: "Required sheet IO_PinMap or IO_Continuity was not found in the test plan.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify that the required sheets exist in the test plan and have not been renamed or removed.");


        /// <summary>Template: "Missing IO_ignore_list sheet in the test plan."</summary>
        public static readonly ErrorCode E_MissingDocument_02 = new(
            enumErrorCategory: EnumErrorCategory.PreAction,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Document,
            code: 2,
            template: "Missing IO_ignore_list sheet in the test plan.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify that the required sheet exists in the test plan.");

        /// <summary>Template: "Cannot find {0}: "{1}"."</summary>
        /// <remarks>{0} = document type, {1} = path</remarks>
        public static readonly ErrorCode E_MissingFile_01 = new(
            enumErrorCategory: EnumErrorCategory.PreAction,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.File,
            code: 1,
            template: "Cannot find {0}: \"{1}\".",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify that the specified path is correct and that the referenced file exists.");

        /// <summary>Template: "Cannot find first header '{0}' in the sheet."</summary>
        /// <remarks>{0} = needed header</remarks>
        public static readonly ErrorCode E_MissingHeader_01 = new(
            enumErrorCategory: EnumErrorCategory.PreAction,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Header,
            code: 1,
            template: "Cannot find first header '{0}' in the sheet.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check that the expected header row exists and the header name matches exactly.");

        /// <summary>Template: "The function: {0} can not find in {1} library!"</summary>
        /// <remarks>{0} = function name, {1} = C#/VBT</remarks>
        public static readonly ErrorCode E_MissingLibrary_01 = new(
            enumErrorCategory: EnumErrorCategory.PreAction,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Library,
            code: 1,
            template: "The function: {0} can not find in {1} library!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please verify that the function exists in the specified library and that the library version and function reference are correct.");

        /// <summary>Template: "Missing Parameter in {0}({1}) : {2}"</summary>
        /// <remarks>{0} = function name, {1} = VBT/.NET, {2} = parameter name</remarks>
        public static readonly ErrorCode E_MissingParameter_01 = new(
            enumErrorCategory: EnumErrorCategory.PreAction,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Parameter,
            code: 1,
            template: "Missing parameter \"{2}\" in function \"{0}\" of library \"({1})\".",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please ensure the parameter exists in the library and that the argument name matches the documented definition.");

        /// <summary>Template: "Non-adjacent DC category: {0}"</summary>
        /// <remarks>{0} = category name</remarks>
        public static readonly ErrorCode E_RuleViolationColumn_01 = new(
            enumErrorCategory: EnumErrorCategory.PreAction,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Column,
            code: 1,
            template: "Non-adjacent DC category: {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Columns belonging to the same DC category must be adjacent. "
                    + "Please check whether any previous column has the same DC category name.");

        /// <summary>Template: "N pin need to behind to P pin of Pin group : {0} for VBT issue"</summary>
        /// <remarks>{0} = pin pair or signal name</remarks>
        public static readonly ErrorCode E_RuleViolationPin_01 = new(
            enumErrorCategory: EnumErrorCategory.PreAction,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 1,
            template: "N pin need to behind to P pin of Pin group : {0} for library issue",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify that the differential pair pins are correctly defined and paired. "
                    + "Ensure the positive and negative pins are assigned to the same differential pair group.");

        /// <summary>Template: "Pin: {0}, Category: {1}, {2}:{3} {4}:{5} {3}{6}{5}"</summary>
        /// <remarks>{0} = pin name, {1} = dc category name, {2} = compared seletor1, {3} = compared value1, {4} = compared seletor2, {5} = compared value2,
        /// {6} = comparison operator</remarks>
        public static readonly ErrorCode E_RuleViolationVoltage_01 = new(
            enumErrorCategory: EnumErrorCategory.PreAction,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Voltage,
            code: 1,
            template: "Pin: {0}, Category: {1}, {2}:{3} {4}:{5} {3}{6}{5}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the voltage table and check the voltage values for the flagged pin and category. "
                    + "Ensure LV <= NV <= HV and all values are self-consistent.");

        /// <summary>Template: "Power pin name "{0}" contains spaces."</summary>
        /// <remarks>{0} = pin name</remarks>
        public static readonly ErrorCode W_InvalidFormat_01 = new(
            enumErrorCategory: EnumErrorCategory.PreAction,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 1,
            template: "Power pin name \"{0}\" contains spaces.",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Remove any spaces from the power pin name and ensure it follows the expected naming convention.");

        /// <summary>Template: "DC category value "{0}" is not a valid number or percentage."</summary>
        /// <remarks>{0} = dc category value</remarks>
        public static readonly ErrorCode W_InvalidFormat_02 = new(
            enumErrorCategory: EnumErrorCategory.PreAction,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 2,
            template: "DC category value \"{0}\" is not a valid number or percentage.",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Verify that the DC category value in the Voltage Table is a valid number (e.g. 1.8) or percentage (e.g. 90%).");

        /// <summary>Template: "DC Category "{0}", Pin "{1}" contains invalid NV/{2} content. NV: {3}, {2}: {4}."</summary>
        /// <remarks>{0} = category name, {1} = pin name, {2} = HV or LV, {3} = NV value, {4} = value of HV/LV</remarks>
        public static readonly ErrorCode W_InvalidVoltage_01 = new(
            enumErrorCategory: EnumErrorCategory.PreAction,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Voltage,
            code: 1,
            template: "DC Category \"{0}\", Pin \"{1}\" contains invalid NV/{2} content. NV: {3}, {2}: {4}.",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Please check the values in the corresponding pin under this category and ensure they are valid numeric values.");

        /// <summary>Template: "Voltage value cannot be less than 0."</summary>
        public static readonly ErrorCode W_InvalidVoltage_02 = new(
            enumErrorCategory: EnumErrorCategory.PreAction,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Voltage,
            code: 2,
            template: "Voltage value cannot be less than 0.",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Verify that the voltage value is a non-negative number.");

        /// <summary>Template: "DC category name "{0}" starts with an unrecognized test block: "{1}"."</summary>
        /// <remarks>{0} = test part name</remarks>
        public static readonly ErrorCode W_RuleViolationDcCategory_01 = new(
            enumErrorCategory: EnumErrorCategory.PreAction,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.DcCategory,
            code: 1,
            template: "DC category name \"{0}\" starts with an unrecognized test block: \"{1}\".",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Verify that the specified DC Category is intended and matches the expected block naming conventions"
                    + "(evs ,ids ,mbist ,hardip ,conti ,nwire ,efuse ,rtos ,rto ,sa ,sachain ,td ,tdchain ,scan ,bincut ,htol).");

        /// <summary>Template: "DC Category "{0}", Pin "{1}": {2} value ({3}) {4} NV."</summary>
        /// <remarks>{0} = category name, {1} = pin name, {2} = HV/LV, {3} = value, {4} = Comparison operator ("<" or ">")</remarks>
        public static readonly ErrorCode W_RuleViolationVoltage_01 = new(
            enumErrorCategory: EnumErrorCategory.PreAction,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Voltage,
            code: 1,
            template: "DC Category \"{0}\", Pin \"{1}\": {2} value ({3}) {4} NV.",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Please review the corresponding pin for the specified category in the Voltage Table and ensure that LV <= NV <= HV.");

        /// <summary>Template: "Pin "{0}" {1} value ({2}) does not match the VAlt {1} value ({3}) defined for Mbist dc category "{4}"."</summary>
        /// <remarks>{0} = pin name, {1} = HV/LV/NV, {2} = vmain value, {3} = valt value, {4} = dc category name</remarks>
        public static readonly ErrorCode W_RuleViolationVoltage_02 = new(
            enumErrorCategory: EnumErrorCategory.PreAction,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Voltage,
            code: 2,
            template: "Pin \"{0}\" {1} vmain value ({2}) does not match the {1} valt value ({3}) defined for Mbist dc category \"{4}\".",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Verify that the specified power pin value matches the VAlt definition for the corresponding MBist DC category.");
    }
}
