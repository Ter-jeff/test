using CommonLib.ErrorReport.Base;

namespace CommonLib.ErrorReport.ErrorCodes
{
    public static class BasicErrorType
    {
        /// <summary>Template: "The pin {0} in IO_PinGroup can not be found in IO_PinMap !!!"</summary>
        /// <remarks>{0} = rule or condition description</remarks>
        public static readonly ErrorCode E_MissingPin_01 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 1,
            template: "The pin {0} in IO_PinGroup can not be found in IO_PinMap !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "The pin specified in IO_PinGroup cannot be found in IO_PinMap. "
                    + "Please ensure that the pin exists in IO_PinMap.");

        /// <summary>Template: "The pin group {0} has more than two pin types !!!"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingPin_02 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 2,
            template: "The pin group {0} has more than two pin types !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "The pin types are defined in the IO_PinMap sheet. "
                    + "Please verify the pin type settings in IO_PinMap.");

        /// <summary>Template: "The pin group name {0} and pin name can not be the same !!!"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingPin_03 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 3,
            template: "The pin group name {0} and pin name can not be the same !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "The pin group name specified in IO_PinGroup conflicts with a pin name. "
                    + "Please ensure that pin group names and pin names are unique.");

        /// <summary>Template: "Should define io infos in {0} for {1}!!!"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_MissingPin_04 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 4,
            template: "Should define io infos in {0} for {1}!!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please check the IO configuration and identify any missing IO definitions. "
                    + "Add the required information, regenerate the project, and confirm that the issue is resolved.");

        /// <summary>Template: "Should define io infos concurrent in {0} for {1}!!!"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_MissingPin_05 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 5,
            template: "Should define io infos concurrent in {0} for {1}!!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please ensure all required concurrent IO information is configured correctly. "
                    + "Add any missing definitions and regenerate the output files.");

        /// <summary>Template: "Missing PinMap in TestPlan"</summary>
        public static readonly ErrorCode E_MissingPin_06 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 6,
            template: "Missing PinMap in TestPlan",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "The PinMap definition is missing in the TestPlan. "
                    + "Please ensure that the corresponding PinMap is configured in the TestPlan.");

        /// <summary>Template: "Column[{0}] The pin: {1} does not exist in the PinMap"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_MissingPin_07 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 7,
            template: "Column[{0}] The pin: {1} does not exist in the PinMap",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please ensure the pin name is defined correctly in the PinMap. "
                    + "Add any missing pin definitions or update the pin name, then regenerate the output files.");

        /// <summary>Template: "Job: {0} does not define in FlowMain(Test Plan)/jobMapping(Setting file) sheet !!!"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingPin_08 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 8,
            template: "Job: {0} does not define in FlowMain(Test Plan)/jobMapping(Setting file) sheet !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the job configuration and ensure the required job is defined correctly. "
                + "Add any missing job entries or update the job reference, then regenerate the output files.");

        /// <summary>Template: "Jobs '{0}' need to define in '{1}' sheet."</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_MissingPin_09 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 9,
            template: "Jobs '{0}' need to define in '{1}' sheet.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the job configuration and ensure all required jobs are defined correctly. "
                    + "Add any missing job entries and regenerate the output files.");

        /// <summary>Template: "The {0} net pin {1} of {2} has existed in TestSetting !!!"</summary>
        /// <remarks>{0} = ?, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_MissingPin_10 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 10,
            template: "The {0} net pin {1} of {2} has existed in TestSetting !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the net pin configuration and ensure each net pin is defined only once. "
                    + "Remove any duplicate definitions and regenerate the output files.");

        /// <summary>Template: "{0} can't be found in powerName list of <Judge_stored_IDS>!!!"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingPin_11 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 11,
            template: "{0} can't be found in powerName list of <Judge_stored_IDS>!!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please verify that the required power name is included in the configured power name list. "
                    + "Add any missing power names and regenerate the output files.");

        /// <summary>Template: "The power name : {0} can't be found in the config setting."</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingPin_12 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 12,
            template: "The power name : {0} can't be found in the config setting.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the power configuration and ensure all required power names are defined correctly. "
                    + "Add any missing power name definitions or update invalid references, then regenerate the output files.");

        /// <summary>Template: "{0} can't be found in flow !!!"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingPin_13 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 13,
            template: "{0} can't be found in flow !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the flow configuration and ensure all required flow entries are defined correctly. "
                    + "Add any missing entries or update invalid references, then regenerate the output files.");

        /// <summary>Template: "{0} can't be found in flowCsharp !!!"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingPin_14 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 14,
            template: "{0} can't be found in flowCsharp !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the FlowCsharp configuration and ensure all required entries are defined correctly. "
                    + "Add any missing entries or update invalid references, then regenerate the output files.");

        /// <summary>Template: "Incorrect Value : {0} in column FS/DD"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingPin_15 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 15,
            template: "Incorrect Value : {0} in column FS/DD",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Please review the FS/DD configuration and ensure all values are specified correctly. "
                    + "Update any invalid entries and regenerate the output files.");

        /// <summary>Template: "Missing conti pin/pin group: {0} in pin map!"</summary>
        /// <remarks>{0} = pin name</remarks>
        public static readonly ErrorCode E_MissingPin_16 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 16,
            template: "Missing conti pin/pin group: {0} in pin map!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify that all required pin names are defined in the pin configuration. "
                    + "Check for typos or missing entries in the pin list.");

        /// <summary>Template: "Missing pin in voltage table: {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingPin_17 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 17,
            template: "Missing pin in voltage table: {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the voltage table configuration and ensure all required pins are defined correctly. "
                    + "Add any missing pin definitions or update invalid references, then regenerate the output files.");

        /// <summary>Template: "Missing pin in PowerInfo for iFold: {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingPin_18 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 18,
            template: "Missing pin in PowerInfo for iFold: {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the PowerInfo configuration and ensure all required pins are defined correctly for iFold settings. "
                    + "Add any missing pin definitions or update invalid references, then regenerate the output files.");

        /// <summary>Template: "The voltage of IO pin group {0} had different voltage"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingPin_19 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 19,
            template: "The voltage of IO pin group {0} had different voltage.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the IO pin group configuration and ensure all pins in the same group use consistent voltage settings. "
                    + "Correct any voltage inconsistencies and regenerate the output files.");

        /// <summary>Template: "Pin (group):{2} in Time Set file:{0};TSet:{1} is missing. "</summary>
        /// <remarks>{0} = ?, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_MissingPin_20 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 20,
            template: "Pin (group):{2} in Time Set file:{0};TSet:{1} is missing.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the Time Set configuration and ensure all required pins and pin groups are defined correctly. "
                    + "Add any missing definitions or update invalid references, then regenerate the output files.");

        /// <summary>Template: "Can't generate power level for pin/pin group: {0} from voltage table, if it just be used to crete variable in DC Specs, please ignore this or make sure {1} had been define this pin/pin group!!!"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_MissingPin_21 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 21,
            template: "Can't generate power level for pin/pin group: {0} from voltage table, if it just be used to crete variable in DC Specs, please ignore this or make sure {1} had been define this pin/pin group!!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Verify that all required pin names are defined in the pin configuration. "
                    + "Check for typos or missing entries in the pin list.");

        /// <summary>Template: "BallName : {0} Not exist in BumpName!!!"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingPin_22 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 22,
            template: "BallName : {0} Not exist in BumpName!!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "A business logic constraint has been violated. "
                    + "Review the rule requirements and correct the input data to satisfy the constraint.");

        /// <summary>Template: "Ifold level {0} defined in PowerInfo is greater than Ifold limit for its instrument, please check!!!"</summary>
        /// <remarks>{0} = condition or context description</remarks>
        public static readonly ErrorCode E_ContiIfold_01 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.IFold,
            code: 1,
            template: "Ifold level {0} defined in PowerInfo is greater than Ifold limit for its instrument, please check!!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the continuous ifold condition definition. "
                    + "Ensure the condition is valid and consistent with the test flow requirements.");

        /// <summary>Template: "TN_Assignment sheet contains duplicate sheets : {0}"</summary>
        /// <remarks>{0} = item name or context</remarks>
        public static readonly ErrorCode E_DuplicateItems_01 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Duplicate,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 1,
            template: "TN_Assignment sheet contains duplicate sheets : {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the TN_Assignment configuration and ensure each sheet is defined only once. "
                    + "Remove any duplicate entries and regenerate the output files.");

        /// <summary>Template: "IoInfo sheet contains duplicate pin level : {0}, remove from {1}"</summary>
        /// <remarks>{0} = item name or context</remarks>
        public static readonly ErrorCode E_DuplicateItems_02 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Duplicate,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 2,
            template: "IoInfo sheet contains duplicate pin level: {0} definition duplicate remove from {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the IoInfo configuration and ensure each pin level is defined only once. "
                    + "Remove any duplicate pin level definitions and regenerate the output files.");

        /// <summary>Template: "Voltage table contains duplicate single pin in : {0}"</summary>
        /// <remarks>{0} = item name or context</remarks>
        public static readonly ErrorCode E_DuplicateItems_03 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Duplicate,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 3,
            template: "Single pin: {0} is duplicate in pin groups: {1} and {2}, please check the sheet: {3}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the pin group configuration and ensure each single pin is assigned to only one pin group. "
                    + "Remove any duplicate pin assignments and regenerate the output files.");

        /// <summary>Template: "Can not find common block."</summary>
        public static readonly ErrorCode E_MissingBlock_01 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Block,
            code: 1,
            template: "Can not find common block.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "A common block cannot be found. "
                    + "Please ensure that the common block is defined correctly.");

        /// <summary>Template: "Can not find common block."</summary>
        public static readonly ErrorCode E_MissingBlock_02 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Block,
            code: 2,
            template: "Can not find common block.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "A common block cannot be found. "
                    + "Please ensure that the common block is defined correctly.");

        /// <summary>Template: "Omit column \"{0}\""</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingBlock_03 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Block,
            code: 3,
            template: "Omit column \"{0}\"",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the input file and ensure all required columns are included. "
                    + "Add any missing columns and regenerate the output files.");

        /// <summary>Template: "{0} in {1} block does not exist in common io pins."</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_MissingBlock_04 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Block,
            code: 4,
            template: "{0} in {1} block does not exist in common io pins.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the common IO pin configuration and ensure all required pins are defined correctly. "
                    + "Add any missing pin definitions or update invalid references, then regenerate the output files.");

        /// <summary>Template: "Omit column \"PowerSequence\""</summary>
        public static readonly ErrorCode E_MissingBlock_05 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Block,
            code: 5,
            template: "Omit column \"PowerSequence\"",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Column PowerSequence is missing. "
                    + "Please ensure that the PowerSequence column is included in the input file.");

        /// <summary>Template: "Pin '{0}' does not exist in TestSetting."</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingBlock_06 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Block,
            code: 6,
            template: "Pin '{0}' does not exist in TestSetting.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the TestSetting configuration and ensure all required pins are defined correctly. "
                    + "Add any missing pin definitions or update invalid references, then regenerate the output files.");

        /// <summary>Template: "Sheet: {0} Not Exist."</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingBlock_07 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Block,
            code: 7,
            template: "Sheet: {0} Not Exist.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the input file and ensure all required sheets are included correctly. "
                    + "Add any missing sheets and regenerate");

        /// <summary>Template: "The TestPlan Error."</summary>
        public static readonly ErrorCode E_FormatError_01 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 1,
            template: "The TestPlan Error.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the field value against the expected format specification. "
                    + "Ensure the value type and range conform to requirements.");

        /// <summary>Template: "The TestPlan Error."</summary>
        public static readonly ErrorCode E_FormatError_02 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 2,
            template: "The TestPlan Error.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the field value against the expected format specification. "
                    + "Ensure the value type and range conform to requirements.");

        /// <summary>Template: "Logic Pin: {0} do not have the *_Valt voltage in {1}"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_FormatError_03 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 3,
            template: "Logic Pin: {0} do not have the *_Valt voltage in {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the voltage configuration and ensure all required *_Valt voltages are defined correctly. "
                    + "Add any missing voltage definitions or update invalid configurations, then regenerate the output files.");

        /// <summary>Template: "Sram Pin: {0} do not have the *_Valt voltage in {1}"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_FormatError_04 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 4,
            template: "Sram Pin: {0} do not have the *_Valt voltage in {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the voltage configuration and ensure all required *_Valt voltages are defined correctly. "
                    + "Add any missing voltage definitions or update invalid configurations, then regenerate the output files.");

        /// <summary>Template: "{0} and {1} should have the different source bit"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_FormatError_05 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 5,
            template: "{0} and {1} should have the different source bit",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the source bit configuration and ensure each entry is assigned a unique source bit value. "
                    + "Update any conflicting assignments and regenerate the output files.");

        /// <summary>Template: "The preserved pin should have the same source bit"</summary>
        public static readonly ErrorCode E_FormatError_06 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 6,
            template: "The preserved pin should have the same source bit",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "The preserved pins have different source bits. "
                    + "Please ensure that all preserved pins use the same source bit.");

        /// <summary>Template: "{0}: The length of Timing set name exceeds the maximum 31 chars "</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_FormatError_07 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 7,
            template: "{0}: The length of Timing set name exceeds the maximum 31 chars ",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the timing set configuration and ensure all timing set names comply with the naming length requirements. "
                    + "Shorten any names that exceed the allowed length and regenerate the output files.");

        /// <summary>Template: "{0}: The length of Flow sheet name exceeds the maximum 31 chars "</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_FormatError_08 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 8,
            template: "{0}: The length of Flow sheet name exceeds the maximum 31 chars ",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the flow sheet configuration and ensure all sheet names comply with the naming length requirements. "
                    + "Shorten any sheet names that exceed the allowed length and regenerate the output files.");

        /// <summary>Template: "{0}: The length of Inst sheet name exceeds the maximum 31 chars "</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_FormatError_09 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 9,
            template: "{0}: The length of Inst sheet name exceeds the maximum 31 chars ",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the Inst sheet configuration and ensure all sheet names comply with the naming length requirements. "
                    + "Shorten any sheet names that exceed the allowed length and regenerate the output files.");

        /// <summary>Template: "The directory is not exist : {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_FormatError_10 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 10,
            template: "The directory is not exist : {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the directory path configuration and ensure all required paths are valid and accessible. "
                    + "Create any missing directories or update invalid paths, then regenerate the output files.");

        /// <summary>Template: "Different pin name from NV: {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_FormatError_11 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 11,
            template: "Different pin name from NV: {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the pin configuration and ensure all pin names match the corresponding NV definitions. "
                    + "Update any inconsistent pin names and regenerate the output files.");

        /// <summary>Template: "Wrong format ratio: {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_FormatError_12 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 12,
            template: "Wrong format ratio: {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the ratio configuration and ensure all ratio values follow the required format. "
                    + "Correct any invalid ratio entries and regenerate the output files.");

        /// <summary>Template: "Different pin name from NV: {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_FormatError_13 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 13,
            template: "Different pin name from NV: {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the pin configuration and ensure all pin names match the corresponding NV definitions. "
                    + "Update any inconsistent pin names and regenerate the output files.");

        /// <summary>Template: "Omit \"{0}\" value for pin: {1}"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_FormatError_14 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 14,
            template: "Omit \"{0}\" value for pin: {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the pin configuration and ensure all required values are specified correctly. "
                    + "Add any missing values and regenerate the output files.");

        /// <summary>Template: "Omit value for pin: {0} in column {1}"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_FormatError_15 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 15,
            template: "Omit value for pin: {0} in column {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the pin configuration and ensure all required values are specified correctly. "
                    + "Add any missing values and regenerate the output files.");

        /// <summary>Template: "Non numerical \"{0}\" value for pin: {1}"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_FormatError_16 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 16,
            template: "Non numerical \"{0}\" value for pin: {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the pin configuration and ensure all required values are specified using valid numeric formats. "
                    + "Correct any non-numeric values and regenerate the output files.");

        /// <summary>Template: "Non numerical \"{0}\" value for pin: {1}"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_FormatError_17 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 17,
            template: "Non numerical \"{0}\" value for pin: {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the pin configuration and ensure all required values are specified using valid numeric formats. "
                    + "Correct any non-numeric values and regenerate the output files.");

        /// <summary>Template: "Non numerical unknown type value for pin: {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_FormatError_18 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 18,
            template: "Non numerical unknown type value for pin: {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the value and type settings for pin {0}. "
                    + "Ensure that a valid numeric value is provided and the type is recognized.");

        /// <summary>Template: "Row:{3} The period [{0}] is illegal for tSet {1} on Pin {2}"</summary>
        /// <remarks>{0} = ?, {1} = ?, {2} = ?, {3} = ?</remarks>
        public static readonly ErrorCode E_FormatError_19 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 19,
            template: "Row:{3} The period [{0}] is illegal for tSet {1} on Pin {2}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the timing configuration and ensure all period values are specified using valid formats and supported ranges. "
                    + "Correct any invalid period settings and regenerate the output files.");

        /// <summary>Template: "Key Don't Exist In {0} : {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_FormatError_20 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 20,
            template: "Key Don't Exist In {0} : {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the configuration and ensure all required keys are defined correctly. "
                    + "Add any missing key definitions or update invalid references, then regenerate the output files.");

        /// <summary>Template: "Must Exist Item :{0} Not Exist."</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_FormatError_21 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 21,
            template: "Must Exist Item :{0} Not Exist.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the configuration and ensure all required items are defined correctly. "
                    + "Add any missing item definitions and regenerate the output files.");

        /// <summary>Template: "Multi-ShiftInFreq TimeSet Sheet problem @ {0} : {1}"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_FormatError_22 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 22,
            template: "Multi-ShiftInFreq TimeSet Sheet problem @ {0} : {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the Multi-ShiftInFreq TimeSet configuration and ensure all settings are defined correctly. "
                    + "Correct any invalid configurations and regenerate the output files.");

        /// <summary>Template: "Can't parse {0} in {1}:Row{2}, it should be a integer."</summary>
        /// <remarks>{0} = ?, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_FormatError_23 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 23,
            template: "Can't parse {0} in {1}:Row{2}, it should be a integer.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the configuration and ensure all required values are specified using valid integer formats. "
                    + "Correct any invalid values and regenerate the output files.");


        /// <summary>Template: "Can't parse {0} in {1}:Row{2}, it should be a integer."</summary>
        /// <remarks>{0} = ?, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_FormatError_24 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 24,
            template: "Can't parse {0} in {1}:Row{2}, it should be a integer.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the configuration and ensure all required values are specified using valid integer formats. "
                    + "Correct any invalid values and regenerate the output files.");


        /// <summary>Template: "Can't parse {0} in {1}:Row{2}, it should be a integer."</summary>
        /// <remarks>{0} = ?, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_FormatError_25 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 25,
            template: "Can't parse {0} in {1}:Row{2}, it should be a integer.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the configuration and ensure all required values are specified using valid integer formats. "
                    + "Correct any invalid values and regenerate the output files.");

        /// <summary>Template: "{0}:{1} in {2}:Row{3}, should be in the range of 1-9999."</summary>
        /// <remarks>{0} = ?, {1} = ?, {2} = ?, {3} = ?</remarks>
        public static readonly ErrorCode E_FormatError_26 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 26,
            template: "{0}:{1} in {2}:Row{3}, should be in the range of 1-9999.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the configuration and ensure all required values are within the supported range. "
                    + "Correct any out-of-range values and regenerate the output files.");

        /// <summary>Template: "{0}:{1} in {2}:Row{3}, should be in the range of 1-9999."</summary>
        /// <remarks>{0} = ?, {1} = ?, {2} = ?, {3} = ?</remarks>
        public static readonly ErrorCode E_FormatError_27 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 27,
            template: "{0}:{1} in {2}:Row{3}, should be in the range of 1-9999.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the configuration and ensure all required values are within the supported range. "
                    + "Correct any out-of-range values and regenerate the output files.");

        /// <summary>Template: "Can't get bin number sheet in test plan, missing {0} header, "</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_FormatError_28 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 28,
            template: "Can't get bin number sheet in test plan, missing {0} header, it should have {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7} ,{8}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the bin number sheet configuration and ensure all required headers are defined correctly. "
                    + "Add any missing headers and regenerate the output files.");

        /// <summary>Template: "Can't get bin number sheet in test plan, missing {0} header."</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_FormatError_29 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 29,
            template: "Can't get bin number sheet in test plan, missing {0} header.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the bin number sheet configuration and ensure all required headers are defined correctly. "
                    + "Add any missing headers and regenerate the output files.");

        /// <summary>Template: "VM Vector name is mismatch, Pattern:{0}, VmVector:{1} "</summary>
        /// <remarks>{0} = field or context description, {1} = ?</remarks>
        public static readonly ErrorCode E_FormatWarning_01 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 1,
            template: "VM Vector name is mismatch, Pattern:{0}, VmVector:{1} ",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the pattern vector name and VM vector name configuration. "
                    + "Ensure that the VM vector name matches the corresponding pattern vector name.");

        /// <summary>Template: "VM Vector name is mismatch, Pattern:{0}, VmVector:{1} "</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_FormatWarning_02 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 2,
            template: "VM Vector name is mismatch, Pattern:{0}, VmVector:{1} ",
            enumErrorLevel: EnumErrorLevel.Error,
             guidance: "Check the pattern vector name and VM vector name configuration. "
                    + "Ensure that the VM vector name matches the corresponding pattern vector name.");

        /// <summary>Template: "The ELB instance {0} can not be found !!!"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_FormatWarning_03 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Instance,
            code: 3,
            template: "The ELB instance {0} can not be found !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Review the flagged value for potential formatting inconsistencies. "
                    + "This warning does not block execution but may cause unexpected behavior.");

        /// <summary>Template: "The ELB instance {0} can not be found !!!"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_FormatWarning_04 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Instance,
            code: 4,
            template: "The ELB instance {0} can not be found !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Review the flagged value for potential formatting inconsistencies. "
                    + "This warning does not block execution but may cause unexpected behavior.");

        /// <summary>Template: "The limit of ELB instance are duplicated -{0} !!!"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_FormatWarning_05 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Duplicate,
            enumErrorTarget: EnumErrorTarget.Limit,
            code: 5,
            template: "The limit of ELB instance are duplicated -{0} !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Review the flagged value for potential formatting inconsistencies. "
                    + "This warning does not block execution but may cause unexpected behavior.");

        /// <summary>Template: "The pin {0} can not be found in sheet {1}"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_FormatWarning_06 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 6,
            template: "The pin {0} can not be found in sheet {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the pin configuration and ensure all required pins are defined correctly. "
                    + "Add any missing pin definitions or update invalid references, then regenerate the output files.");

        /// <summary>Template: "The I/O pin {0} in {1} can not be found in {2} ({3}, {4})"</summary>
        /// <remarks>{0} = ?, {1} = ?, {2} = ?, {3} = ?, {4} = ?</remarks>
        public static readonly ErrorCode E_FormatWarning_07 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 7,
            template: "The I/O pin {0} in {1} can not be found in {2} ({3}, {4})",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the I/O pin configuration and ensure all required pin definitions are defined correctly. "
                    + "Add any missing pin definitions or update invalid references, then regenerate the output files.");

        /// <summary>Template: "The I/O pin {0} in {1} can not be found in {2} ({3}, {4})"</summary>
        /// <remarks>{0} = ?, {1} = ?, {2} = ?, {3} = ?, {4} = ?</remarks>
        public static readonly ErrorCode E_FormatWarning_08 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 8,
            template: "The I/O pin {0} in {1} can not be found in {2} ({3}, {4})",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the I/O pin configuration and ensure all required pin definitions are defined correctly. "
                    + "Add any missing pin definitions or update invalid references, then regenerate the output files.");

        /// <summary>Template: "The I/O pin {0} in {1} can not be found in {2} ({3}, {4})"</summary>
        /// <remarks>{0} = ?, {1} = ?, {2} = ?, {3} = ?, {4} = ?</remarks>
        public static readonly ErrorCode E_FormatWarning_09 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 9,
            template: "The I/O pin {0} in {1} can not be found in {2} ({3}, {4})",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the I/O pin configuration and ensure all required pin definitions are defined correctly. "
                    + "Add any missing pin definitions or update invalid references, then regenerate the output files.");

        /// <summary>Template: "The pin {0} in {1} can not be found in {2}"</summary>
        /// <remarks>{0} = ?, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_FormatWarning_10 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 10,
            template: "The pin {0} in {1} can not be found in {2}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the pin configuration and ensure all required pins are defined correctly. "
                    + "Add any missing pin definitions or update invalid references, then regenerate the output files.");

        /// <summary>Template: "The pin {0} can not be found in sheet {1}"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_FormatWarning_11 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 11,
            template: "The pin {0} can not be found in sheet {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the pin configuration and ensure all required pins are defined correctly. "
                    + "Add any missing pin definitions or update invalid references, then regenerate the output files.");

        /// <summary>Template: "The pin {0} can not be found in sheet {1}"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_FormatWarning_12 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 12,
            template: "The pin {0} can not be found in sheet {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the pin configuration and ensure all required pins are defined correctly. "
                    + "Add any missing pin definitions or update invalid references, then regenerate the output files.");

        /// <summary>Template: "Needs to fill FS or DD in FS/DD Column !!!"</summary>
        public static readonly ErrorCode E_FSDDIssue_01 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Column,
            code: 1,
            template: "Needs to fill FS or DD in FS/DD Column !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the FSDD (Functional Signal Drive/Detect) configuration for the specified component. "
                    + "Verify that all FSDD settings are valid and consistent with the test requirements.");

        /// <summary>Template: "Missing Parameter in {0}({1}) : {2}"</summary>
        /// <remarks>{0} = parameter name, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_MissingParameter_01 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 1,
            template: "Missing Parameter in {0}({1}) : {2}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the configuration sheet and add the required parameter. "
                    + "Check the specification to confirm which parameters are mandatory.");

        /// <summary>Template: "The VBT function: {0} can not find in {1} library!"</summary>
        /// <remarks>{0} = module name, {1} = ?</remarks>
        public static readonly ErrorCode E_MissVbtModule_01 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Module,
            code: 1,
            template: "The VBT function: {0} can not find in {1} library!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please review the VBT library configuration and ensure all required functions are defined correctly. "
                    + "Add any missing function definitions or update invalid references, then regenerate the output files.");

        /// <summary>Template: "Line:{0} format is empty line"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode W_FormatError_01 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 1,
            template: "Line:{0} format is empty line",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Please review the input file and ensure all required entries are provided correctly. "
                    + "Remove any unexpected empty lines or add the required content, then regenerate the output files.");

        /// <summary>Template: "Please check {0} !!! Pin OutClkVoltage value is 0"</summary>
        /// <remarks>{0} = context or instance name</remarks>
        public static readonly ErrorCode W_NwireConfigMismatch_01 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Nwire,
            code: 1,
            template: "Please check {0} !!! Pin OutClkVoltage value is 0",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Review the N-wire output clock voltage setting for the specified instance. "
                    + "Verify the voltage level is within the acceptable range for the target device.");

        /// <summary>Template: "PowerUp instance exist in UF_Instance sheet, Autogen will not generate default power up instance \"PowerUp\", please make sure to put this instance in Flow_Main"</summary>
        public static readonly ErrorCode E_DefaultPowerUpInstanceExist_01 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Instance,
            code: 1,
            template: "PowerUp instance exist in UF_Instance sheet, Autogen will not generate default power up instance \"PowerUp\", please make sure to put this instance in Flow_Main",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Check the UF instance list for an existing default power-up instance. "
                    + "Only one default power-up instance is permitted; remove the duplicate or rename it.");

        public static readonly ErrorCode E_TimeSetOverride_01 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Header,
            code: 1,
            template: "Invalid `TimeSetClockOverride` sheet header. The following columns are required: {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "One or more required columns are missing from the `TimeSetClockOverride` sheet header. "
                    + "Verify that all required columns are present and correctly named.");

        public static readonly ErrorCode E_TimeSetOverride_02 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Column,
            code: 2,
            template: "Columns {0} are required at row {1} in TimeSetClockOverride sheet",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Make sure required columns are not empty on every non-empty rows.");

        public static readonly ErrorCode E_TimeSetOverride_03 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Argument,
            code: 3,
            template: "Frequency variable name [{0}] is invalid in TimeSetClockOverride sheet",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Variable names must start with a letter and may only contain letters, numbers, and underscores");

        public static readonly ErrorCode E_TimeSetOverride_04 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 30,
            template: "Incorrect Frequency format [{0}]",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Make sure Frequency format looks like this \"SOME_VAR\" or \"SOME_VAR = VALUE\"");

        /// <summary>
        /// Frequency value is invalid in TimeSetOverride sheet
        /// </summary>
        public static readonly ErrorCode E_TimeSetOverride_05 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Argument,
            code: 5,
            template: "Frequency value [{0}] is invalid in TimeSetClockOverride sheet",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Frequency value must be a number with optional frequency unit MHz or KHz");

        /// <summary>
        /// Frequency variable is assigned multiple conflicting values
        /// </summary>
        public static readonly ErrorCode E_TimeSetOverride_06 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Argument,
            code: 6,
            template: "Frequency variable `{0}` is assigned multiple conflicting values ({1}) in TimeSetClockOverride sheet",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Ensure that each variable is assigned a single consistent value throughout the sheet." +
            " If the variable is defined multiple times, all assignments must use the same value.");

        /// <summary>
        /// Pin/Group Setup does not allow empty cell
        /// </summary>
        public static readonly ErrorCode E_TimeSetOverride_07 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Argument,
            code: 7,
            template: "Pin/Group Setup does not allow empty cell in TimeSetClockOverride sheet",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Pin/Group Setup cannot be empty. Only support `io`, `clock` and `clock_2x`");

        /// <summary>
        /// Pin/Group Setup value is invalid
        /// </summary>
        public static readonly ErrorCode E_TimeSetOverride_08 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Argument,
            code: 8,
            template: "Pin/Group Setup value {0} is invalid in TimeSetClockOverride sheet",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Pin/Group Setup only support `io`, `clock` and `clock_2x`");

        /// <summary>
        /// Data Src value is invalid
        /// </summary>
        public static readonly ErrorCode E_TimeSetOverride_09 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Argument,
            code: 9,
            template: "Data Src value {0} is invalid",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Data Src only support `ALLHI`, `ALLLO`, `PA`, `PAT`, `PATHI`, `PATLO`, `PATNOT`");

        /// <summary>
        /// Data Fmt value is invalid
        /// </summary>
        public static readonly ErrorCode E_TimeSetOverride_10 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Argument,
            code: 10,
            template: "Data Fmt value {0} is invalid",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Data Fmt only support `NR`, `RH`, `RL`, `STAY`." +
            " And suffix `-2X` is only allowed when Setup is `clock_2x`");

        /// <summary>
        /// Data Src ALLHI must return low `RL` in TimeSetClockOverride sheet
        /// </summary>
        public static readonly ErrorCode E_TimeSetOverride_11 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Argument,
            code: 11,
            template: "Data Src ALLHI must return low `RL` in TimeSetClockOverride sheet",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Make sure TimeSetClockOverride sheet has correct Data Src and Fmt");

        /// <summary>
        /// Data Src ALLLO must return high `RH` in TimeSetOverride sheet
        /// </summary>
        public static readonly ErrorCode E_TimeSetOverride_12 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Argument,
            code: 12,
            template: "Data Src ALLLO must return high `RH` in TimeSetClockOverride sheet",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Make sure TimeSetClockOverride sheet has correct Data Src and Fmt");

        public static readonly ErrorCode E_TimeSetOverride_13 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.File,
            code: 13,
            template: "Cannot override TimeSet `{0}` because it is not used in the Pattern Dashboard.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Make sure TimeSetClockOverride sheet only override TimeSets inside Pattern Dashboard");

        public static readonly ErrorCode E_TimeSetOverride_14 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Duplicate,
            enumErrorTarget: EnumErrorTarget.Row,
            code: 14,
            template: "TimeSet: `{0}` Pin/Group: `{1}` is already defined at row {2}.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Make sure TimeSetClockOverride sheet has no duplicate `TimeSet` and `Pin/Group` pair under the same `TimeSet File`");

        public static readonly ErrorCode E_TimeSetOverride_15 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Timing,
            code: 15,
            template: "Cannot find `Pin/Group` `{1}` under `TimeSet` `{0}` in TimeSet File `{2}`.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify that the specified `TimeSet` and `Pin/Group` exist in the target TimeSet File" +
            " and match the values defined in the `TimeSetClockOverride` sheet.");

        public static readonly ErrorCode E_TimeSetOverride_16 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Timing,
            code: 16,
            template: "The `Pin/Group Clock Period` formula: `{0}` at row {1} of TimeSet File `{2}` does not contain any variable.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "This might be the TimeSet File error. Specify exactly one variable in the formula. For example: `=(1/_TCK_Freq_Var)`.");

        public static readonly ErrorCode E_TimeSetOverride_17 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Timing,
            code: 17,
            template: "Multiple variables were found in the `Pin/Group Clock Period` formula: `{0}` at row {1} of TimeSet File `{2}`",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Ensure the formula contains exactly one variable reference.");

        public static readonly ErrorCode E_TimeSetOverride_18 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Argument,
            code: 18,
            template: "An unexpected error occurred while applying invalid Pin/Group Setup `{0}` to TimeSet File `{1}` at row {2}." +
            " This value should have been rejected during validation.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "This error indicates an internal consistency issue in Autogen." +
            " Please contact the HardIp owner or support team and provide the generated error report for further investigation.");

        public static readonly ErrorCode E_TimeSetOverride_19 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Argument,
            code: 19,
            template: "While modifying TimeSet File: `{0}` at row {1}, Data Src is set to `ALLHI` but Data Fmt did not set to `RL`",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the final Pin/Group configuration in both the TimeSetClockOverride sheet and the target TimeSet File." +
            " Ensure that any Pin/Group using Data Src `ALLHI` has Data Fmt set to `RL`.");

        public static readonly ErrorCode E_TimeSetOverride_20 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Argument,
            code: 20,
            template: "While modifying TimeSet File: `{0}` at row {1}, Data Src is set to `ALLLO` but Data Fmt did not set to `RH`",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the final Pin/Group configuration in both the TimeSetClockOverride sheet and the target TimeSet File." +
            " Ensure that any Pin/Group using Data Src `ALLLO` has Data Fmt set to `RH`.");

        public static readonly ErrorCode E_TimeSetOverride_21 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Timing,
            code: 21,
            template: "Cannot find `TimeSet` `{0}` in TimeSet File `{1}`.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify that the specified `TimeSet` exist in the target TimeSet File" +
            " and match the values defined in the `TimeSetClockOverride` sheet.");

        public static readonly ErrorCode E_TimeSetOverride_22 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Argument,
            code: 22,
            template: "Brand new Frequency variable {0} must assign a value",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "If a variable does not exist in TimeSet VAR Definition and is assigned in Frequency cell," +
            " you must assign a value to it");

        /// <summary>Template: "Format error in TimeSet sheet '{0}'."</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_InvalidTiming_01 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Timing,
            code: 1,
            template: "Parsing error in TimeSet: {0} Row: {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Review the TimeSet sheet for structural issues and correct the flagged entry.");

        /// <summary>Template: "Equation base variable '{0}' used in Time Set file {1} is not assigned an initial value"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_InvalidTiming_02 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Timing,
            code: 2,
            template: "Equation base variable '{0}' used in Time Set file {1} is not assigned an initial value",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Review the TimeSet sheet for structural issues and correct the flagged entry.");

        /// <summary>Template: "Equation base variable '{0}' used in the context of Time Set file {1} is not assigned value in comment"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_InvalidTiming_03 = new(
            enumErrorCategory: EnumErrorCategory.Basic,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Timing,
            code: 3,
            template: "Equation base variable '{0}' used in the context of Time Set file {1} is not assigned value in comment",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Review the TimeSet sheet for structural issues and correct the flagged entry.");
    }
}
