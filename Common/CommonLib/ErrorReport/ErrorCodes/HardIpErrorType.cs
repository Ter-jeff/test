using CommonLib.ErrorReport.Base;

namespace CommonLib.ErrorReport.ErrorCodes
{
    public static class HardIpErrorType
    {
        /// <summary>Template: "Limit count not enough for ADC convertor use"</summary>
        public static readonly ErrorCode E_ADC_Convertor_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Limit,
            code: 1,
            template: "Limit count not enough for ADC convertor use",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the ADC converter definition for the specified pin or instance. "
                    + "Verify the conversion formula, reference voltage, and scaling factors.");

        /// <summary>Template: "Previous use-limits count:{0} contains simple calc equation usage, need to check"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_ADC_Convertor_02 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Limit,
            code: 2,
            template: "Previous use-limits count:{0} contains simple calc equation usage, need to check",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the ADC converter definition for the specified pin or instance. "
                    + "Verify the conversion formula, reference voltage, and scaling factors.");

        /// <summary>Template: "Calculation equation check failed for '{0}': {1}"</summary>
        /// <remarks>{0} = equation or test name, {1} = details</remarks>
        public static readonly ErrorCode E_CalcEqnCheck_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Equation,
            code: 1,
            template: "CalcEquation :{0} ,convert to Cus_Str_MainProgram:{1} for TTR usage",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Review the calculation equation for syntax and semantic errors. "
                    + "Ensure all referenced variables are defined and the result is within expected bounds.");

        /// <summary>Template: "Calculation equation check failed for '{0}': {1}"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_CalcEqnCheck_02 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Equation,
            code: 2,
            template: "CalcEquation :{0} , can not convert to Cus_Str_MainProgram:{1} for TTR usage, need check",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Review the calculation equation for syntax and semantic errors. "
                    + "Ensure all referenced variables are defined and the result is within expected bounds.");

        /// <summary>Template: "Cannot get SELSRAM setting for '{0}'."</summary>
        /// <remarks>{0} = instance or context name</remarks>
        public static readonly ErrorCode E_CanNotGetSelsramSetting_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Setting,
            code: 1,
            template: "Can not get selsrm setting by Block: {0}; pattern: {1} in {2}.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the SELSRAM instance name and configuration. "
                    + "Ensure the SELSRAM setting is properly defined and accessible for the specified instance.");

        /// <summary>Template: "Error: {0} is Duplicated in PatInfoFile.txt"</summary>
        /// <remarks>{0} = pattern name</remarks>
        public static readonly ErrorCode E_DuplicatePatternInPatInfo_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Duplicate,
            enumErrorTarget: EnumErrorTarget.PatternInfo,
            code: 1,
            template: "Error: {0} is Duplicated in PatInfoFile.txt",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the PatInfo sheet for duplicate pattern entries. "
                    + "Remove the duplicate or rename one entry so each pattern name is unique.");

        /// <summary>Template: "We found duplicate Store name \"{0}\" in HardIP multiple pattern."</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_DuplicateStoreName_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Duplicate,
            enumErrorTarget: EnumErrorTarget.StoreName,
            code: 1,
            template: "We found duplicate Store name \"{0}\" in HardIP multiple pattern.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the Hard IP store name list for duplicate entries. "
                    + "Rename one of the duplicate store names so each name is unique.");

        /// <summary>Template: "Cannot find \"{0}\" in EFUSE_BitDef_Table"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_HIPeFuseCatename_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Argument,
            code: 1,
            template: "Cannot find \"{0}\" in EFUSE_BitDef_Table",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the Hard IP eFuse category name against the EFuse category list. "
                    + "Check for typos or missing category definitions.");

        /// <summary>Template: "Not allow syntax:{0}"</summary>
        /// <remarks>{0} = DSP wave name or context</remarks>
        public static readonly ErrorCode E_HIPeFuseDSPWAVE_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.DSPWave,
            code: 1,
            template: "Not allow syntax:{0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the Hard IP eFuse DSP wave configuration. "
                    + "Ensure the DSP wave name and parameters are consistent with the eFuse specification.");

        /// <summary>Template: "\"{0}\" dspwave size :{1} is not same to efuse_bitdef: {2}"</summary>
        /// <remarks>{0} = ?, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_HIPeFuseDSPWAVE_02 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.DSPWave,
            code: 2,
            template: "\"{0}\" dspwave size :{1} is not same to efuse_bitdef: {2}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the Hard IP eFuse DSP wave configuration. "
                    + "Ensure the DSP wave name and parameters are consistent with the eFuse specification.");

        /// <summary>Template: "Please provide EFUSE_BitDef_Table to do the cross-checking of efuse and hardip."</summary>
        public static readonly ErrorCode E_HIPeFuseInput_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Document,
            code: 1,
            template: "Please provide EFUSE_BitDef_Table to do the cross-checking of efuse and hardip.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the Hard IP eFuse input definition against the eFuse specification. "
                    + "Check the input name, bit address, and allowed values.");

        /// <summary>Template: "Please provide eFuse_HardIP_Table to do the cross-checking of efuse and hardip."</summary>
        public static readonly ErrorCode E_HIPeFuseInput_02 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Document,
            code: 2,
            template: "Please provide eFuse_HardIP_Table to do the cross-checking of efuse and hardip.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the Hard IP eFuse input definition against the eFuse specification. "
                    + "Check the input name, bit address, and allowed values.");

        /// <summary>Template: "{0} func but not found \"EFUSE_BitDef_Table\""</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_HIPeFuseInput_03 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Document,
            code: 3,
            template: "{0} func but not found \"EFUSE_BitDef_Table\"",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the Hard IP eFuse input definition against the eFuse specification. "
                    + "Check the input name, bit address, and allowed values.");

        /// <summary>Template: "{0} func cannot find \"m_catename\" or \"fieName\"(.NET)"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_HIPeFuseInput_04 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Document,
            code: 4,
            template: "{0} func cannot find \"m_catename\" or \"fieName\"(.NET)",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the Hard IP eFuse input definition against the eFuse specification. "
                    + "Check the input name, bit address, and allowed values.");

        /// <summary>Template: "{0} func cannot find \"DSPWAVESIZE\" or \"sampleSize\"(.NET)"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_HIPeFuseInput_05 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Document,
            code: 5,
            template: "{0} func cannot find \"DSPWAVESIZE\" or \"sampleSize\"(.NET)",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the Hard IP eFuse input definition against the eFuse specification. "
                    + "Check the input name, bit address, and allowed values.");

        /// <summary>Template: "{0} func but not found \"EFUSE_BitDef_Table\""</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_HIPeFuseInput_06 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Document,
            code: 6,
            template: "{0} func but not found \"EFUSE_BitDef_Table\"",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the Hard IP eFuse input definition against the eFuse specification. "
                    + "Check the input name, bit address, and allowed values.");

        /// <summary>Template: "IG-XL 9.0 need to keep the String length &lt; 8000 characters ;Arg:{0};Value:{1}"</summary>
        /// <remarks>{0} = limitation description, {1} = ?</remarks>
        public static readonly ErrorCode E_IgxlLimitation_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Igxl,
            code: 1,
            template: "IG-XL 9.0 need to keep the String length < 8000 characters ;Arg:{0};Value:{1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Review the flagged IGXL limitation and restructure the test to stay within the allowed bounds. "
                    + "Consult the IGXL platform documentation for the specific constraint.");

        /// <summary>Template: "Missing Dictionary key in calculation does not exist above test items : {0}"</summary>
        /// <remarks>{0} = parameter name</remarks>
        public static readonly ErrorCode E_MisCalculationParaDefine_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.StoreName,
            code: 1,
            template: "Missing Dictionary key in calculation does not exist above test items : {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Add the missing calculation parameter definition. "
                    + "Please make sure either pin or storename exist in calculation");

        /// <summary>Template: "The job of the {0} is not match to the Efuse_Bit_Def definition."</summary>
        /// <remarks>{0} = field</remarks>
        public static readonly ErrorCode E_MismatchFieldJob_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Job,
            code: 1,
            template: "The job of the {0} is not match to the Efuse_Bit_Def definition.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Compare the field job assignment against the expected job definition. "
                    + "Update the Hard IP sheet so the field job is consistent with the specification.");

        /// <summary>Template: "The bit width of the {0} is not match to the Efuse_Bit_Def definition."</summary>
        /// <remarks>{0} = field</remarks>
        public static readonly ErrorCode E_MismatchFieldWidth_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Bit,
            code: 1,
            template: "The bit width of the {0} is not match to the Efuse_Bit_Def definition.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Compare the field width in the Hard IP sheet against the EFuse bit definition. "
                    + "Correct the width so both definitions are consistent.");

        /// <summary>Template: "Measure type is different from limit unit"</summary>
        public static readonly ErrorCode E_MismatchLimitUnit_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Unit,
            code: 1,
            template: "Measure type is different from limit unit",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the unit specification in the limit definition. "
                    + "Ensure the unit matches the measurement type (e.g. V, A, Hz).");

        /// <summary>Template: "The field of the {0} is not real, please check with the Efuse_Bit_Def definition."</summary>
        /// <remarks>{0} = field</remarks>
        public static readonly ErrorCode E_MismatchRealValue_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Field,
            code: 1,
            template: "The field of the {0} is not real, please check with the Efuse_Bit_Def definition.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Compare the real value in the Hard IP sheet against the expected definition. "
                    + "Correct the value in the Hard IP sheet to match the specification.");

        /// <summary>Template: "Please Check the pattern \"{0}\" , the sequence count and call subrs cnt is not match."</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MismatchSequenceAndCallSubrsCnt_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 1,
            template: "Please Check the pattern \"{0}\" , the sequence count and call subrs cnt is not match.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check that the number of measurement sequence entries equals the number of call subroutine entries. "
                    + "Add or remove entries to restore the count balance.");

        /// <summary>Template: "Missing pattern for measuments"</summary>
        public static readonly ErrorCode E_MisPatternForMeasument_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 1,
            template: "Missing pattern for measuments",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check whether the pattern is configured for measurement operations. "
                    + "Use a pattern that includes measurement sequences for this test.");

        /// <summary>Template: "Missing bin number setting"</summary>
        public static readonly ErrorCode E_MissingBinNum_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Bin,
            code: 1,
            template: "Missing bin number setting",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the bin definition and assign a bin number for the specified test. "
                    + "Ensure the bin number is unique and within the valid bin range.");

        /// <summary>Template: "Missing Dictionary key in calculation does not exist above calculation funtion : {0}"</summary>
        /// <remarks>{0} = function name</remarks>
        public static readonly ErrorCode E_MissingCalcFunction_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.StoreName,
            code: 1,
            template: "Missing Dictionary key in calculation does not exist above calculation funtion : {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Add the missing calculation function definition to the Hard IP configuration. "
                    + "Verify the function name matches the reference in the sheet.");

        /// <summary>Template: "Please Check the pattern \"{0}\" have no call_subr but have meas seq info."</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingCallSubroutine_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 1,
            template: "Please Check the pattern \"{0}\" have no call_subr but have meas seq info.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Add the required call subroutine entry for the pattern in PatInfo. "
                    + "Ensure the subroutine name matches the one used in the pattern file.");

        /// <summary>Template: "Can not find {0} in Efuse_Bit_Def definition."</summary>
        /// <remarks>{0} = field name</remarks>
        public static readonly ErrorCode E_MissingFieldName_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Argument,
            code: 2,
            template: "Can not find {0} in Efuse_Bit_Def definition.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the Hard IP eFuse sheet and add the required field name. "
                    + "Verify the field name matches the EFuse bit definition.");

        /// <summary>Template: "Can not find HardIp category: {0} from TestSettting sheet"</summary>
        /// <remarks>{0} = category name</remarks>
        public static readonly ErrorCode E_MissingHardIpCategory_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.DcCategory,
            code: 1,
            template: "Can not find HardIp category: {0} from TestSettting sheet",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Add the missing Hard IP category definition. "
                    + "Verify the category name against the expected Hard IP configuration.");

        /// <summary>Template: "Missing HardIpDc pin in Pinmap"</summary>
        public static readonly ErrorCode E_MissingHardIpDcPin_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 1,
            template: "Missing HardIpDc pin in Pinmap",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Add the missing DC pin to the Hard IP pin list. "
                    + "Verify the pin name against the device pin map.");

        /// <summary>Template: "Here is one HardIP sheet :\"{0}\" but sheet name is not start with \"HardIP_\""</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingHardIpSheet_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 1,
            template: "Here is one HardIP sheet :\"{0}\" but sheet name is not start with \"HardIP_\"",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Add \"HardIP_\" prefix to the sheet to solve the problem");

        /// <summary>Template: "Missing header: {0} in sheet {1}"</summary>
        /// <remarks>{0} = header name, {1} = sheet name</remarks>
        public static readonly ErrorCode E_MissingHeader_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Header,
            code: 1,
            template: "Missing header: {0} in sheet {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the sheet and add the required header column. "
                    + "Compare the column headers against the specification to identify which ones are missing.");

        /// <summary>Template: "The field {0} is not defined in the prewrite HardIP testplan."</summary>
        /// <remarks>{0} = entry name</remarks>
        public static readonly ErrorCode E_MissingInPreWrite_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 2,
            template: "The field {0} is not defined in the prewrite HardIP testplan.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Add the missing entry to the pre-write definition in the Hard IP sheet. "
                    + "Ensure all required fields are present before the write sequence.");

        /// <summary>Template: "Please Check the pattern \"{0}\" have no meas seq info but have call_subr."</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingMeasureSequence_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 3,
            template: "Please Check the pattern \"{0}\" have no meas seq info but have call_subr.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the PatInfo sheet and add the required measurement sequence for the pattern. "
                    + "Verify the sequence is needed and add all required measurement entries.");

        /// <summary>Template: "Missing IO_PinMap sheet in TestPlan"</summary>
        public static readonly ErrorCode E_MissingNeededSheets_02 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Document,
            code: 7,
            template: "Missing {0} sheet in TestPlan",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Add the missing sheet to the workbook or check whether it was renamed. "
                    + "Compare the workbook contents against the expected sheet list.");

        /// <summary>Template: "Missing Parameter in {0} : {1} Or wrong key word in misc-info"</summary>
        /// <remarks>{0} = parameter name, {1} = ?</remarks>
        public static readonly ErrorCode E_MissingParameter_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Argument,
            code: 3,
            template: "Missing Parameter in {0} : {1} Or wrong key word in misc-info",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the Hard IP sheet and add the required parameter. "
                    + "Check the Hard IP specification to confirm which parameters are mandatory.");

        /// <summary>Template: "Missing Parameter in {0}({1}) : {2}"</summary>
        /// <remarks>{0} = ?, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_MissingParameter_03 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Argument,
            code: 7,
            template: "Missing Parameter in {0}({1}) : {2}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the Hard IP sheet and add the required parameter. "
                    + "Check the Hard IP specification to confirm which parameters are mandatory.");

        /// <summary>Template: "PatInfo pin '{0}' is not found in the pin map."</summary>
        /// <remarks>{0} = pin name</remarks>
        public static readonly ErrorCode E_MissingPatInfoPinInPinMap_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 2,
            template: "Missing MeasC pin in Pinmap:{0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the pin map and add the missing PatInfo pin entry. "
                    + "Ensure the pin name in PatInfo exactly matches the pin map definition.");

        /// <summary>Template: "Missing patten in PatInfo file: {0}"</summary>
        /// <remarks>{0} = pattern name</remarks>
        public static readonly ErrorCode E_MissingPatternInPatInfo_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.PatternInfo,
            code: 1,
            template: "Missing patten in PatInfo file: {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Add the missing pattern definition to the PatInfo sheet. "
                    + "Ensure all patterns used in the test plan have a corresponding PatInfo entry.");

        /// <summary>Template: "Missing pattern in \"Pattern\" Column"</summary>
        public static readonly ErrorCode E_MissingPatternInTestPlan_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 4,
            template: "Missing pattern in \"Pattern\" Column",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Add the missing pattern entry to the test plan. "
                    + "Verify the pattern file exists and is included in the pattern compilation step.");

        /// <summary>Template: "Missing Hi limit pin in Pinmap : {0}"</summary>
        /// <remarks>{0} = row number or context</remarks>
        public static readonly ErrorCode E_MissingPinName_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 3,
            template: "Missing Hi limit pin in Pinmap : {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the Hard IP sheet and fill in the pin name for the flagged row. "
                    + "All pin rows must have a non-empty pin name.");

        /// <summary>Template: "Missing Lo limit pin in Pinmap : {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingPinName_02 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 4,
            template: "Missing Lo limit pin in Pinmap : {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the Hard IP sheet and fill in the pin name for the flagged row. "
                    + "All pin rows must have a non-empty pin name.");

        /// <summary>Template: "Missing force pin in Pinmap : {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingPinName_03 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 5,
            template: "Missing force pin in Pinmap : {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the Hard IP sheet and fill in the pin name for the flagged row. "
                    + "All pin rows must have a non-empty pin name.");

        /// <summary>Template: "Missing sweep pin in Pinmap : {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingPinName_04 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 6,
            template: "Missing sweep pin in Pinmap : {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the Hard IP sheet and fill in the pin name for the flagged row. "
                    + "All pin rows must have a non-empty pin name.");

        /// <summary>Template: "Missing meas pins in Pinmap"</summary>
        public static readonly ErrorCode E_MissingPinName_05 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 7,
            template: "Missing meas pins in Pinmap : {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the Hard IP sheet and fill in the pin name for the flagged row. "
                    + "All pin rows must have a non-empty pin name.");

        /// <summary>Template: "Missing force pins of MeasPin in Pin map : {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingPinName_06 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 8,
            template: "Missing force pins of MeasPin in Pin map : {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the Hard IP sheet and fill in the pin name for the flagged row. "
                    + "All pin rows must have a non-empty pin name.");

        /// <summary>Template: "Empty Meas PinName but contains Limits"</summary>
        public static readonly ErrorCode E_MissingPinName_07 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 9,
            template: "Empty Meas PinName but contains Limits",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the Hard IP sheet and fill in the pin name for the flagged row. "
                    + "All pin rows must have a non-empty pin name.");

        /// <summary>Template: "Missing \"{0}\" in pin's misc info to ingore limit for trim instance."</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingPinName_08 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 10,
            template: "Missing \"Ignore_Flow_Limit\" in pin's misc info to ingore limit for trim instance : {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the Hard IP sheet and fill in the pin name for the flagged row. "
                    + "All pin rows must have a non-empty pin name.");

        /// <summary>Template: "Missing {0} sheet in TestPlan"</summary>
        /// <remarks>{0} = table name</remarks>
        public static readonly ErrorCode E_MissingSweepSingleloopTable_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Document,
            code: 8,
            template: "Missing sweep or single-loop table definition sheet {0} in TestPlan",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Add the missing sweep or single-loop table definition. "
                    + "Verify the table name against the expected Hard IP register configuration.");

        /// <summary>Template: "The function: {0} can not find in library!"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissVbtModule_02 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Module,
            code: 2,
            template: "The function: {0} can not find in library!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the VBT module name against the expected module list. "
                    + "Add the missing module definition or correct the reference.");

        /// <summary>Template: "Opposite limit value"</summary>
        public static readonly ErrorCode E_OppositeLimit_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Limit,
            code: 1,
            template: "Opposite limit value",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check that the lower limit is less than the upper limit. "
                    + "Swap the values if they were accidentally reversed.");

        /// <summary>Template: "pattern row exist measuments"</summary>
        public static readonly ErrorCode E_PatternExistMeasument_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Row,
            code: 1,
            template: "pattern row exist measuments",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the pattern list for an existing measurement pattern with the same name. "
                    + "Remove the duplicate reference or rename the new pattern.");

        /// <summary>Template: "Register Assign Type {0} is not support to split"</summary>
        /// <remarks>{0} = register or assignment</remarks>
        public static readonly ErrorCode E_RegAssignError_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Assignment,
            code: 1,
            template: "Register Assign Type {0} is not support to split",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the register assignment definition and correct the error. "
                    + "Verify the assignment value is within the allowed range and format.");

        /// <summary>Template: "The count of Register Assign {0} differs from measureSequence"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_RegAssignError_02 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Assignment,
            code: 2,
            template: "The count of Register Assign {0} differs from measureSequence",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the register assignment definition and correct the error. "
                    + "Verify the assignment value is within the allowed range and format.");

        /// <summary>Template: "Register Assign Type {0} is not support to split"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_RegAssignError_03 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Assignment,
            code: 3,
            template: "Register Assign Type {0} is not support to split",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the register assignment definition and correct the error. "
                    + "Verify the assignment value is within the allowed range and format.");

        /// <summary>Template: "Repeat SubBlock Name:{0}"</summary>
        /// <remarks>{0} = sub-block name</remarks>
        public static readonly ErrorCode E_RepeatSubBlock_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Duplicate,
            enumErrorTarget: EnumErrorTarget.Subblock,
            code: 1,
            template: "Repeat SubBlock Name:{0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the Hard IP sub-block list for duplicate entries. "
                    + "Remove the repeated sub-block reference or consolidate the duplicates.");

        /// <summary>Template: "SELSRAM digital source assignment '{0}' is not defined in the instance."</summary>
        /// <remarks>{0} = assignment name</remarks>
        public static readonly ErrorCode E_SelsramDigSrcAssignmentNotDefineInInstance_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Assignment,
            code: 1,
            template: "SELSRAM digital source assignment '{0}' is not defined in the instance.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Add the missing SELSRAM digital source assignment to the instance definition. "
                    + "Ensure the assignment matches the one referenced in the mapping table.");

        /// <summary>Template: "Block: {0} Pattern: {1} Alpha: {2} of DigSrc_Assignment is not defined in SELSRM_Mapping_Table."</summary>
        /// <remarks>{0} = assignment name, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_SelsramDigSrcAssignmentNotDefineInTable_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Assignment,
            code: 2,
            template: "Block: {0} Pattern: {1} Alpha: {2} of DigSrc_Assignment is not defined in {3}.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Add the missing SELSRAM digital source assignment to the mapping table. "
                    + "Verify the assignment name against the table definition.");

        /// <summary>Template: "Can't get SELSRAM_Mapping_Table in the test plan."</summary>
        public static readonly ErrorCode E_SelsramMappingTableError_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Document,
            code: 9,
            template: "Can't get SELSRAM_Mapping_Table in the test plan.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the SELSRAM mapping table for invalid or inconsistent entries. "
                    + "Verify the mapping conforms to the SELSRAM specification.");

        /// <summary>Template: "DigSrc_Assignment column didn't exist in SELSRAM_Mapping_Table."</summary>
        public static readonly ErrorCode E_SelsramMappingTableError_02 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Column,
            code: 2,
            template: "DigSrc_Assignment column didn't exist in SELSRAM_Mapping_Table.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the SELSRAM mapping table for invalid or inconsistent entries. "
                    + "Verify the mapping conforms to the SELSRAM specification.");

        /// <summary>Template: "SubBlock :{0}, not found in {1} sheet."</summary>
        /// <remarks>{0} = sub-block name, {1} = ?</remarks>
        public static readonly ErrorCode E_SubBlockNotFound_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Subblock,
            code: 1,
            template: "SubBlock :{0}, not found in {1} sheet.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify that the sub-block name is spelled correctly and defined in the Hard IP sheet. "
                    + "Add the missing sub-block or correct the reference.");

        /// <summary>Template: "Sub-block usage error for '{0}'"</summary>
        /// <remarks>{0} = sub-block name, {1} = details</remarks>
        public static readonly ErrorCode E_SubBlockUsage_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Subblock,
            code: 1,
            template: "Invalid SubBlock Usage:{0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the sub-block usage in the Hard IP configuration. "
                    + "Ensure the sub-block is referenced correctly and all required usage rules are satisfied.");

        /// <summary>Template: "DigSrc Signal Name \"{0}\" is different from VM_Vector \"{1} \" "</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_WrongDigSrcSignalName_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.PatternInfo,
            code: 1,
            template: "DigSrc Signal Name \"{0}\" is different from VM_Vector \"{1} \" ",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the digital source signal name against the defined signal list. "
                    + "Check for typos or renamed signals in the source definition.");

        /// <summary>Template: "Wrong Force Type for InterPose_PreMeas:{0}"</summary>
        /// <remarks>{0} = instance name</remarks>
        public static readonly ErrorCode E_WrongForceCondition_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Force,
            code: 1,
            template: "Wrong Force Type for InterPose_PreMeas:{0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the Hard IP force condition definition and verify the values. "
                    + "Ensure the condition is compatible with the instance type and test requirements.");

        /// <summary>Template: "Wrong Force Type for others force condition:{0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_WrongForceCondition_02 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Force,
            code: 2,
            template: "Wrong Force Type for others force condition:{0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the Hard IP force condition definition and verify the values. "
                    + "Ensure the condition is compatible with the instance type and test requirements.");

        /// <summary>Template: "{0} could not support"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_WrongForceCondition_03 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Force,
            code: 3,
            template: "{0} could not support",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the Hard IP force condition definition and verify the values. "
                    + "Ensure the condition is compatible with the instance type and test requirements.");

        /// <summary>Template: "Wrong ForceCondition for {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_WrongForceCondition_04 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Force,
            code: 4,
            template: "Wrong ForceCondition for {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the Hard IP force condition definition and verify the values. "
                    + "Ensure the condition is compatible with the instance type and test requirements.");

        /// <summary>Template: "Wrong pins of AC special setting: {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_WrongForceCondition_05 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 5,
            template: "Wrong pins of AC special setting: {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the Hard IP force condition definition and verify the values. "
                    + "Ensure the condition is compatible with the instance type and test requirements.");

        /// <summary>Template: "Wrong voltage format of AcSelector setting: {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_WrongForceCondition_06 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 6,
            template: "Wrong voltage format of AcSelector setting: {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the Hard IP force condition definition and verify the values. "
                    + "Ensure the condition is compatible with the instance type and test requirements.");

        /// <summary>Template: "Wrong type format of AcSelector setting: {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_WrongForceCondition_07 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 7,
            template: "Wrong type format of AcSelector setting: {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the Hard IP force condition definition and verify the values. "
                    + "Ensure the condition is compatible with the instance type and test requirements.");

        /// <summary>Template: "Wrong format of AcSelector setting: {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_WrongForceCondition_08 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 8,
            template: "Wrong format of AcSelector setting: {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the Hard IP force condition definition and verify the values. "
                    + "Ensure the condition is compatible with the instance type and test requirements.");

        /// <summary>Template: "Wrong ForceCondition for {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_WrongForceCondition_09 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Force,
            code: 9,
            template: "Wrong ForceCondition for {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the Hard IP force condition definition and verify the values. "
                    + "Ensure the condition is compatible with the instance type and test requirements.");

        /// <summary>Template: "Wrong ForceCondition for {0}"</summary>
        /// <remarks>This looks exactly the same as E_WrongForceCondition_09. 
        /// But 09 is for pin name, this is for pinType. So I still keep it
        /// </remarks>
        public static readonly ErrorCode E_WrongForceCondition_10 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Force,
            code: 10,
            template: "Wrong ForceCondition for {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the Hard IP force condition definition and verify the values. "
                    + "Ensure the condition is compatible with the instance type and test requirements.");

        /// <summary>Template: "Wrong format of force condition in Test Plan"</summary>
        public static readonly ErrorCode E_WrongForceCondition_11 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 11,
            template: "Wrong format of force condition in Test Plan",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the Hard IP force condition definition and verify the values. "
                    + "Ensure the condition is compatible with the instance type and test requirements.");

        /// <summary>Template: "Unrecognied limit value"</summary>
        public static readonly ErrorCode E_WrongLimitValue_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Limit,
            code: 2,
            template: "Unrecognied limit value",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the limit value against the device specification. "
                    + "Ensure the value does not exceed the allowed min/max bounds.");

        /// <summary>Template: "Can't get cap pin from pattern info, pattern : {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_WrongMeasC_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.PatternInfo,
            code: 2,
            template: "Can't get cap pin from pattern info, pattern : {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Compare the MeasC pin assignment with the cap pin defined in PatInfo. "
                    + "Update one of the definitions so both references point to the same physical pin.");

        /// <summary>Template: "Define MeasC, but not define number of capture bits"</summary>
        public static readonly ErrorCode E_WrongMeasC_02 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 5,
            template: "Define MeasC, but not define number of capture bits",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Compare the MeasC pin assignment with the cap pin defined in PatInfo. "
                    + "Update one of the definitions so both references point to the same physical pin.");

        /// <summary>Template: "MeasC capture bit mismatch between test plan {0} and patInfo {1}"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_WrongMeasC_03 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Measurement,
            code: 3,
            template: "MeasC capture bit mismatch between test plan {0} and patInfo {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Compare the MeasC pin assignment with the cap pin defined in PatInfo. "
                    + "Update one of the definitions so both references point to the same physical pin.");

        /// <summary>Template: "Wrong measure data in 'Meas' column"</summary>
        public static readonly ErrorCode E_WrongMeasContent_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Measurement,
            code: 1,
            template: "Wrong measure data in 'Meas' column",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the measurement content field in PatInfo for the specified entry. "
                    + "Ensure the content value is within the allowed set for this measurement type.");

        /// <summary>Template: "CusStr of\"{0}\" is not defined in this test item, and default calc type would be generated with C"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_WrongMeasContent_02 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Measurement,
            code: 2,
            template: "CusStr of\"{0}\" is not defined in this test item, and default calc type would be generated with C",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the measurement content field in PatInfo for the specified entry. "
                    + "Ensure the content value is within the allowed set for this measurement type.");

        /// <summary>Template: "Measurement pin '{0}' in PatInfo is invalid."</summary>
        /// <remarks>{0} = pin name</remarks>
        public static readonly ErrorCode E_WrongMeasPinInPatInfo_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 1,
            template: "MeasC pin:{0} is different from pat info cap pin :{1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the measurement pin name against the pin map and PatInfo definition. "
                    + "Ensure the pin is a valid measurement pin for this Hard IP type.");

        /// <summary>Template: "pattern info pin:{0}, can't match to test plan"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_WrongMeasPinInPatInfo_02 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 2,
            template: "pattern info pin:{0}, can't match to test plan",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the measurement pin name against the pin map and PatInfo definition. "
                    + "Ensure the pin is a valid measurement pin for this Hard IP type.");

        /// <summary>Template: "WrongMeasType of '{0}'"</summary>
        /// <remarks>{0} = meas type</remarks>
        public static readonly ErrorCode E_WrongMeasType_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Measurement,
            code: 3,
            template: "WrongMeasType of '{0}'",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the measurement type assignment in the PatInfo sheet. "
                    + "Ensure the type is supported for the specified pin and Hard IP category.");

        /// <summary>Template: "One of RegisterAssignment or digsrc of pattern does not have information, please check"</summary>
        public static readonly ErrorCode E_WrongRegisterAssignment_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Assignment,
            code: 1,
            template: "One of RegisterAssignment or digsrc of pattern does not have information, please check",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the register assignment definition and verify the value against the register specification. "
                    + "Ensure the assigned value is within the bit-width and allowed range.");

        /// <summary>Template: "Send bit: \"{0}\", Mismatch bit width between pattern send bit: {1} and Register Assignment in test plan: {2} ."</summary>
        /// <remarks>{0} = ?, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_WrongRegisterAssignment_02 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Bit,
            code: 2,
            template: "Send bit: \"{0}\", Mismatch bit width between pattern send bit: {1} and Register Assignment in test plan: {2} .",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the register assignment definition and verify the value against the register specification. "
                    + "Ensure the assigned value is within the bit-width and allowed range.");

        /// <summary>Template: "\"{0}\" is not defined before use it in register assignment"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_WrongRegisterAssignment_04 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.StoreName,
            code: 4,
            template: "\"{0}\" is not defined before use it in register assignment",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the register assignment definition and verify the value against the register specification. "
                    + "Ensure the assigned value is within the bit-width and allowed range.");

        /// <summary>Template: "Can't get assignment: {0} from pattern info: {1}"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_WrongRegisterAssignment_05 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Assignment,
            code: 5,
            template: "{0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the register assignment definition and verify the value against the register specification. "
                    + "Ensure the assigned value is within the bit-width and allowed range.");

        /// <summary>Template: "There is a Sweep format issue : {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_WrongRegisterAssignment_06 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Assignment,
            code: 6,
            template: "There is a Sweep format issue : {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the register assignment definition and verify the value against the register specification. "
                    + "Ensure the assigned value is within the bit-width and allowed range.");

        /// <summary>Template: "Can't get assignment: {0} from pattern info: {1}"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_WrongRegisterAssignment_07 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Assignment,
            code: 7,
            template: "Burst register assignment length {0} is different from burst pattern count {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the register assignment definition and verify the value against the register specification. "
                    + "Ensure the assigned value is within the bit-width and allowed range.");

        /// <summary>Template: "Can't get assignment: {0} from pattern info: {1}"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_WrongRegisterAssignment_08 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.PatternInfo,
            code: 8,
            template: "Can't get assignment: {0} from pattern info: {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the register assignment definition and verify the value against the register specification. "
                    + "Ensure the assigned value is within the bit-width and allowed range.");

        /// <summary>Template: "Can't get assignment: {0} from pattern info: {1}"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_WrongRegisterAssignment_09 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Assignment,
            code: 9,
            template: "Lose pattern info assignment in plan: {0} from pattern info: {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the register assignment definition and verify the value against the register specification. "
                    + "Ensure the assigned value is within the bit-width and allowed range.");

        /// <summary>Template: "# of registers between patinfo {0} and test plan {1} not match"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_WrongRegisterAssignment_10 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Assignment,
            code: 10,
            template: "# of registers between patinfo {0} and test plan {1} not match",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the register assignment definition and verify the value against the register specification. "
                    + "Ensure the assigned value is within the bit-width and allowed range.");

        /// <summary>Template: "patinfo contains registers: {0}, not exist in test plan"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_WrongRegisterAssignment_11 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Assignment,
            code: 11,
            template: "patinfo contains registers: {0}, not exist in test plan",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the register assignment definition and verify the value against the register specification. "
                    + "Ensure the assigned value is within the bit-width and allowed range.");

        /// <summary>Template: "test plan contains registers: {0}, not exist in patinfo"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_WrongRegisterAssignment_12 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Assignment,
            code: 12,
            template: "test plan contains registers: {0}, not exist in patinfo",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the register assignment definition and verify the value against the register specification. "
                    + "Ensure the assigned value is within the bit-width and allowed range.");

        /// <summary>Template: "Same name in \"Send Bit Name\" but with different num of bit in \"Send Bit Str\" -- {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_WrongSendInformation_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Bit,
            code: 3,
            template: "Same name in \"Send Bit Name\" but with different num of bit in \"Send Bit Str\" -- {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the send information fields in PatInfo for the specified pattern. "
                    + "Verify all values conform to the expected format and bit-width constraints.");

        /// <summary>Template: "The step for sweep voltage is incorrect"</summary>
        public static readonly ErrorCode E_WrongSweepStep_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Sweep,
            code: 1,
            template: "The step for sweep voltage is incorrect due to {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the sweep step value against the specification. "
                    + "Ensure the step size is positive, within range, and compatible with the sweep range.");

        /// <summary>Template: "The Payload of multiple init : "</summary>
        public static readonly ErrorCode E_WrongTimeSet_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Tset,
            code: 1,
            template: "The Payload of multiple init : {0} is not the same timeSet with pattern : {1}, please check",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the time set name against the defined time set list. "
                    + "Ensure the time set is compatible with the Hard IP instance requirements.");

        /// <summary>Template: "We found duplicate Store name \"{0}\" in HardIP single pattern."</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode I_DuplicateStoreName_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Duplicate,
            enumErrorTarget: EnumErrorTarget.StoreName,
            code: 1,
            template: "We found duplicate Store name \"{0}\" in HardIP single pattern.",
            enumErrorLevel: EnumErrorLevel.Info,
            guidance: "Check the Hard IP store name list for duplicate entries. "
                    + "Rename one of the duplicate store names so each name is unique.");

        /// <summary>Template: "We have found an invalid Opcode:{0}."</summary>
        /// <remarks>{0} = opcode or pattern name</remarks>
        public static readonly ErrorCode I_UnsupportedOpcode_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Opcode,
            code: 1,
            template: "We have found an invalid Opcode:{0}.",
            enumErrorLevel: EnumErrorLevel.Info,
            guidance: "Check the IGXL version for supported opcodes. "
                    + "Replace the unsupported opcode with a supported equivalent.");

        /// <summary>Template: "Burst pattern judged by {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode W_BurstPatJudgeFlag_02 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 2,
            template: "{0}: Burst pattern judged by {1}",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Verify the burst pattern judge flag definition. "
                    + "Ensure the flag name and condition are valid and match the burst pattern flow.");

        /// <summary>Template: "Limit clear by tool to let VBT apply limit directly"</summary>
        public static readonly ErrorCode W_ClearLimit_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Limit,
            code: 1,
            template: "Limit clear by tool to let VBT apply limit directly",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Verify whether the limit clearing is intentional. "
                    + "If not, restore the limit definition from the specification.");

        /// <summary>Template: "We found duplicate instance but we add index to avoid validate fail, please check"</summary>
        public static readonly ErrorCode W_DuplicateInstance_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Duplicate,
            enumErrorTarget: EnumErrorTarget.Instance,
            code: 1,
            template: "We found duplicate instance but we add index to avoid validate fail, please check",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Check the instance list for duplicate entries. "
                    + "Remove or rename the duplicate so each instance name is unique.");

        /// <summary>Template: "SubBlock:{0}, existed on other sheet"</summary>
        /// <remarks>{0} = sub-block name</remarks>
        public static readonly ErrorCode W_ExistedSubBlock_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Duplicate,
            enumErrorTarget: EnumErrorTarget.Subblock,
            code: 1,
            template: "SubBlock:{0}, existed on other sheet",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Check whether the sub-block has already been defined. "
                    + "Remove the duplicate definition or update the existing one.");

        /// <summary>Template: "Judged by {0}"</summary>
        /// <remarks>{0} = flag name</remarks>
        public static readonly ErrorCode W_FuseWriteJudgeFlag_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Assignment,
            code: 2,
            template: "Need to perform eFuse pre-write, but flag is missing in MiscInfo, Autogen judged that by the latest fail flag: {0}",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Verify the fuse write judge flag definition. "
                    + "Ensure the flag name and condition are valid and consistent with the fuse write flow.");

        /// <summary>Template: "Manual items that Autogen cannot support yet!"</summary>
        public static readonly ErrorCode W_ManualItems_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 1,
            template: "Manual items that Autogen cannot support yet!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Review the flagged manual item and take the described action. "
                    + "Manual items are reminders for configuration steps that cannot be automated.");

        /// <summary>Template: "Measure pin had no limit"</summary>
        public static readonly ErrorCode W_MissingLimit_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Limit,
            code: 1,
            template: "Measure pin had no limit",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Add the required limit definition for the specified test. "
                    + "Check whether the test requires upper, lower, or both limits.");

        /// <summary>Template: "Measure pin had no limit unit"</summary>
        public static readonly ErrorCode W_MissingLimitUnit_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Unit,
            code: 1,
            template: "Measure pin had no limit unit",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the limit definition and add the required unit. "
                    + "All limits must specify a unit consistent with the measurement type.");

        /// <summary>Template: "No measure sequence in PatInfo file, but specified measure pins in testPlan"</summary>
        public static readonly ErrorCode W_MissingPinSeq_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Measurement,
            code: 1,
            template: "No measure sequence in PatInfo file, but specified measure pins in testPlan",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Add the missing pin sequence definition. "
                    + "Ensure the sequence covers all required pins in the correct order.");

        /// <summary>Template: "{0} is real value in Efuse_Bit_Def but can not find in eFuse_HardIP_Table."</summary>
        /// <remarks>{0} = field name</remarks>
        public static readonly ErrorCode W_MissingRealFieldNameInEfuseHardIP_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Field,
            code: 1,
            template: "{0} is real value in Efuse_Bit_Def but can not find in eFuse_HardIP_Table.",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Add the missing real field name to the eFuse Hard IP sheet. "
                    + "Verify the field name matches exactly the name in the EFuse bit definition.");

        /// <summary>Template: "this pin :{0} can not find in pattern info"</summary>
        /// <remarks>{0} = pin name</remarks>
        public static readonly ErrorCode W_MissingTestplanPin_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 3,
            template: "this pin :{0} can not find in pattern info {1}",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Please compare the pin you assigned with pins in PatternInfo.");

        /// <summary>Template: "The SubBlock:{0}, not be referenced by burst item"</summary>
        /// <remarks>{0} = sub-block or instance name</remarks>
        public static readonly ErrorCode W_NoBurstSubBlock_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Subblock,
            code: 1,
            template: "The SubBlock:{0}, not be referenced by burst item",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Add a burst sub-block definition for the specified instance. "
                    + "Check whether the instance requires a burst configuration and add it accordingly.");

        /// <summary>Template: "The legnth of {0} content is over 6000, move to reg assign table"</summary>
        /// <remarks>{0} = register or assignment</remarks>
        public static readonly ErrorCode W_RegAssignWarning_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 3,
            template: "The legnth of {0} content is over 6000, move to reg assign table",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Review the register assignment warning. "
                    + "Confirm the assignment is intentional; it may produce unexpected device behavior.");

        /// <summary>Template: "RelayOn/Off would not restore in PrePat, suggest fill in Misc info with 'RelayOn:{0}' or 'RelayOff:{0}'"</summary>
        /// <remarks>{0} = relay or context description</remarks>
        public static readonly ErrorCode W_RelayRestore_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 4,
            template: "RelayOn/Off would not restore in PrePat, suggest fill in Misc info with 'RelayOn:{0}' or 'RelayOff:{0}'",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "If use [PinName]:RelayOn in force condition, highly recommended to use RelayOff:[PinName] " +
            "in MiscInfo to restore relay state. Vise versa");

        /// <summary>Template: "Invalid SubBlock Name:{0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode W_RepeatSubBlock_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Subblock,
            code: 1,
            template: "Invalid SubBlock Name:{0}",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Check the Hard IP sub-block list for duplicate entries. "
                    + "Remove the repeated sub-block reference or consolidate the duplicates.");

        /// <summary>Template: "Unrecognied header : {0}"</summary>
        /// <remarks>{0} = header name</remarks>
        public static readonly ErrorCode W_UnrecognisedHeader_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Header,
            code: 1,
            template: "Unrecognied header : {0}",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Remove or rename the unrecognised column header. "
                    + "Check the specification for the list of accepted header names.");

        /// <summary>Template: "The capture store name : {0} is repeated with {1} times in this item! "</summary>
        /// <remarks>{0} = MeasC pin, {1} = PatInfo cap pin</remarks>
        public static readonly ErrorCode W_WrongMeasC_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Duplicate,
            enumErrorTarget: EnumErrorTarget.StoreName,
            code: 1,
            template: "The capture store name : {0} is repeated with {1} times in this item! ",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Compare the MeasC pin assignment with the cap pin defined in PatInfo. "
                    + "Update one of the definitions so both references point to the same physical pin.");

        /// <summary>Template: "Patterns that not start with “dd_”, ”cz_” or “pp_” will be ignored."</summary>
        public static readonly ErrorCode W_WrongPatternName_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 5,
            template: "Patterns that not start with “dd_”, ”cz_” or “pp_” will be ignored.",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Please make sure again. Do you really want to use pattern naming that will be ignored by default?");

        /// <summary>Template: "start:{0} is larger than stop:{1}"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode W_WrongRegisterAssignment_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Assignment,
            code: 3,
            template: "start:{0} is larger than stop:{1}",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "RegisterAssign syntax <KEY>:<START>:<END> or <KEY>[<START>:<END>]. "
                    + "<START> must not larger than <END>");

        /// <summary>Template: "Total MeasCount different, TestPlan:{0}, PatInfo:{1}"</summary>
        /// <remarks>{0} = expected count, {1} = actual count</remarks>
        public static readonly ErrorCode W_WrongTotalMeasCount_01 = new(
            enumErrorCategory: EnumErrorCategory.HardIp,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 4,
            template: "Total MeasCount different, TestPlan:{0}, PatInfo:{1}",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Compare the expected and actual measurement counts in the PatInfo sheet. "
                    + "Add or remove measurement entries to match the expected count.");
    }
}
