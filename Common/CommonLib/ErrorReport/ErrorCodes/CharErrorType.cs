using CommonLib.ErrorReport.Base;

namespace CommonLib.ErrorReport.ErrorCodes
{
    public static class CharErrorType
    {
        ///
        ///
        public static readonly ErrorCode E_IllegalChar_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 1,
            template: "There is illegal char in {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Remove all the underline for each USERDEF.");

        ///
        ///
        public static readonly ErrorCode E_DuplicateTpName_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Duplicate,
            enumErrorTarget: EnumErrorTarget.Instance,
            code: 1,
            template: "Duplicate TP Name",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please remove one of duplicate case from char plan. "
                    + "Or re-name to another TP name (if the pattern set is different).");

        ///
        ///
        public static readonly ErrorCode E_DuplicatePattern_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Duplicate,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 1,
            template: "duplicate patterns",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "There is duplicate pattern in a char item. "
                    + "Please check the correct pattern is fill in char plan");

        ///
        ///
        public static readonly ErrorCode E_ErrorShmooRange_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Range,
            code: 1,
            template: "Step is not empty besides primary shmoo",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Only the primary shmoo can define a Step value. The Step field of all other shmoo settings must be empty.");

        ///
        ///
        public static readonly ErrorCode E_ErrorShmooRange_02 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Range,
            code: 2,
            template: "NumStepNotMatch",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Shmoo name not match, need to check.");

        ///
        ///
        public static readonly ErrorCode E_ErrorMethod_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Sweep,
            code: 1,
            template: "Error Search Method",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Pre check search method should be liner for sweep on freq char row.");

        ///
        ///
        public static readonly ErrorCode E_OppositeUsl_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Limit,
            code: 1,
            template: "Opposite USL/LSL",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "USL must be greater than LSL.");

        ///
        ///
        public static readonly ErrorCode E_MissingPinName_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 1,
            template: "Missing (pin name)/(pin group) in PinMap.txt/pinList sheet",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check if pin name exists in the PinList from PinMap.");

        ///
        ///
        public static readonly ErrorCode E_MissingPinName_02 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 2,
            template: "Missing pin in pinmap file in {0}, Row {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Pin in char plan but not in pin map.");

        ///
        ///
        public static readonly ErrorCode E_MissingPinName_03 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 3,
            template: "Missing VDD name in PinMap.txt",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Pin in char plan but not in pin map.");

        ///
        ///
        public static readonly ErrorCode W_WrongGroup_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.PerformanceMode,
            code: 1,
            template: "Wrong group for {0}",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "P-mode and Category mismatch. "
                    + "The p-mode and category should be the same block. "
                    + "You can ignore this error, if you want to use CPU P-mode to test SOC pattern.");

        ///
        ///
        public static readonly ErrorCode E_MissingHeader_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Header,
            code: 1,
            template: "Missing Header: {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Wrong header in input sheet, please check!");

        ///
        ///
        public static readonly ErrorCode W_WrongRetention_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 1,
            template: "Retention time format is wrong, the commaon count must be  in [0, 4, 14]",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "If you want the wait time after payload1. "
                    + "You should fill the comma in the corresponding position.");

        ///
        ///
        public static readonly ErrorCode W_WrongRetention_02 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 1,
            template: "Missing Retention time or PowerRunScenario for Category 'EXTRET' or 'NAPRET'",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "For category \"EXTRET\", must specifiy retention and power-run-scenario.");

        ///
        ///
        public static readonly ErrorCode W_WrongRetention_03 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 2,
            template: "Missing PowerRunScenario for Category 'INTRET'",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "For category \"INTRET\", must specify power-run-scenario.");

        ///
        ///
        public static readonly ErrorCode W_WrongRetention_04 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 3,
            template: "Missing PowerRunScenario for Category 'DISTURB'",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "For category \"DISTURB\", must specify power-run-scenario.");

        ///
        ///
        public static readonly ErrorCode W_WrongRetention_05 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 1,
            template: "Wait time exists but Category is not 'EXTRET'",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "If specified retention, the category must be EXTRET.");

        ///
        ///
        public static readonly ErrorCode E_WrongRetention_06 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 1,
            template: "No matching write pattern for read pattern: {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "No matching write pattern for read pattern.");

        ///
        ///
        public static readonly ErrorCode E_WrongRetention_07 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 2,
            template: "Unmatched write pattern {0} no found with read pattern",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "No matching write pattern for read pattern.");

        ///
        ///
        public static readonly ErrorCode W_WrongRetention_08 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 2,
            template: "Retention time \" {0} \" is not same to tool expect value \" {1} \"",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Retention time is not same to tool expect value.");

        ///
        ///
        public static readonly ErrorCode W_WrongRetention_09 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 3,
            template: "Retention time due to exist extra inits(10) or payloads(5)",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Retention time due to exist extra inits(10) or payloads(5).");

        ///
        ///
        public static readonly ErrorCode E_MissingShmooValue_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 4,
            template: "{0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Missing start or stop value.");

        ///
        ///
        public static readonly ErrorCode E_MissingShmooCondition_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 5,
            template: "Missing shmoo condition in {0}, Row: {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Shmoo Condition must to define both start and stop value.");

        ///
        ///
        public static readonly ErrorCode E_MissingShmooCondition_02 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 6,
            template: "Missing VIH/VIL shmoo for HIO in {0}, Row: {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "For HIO, should have some pin sweep setting.");

        ///
        ///
        public static readonly ErrorCode W_DummyShmooCondition_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 2,
            template: "Dummy shmoo condition for HAC in {0}, Row: {1}",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "For HAC, start and stop values of power/IO pins should be the same under normal conditions.");

        ///
        ///
        public static readonly ErrorCode E_WrongTpNameEquation_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Equation,
            code: 1,
            template: "Wrong equation of TP Name",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please check the correct TP name in char plan should be equation base, "
                    + "and the any space in userdef1-userdef9, group, category should be removed.");

        ///
        ///
        public static readonly ErrorCode E_WrongUserdef1_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 3,
            template: "Wrong USERDEF1 in {0} sheets in {1}, Row: {2}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "The value specified in UserDef1 must be defined in the corresponding sheet configuration. "
                    + "(HAC|HFL|HFH|HFLH|HIO) or (DFTL|DFTH|DFTLH|MCL|MCH|MCLH)");

        ///
        ///
        public static readonly ErrorCode E_WrongMeasOfHac_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Measurement,
            code: 1,
            template: "Wrong MeasType for {0} in {1}, Row: {2}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Should have correct meas type in user_def2 column.");

        ///
        ///
        public static readonly ErrorCode E_WrongUserdef3OfHac_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 2,
            template: "Wrong USERDEF3 for HAC in {0}, Row: {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Should have correct format in USERDEF3 column.");

        ///
        ///
        public static readonly ErrorCode E_WrongMeasCount_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Measurement,
            code: 2,
            template: "{0} for HAC in {1}, Row: {2}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Wrong MeasCount.");

        ///
        ///
        public static readonly ErrorCode E_WrongMeasCount_02 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Measurement,
            code: 3,
            template: "{0} total MeasCount in Char_Plan for HAC in {1} than in HardIP_info",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Measure counts are not equal.");

        ///
        ///
        public static readonly ErrorCode E_WrongMeasPin_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Measurement,
            code: 4,
            template: "{0} for HAC in {1}, Row: {2}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Wrong MeasPin.");

        ///
        ///
        public static readonly ErrorCode E_MissingPinSeq_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 4,
            template: "Missing pin seq or specified of pattern in PatInfo file",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Measure count are different and pat info meas count == 0.");

        ///
        ///
        public static readonly ErrorCode E_MissingPatternInPatInfo_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.PatternInfo,
            code: 1,
            template: "Missing patten in PatInfo file in {0}, Row: {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Cannot find patternName in PatInfoDict.");

        ///
        ///
        public static readonly ErrorCode E_WrongUserdef1OfVih_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 4,
            template: "TpName Contains VIH/VIL but USERDEF1 is not HFL|HFH|HIO in {0}, Row: {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "For item in hard ip sheets and tp name contains vih | vil, their user def 1 should be HFH | HFL | HIO.");

        ///
        ///
        public static readonly ErrorCode E_WrongShmooNv_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 5,
            template: "NV GlobalSpecs not within Start/Stop",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "The NV value defined in GlobalSpecs must be within the Start and Stop range of the shmoo setting.");

        ///
        ///
        public static readonly ErrorCode E_WrongShmooNv_02 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 6,
            template: "VDD power setting over than 1.7 of  nominal voltage ",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "The Start and Stop values of a VDD power shmoo must not exceed 1.7 times the nominal voltage (NV) defined in GlobalSpecs.");

        ///
        ///
        public static readonly ErrorCode E_MissingPinInGlobalSpecs_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 5,
            template: "Shmoo pin does not exist in GlobalSpecs",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Shmoo pin does not exist in GlobalSpecs.");

        ///
        ///
        public static readonly ErrorCode E_NoOtherSupplies_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 3,
            template: "missing or wrong other supplies in {0}, Row: {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please must define the other supply power condition. " +
                    "The other supply power should be NV or HV or LV.");

        ///
        ///
        public static readonly ErrorCode E_WrongVddInPinColumn_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 1,
            template: "Error VDD pins in Pin column in {0}, Row {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "The VDD pin specified in the Pin column is invalid. Please verify that the pin name and configuration are defined correctly.");

        ///
        ///
        public static readonly ErrorCode E_PatternOtherThanInPayload1_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 1,
            template: "There is pattern other than in payload1 in {0}, With: {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Only patterns defined in Payload1 are supported. Please remove any patterns assigned to other payloads.");

        ///
        ///
        public static readonly ErrorCode E_PatternsWithoutPayload_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 1,
            template: "There is no payload in these patterns: {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "For non-Rtos cmd char row, there should be at least one payload in the patterns");

        ///
        ///
        public static readonly ErrorCode W_MissingHlnCondition_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Setting,
            code: 1,
            template: "Missing HLN condition in {0}",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Each H, L, N, and N1 condition must have a corresponding counter row. " +
                    "Please verify that all required HLN conditions are defined correctly.");

        ///
        ///
        public static readonly ErrorCode W_TooManyInstanceName_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Instance,
            code: 1,
            template: "Too many instance names",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "The number of instance names associated with the same Payload1 must not exceed 200.");

        /// 1
        ///
        public static readonly ErrorCode E_MixedSIandDm_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 3,
            template: "Init/Payload patterns is mixed with SI/DM in {0}, Row {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "SI/DM patterns must not be mixed with init or payload patterns. " +
                    "Please separate SI/DM patterns from the init/payload configuration.");

        ///
        ///
        public static readonly ErrorCode E_PinGroupNotMatch_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.PinGroup,
            code: 1,
            template: "Pin group in PinList sheet not match pinmap file",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "The pin group definition in the PinList sheet must match the corresponding pin group definition in the pinmap file. " +
                    "Please verify that both the pin count and pin members are consistent.");

        ///
        ///
        public static readonly ErrorCode E_WrongUsllslRange_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Range,
            code: 1,
            template: "{0} value is invalid or Wrong {0} for shmoo range in \" {1} \"",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Invalid USL/LSL: the value is not within the shmoo range");

        ///
        ///
        public static readonly ErrorCode E_MissingForceCondition_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 6,
            template: "USERDEF6 contains Vih|Vil|Vicm|Vid, But do not has pin sweep {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "For USERDEF6 contains Vih|Vil|Vicm|Vid, " +
                    "should also specified the pin sweep with the corresponding type.");

        ///
        ///
        public static readonly ErrorCode E_WrongCapture_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Count,
            code: 1,
            template: "Wrong Capture count in USERDEF6",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Need to check plan measC count matches with pattern info.");

        ///
        ///
        public static readonly ErrorCode E_WrongCapture_02 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Sequence,
            code: 1,
            template: "Wrong Capture sequence in USERDEF6 for MeasC",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Need to check the measC sequence matches with pattern info.");

        ///
        ///
        public static readonly ErrorCode E_Userdef3MismatchShmoo_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 4,
            template: "There is an issue with USERDEF3 related Vmain or Valt pin shmoo setting",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "PE will base on the TP name to analysis the char data, " +
                    "so need to check if the shmoo setting is sweep.");

        ///
        ///
        public static readonly ErrorCode E_Userdef3MismatchShmoo_02 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 5,
            template: "USERDEF3 power does not match with shmoo setting",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "PE will base on the TP name to analysis the char data, " +
                    "so the information of shmoo sweep pins and userdef3 must be the same.");

        ///
        ///
        public static readonly ErrorCode E_WrongShmooSteps_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 7,
            template: "Start didn't equal Stop but step is 0",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "The shmoo steps from this char looks not reasonable. " +
                    "Start and Stop must be the same when Step is 0.");

        ///
        ///
        public static readonly ErrorCode W_WrongShmooSteps_02 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 8,
            template: "Shmoo step > 0.005",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "The shmoo steps from this char looks not reasonable, " +
                    "Step value must be less than 0.005.");

        ///
        ///
        public static readonly ErrorCode W_WrongShmooSteps_03 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 9,
            template: "Shmoo point is not integer",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "There is a shmoo point which is not integer in a shmoo setting.");

        ///
        ///
        public static readonly ErrorCode W_HeatAlarm_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 10,
            template: "Stop voltage higher than start voltage",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Stop voltage higher than start voltage. It might cause heat alarm, please check.");

        ///
        ///
        public static readonly ErrorCode E_LessThan6ShmooPoints_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 11,
            template: "Less than 6sShmoo points",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "A shmoo sweep must contain at least 6 points. " +
                    "Please adjust the Start, Stop, or Step value to increase the number of shmoo points.");

        ///
        ///
        public static readonly ErrorCode E_WrongNetName_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Header,
            code: 1,
            template: "Header power name \" {0} \" is not a Net Name",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "VDD power on the header / USERDEF3 is not net name.");

        ///
        ///
        public static readonly ErrorCode E_WrongLimitFormat_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 4,
            template: "Wrong format of USL/LSL",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Invalid USL/LSL format. The value must be a valid double number.");

        ///
        ///
        public static readonly ErrorCode E_WrongDigSrc_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Assignment,
            code: 1,
            template: "Missing digsrc assignment in Userdef6/7/8",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "There is one pattern specified with ':DigSrc' but no digsrc assignment in Userdef6/7/8.");

        ///
        ///
        public static readonly ErrorCode E_WrongDigSrc_02 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Assignment,
            code: 2,
            template: "Missing digsrc assignment in Userdef6/7",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "There are two patterns specified with ':DigSrc' but no digsrc assignment in Userdef6/7.");

        ///
        ///
        public static readonly ErrorCode E_WrongDigSrc_03 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Assignment,
            code: 1,
            template: "There are more than two patterns specified with ':DigSrc'",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "More than two patterns are specified with ':DigSrc'. Please limit the configuration to two patterns.");

        ///
        ///
        public static readonly ErrorCode E_WrongDigSrc_04 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Assignment,
            code: 3,
            template: "Missing digsrc assignment in Userdef6",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "There are two patterns specified with ':DigSrc' but no digsrc assignment in Userdef6.");

        ///
        ///
        public static readonly ErrorCode E_WrongDigSrc_05 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Assignment,
            code: 4,
            template: "Missing digsrc assignment in Userdef7",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "There are two patterns specified with ':DigSrc' but no digsrc assignment in Userdef7.");

        ///
        ///
        public static readonly ErrorCode E_WrongDigSrc_06 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Assignment,
            code: 2,
            template: "There are more than two MTD patterns specified with ':DigSrc'",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "More than two MTD patterns are specified with ':DigSrc'. Please limit the configuration to two patterns.");

        ///
        ///
        public static readonly ErrorCode E_WrongDigSrc_07 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 12,
            template: "S and binary can't be mixed in \"{0}\"}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please remove the mixed usage of 'S' and binary digits in the same segment.");

        ///
        ///
        public static readonly ErrorCode E_WrongDigSrc_08 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 7,
            template: "\"{0}\" doesn't exist in Send Bit Str of pattern info",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "All DigSrc segments must have a corresponding definition in the Send Bit Str of pattern info. " +
                    "Please make sure the segment name is defined correctly.");

        ///
        ///
        public static readonly ErrorCode E_WrongDigSrc_09 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 13,
            template: "\"{0}\" has S that exceed the length of \" {1} \"",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "The total number of 'S' bits assigned to DigSrc segments cannot exceed the available S length defined in UserDef9. " +
                    "Please reduce the number of 'S' bits or update UserDef9 accordingly.");

        ///
        ///
        public static readonly ErrorCode E_WrongDigSrc_10 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 14,
            template: "S Length of {0} ({1}) isn't the same as segment length in pattern info ({2})",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "The number of 'S' bits defined in the DigSrc segment must match the segment length specified in the pattern information. " +
                    "Please ensure both definitions use the same length.");

        ///
        ///
        public static readonly ErrorCode E_WrongDigSrc_11 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 15,
            template: "{0} over the S length of {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "The combined length of all DigSrc segments using 'S' exceeds the S length defined in UserDef9. " +
                    "Please adjust the segment assignments or update UserDef9 to provide sufficient S bits.");

        ///
        ///
        public static readonly ErrorCode E_WrongDigSrc_12 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Assignment,
            code: 3,
            template: "{0}: segments for S are not continuous",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "All DigSrc segments assigned with 'S' must be placed in consecutive order. " +
                    "Please ensure the S-related segments are continuous without interruption from other segment assignments.");

        ///
        ///
        public static readonly ErrorCode E_WrongDigSrc_13 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 5,
            template: "Data in {0} must be 0, 1, S",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "For segments with the 'F' format, only '0', '1', and 'S' are allowed. " +
                    "Please remove any unsupported characters from the segment definition.");

        ///
        ///
        public static readonly ErrorCode E_WrongDigSrc_14 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 6,
            template: "Data in {0} must be 0-9 and A-F",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "For segments with the 'G' format, only hexadecimal characters (0-9 and A-F) are allowed. " +
                    "Please update the segment definition to use valid hexadecimal values.");

        ///
        ///
        public static readonly ErrorCode E_WrongDigSrc_15 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 7,
            template: "\" {0} \" didn't be recognized",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "The segment format is not recognized. " +
                    "Please use a supported segment type and ensure the segment definition follows the expected syntax.");

        ///
        ///
        public static readonly ErrorCode W_WrongDigSrc_16 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 8,
            template: "Sgmt default value is not found, use '0' as sgmt default value",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Sgmt default value is not found, use '0' as sgmt default value. " +
                    "If wish to define the sgmt default value, just add 'sgmt_default=0', or 'sgmt_default=1' in the Revision sheet.");

        ///
        ///
        public static readonly ErrorCode W_MissingRevisonSheet_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.File,
            code: 1,
            template: "Missing revision sheet in CharPlan",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Need to add Revision sheet in CharPlan to specify default digsrc value.");

        ///
        ///
        public static readonly ErrorCode E_MisMatchDigSrc_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 6,
            template: "Pattern ends with DigSrc without SendBitStr in pat info",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "pattern:DigSrc need to have sendBitStr in pattern info.");

        ///
        ///
        public static readonly ErrorCode E_TestNameOverLength_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 9,
            template: "Test Name length over 255 bytes",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Test Name length need less than 255 bytes, please check it.");

        ///
        ///
        public static readonly ErrorCode E_TestNameNotEndswithSingleUnderline_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 10,
            template: "Test Name not ends with single underline",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please check userdef9 and TP name cell if existed only one underline at the last.");

        ///
        ///
        public static readonly ErrorCode W_EmptySheet_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.File,
            code: 1,
            template: "sheet {0} is empty, all items are unused!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Sheet is empty, all items are unused!");

        ///
        ///
        public static readonly ErrorCode W_EmptyUse_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 11,
            template: "Use/Not Use of this Char Item is empty, please check",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Use/Not Use of this Char Item is empty, please check.");

        ///
        ///
        public static readonly ErrorCode E_WrongPowerRunScenario_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 16,
            template: "Wrong PowerRunScenario ({0}): {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "PowerRunScenario format check fail, need to modify.");

        ///
        ///
        public static readonly ErrorCode W_WrongPowerRunScenario_02 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 17,
            template: "Wrong PowerRunScenario : {0}",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "PowerRunScenario should be init_(NV|Sweep|VRS)_PL_(NV|Sweep|VRS) due to exist extra inits or payloads.");

        ///
        ///
        public static readonly ErrorCode E_WrongPowerRunScenario_03 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 12,
            template: "Spell wrong PowerRunScenario",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "The PowerRunScenario header is misspelled. Please use the exact header name 'PowerRunScenario'.");

        ///
        ///
        public static readonly ErrorCode E_EmptyHeader_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 13,
            template: "EmptyColumn Need Check: {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "An empty column header was detected. Please enter a valid header name or remove the unused column.");

        ///
        ///
        public static readonly ErrorCode W_PowerGroupApply_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 18,
            template: "Suggest tie all {0} related power together",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Sweep condition should apply to pin group for vdd_low and vdd_fixed, instead of indivisual of them.");

        ///
        ///
        public static readonly ErrorCode W_RtosUserdef6Syntax_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Assignment,
            code: 5,
            template: "There is no RTOS_Userdef6_Syntax in Revision sheet",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "RTOS Userdef6 Syntax is not assigned, please check it.");

        ///
        ///
        public static readonly ErrorCode E_RtosUserdef6Syntax_02 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 19,
            template: "RTOS_Userdef6_Syntax: {0} is wrong",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "The RTOS command format in UserDef6 is invalid. Please verify that the syntax follows the supported RTOS command format.");

        ///
        ///
        public static readonly ErrorCode E_TrackingPinMismatchPrimaryStep_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 1,
            template: "Mismatch tracking pins: {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Tracking pins with the same Order must have the same number of shmoo points. " +
                    " Please verify that the Start, Stop, and Step settings generate consistent shmoo point counts.");

        ///
        ///
        public static readonly ErrorCode E_PrimaryTrackingPinValtInconsist_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Setting,
            code: 1,
            template: "Shmoo Setting Contains Valt and Vmain",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "A shmoo tracking group must use either Valt or Vmain consistently. " +
                    " Please verify the voltage type setting for all shmoos with the same Order.");

        ///
        ///
        public static readonly ErrorCode W_SramPinNotTieLogic_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Pin,
            code: 2,
            template: "Shmoo pin only define: {0}; suggest to tie: {1}, {2}",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "SELSRAM-related shmoo pins must be defined together. Please include all associated pins in the same shmoo setting.");

        ///
        ///
        public static readonly ErrorCode W_InitPatternbehindSelsram_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 2,
            template: "Exist init pattern behind selsram pattern",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Check if selsram pattern is always the last init pattern in a single char row.");

        ///
        ///
        public static readonly ErrorCode E_SelsramPatternNotFoundInMappingTable_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 3,
            template: "Cannot find any init patterns in mapping table",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Confirm the selsram pattern is exist in mapping table.");

        ///
        ///
        public static readonly ErrorCode E_SelsramBitMismatch_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 7,
            template: "UserDef9 data mismatch with mapping table",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Confirm the selsram refer length in char plan and selsram pattern bits in mapping table is match.");

        ///
        ///
        public static ErrorCode W_NotInHarvTable_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 8,
            template: "{0} in {1}, Row: {2} is not defined in Harv_Mapping_Table",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "The specified HARV FSTP was not found in the Harv_Mapping_Table. " +
                    "Please verify the HARV FSTP name and ensure a corresponding mapping entry exists.");

        ///
        ///
        public static readonly ErrorCode W_MissingManualAcInProgram_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.AcCategory,
            code: 1,
            template: "Can't serach category: {0} in program ac specs, use {1} as base to generate",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "The specified AC category was not found in the program AC specs. " +
                    "Please verify the category name or add the corresponding AC category definition.");

        ///
        ///
        public static readonly ErrorCode W_MissingManualAcInTimeSettingsSheet_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.AcCategory,
            code: 2,
            template: "Char manual ac: {0} is not in timesettings sheet, directly use program ac spec",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "The Manual AC setting is not defined in the TimeSettings sheet. " +
                    "The program AC spec will be used as the fallback source.");

        ///
        ///
        public static readonly ErrorCode E_MissingManualAcInProgAndTsSheet_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.AcCategory,
            code: 3,
            template: "Char manual ac: {0} is not in timesettings sheet and program ac spec",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "The Manual AC setting must exist in the TimeSettings sheet or program AC specs.");

        ///
        ///
        public static readonly ErrorCode E_IllegalShmooValue_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 20,
            template: "Shmoo value is non-numeric string",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Shmoo Start, Stop, and Step values must be valid numeric values. " +
                    "Please verify that all shmoo settings contain numeric data only.");

        ///
        ///
        public static readonly ErrorCode W_MultiUsedAcForPayload_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.AcCategory,
            code: 1,
            template: "Multi ac category for payload({0}) in base program: {1}",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Each payload in the base program must be associated with only one AC category. " +
                    "Please verify the payload configuration and remove any unexpected AC category assignments.");

        ///
        ///
        public static readonly ErrorCode W_MultiUsedTimesetForPayload_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Timing,
            code: 1,
            template: "Multi timeset for payload({0}) in base program: {1}",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Each payload in the base program must be associated with only one timeset. " +
                    "Please verify the payload configuration and remove any unexpected timeset assignments.");

        ///
        ///
        public static readonly ErrorCode W_MultiUsedDcForPayload_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.DcCategory,
            code: 1,
            template: "Multi dc category for payload({0}) in base program: {1}",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Each payload in the base program must be associated with only one DC category. " +
                    "Please verify the payload configuration and remove any unexpected DC category assignments.");

        ///
        ///
        public static readonly ErrorCode E_MissingUsedAcForPayload_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.AcCategory,
            code: 4,
            template: "None of ac category for payload({0}) in base program.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Each payload in the base program must have a corresponding AC category. " +
                    "Please verify the payload configuration and ensure an AC category is defined.");

        ///
        ///
        public static readonly ErrorCode E_MissingUsedTimesetForPayload_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Timing,
            code: 1,
            template: "None of timeset for payload({0}) in base program",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Each payload in the base program must have a corresponding timeset. " +
                    "Please verify the payload configuration and ensure an timeset is defined.");

        ///
        ///
        public static readonly ErrorCode E_MissingUsedDcForPayload_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.DcCategory,
            code: 1,
            template: "None of dc category for payload({0}) in base program",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Each payload in the base program must have a corresponding DC category. " +
                    "Please verify the payload configuration and ensure an DC category is defined.");

        ///
        ///
        public static readonly ErrorCode W_MissingDigSrcSgmtInPattern_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Source,
            code: 1,
            template: "Pattern: \"{0}\" has not \"{1}\"",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "The specified segment was not found in the pattern's Send Bit String definition. " +
                    "Please verify the segment name and ensure it is defined in the corresponding pattern information.");

        ///
        ///
        public static readonly ErrorCode W_VoltageHigherThan1P3_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Voltage,
            code: 1,
            template: "Shmoo {0} value >= 1.3",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "The absolute value of shmoo value should be less than 1.3. " +
                    "Please verify the Start/STOP setting and adjust it to a valid range.");

        ///
        ///
        public static readonly ErrorCode E_VddShmooLowToHigh_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Setting,
            code: 2,
            template: "Shmoo range is from low to high",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "The Shmoo Start value must be greater than the Stop value. " +
                    "Please define the shmoo range from high to low.");

        ///
        ///
        public static readonly ErrorCode E_MissingPinMap_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.File,
            code: 2,
            template: "Missing PinMap.txt",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Missing PinMap.txt. Please Reload PinMap or Test Program.");

        ///
        ///
        public static readonly ErrorCode W_MissingPatInfo_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.File,
            code: 3,
            template: "Pattern info is empty, please check!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Pattern info is empty, please check!");

        ///
        ///
        public static readonly ErrorCode E_IllegalForNewTChar_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 14,
            template: "Define multi patterns",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Can't define multi patterns here, please define one cell by one pattern.");

        ///
        ///
        public static readonly ErrorCode E_IllegalForNewTChar_02 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 15,
            template: "Retention format error",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please follow {INIT/PL}{Index}:{RetentionTime}:{GuardBand}mV, e.g INIT2:0.02:+25mV.");

        ///
        ///
        public static readonly ErrorCode E_IllegalForNewTChar_03 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 16,
            template: "{0} pattern didn't be defined, please check.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "INIT/Payload pattern in retention cannot find corresponding INIT/Payload pattern in char plan.");

        ///
        ///
        public static readonly ErrorCode E_IllegalCharUslLsl_01 = new ErrorCode(
            enumErrorCategory: EnumErrorCategory.Char,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 17,
            template: "Can not convert {0} to a value for limit!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Invalid USL/LSL format. The value must be a valid double number.");
    }
}
