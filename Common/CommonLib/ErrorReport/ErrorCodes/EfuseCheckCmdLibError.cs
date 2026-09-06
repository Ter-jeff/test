using CommonLib.ErrorReport.Base;

namespace CommonLib.ErrorReport.ErrorCodes
{
    public static class EfuseCheckCmdLibError
    {
        /// <summary>Template: ""</summary>
        public static readonly ErrorCode E_MismatchFormat_01 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 1,
            template: "The bool value sould be true",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please check the value should be true");

        /// <summary>Template: ""</summary>
        public static readonly ErrorCode E_MismatchFormat_02 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 2,
            template: "The bool value sould be false",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please check the value should be false");

        /// <summary>Template: ""</summary>
        public static readonly ErrorCode E_MismatchConfig_01 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Config,
            code: 1,
            template: "",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "");

        /// <summary>Template: ""</summary>
        public static readonly ErrorCode E_MismatchConfig_02 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Config,
            code: 2,
            template: "",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "");

        /// <summary>Template: "Value should be 0. "</summary>
        public static readonly ErrorCode E_MismatchValue_16 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 16,
            template: "Value should be 0. ",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "");

        /// <summary>Template: "{0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MismatchValue_17 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 17,
            template: "{0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please check the value.");

        /// <summary>Template: ""</summary>
        public static readonly ErrorCode E_MismatchValue_18 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 18,
            template: "{0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please check the check sum value");

        /// <summary>Template: "Compare item not match. "</summary>
        public static readonly ErrorCode E_MismatchValue_19 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 19,
            template: "Compared item not match.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please check Compared item value does not match");

        /// <summary>Template: ""</summary>
        public static readonly ErrorCode E_MismatchValue_20 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 20,
            template: "",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please check CRC item value");

        /// <summary>Template: "Datalog: {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MismatchValue_21 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 21,
            template: "Datalog: {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please check High Limit value in Data Log");

        /// <summary>Template: "Datalog: {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MismatchValue_22 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 22,
            template: "Datalog: {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please check Low Limit value in Data Log");

        /// <summary>Template: "CalcEccValue should be {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MismatchValue_23 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 23,
            template: "CalcEccValue should be {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please check the CalcEcc Value");

        /// <summary>Template: "Not Match with fuse_revision. "</summary>
        public static readonly ErrorCode E_MismatchValue_24 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 24,
            template: "Not Match with fuse_revision.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please check the fuse_revision");

        /// <summary>Template: "Ids compare item not match."</summary>
        public static readonly ErrorCode E_MismatchValue_25 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 25,
            template: "Ids compared item not match.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please check Ids compared item value");

        /// <summary>Template: ""</summary>
        public static readonly ErrorCode E_MismatchValue_26 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 26,
            template: "Ids value not match.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please check Ids value");

        /// <summary>Template: "Prober:{0} not same"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MismatchValue_27 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 27,
            template: "Prober:{0} not same",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please check the value between Prober and DSSC");

        /// <summary>Template: "The programming stage is same between {0} &amp; {1}"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_MismatchValue_28 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 28,
            template: "The programming stage is same between {0} & {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please check The programming stage");

        /// <summary>Template: "{0}_W{1}_X{2}_Y{3},Calculate_PRR:{4}"</summary>
        /// <remarks>{0} = ?, {1} = ?, {2} = ?, {3} = ?, {4} = ?</remarks>
        public static readonly ErrorCode E_MismatchValue_29 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 29,
            template: "{0}_W{1}_X{2}_Y{3},Calculate_PRR:{4}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please check PRR value");

        /// <summary>Template: "Not Match with PRR_Code. "</summary>
        public static readonly ErrorCode E_MismatchValue_30 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 30,
            template: "Not Match with PRR_Code.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please check PRR value");

        /// <summary>Template: "Value Not Same"</summary>
        public static readonly ErrorCode E_MismatchValue_31 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 31,
            template: "CFG Value is not the same",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please check CFG value");

        /// <summary>Template: "X{0}_Y{1}, from dssc value (x_coo, y_coo):X{2}_Y{3}"</summary>
        /// <remarks>{0} = ?, {1} = ?, {2} = ?, {3} = ?</remarks>
        public static readonly ErrorCode E_MismatchValue_32 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 32,
            template: "X{0}_Y{1}, from dssc value (x_coo, y_coo):X{2}_Y{3}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please check the XY value with DSSC");

        /// <summary>Template: "The real value of the field are all the same, please check!!"</summary>
        public static readonly ErrorCode I_MismatchValue_01 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 1,
            template: "The real value of the field are all the same, please check!!",
            enumErrorLevel: EnumErrorLevel.Info,
            guidance: "Please check the scenario");

        /// <summary>Template: "Not Match to Config BaseVoltage {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode W_MismatchValue_01 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 1,
            template: "Not Match to Config BaseVoltage {0}",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Please check Config BaseVoltage value");

        /// <summary>Template: "Fail"</summary>
        public static readonly ErrorCode W_MismatchValue_02 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 2,
            template: "Check SVM_CFuse Value Fail",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Please Check SVM_CFuse Value Fail");

        /// <summary>Template: "{0}:Sum of related pin limit value exceed table range"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode W_MismatchValue_03 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 3,
            template: "{0}:Sum of related pin limit value exceed table range",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Please check the limit value in the range");

        /// <summary>Template: "{0}:Meas limit:{1} exceed table"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode W_MismatchValue_04 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 4,
            template: "{0}:Meas limit:{1} exceed table",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Please check the Meas limit value in the range");

        /// <summary>Template: "Not Found in Datalog"</summary>
        public static readonly ErrorCode E_MissingValue_01 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 1,
            template: "Not Found in Datalog.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please check this filed in datalog");

        /// <summary>Template: "Null in Datalog. "</summary>
        public static readonly ErrorCode E_MissingValue_02 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 2,
            template: "Null in Datalog.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please check this filed in datalog");

        /// <summary>Template: "Scenario:{0} Not Found"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingValue_03 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 3,
            template: "Scenario:{0} Not Found",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please check the Scenario value");

        /// <summary>Template: "Can not find x_coo or y_coo"</summary>
        public static readonly ErrorCode E_MissingValue_04 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 4,
            template: "Can not find x_coo or y_coo",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Please check x_coo and y_coo values");

        /// <summary>Template: "X or Y value not found from dssc value, x: {0}, y: {1}"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_MissingValue_05 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 5,
            template: "X or Y value not found from dssc value, x: {0}, y: {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "X or Y value not found from dssc value");

        /// <summary>Template: "SetWriteVariable mismatch for '{0}': value in sheet does not match extracted value."</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MismatchValue_01 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 1,
            template: "SetWriteVariable mismatch for '{0}': value in sheet does not match extracted value.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the SetWriteVariable sheet row against the expected extracted value.");

        /// <summary>Template: "No matching bin results found with FuseCheck table."</summary>
        public static readonly ErrorCode E_MismatchValue_02 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 2,
            template: "No matching bin results found with FuseCheck table.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify that the FuseCheck table contains entries matching the current bin configuration.");

        /// <summary>Template: "FuseCheck mismatch '{0}': real value does not match table value."</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MismatchValue_03 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 3,
            template: "FuseCheck mismatch '{0}': real value '{1}' does not match table value '{2}'.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Compare the real fuse value against the FuseCheck table value for the flagged cell.");

        /// <summary>Template: "FuseCheck mismatch '{0}': real value does not match table value (second check)."</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MismatchValue_04 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 4,
            template: "FuseCheck mismatch '{0}': real value '{1}' does not match table value '{2}' (second check).",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Compare the real fuse value against the FuseCheck table value for the flagged cell.");

        /// <summary>Template: "FuseCheck mismatch '{0}': real value does not match table value (third check)."</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MismatchValue_05 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 5,
            template: "FuseCheck mismatch '{0}': real value '{1}' does not match table value '{2}' (third check).",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Compare the real fuse value against the FuseCheck table value for the flagged cell.");

        /// <summary>Template: "FuseCheck mismatch in Bank and Field for entry '{0}'."</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MismatchValue_06 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 6,
            template: "FuseCheck mismatch in Bank and Field.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the Bank and Field values in the FuseCheck table match the expected configuration.");

        /// <summary>Template: "No DRAM files found in EnableWords for entry '{0}'."</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MismatchValue_07 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 7,
            template: "No such DRAM files in EnableWords",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Ensure that DRAM files are present in EnableWords for the flagged entry.");

        /// <summary>Template: "DRAM type mismatch '{0}': real value does not match table value."</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MismatchValue_08 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 8,
            template: "DRAM type mismatch '{0}': real value '{1}' does not match table value '{2}'.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Compare the DRAM type value against the EnableWords table value for the flagged cell.");

        /// <summary>Template: "No matching DRAM type found in EnableWords for entry '{0}'."</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MismatchValue_09 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 9,
            template: "No matching DRAM type in EnableWords",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify that the DRAM type is defined in the EnableWords configuration.");

        /// <summary>Template: "CRC mismatch for sub-config '{0}': value does not match expected."</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MismatchValue_10 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 10,
            template: "CRC mismatch for sub-config '{0}': value '{1}' does not match expected value '{2}'.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the CRC value in the bit definition row against the expected CRC for the sub-config.");

        /// <summary>Template: "Default value mismatch for sub-config '{0}': value does not match bit definition default."</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MismatchValue_11 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 11,
            template: "Default value mismatch for sub-config '{0}': value '{1}' does not match bit definition default value '{2}'.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the default value in the bit definition row matches the expected default for the sub-config.");

        /// <summary>Template: "After programming stage, value for sub-config '{0}' does not match expected default."</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MismatchValue_12 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 12,
            template: "After programming stage, value '{1}' for sub-config '{0}' does not match expected default value '{2}'.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Review the programming stage result for the sub-config and compare against the bit definition default.");

        /// <summary>Template: "Log entry missing for '{0}' not found in the programming log."</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MismatchValue_13 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 13,
            template: "Missing in log : {0} {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Ensure the bit definition entry is present in the programming log output.");

        /// <summary>Template: "The bit width of {0} : {1} is not match table {2} !!!"</summary>
        /// <remarks>{0} = ?, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_MismatchValue_14 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 14,
            template: "The bit width of {0} : {1} is not match table {2} !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Review the formula output and verify it matches the expected result.");

        /// <summary>Template: "Can not find {0} in log !!!"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MismatchValue_15 = new(
            enumErrorCategory: EnumErrorCategory.EfuseCheck,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 15,
            template: "Can not find {0} in log !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Review the formula output and verify it matches the expected result.");
    }
}
