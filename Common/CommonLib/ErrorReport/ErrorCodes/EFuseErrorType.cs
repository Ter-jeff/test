using CommonLib.ErrorReport.Base;

namespace CommonLib.ErrorReport.ErrorCodes
{
    public static class EFuseErrorType
    {
        /// <summary>Template: "LSB BIT isn't a Number. LSB BIT : {0}"</summary>
        /// <remarks>{0} = bit name or address</remarks>
        public static readonly ErrorCode E_RuleViolationBit_01 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Bit,
            code: 1,
            template: "LSB BIT isn't a Number. LSB BIT : {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the EFuse bit definition for the specified bit. "
                    + "Verify the bit address, width, and value are all within valid ranges.");

        /// <summary>Template: "MSB BIT isn't a Number. MSB BIT : {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_RuleViolationBit_02 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Bit,
            code: 2,
            template: "MSB BIT isn't a Number. MSB BIT : {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the EFuse bit definition for the specified bit. "
                    + "Verify the bit address, width, and value are all within valid ranges.");

        /// <summary>Template: "Bit Width isn't a Number. Bit Width : {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_RuleViolationBit_03 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Bit,
            code: 3,
            template: "Bit Width isn't a Number. Bit Width : {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the EFuse bit definition for the specified bit. "
                    + "Verify the bit address, width, and value are all within valid ranges.");

        /// <summary>Template: "Bit Width doesn't Equals with End Bit - Start Bit + 1. Result is {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_RuleViolationBit_04 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Bit,
            code: 4,
            template: "Bit Width doesn't Equals with End Bit - Start Bit + 1. Result is {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the EFuse bit definition for the specified bit. "
                    + "Verify the bit address, width, and value are all within valid ranges.");

        /// <summary>Template: "Incorrect bit width in the column Default Value. Bit Width (Math.Ceiling({0}/4) = {1}) &lt; Default value ({2})"</summary>
        /// <remarks>{0} = ?, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_RuleViolationBit_05 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Bit,
            code: 5,
            template: "Incorrect bit width in the column Default Value. Bit Width (Math.Ceiling({0}/4) = {1}) < Default value ({2})",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the EFuse bit definition for the specified bit. "
                    + "Verify the bit address, width, and value are all within valid ranges.");

        /// <summary>Template: "Incorrect bit width in the column Default Value. Bit Width ({0}) &lt; Default value ({1})"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_RuleViolationBit_06 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Bit,
            code: 6,
            template: "Incorrect bit width in the column Default Value. Bit Width ({0}) < Default value ({1})",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the EFuse bit definition for the specified bit. "
                    + "Verify the bit address, width, and value are all within valid ranges.");

        /// <summary>Template: "The content of this column must be [MSB:LSB] or [LSB:MSB]. Content {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_RuleViolationBit_07 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Bit,
            code: 7,
            template: "The content of this column must be [MSB:LSB] or [LSB:MSB]. Content {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the EFuse bit definition for the specified bit. "
                    + "Verify the bit address, width, and value are all within valid ranges.");

        /// <summary>Template: "Non-numeric string exists in Fuse column. Content {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_RuleViolationBit_08 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Bit,
            code: 8,
            template: "Non-numeric string exists in Fuse column. Content {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the EFuse bit definition for the specified bit. "
                    + "Verify the bit address, width, and value are all within valid ranges.");

        /// <summary>Template: "This row format is MSB:LSB is inconsistent with the LSB:MSB used by other rows. Content {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_RuleViolationBit_09 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Bit,
            code: 9,
            template: "This row format is MSB:LSB is inconsistent with the LSB:MSB used by other rows. Content {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the EFuse bit definition for the specified bit. "
                    + "Verify the bit address, width, and value are all within valid ranges.");

        /// <summary>Template: "This row format is LSB:MSB is inconsistent with the MSB:LSB used by other rows. Content {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_RuleViolationBit_10 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Bit,
            code: 10,
            template: "This row format is LSB:MSB is inconsistent with the MSB:LSB used by other rows. Content {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the EFuse bit definition for the specified bit. "
                    + "Verify the bit address, width, and value are all within valid ranges.");

        /// <summary>Template: "The item {0}'s Bitwidth of the CMP block is not same as UDR Block"</summary>
        /// <remarks>{0} = CMP field or context</remarks>
        public static readonly ErrorCode E_RuleViolationCMP_01 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Block,
            code: 1,
            template: "The item {0}'s Bitwidth of the CMP block is not same as UDR Block",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the EFuse comparator configuration and value. "
                    + "Ensure the comparator threshold is valid for the specified fuse field.");

        /// <summary>Template: "Crc bit format needs to use [number:number]"</summary>
        public static readonly ErrorCode E_InvalidCRC_01 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 1,
            template: "Crc bit format needs to use [number:number]",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Recalculate the EFuse CRC for the specified block. "
                    + "Ensure the input data matches the programmed fuse values before recomputing.");

        /// <summary>Template: "Item {0} bit range {1}:{2} exist itself"</summary>
        /// <remarks>{0} = ?, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_InvalidCRC_02 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 2,
            template: "Item {0} bit range {1}:{2} exist itself",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Recalculate the EFuse CRC for the specified block. "
                    + "Ensure the input data matches the programmed fuse values before recomputing.");

        /// <summary>Template: "Item {0} bit range {1}:{2} does not exist in ignore bit"</summary>
        /// <remarks>{0} = ?, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_InvalidCRC_03 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 3,
            template: "Item {0} bit range {1}:{2} does not exist in ignore bit",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Recalculate the EFuse CRC for the specified block. "
                    + "Ensure the input data matches the programmed fuse values before recomputing.");

        /// <summary>Template: "The CRC bit range {0}:{1} exist in itself and the programming stage {2} is untested job. CRC item: {3}, programming stage: {4}"</summary>
        /// <remarks>{0} = ?, {1} = ?, {2} = ?, {3} = ?, {4} = ?</remarks>
        public static readonly ErrorCode E_InvalidCRC_04 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 4,
            template: "The CRC bit range {0}:{1} exist in itself and the programming stage {2} is untested job. CRC item: {3}, programming stage: {4}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Recalculate the EFuse CRC for the specified block. "
                    + "Ensure the input data matches the programmed fuse values before recomputing.");

        /// <summary>Template: "The bit range {0}:{1} does not exist in the ignore bit and the programming stage {2} is untested job. CRC item: {3}, programming stage: {4}"</summary>
        /// <remarks>{0} = ?, {1} = ?, {2} = ?, {3} = ?, {4} = ?</remarks>
        public static readonly ErrorCode E_InvalidCRC_05 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 5,
            template: "The bit range {0}:{1} does not exist in the ignore bit and the programming stage {2} is untested job. CRC item: {3}, programming stage: {4}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Recalculate the EFuse CRC for the specified block. "
                    + "Ensure the input data matches the programmed fuse values before recomputing.");

        /// <summary>Template: "The CRC bit range {0}:{1} exist in itself and the programming stage {3} is untested job. CRC item: {4}, programming stage: {5}"</summary>
        /// <remarks>{0} = ?, {1} = ?, {3} = ?, {4} = ?, {5} = ?</remarks>
        public static readonly ErrorCode E_InvalidCRC_06 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 6,
            template: "The CRC bit range {0}:{1} exist in itself and the programming stage {3} is untested job. CRC item: {4}, programming stage: {5}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Recalculate the EFuse CRC for the specified block. "
                    + "Ensure the input data matches the programmed fuse values before recomputing.");

        /// <summary>Template: "The {0} bit range {1}:{2} does not exist in ignore bit and the programming stage {3} is untested job. CRC item: {4}, programming stage: {5}"</summary>
        /// <remarks>{0} = ?, {1} = ?, {2} = ?, {3} = ?, {4} = ?, {5} = ?</remarks>
        public static readonly ErrorCode E_InvalidCRC_07 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Format,
            code: 7,
            template: "The {0} bit range {1}:{2} does not exist in ignore bit and the programming stage {3} is untested job. CRC item: {4}, programming stage: {5}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Recalculate the EFuse CRC for the specified block. "
                    + "Ensure the input data matches the programmed fuse values before recomputing.");

        /// <summary>Template: "Missing {0} name in the config table"</summary>
        /// <remarks>{0} = field or context description</remarks>
        public static readonly ErrorCode E_MissingField_01 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Field,
            code: 1,
            template: "Missing {0} name in the config table",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the EFuse default value against the allowed bit-width for the field. "
                    + "Reduce the value or increase the allocated width to accommodate it.");

        /// <summary>Template: "ECID Programming Stage is {0}"</summary>
        /// <remarks>{0} = stage name or context</remarks>
        public static readonly ErrorCode E_RuleViolationStage_01 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Stage,
            code: 1,
            template: "ECID Programming Stage is {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the ECID non-CP programming stage against the expected stage definition. "
                    + "Ensure the stage is correctly configured for non-CP (non-circuit probe) testing.");

        /// <summary>Template: "{0} is integer, tool change to {1}"</summary>
        /// <remarks>{0} = sheet or resource name, {1} = ?</remarks>
        public static readonly ErrorCode E_InvalidValue_01 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 1,
            template: "{0} is integer, tool change to {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify that the EFuse sheet or referenced resource is present in the workbook. "
                    + "Check the sheet name spelling and ensure it was not accidentally deleted.");

        /// <summary>Template: "Fuse Blow Location is FTF"</summary>
        public static readonly ErrorCode W_RuleViolationFuseBlowLocation_01 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 1,
            template: "Fuse Blow Location is FTF",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Verify the EFuse blow location address against the device fuse map. "
                    + "Ensure the location is within the valid fuse address range.");

        /// <summary>Template: "Low Limit isn't a Number. Low Limit : {0}"</summary>
        /// <remarks>{0} = limit name or value</remarks>
        public static readonly ErrorCode E_InvalidLimit_01 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Limit,
            code: 1,
            template: "Low Limit isn't a Number. Low Limit : {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the EFuse limit value against the specification. "
                    + "Ensure the limit is within the allowed range for the fuse type.");

        /// <summary>Template: "High Limit isn't a Number. High Limit : {0}"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_InvalidLimit_02 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Limit,
            code: 2,
            template: "High Limit isn't a Number. High Limit : {0}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the EFuse limit value against the specification. "
                    + "Ensure the limit is within the allowed range for the fuse type.");

        /// <summary>Template: "When DefalutOrReal is Real, High Limit must large than Low Limit."</summary>
        public static readonly ErrorCode E_InvalidLimit_03 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Limit,
            code: 3,
            template: "When DefalutOrReal is Real, High Limit must large than Low Limit.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the EFuse limit value against the specification. "
                    + "Ensure the limit is within the allowed range for the fuse type.");

        /// <summary>Template: "IDS resolution * 2^(Bit Range) &lt; High Limit."</summary>
        public static readonly ErrorCode E_InvalidLimit_04 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Limit,
            code: 4,
            template: "IDS resolution * 2^(Bit Range) < High Limit.",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the EFuse limit value against the specification. "
                    + "Ensure the limit is within the allowed range for the fuse type.");

        /// <summary>Template: "The lsb bit can not be recognized"</summary>
        public static readonly ErrorCode E_InvalidLsbMsb_01 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Type,
            code: 1,
            template: "The lsb bit can not be recognized",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the bit ordering (LSB/MSB) configuration for the specified EFuse field. "
                    + "Ensure the ordering matches the device specification and programming tool expectation.");

        /// <summary>Template: "The msb bit can not be recognized"</summary>
        public static readonly ErrorCode E_InvalidLsbMsb_02 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Type,
            code: 2,
            template: "The msb bit can not be recognized",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the bit ordering (LSB/MSB) configuration for the specified EFuse field. "
                    + "Ensure the ordering matches the device specification and programming tool expectation.");

        /// <summary>Template: "The last config condition MSB ({0}) &lt; Maximum Bit ({1})"</summary>
        /// <remarks>{0} = context or count description, {1} = ?</remarks>
        public static readonly ErrorCode E_InvalidMaximumBits_01 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Limit,
            code: 5,
            template: "The last config condition MSB ({0}) < Maximum Bit ({1})",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the total EFuse bit allocation does not exceed the device maximum. "
                    + "Reduce the number of fuse bits or reorganize the fuse map.");

        /// <summary>Template: "The database revision is not same with the efuse database revision sheet content, please check it!"</summary>
        public static readonly ErrorCode E_MismatchRevision_01 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Revision,
            code: 1,
            template: "The database revision is not same with the efuse database revision sheet content, please check it!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Compare the EFuse value against the expected definition. "
                    + "Check that the programmed value matches the test plan specification.");

        /// <summary>Template: "The database revision should be defined as default value, please check it!"</summary>
        public static readonly ErrorCode E_MismatchDefaultType_01 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Default,
            code: 1,
            template: "The database revision should be defined as default value, please check it!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the EFuse default type definition against the expected type. "
                    + "Ensure the type is consistent across all references to this field.");

        /// <summary>Template: "Missing Parameter in {0}({1}) : {2}"</summary>
        /// <remarks>{0} = parameter name, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode E_MissingParameter_01 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Type,
            code: 1,
            template: "Missing Parameter in {0}({1}) : {2}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Open the EFuse sheet and add the required parameter. "
                    + "Check the EFuse specification to confirm which parameters are mandatory.");

        /// <summary>Template: "The VBT function: {0} can not find in {1} library!"</summary>
        /// <remarks>{0} = module name, {1} = ?</remarks>
        public static readonly ErrorCode E_MissVbtModule_01 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Module,
            code: 1,
            template: "The VBT function: {0} can not find in {1} library!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the VBT module name against the expected module list. "
                    + "Add the missing module definition or correct the reference.");

        /// <summary>Template: "The pattern type is match to the BDF, pattern is {0}, but BDF is {1}"</summary>
        /// <remarks>{0} = field or entry name, {1} = ?</remarks>
        public static readonly ErrorCode E_NotMatchToBDF_01 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Type,
            code: 1,
            template: "The pattern type is match to the BDF, pattern is {0}, but BDF is {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Compare the EFuse definition against the BDF. "
                    + "Update either the EFuse configuration or the BDF so the definitions are consistent.");

        /// <summary>Template: "This bank count should be mod by 16 due to the ECC"</summary>
        public static readonly ErrorCode E_MismatchBit_01 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Bit,
            code: 1,
            template: "This bank count should be mod by 16 due to the ECC",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the EFuse bit count against the device specification. "
                    + "Ensure the total bits allocated do not exceed the physical fuse capacity.");

        /// <summary>Template: "This column cannot be empty"</summary>
        public static readonly ErrorCode E_MismatchBit_02 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Bit,
            code: 2,
            template: "This column cannot be empty",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the EFuse bit count against the device specification. "
                    + "Ensure the total bits allocated do not exceed the physical fuse capacity.");

        /// <summary>Template: "This column must be a number"</summary>
        public static readonly ErrorCode E_MismatchBit_03 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Bit,
            code: 3,
            template: "This column must be a number",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the EFuse bit count against the device specification. "
                    + "Ensure the total bits allocated do not exceed the physical fuse capacity.");

        /// <summary>Template: "The pattern {0} is not found in pattern directory!!!"</summary>
        /// <remarks>{0} = pattern name</remarks>
        public static readonly ErrorCode E_MissingPattern_01 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 1,
            template: "The pattern {0} is not found in pattern directory!!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the EFuse pattern name and its definition. "
                    + "Check that the pattern exists in the pattern folder and is correctly referenced.");

        /// <summary>Template: "The pattern {0} can't be found in the CSV !!!"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingPattern_02 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 2,
            template: "The pattern {0} can't be found in the CSV !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the EFuse pattern name and its definition. "
                    + "Check that the pattern exists in the pattern folder and is correctly referenced.");

        /// <summary>Template: "EFuse pattern '{0}' has an error."</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingPattern_03 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 3,
            template: "This pattern {0} is \"Dont_useInCsv\" !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the EFuse pattern name. "
                    + "Check that the pattern exists in the pattern folder and is correctly referenced.");

        /// <summary>Template: "EFuse pattern '{0}' has an error."</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingPattern_04 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 4,
            template: "This pattern {0} of the FileVersion column is \"n/a\" !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the EFuse pattern name and its definition. "
                    + "Check that the pattern exists in the pattern folder and is correctly referenced.");

        /// <summary>Template: "This pattern {0} can't get Read/Write pin !!!"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingPattern_05 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 5,
            template: "This pattern {0} can't get Read/Write pin !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the EFuse pattern name and its definition. "
                    + "Check that the pattern exists in the pattern folder and is correctly referenced.");

        /// <summary>Template: "The pattern {0} can't be found in the CSV !!!"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode E_MissingPattern_06 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Missing,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 6,
            template: "The pattern {0} can't be found in the CSV !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the EFuse pattern name and its definition. "
                    + "Check that the pattern exists in the pattern folder and is correctly referenced.");

        /// <summary>Template: "The pattern {0} send count {1} {1} is not equal to {2} in the bank {3}, {4}*{5} !!!!!"</summary>
        /// <remarks>{0} = ?, {1} = ?, {2} = ?, {3} = ?, {4} = ?, {5} = ?</remarks>
        public static readonly ErrorCode E_MismatchPattern_01 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 1,
            template: "The pattern {0} send count {1} {1} is not equal to {2} in the bank {3}, {4}*{5} !!!!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the EFuse pattern name and its definition. "
                    + "Check that the pattern exists in the pattern folder and is correctly referenced.");

        /// <summary>Template: "The pattern {0} access mode (DAA) is not equal to BDF file {1}"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_MismatchPattern_02 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 2,
            template: "The pattern {0} access mode (DAA) is not equal to BDF file {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the EFuse pattern name and its definition. "
                    + "Check that the pattern exists in the pattern folder and is correctly referenced.");

        /// <summary>Template: "The pattern {0} access mode (JTG) is not equal to BDF file {1}"</summary>
        /// <remarks>{0} = ?, {1} = ?</remarks>
        public static readonly ErrorCode E_MismatchPattern_03 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 3,
            template: "The pattern {0} access mode (JTG) is not equal to BDF file {1}",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the EFuse pattern name and its definition. "
                    + "Check that the pattern exists in the pattern folder and is correctly referenced.");

        /// <summary>Template: "The pattern {0} ({1}) store count {2} is not equal to {3} in the bank {4} ({5}), {6}*{7} !!!"</summary>
        /// <remarks>{0} = ?, {1} = ?, {2} = ?, {3} = ?, {4} = ?, {5} = ?, {6} = ?, {7} = ?</remarks>
        public static readonly ErrorCode E_MismatchPattern_04 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 4,
            template: "The pattern {0} ({1}) store count {2} is not equal to {3} in the bank {4} ({5}), {6}*{7} !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the EFuse pattern name and its definition. "
                    + "Check that the pattern exists in the pattern folder and is correctly referenced.");

        /// <summary>Template: "This pattern {0} can't be discerned in what mode !!!"</summary>
        /// <remarks>{0} = pattern or test mode name</remarks>
        public static readonly ErrorCode E_MismatchPattern_05 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 5,
            template: "This pattern {0} can't be discerned in what mode !!!",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the EFuse pattern test mode against the supported mode list. "
                    + "Update the configuration to use a valid test mode.");

        /// <summary>Template: "Programming Stage is FTF"</summary>
        public static readonly ErrorCode E_RuleViolationStage_02 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Stage,
            code: 2,
            template: "Programming Stage is FTF",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Verify the EFuse programming stage definition against the expected stage list. "
                    + "Ensure the stage name and sequence are consistent with the test plan.");

        /// <summary>Template: "The programming stage is same between {0} &amp; {1}"</summary>
        /// <remarks>{0} = stage name or context, {1} = ?</remarks>
        public static readonly ErrorCode W_RuleViolationStage_03 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Stage,
            code: 3,
            template: "The programming stage is same between {0} & {1}",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Verify the shadow programming stage definition against the expected stage list. "
                    + "Ensure the shadow stage is correctly configured for the fuse type.");

        /// <summary>Template: "The count of rows in sheet {0} ({1}) is inconsistent with sheet {2} ({3})"</summary>
        /// <remarks>{0} = row identifier or context, {1} = ?, {2} = ?, {3} = ?</remarks>
        public static readonly ErrorCode E_MismatchRow_01 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Row,
            code: 1,
            template: "The count of rows in sheet {0} ({1}) is inconsistent with sheet {2} ({3})",
            enumErrorLevel: EnumErrorLevel.Error,
            guidance: "Check the EFuse row information against the device fuse map layout. "
                    + "Ensure row addresses and counts are consistent with the specification.");

        /// <summary>Template: "Tool filter out the bank name {0} which contains the keyword bira!!!"</summary>
        /// <remarks>{0} = bank name or number</remarks>
        public static readonly ErrorCode I_RuleViolationBank_01 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Bank,
            code: 1,
            template: "Tool filter out the bank name {0} which contains the keyword bira!!!",
            enumErrorLevel: EnumErrorLevel.Info,
            guidance: "Open the EFuse bank definition and inspect it for invalid fields. "
                    + "Verify bank address, width, and bit assignments against the device specification.");

        /// <summary>Template: "Tool filter out the bank name {0} which job is SLT!!!"</summary>
        /// <remarks>{0} = ?</remarks>
        public static readonly ErrorCode I_RuleViolationBank_02 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Bank,
            code: 2,
            template: "Tool filter out the bank name {0} which job is SLT!!!",
            enumErrorLevel: EnumErrorLevel.Info,
            guidance: "Open the EFuse bank definition and inspect it for invalid fields. "
                    + "Verify bank address, width, and bit assignments against the device specification.");

        /// <summary>Template: "Bank Name ({0}) is not recognized"</summary>
        /// <remarks>{0} = bank name or number</remarks>
        public static readonly ErrorCode W_RuleViolationBank_01 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Bank,
            code: 1,
            template: "Bank Name ({0}) is not recognized",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Check the EFuse bank identifier against the defined bank list. "
                    + "Add the missing bank definition or correct the reference to an existing bank.");

        /// <summary>Template: "The base row has different value from other base rows"</summary>
        public static readonly ErrorCode W_MismatchRow_01 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Mismatch,
            enumErrorTarget: EnumErrorTarget.Row,
            code: 1,
            template: "The base row has different value from other base rows",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Verify the EFuse base voltage against the device electrical specifications. "
                    + "Ensure the voltage level is within the allowed operating range.");

        /// <summary>Template: "Bit width over 1024 is not allowed."</summary>
        public static readonly ErrorCode W_InvalidBit_01 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Bit,
            code: 1,
            template: "Bit width over 1024 is not allowed.",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the EFuse bit definition table and verify the flagged bit entry. "
                    + "Check the bit address, width, and value are within valid ranges.");

        /// <summary>Template: "Default Value isn't a Number. Default Value : {0}"</summary>
        /// <remarks>{0} = field or context</remarks>
        public static readonly ErrorCode W_InvalidValue_01 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Invalid,
            enumErrorTarget: EnumErrorTarget.Value,
            code: 1,
            template: "Default Value isn't a Number. Default Value : {0}",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Open the EFuse definition and verify the default value for the second occurrence. "
                    + "Ensure the value is within the bit-width range for the specified field.");

        /// <summary>Template: "The database revision is judge by tool, please check it!"</summary>
        public static readonly ErrorCode W_RuleViolationRevision_01 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Revision,
            code: 1,
            template: "The database revision is judge by tool, please check it!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Check the EFuse data revision logic in the tool configuration. "
                    + "Ensure the revision comparison rules and input data are correct.");

        /// <summary>Template: "If Algorithm is base, Low Limit and High Limit and Default Value must be same."</summary>
        public static readonly ErrorCode W_RuleViolationLimit_01 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.RuleViolation,
            enumErrorTarget: EnumErrorTarget.Limit,
            code: 1,
            template: "If Algorithm is base, Low Limit and High Limit and Default Value must be same.",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Check the EFuse limit value against the specification. "
                    + "Ensure the limit is within the allowed range for the fuse type.");

        /// <summary>Template: "{0} is unused pattern (bank name: {1}, test mode: {2}) !!!"</summary>
        /// <remarks>{0} = pattern name, {1} = ?, {2} = ?</remarks>
        public static readonly ErrorCode W_RedundantPattern_01 = new(
            enumErrorCategory: EnumErrorCategory.EFuse,
            enumErrorBehavior: EnumErrorBehavior.Redundant,
            enumErrorTarget: EnumErrorTarget.Pattern,
            code: 1,
            template: "{0} is unused pattern (bank name: {1}, test mode: {2}) !!!",
            enumErrorLevel: EnumErrorLevel.Warning,
            guidance: "Review the EFuse pattern list and remove unused patterns, "
                    + "or add the missing test references to consume them.");
    }
}
