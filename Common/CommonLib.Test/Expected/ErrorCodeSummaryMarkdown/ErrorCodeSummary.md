## AutoAiErrorType (`AutoAiErrorType.cs`)

### E_MissingTimesetFile_01
- **FullCode**: EAI476001
- **EnumErrorCategory**: AutoAi
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Timing
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: [TimeSet Latest] The timeset "{0}" in pattern dash board is not existed in pattern folder
- **Guidance**: The TimeSet specified in the Pattern Dashboard must exist in the pattern folder. Please verify the TimeSet name and ensure the corresponding file is available.

### E_MissingPatternFile_01
- **FullCode**: EAI452001
- **EnumErrorCategory**: AutoAi
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pattern
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: [File Versions] The pattern "{0}" in pattern dash board is not existed in pattern folder
- **Guidance**: The pattern specified in the Pattern Dashboard must exist in the pattern folder. Please verify the pattern name and ensure the corresponding pattern file is available.

### E_MissingHeader_01
- **FullCode**: EAI433001
- **EnumErrorCategory**: AutoAi
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Header
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: [Header] Missing the header: " {0} "
- **Guidance**: Wrong header in input sheet, please check!

### W_MissingHeader_02
- **FullCode**: WAI433002
- **EnumErrorCategory**: AutoAi
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Header
- **Code**: 2
- **Level**: Warning
- **#Args**: 1
- **Template**: [Header] Missing the header: " {0} "
- **Guidance**: Wrong header in input sheet, please check!

### E_MissingSheet_01
- **FullCode**: EAI425001
- **EnumErrorCategory**: AutoAi
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: File
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Can not find " {0} " sheet in input file !!!
- **Guidance**: TBD

### W_FormatError_01
- **FullCode**: WAI229001
- **EnumErrorCategory**: AutoAi
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 1
- **Level**: Warning
- **#Args**: 1
- **Template**: [UseNotUse] The syntax "{0}" can not be identified, and this row will be ignored !!!
- **Guidance**: The UseNotUse field must be set to either 'Use' or 'Not Use'. Please correct the value to ensure the row is processed.

### E_FormatError_02
- **FullCode**: EAI229002
- **EnumErrorCategory**: AutoAi
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 2
- **Level**: Error
- **#Args**: 1
- **Template**: [AIType] The syntax "{0}" can not be identified, and treat this as Data log !!!
- **Guidance**: The AIType field must be set to 'Data log', '1D', or '2D'. Unrecognized values will be treated as 'Data log'.

### E_FormatError_03
- **FullCode**: EAI229003
- **EnumErrorCategory**: AutoAi
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 3
- **Level**: Error
- **#Args**: 1
- **Template**: [Data Logging Setting] The syntax "{0}" can not be identified, and treat this as NA !!!
- **Guidance**: The Data Logging Setting field must be set to 'NA', 'DFCList...', or 'DFCStep...'. Unrecognized values will be treated as 'NA'.

### E_FormatError_04
- **FullCode**: EAI229004
- **EnumErrorCategory**: AutoAi
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 4
- **Level**: Error
- **#Args**: 1
- **Template**: [Data Logging Setting] The DFC syntax "{0}" can not be identified, and treat this as NA !!!
- **Guidance**: The DFCList setting must follow the format DFCList(value1,value2,...). Please verify the syntax and ensure the values are enclosed in parentheses.

### E_FormatError_05
- **FullCode**: EAI229005
- **EnumErrorCategory**: AutoAi
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 5
- **Level**: Error
- **#Args**: 1
- **Template**: [Data Logging Setting] The DFC arg syntax "{0}" can not be identified, and treat this as NA !!!
- **Guidance**: All values specified in DFCList must be valid integer numbers. Please verify the DFC values and remove any unsupported characters or formats.

### E_FormatError_06
- **FullCode**: EAI229006
- **EnumErrorCategory**: AutoAi
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 6
- **Level**: Error
- **#Args**: 1
- **Template**: [Data Logging Setting] The DFC syntax "{0}" can not be identified, and treat this as NA !!!
- **Guidance**: The DFCStep setting must follow the format DFCStep(start,step). Please verify the syntax and ensure the values are enclosed in parentheses.

### E_FormatError_07
- **FullCode**: EAI229007
- **EnumErrorCategory**: AutoAi
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 7
- **Level**: Error
- **#Args**: 1
- **Template**: [Data Logging Setting] The DFC arg syntax "{0}" can not be identified, and treat this as NA !!!
- **Guidance**: The DFCStep setting must contain exactly two integer values in the format DFCStep(start,step). Please verify the number and format of the arguments.

### E_FormatError_08
- **FullCode**: EAI229008
- **EnumErrorCategory**: AutoAi
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 8
- **Level**: Error
- **#Args**: 1
- **Template**: [SELSRM] The syntax "{0}" can not be identified, and this will be ignored !!!
- **Guidance**: The SELSRAM DSSC setting must start with 'SELSRM' or 'DSELSRM'. Please verify the number and format of the arguments.

### W_CanNotDetermineWhichSpecToUse_01
- **FullCode**: WAI376001
- **EnumErrorCategory**: AutoAi
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Timing
- **Code**: 1
- **Level**: Warning
- **#Args**: 2
- **Template**: "[Mapping] Multi timeset for payload({0}) in base program: {1}
- **Guidance**: Each payload in the base program must be associated with only one timeset. Please verify the payload configuration and remove any unexpected timeset assignments.

### E_CanNotDetermineWhichSpecToUse_02
- **FullCode**: EAI476002
- **EnumErrorCategory**: AutoAi
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Timing
- **Code**: 2
- **Level**: Error
- **#Args**: 1
- **Template**: [Mapping] None of timeset for payload({0}) in base program
- **Guidance**: Each payload in the base program must have a corresponding timeset. Please verify the payload configuration and ensure an timeset is defined.

### W_CanNotDetermineWhichSpecToUse_03
- **FullCode**: WAI318001
- **EnumErrorCategory**: AutoAi
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: DcCategory
- **Code**: 1
- **Level**: Warning
- **#Args**: 2
- **Template**: [Mapping] Multi dc category for payload({0}) in base program: {1}
- **Guidance**: Each payload in the base program must be associated with only one DC category. Please verify the payload configuration and remove any unexpected DC category assignments.

### E_CanNotDetermineWhichSpecToUse_04
- **FullCode**: EAI418001
- **EnumErrorCategory**: AutoAi
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: DcCategory
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: [Mapping] None of dc category for payload({0}) in base program
- **Guidance**: Each payload in the base program must have a corresponding DC category. Please verify the payload configuration and ensure an DC category is defined.

## BasicErrorType (`BasicErrorType.cs`)

### E_MissingPin_01
- **FullCode**: EBA455001
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: The pin {0} in IO_PinGroup can not be found in IO_PinMap !!!
- **Guidance**: The pin specified in IO_PinGroup cannot be found in IO_PinMap. Please ensure that the pin exists in IO_PinMap.

### E_MissingPin_02
- **FullCode**: EBA455002
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 2
- **Level**: Error
- **#Args**: 1
- **Template**: The pin group {0} has more than two pin types !!!
- **Guidance**: The pin types are defined in the IO_PinMap sheet. Please verify the pin type settings in IO_PinMap.

### E_MissingPin_03
- **FullCode**: EBA455003
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 3
- **Level**: Error
- **#Args**: 1
- **Template**: The pin group name {0} and pin name can not be the same !!!
- **Guidance**: The pin group name specified in IO_PinGroup conflicts with a pin name. Please ensure that pin group names and pin names are unique.

### E_MissingPin_04
- **FullCode**: EBA455004
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 4
- **Level**: Error
- **#Args**: 2
- **Template**: Should define io infos in {0} for {1}!!!
- **Guidance**: Please check the IO configuration and identify any missing IO definitions. Add the required information, regenerate the project, and confirm that the issue is resolved.

### E_MissingPin_05
- **FullCode**: EBA455005
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 5
- **Level**: Error
- **#Args**: 2
- **Template**: Should define io infos concurrent in {0} for {1}!!!
- **Guidance**: Please ensure all required concurrent IO information is configured correctly. Add any missing definitions and regenerate the output files.

### E_MissingPin_06
- **FullCode**: EBA455006
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 6
- **Level**: Error
- **#Args**: 0
- **Template**: Missing PinMap in TestPlan
- **Guidance**: The PinMap definition is missing in the TestPlan. Please ensure that the corresponding PinMap is configured in the TestPlan.

### E_MissingPin_07
- **FullCode**: EBA455007
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 7
- **Level**: Error
- **#Args**: 2
- **Template**: Column[{0}] The pin: {1} does not exist in the PinMap
- **Guidance**: Please ensure the pin name is defined correctly in the PinMap. Add any missing pin definitions or update the pin name, then regenerate the output files.

### E_MissingPin_08
- **FullCode**: EBA455008
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 8
- **Level**: Error
- **#Args**: 1
- **Template**: Job: {0} does not define in FlowMain(Test Plan)/jobMapping(Setting file) sheet !!!
- **Guidance**: Please review the job configuration and ensure the required job is defined correctly. Add any missing job entries or update the job reference, then regenerate the output files.

### E_MissingPin_09
- **FullCode**: EBA455009
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 9
- **Level**: Error
- **#Args**: 2
- **Template**: Jobs '{0}' need to define in '{1}' sheet.
- **Guidance**: Please review the job configuration and ensure all required jobs are defined correctly. Add any missing job entries and regenerate the output files.

### E_MissingPin_10
- **FullCode**: EBA455010
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 10
- **Level**: Error
- **#Args**: 3
- **Template**: The {0} net pin {1} of {2} has existed in TestSetting !!!
- **Guidance**: Please review the net pin configuration and ensure each net pin is defined only once. Remove any duplicate definitions and regenerate the output files.

### E_MissingPin_11
- **FullCode**: EBA455011
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 11
- **Level**: Error
- **#Args**: 1
- **Template**: {0} can't be found in powerName list of <Judge_stored_IDS>!!!
- **Guidance**: Please verify that the required power name is included in the configured power name list. Add any missing power names and regenerate the output files.

### E_MissingPin_12
- **FullCode**: EBA455012
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 12
- **Level**: Error
- **#Args**: 1
- **Template**: The power name : {0} can't be found in the config setting.
- **Guidance**: Please review the power configuration and ensure all required power names are defined correctly. Add any missing power name definitions or update invalid references, then regenerate the output files.

### E_MissingPin_13
- **FullCode**: EBA455013
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 13
- **Level**: Error
- **#Args**: 1
- **Template**: {0} can't be found in flow !!!
- **Guidance**: Please review the flow configuration and ensure all required flow entries are defined correctly. Add any missing entries or update invalid references, then regenerate the output files.

### E_MissingPin_14
- **FullCode**: EBA455014
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 14
- **Level**: Error
- **#Args**: 1
- **Template**: {0} can't be found in flowCsharp !!!
- **Guidance**: Please review the FlowCsharp configuration and ensure all required entries are defined correctly. Add any missing entries or update invalid references, then regenerate the output files.

### E_MissingPin_15
- **FullCode**: WBA455015
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 15
- **Level**: Warning
- **#Args**: 1
- **Template**: Incorrect Value : {0} in column FS/DD
- **Guidance**: Please review the FS/DD configuration and ensure all values are specified correctly. Update any invalid entries and regenerate the output files.

### E_MissingPin_16
- **FullCode**: EBA455016
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 16
- **Level**: Error
- **#Args**: 1
- **Template**: Missing conti pin/pin group: {0} in pin map!
- **Guidance**: Verify that all required pin names are defined in the pin configuration. Check for typos or missing entries in the pin list.

### E_MissingPin_17
- **FullCode**: EBA455017
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 17
- **Level**: Error
- **#Args**: 1
- **Template**: Missing pin in voltage table: {0}
- **Guidance**: Please review the voltage table configuration and ensure all required pins are defined correctly. Add any missing pin definitions or update invalid references, then regenerate the output files.

### E_MissingPin_18
- **FullCode**: EBA455018
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 18
- **Level**: Error
- **#Args**: 1
- **Template**: Missing pin in PowerInfo for iFold: {0}
- **Guidance**: Please review the PowerInfo configuration and ensure all required pins are defined correctly for iFold settings. Add any missing pin definitions or update invalid references, then regenerate the output files.

### E_MissingPin_19
- **FullCode**: EBA455019
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 19
- **Level**: Error
- **#Args**: 1
- **Template**: The voltage of IO pin group {0} had different voltage.
- **Guidance**: Please review the IO pin group configuration and ensure all pins in the same group use consistent voltage settings. Correct any voltage inconsistencies and regenerate the output files.

### E_MissingPin_20
- **FullCode**: EBA455020
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 20
- **Level**: Error
- **#Args**: 3
- **Template**: Pin (group):{2} in Time Set file:{0};TSet:{1} is missing.
- **Guidance**: Please review the Time Set configuration and ensure all required pins and pin groups are defined correctly. Add any missing definitions or update invalid references, then regenerate the output files.

### E_MissingPin_21
- **FullCode**: WBA455021
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 21
- **Level**: Warning
- **#Args**: 2
- **Template**: Can't generate power level for pin/pin group: {0} from voltage table, if it just be used to crete variable in DC Specs, please ignore this or make sure {1} had been define this pin/pin group!!!
- **Guidance**: Verify that all required pin names are defined in the pin configuration. Check for typos or missing entries in the pin list.

### E_MissingPin_22
- **FullCode**: WBA455022
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 22
- **Level**: Warning
- **#Args**: 1
- **Template**: BallName : {0} Not exist in BumpName!!!
- **Guidance**: A business logic constraint has been violated. Review the rule requirements and correct the input data to satisfy the constraint.

### E_ContiIfold_01
- **FullCode**: EBA235001
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: IFold
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Ifold level {0} defined in PowerInfo is greater than Ifold limit for its instrument, please check!!!
- **Guidance**: Check the continuous ifold condition definition. Ensure the condition is valid and consistent with the test flow requirements.

### E_DuplicateItems_01
- **FullCode**: EBA155001
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Duplicate
- **EnumErrorTarget**: Pin
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: TN_Assignment sheet contains duplicate sheets : {0}
- **Guidance**: Please review the TN_Assignment configuration and ensure each sheet is defined only once. Remove any duplicate entries and regenerate the output files.

### E_DuplicateItems_02
- **FullCode**: EBA155002
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Duplicate
- **EnumErrorTarget**: Pin
- **Code**: 2
- **Level**: Error
- **#Args**: 2
- **Template**: IoInfo sheet contains duplicate pin level: {0} definition duplicate remove from {1}
- **Guidance**: Please review the IoInfo configuration and ensure each pin level is defined only once. Remove any duplicate pin level definitions and regenerate the output files.

### E_DuplicateItems_03
- **FullCode**: EBA155003
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Duplicate
- **EnumErrorTarget**: Pin
- **Code**: 3
- **Level**: Error
- **#Args**: 4
- **Template**: Single pin: {0} is duplicate in pin groups: {1} and {2}, please check the sheet: {3}
- **Guidance**: Please review the pin group configuration and ensure each single pin is assigned to only one pin group. Remove any duplicate pin assignments and regenerate the output files.

### E_MissingBlock_01
- **FullCode**: EBA407001
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Block
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: Can not find common block.
- **Guidance**: A common block cannot be found. Please ensure that the common block is defined correctly.

### E_MissingBlock_02
- **FullCode**: EBA407002
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Block
- **Code**: 2
- **Level**: Error
- **#Args**: 0
- **Template**: Can not find common block.
- **Guidance**: A common block cannot be found. Please ensure that the common block is defined correctly.

### E_MissingBlock_03
- **FullCode**: EBA407003
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Block
- **Code**: 3
- **Level**: Error
- **#Args**: 1
- **Template**: Omit column "{0}"
- **Guidance**: Please review the input file and ensure all required columns are included. Add any missing columns and regenerate the output files.

### E_MissingBlock_04
- **FullCode**: EBA407004
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Block
- **Code**: 4
- **Level**: Error
- **#Args**: 2
- **Template**: {0} in {1} block does not exist in common io pins.
- **Guidance**: Please review the common IO pin configuration and ensure all required pins are defined correctly. Add any missing pin definitions or update invalid references, then regenerate the output files.

### E_MissingBlock_05
- **FullCode**: EBA407005
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Block
- **Code**: 5
- **Level**: Error
- **#Args**: 0
- **Template**: Omit column "PowerSequence"
- **Guidance**: Column PowerSequence is missing. Please ensure that the PowerSequence column is included in the input file.

### E_MissingBlock_06
- **FullCode**: EBA407006
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Block
- **Code**: 6
- **Level**: Error
- **#Args**: 1
- **Template**: Pin '{0}' does not exist in TestSetting.
- **Guidance**: Please review the TestSetting configuration and ensure all required pins are defined correctly. Add any missing pin definitions or update invalid references, then regenerate the output files.

### E_MissingBlock_07
- **FullCode**: EBA407007
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Block
- **Code**: 7
- **Level**: Error
- **#Args**: 1
- **Template**: Sheet: {0} Not Exist.
- **Guidance**: Please review the input file and ensure all required sheets are included correctly. Add any missing sheets and regenerate

### E_FormatError_01
- **FullCode**: EBA229001
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: The TestPlan Error.
- **Guidance**: Check the field value against the expected format specification. Ensure the value type and range conform to requirements.

### E_FormatError_02
- **FullCode**: EBA229002
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 2
- **Level**: Error
- **#Args**: 0
- **Template**: The TestPlan Error.
- **Guidance**: Check the field value against the expected format specification. Ensure the value type and range conform to requirements.

### E_FormatError_03
- **FullCode**: EBA229003
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 3
- **Level**: Error
- **#Args**: 2
- **Template**: Logic Pin: {0} do not have the *_Valt voltage in {1}
- **Guidance**: Please review the voltage configuration and ensure all required *_Valt voltages are defined correctly. Add any missing voltage definitions or update invalid configurations, then regenerate the output files.

### E_FormatError_04
- **FullCode**: EBA229004
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 4
- **Level**: Error
- **#Args**: 2
- **Template**: Sram Pin: {0} do not have the *_Valt voltage in {1}
- **Guidance**: Please review the voltage configuration and ensure all required *_Valt voltages are defined correctly. Add any missing voltage definitions or update invalid configurations, then regenerate the output files.

### E_FormatError_05
- **FullCode**: EBA229005
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 5
- **Level**: Error
- **#Args**: 2
- **Template**: {0} and {1} should have the different source bit
- **Guidance**: Please review the source bit configuration and ensure each entry is assigned a unique source bit value. Update any conflicting assignments and regenerate the output files.

### E_FormatError_06
- **FullCode**: EBA229006
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 6
- **Level**: Error
- **#Args**: 0
- **Template**: The preserved pin should have the same source bit
- **Guidance**: The preserved pins have different source bits. Please ensure that all preserved pins use the same source bit.

### E_FormatError_07
- **FullCode**: EBA229007
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 7
- **Level**: Error
- **#Args**: 1
- **Template**: {0}: The length of Timing set name exceeds the maximum 31 chars 
- **Guidance**: Please review the timing set configuration and ensure all timing set names comply with the naming length requirements. Shorten any names that exceed the allowed length and regenerate the output files.

### E_FormatError_08
- **FullCode**: EBA229008
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 8
- **Level**: Error
- **#Args**: 1
- **Template**: {0}: The length of Flow sheet name exceeds the maximum 31 chars 
- **Guidance**: Please review the flow sheet configuration and ensure all sheet names comply with the naming length requirements. Shorten any sheet names that exceed the allowed length and regenerate the output files.

### E_FormatError_09
- **FullCode**: EBA229009
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 9
- **Level**: Error
- **#Args**: 1
- **Template**: {0}: The length of Inst sheet name exceeds the maximum 31 chars 
- **Guidance**: Please review the Inst sheet configuration and ensure all sheet names comply with the naming length requirements. Shorten any sheet names that exceed the allowed length and regenerate the output files.

### E_FormatError_10
- **FullCode**: EBA229010
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 10
- **Level**: Error
- **#Args**: 1
- **Template**: The directory is not exist : {0}
- **Guidance**: Please review the directory path configuration and ensure all required paths are valid and accessible. Create any missing directories or update invalid paths, then regenerate the output files.

### E_FormatError_11
- **FullCode**: EBA229011
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 11
- **Level**: Error
- **#Args**: 1
- **Template**: Different pin name from NV: {0}
- **Guidance**: Please review the pin configuration and ensure all pin names match the corresponding NV definitions. Update any inconsistent pin names and regenerate the output files.

### E_FormatError_12
- **FullCode**: EBA229012
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 12
- **Level**: Error
- **#Args**: 1
- **Template**: Wrong format ratio: {0}
- **Guidance**: Please review the ratio configuration and ensure all ratio values follow the required format. Correct any invalid ratio entries and regenerate the output files.

### E_FormatError_13
- **FullCode**: EBA229013
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 13
- **Level**: Error
- **#Args**: 1
- **Template**: Different pin name from NV: {0}
- **Guidance**: Please review the pin configuration and ensure all pin names match the corresponding NV definitions. Update any inconsistent pin names and regenerate the output files.

### E_FormatError_14
- **FullCode**: EBA229014
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 14
- **Level**: Error
- **#Args**: 2
- **Template**: Omit "{0}" value for pin: {1}
- **Guidance**: Please review the pin configuration and ensure all required values are specified correctly. Add any missing values and regenerate the output files.

### E_FormatError_15
- **FullCode**: EBA229015
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 15
- **Level**: Error
- **#Args**: 2
- **Template**: Omit value for pin: {0} in column {1}
- **Guidance**: Please review the pin configuration and ensure all required values are specified correctly. Add any missing values and regenerate the output files.

### E_FormatError_16
- **FullCode**: EBA229016
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 16
- **Level**: Error
- **#Args**: 2
- **Template**: Non numerical "{0}" value for pin: {1}
- **Guidance**: Please review the pin configuration and ensure all required values are specified using valid numeric formats. Correct any non-numeric values and regenerate the output files.

### E_FormatError_17
- **FullCode**: EBA229017
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 17
- **Level**: Error
- **#Args**: 2
- **Template**: Non numerical "{0}" value for pin: {1}
- **Guidance**: Please review the pin configuration and ensure all required values are specified using valid numeric formats. Correct any non-numeric values and regenerate the output files.

### E_FormatError_18
- **FullCode**: EBA229018
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 18
- **Level**: Error
- **#Args**: 1
- **Template**: Non numerical unknown type value for pin: {0}
- **Guidance**: Check the value and type settings for pin {0}. Ensure that a valid numeric value is provided and the type is recognized.

### E_FormatError_19
- **FullCode**: EBA229019
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 19
- **Level**: Error
- **#Args**: 4
- **Template**: Row:{3} The period [{0}] is illegal for tSet {1} on Pin {2}
- **Guidance**: Please review the timing configuration and ensure all period values are specified using valid formats and supported ranges. Correct any invalid period settings and regenerate the output files.

### E_FormatError_20
- **FullCode**: EBA229020
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 20
- **Level**: Error
- **#Args**: 2
- **Template**: Key Don't Exist In {0} : {1}
- **Guidance**: Please review the configuration and ensure all required keys are defined correctly. Add any missing key definitions or update invalid references, then regenerate the output files.

### E_FormatError_21
- **FullCode**: EBA229021
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 21
- **Level**: Error
- **#Args**: 1
- **Template**: Must Exist Item :{0} Not Exist.
- **Guidance**: Please review the configuration and ensure all required items are defined correctly. Add any missing item definitions and regenerate the output files.

### E_FormatError_22
- **FullCode**: EBA229022
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 22
- **Level**: Error
- **#Args**: 2
- **Template**: Multi-ShiftInFreq TimeSet Sheet problem @ {0} : {1}
- **Guidance**: Please review the Multi-ShiftInFreq TimeSet configuration and ensure all settings are defined correctly. Correct any invalid configurations and regenerate the output files.

### E_FormatError_23
- **FullCode**: EBA229023
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 23
- **Level**: Error
- **#Args**: 3
- **Template**: Can't parse {0} in {1}:Row{2}, it should be a integer.
- **Guidance**: Please review the configuration and ensure all required values are specified using valid integer formats. Correct any invalid values and regenerate the output files.

### E_FormatError_24
- **FullCode**: EBA229024
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 24
- **Level**: Error
- **#Args**: 3
- **Template**: Can't parse {0} in {1}:Row{2}, it should be a integer.
- **Guidance**: Please review the configuration and ensure all required values are specified using valid integer formats. Correct any invalid values and regenerate the output files.

### E_FormatError_25
- **FullCode**: EBA229025
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 25
- **Level**: Error
- **#Args**: 3
- **Template**: Can't parse {0} in {1}:Row{2}, it should be a integer.
- **Guidance**: Please review the configuration and ensure all required values are specified using valid integer formats. Correct any invalid values and regenerate the output files.

### E_FormatError_26
- **FullCode**: EBA229026
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 26
- **Level**: Error
- **#Args**: 4
- **Template**: {0}:{1} in {2}:Row{3}, should be in the range of 1-9999.
- **Guidance**: Please review the configuration and ensure all required values are within the supported range. Correct any out-of-range values and regenerate the output files.

### E_FormatError_27
- **FullCode**: EBA229027
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 27
- **Level**: Error
- **#Args**: 4
- **Template**: {0}:{1} in {2}:Row{3}, should be in the range of 1-9999.
- **Guidance**: Please review the configuration and ensure all required values are within the supported range. Correct any out-of-range values and regenerate the output files.

### E_FormatError_28
- **FullCode**: EBA229028
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 28
- **Level**: Error
- **#Args**: 9
- **Template**: Can't get bin number sheet in test plan, missing {0} header, it should have {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7} ,{8}
- **Guidance**: Please review the bin number sheet configuration and ensure all required headers are defined correctly. Add any missing headers and regenerate the output files.

### E_FormatError_29
- **FullCode**: EBA229029
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 29
- **Level**: Error
- **#Args**: 1
- **Template**: Can't get bin number sheet in test plan, missing {0} header.
- **Guidance**: Please review the bin number sheet configuration and ensure all required headers are defined correctly. Add any missing headers and regenerate the output files.

### E_FormatWarning_01
- **FullCode**: EBA329001
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Format
- **Code**: 1
- **Level**: Error
- **#Args**: 2
- **Template**: VM Vector name is mismatch, Pattern:{0}, VmVector:{1} 
- **Guidance**: Check the pattern vector name and VM vector name configuration. Ensure that the VM vector name matches the corresponding pattern vector name.

### E_FormatWarning_02
- **FullCode**: EBA329002
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Format
- **Code**: 2
- **Level**: Error
- **#Args**: 2
- **Template**: VM Vector name is mismatch, Pattern:{0}, VmVector:{1} 
- **Guidance**: Check the pattern vector name and VM vector name configuration. Ensure that the VM vector name matches the corresponding pattern vector name.

### E_FormatWarning_03
- **FullCode**: EBA437003
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Instance
- **Code**: 3
- **Level**: Error
- **#Args**: 1
- **Template**: The ELB instance {0} can not be found !!!
- **Guidance**: Review the flagged value for potential formatting inconsistencies. This warning does not block execution but may cause unexpected behavior.

### E_FormatWarning_04
- **FullCode**: EBA437004
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Instance
- **Code**: 4
- **Level**: Error
- **#Args**: 1
- **Template**: The ELB instance {0} can not be found !!!
- **Guidance**: Review the flagged value for potential formatting inconsistencies. This warning does not block execution but may cause unexpected behavior.

### E_FormatWarning_05
- **FullCode**: EBA142005
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Duplicate
- **EnumErrorTarget**: Limit
- **Code**: 5
- **Level**: Error
- **#Args**: 1
- **Template**: The limit of ELB instance are duplicated -{0} !!!
- **Guidance**: Review the flagged value for potential formatting inconsistencies. This warning does not block execution but may cause unexpected behavior.

### E_FormatWarning_06
- **FullCode**: EBA329006
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Format
- **Code**: 6
- **Level**: Error
- **#Args**: 2
- **Template**: The pin {0} can not be found in sheet {1}
- **Guidance**: Please review the pin configuration and ensure all required pins are defined correctly. Add any missing pin definitions or update invalid references, then regenerate the output files.

### E_FormatWarning_07
- **FullCode**: EBA329007
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Format
- **Code**: 7
- **Level**: Error
- **#Args**: 5
- **Template**: The I/O pin {0} in {1} can not be found in {2} ({3}, {4})
- **Guidance**: Please review the I/O pin configuration and ensure all required pin definitions are defined correctly. Add any missing pin definitions or update invalid references, then regenerate the output files.

### E_FormatWarning_08
- **FullCode**: EBA329008
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Format
- **Code**: 8
- **Level**: Error
- **#Args**: 5
- **Template**: The I/O pin {0} in {1} can not be found in {2} ({3}, {4})
- **Guidance**: Please review the I/O pin configuration and ensure all required pin definitions are defined correctly. Add any missing pin definitions or update invalid references, then regenerate the output files.

### E_FormatWarning_09
- **FullCode**: EBA329009
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Format
- **Code**: 9
- **Level**: Error
- **#Args**: 5
- **Template**: The I/O pin {0} in {1} can not be found in {2} ({3}, {4})
- **Guidance**: Please review the I/O pin configuration and ensure all required pin definitions are defined correctly. Add any missing pin definitions or update invalid references, then regenerate the output files.

### E_FormatWarning_10
- **FullCode**: EBA329010
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Format
- **Code**: 10
- **Level**: Error
- **#Args**: 3
- **Template**: The pin {0} in {1} can not be found in {2}
- **Guidance**: Please review the pin configuration and ensure all required pins are defined correctly. Add any missing pin definitions or update invalid references, then regenerate the output files.

### E_FormatWarning_11
- **FullCode**: EBA329011
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Format
- **Code**: 11
- **Level**: Error
- **#Args**: 2
- **Template**: The pin {0} can not be found in sheet {1}
- **Guidance**: Please review the pin configuration and ensure all required pins are defined correctly. Add any missing pin definitions or update invalid references, then regenerate the output files.

### E_FormatWarning_12
- **FullCode**: EBA329012
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Format
- **Code**: 12
- **Level**: Error
- **#Args**: 2
- **Template**: The pin {0} can not be found in sheet {1}
- **Guidance**: Please review the pin configuration and ensure all required pins are defined correctly. Add any missing pin definitions or update invalid references, then regenerate the output files.

### E_FSDDIssue_01
- **FullCode**: EBA211001
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Column
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: Needs to fill FS or DD in FS/DD Column !!!
- **Guidance**: Check the FSDD (Functional Signal Drive/Detect) configuration for the specified component. Verify that all FSDD settings are valid and consistent with the test requirements.

### E_MissingParameter_01
- **FullCode**: EBA482001
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Value
- **Code**: 1
- **Level**: Error
- **#Args**: 3
- **Template**: Missing Parameter in {0}({1}) : {2}
- **Guidance**: Open the configuration sheet and add the required parameter. Check the specification to confirm which parameters are mandatory.

### E_MissVbtModule_01
- **FullCode**: EBA447001
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Module
- **Code**: 1
- **Level**: Error
- **#Args**: 2
- **Template**: The VBT function: {0} can not find in {1} library!
- **Guidance**: Please review the VBT library configuration and ensure all required functions are defined correctly. Add any missing function definitions or update invalid references, then regenerate the output files.

### W_FormatError_01
- **FullCode**: WBA229001
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 1
- **Level**: Warning
- **#Args**: 1
- **Template**: Line:{0} format is empty line
- **Guidance**: Please review the input file and ensure all required entries are provided correctly. Remove any unexpected empty lines or add the required content, then regenerate the output files.

### W_NwireConfigMismatch_01
- **FullCode**: WBA448001
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Nwire
- **Code**: 1
- **Level**: Warning
- **#Args**: 1
- **Template**: Please check {0} !!! Pin OutClkVoltage value is 0
- **Guidance**: Review the N-wire output clock voltage setting for the specified instance. Verify the voltage level is within the acceptable range for the target device.

### E_DefaultPowerUpInstanceExist_01
- **FullCode**: WBA437001
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Instance
- **Code**: 1
- **Level**: Warning
- **#Args**: 0
- **Template**: PowerUp instance exist in UF_Instance sheet, Autogen will not generate default power up instance "PowerUp", please make sure to put this instance in Flow_Main
- **Guidance**: Check the UF instance list for an existing default power-up instance. Only one default power-up instance is permitted; remove the duplicate or rename it.

### E_TimeSetOverride_01
- **FullCode**: EBA233001
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Header
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Invalid `TimeSetClockOverride` sheet header. The following columns are required: {0}
- **Guidance**: One or more required columns are missing from the `TimeSetClockOverride` sheet header. Verify that all required columns are present and correctly named.

### E_TimeSetOverride_02
- **FullCode**: EBA211002
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Column
- **Code**: 2
- **Level**: Error
- **#Args**: 2
- **Template**: Columns {0} are required at row {1} in TimeSetClockOverride sheet
- **Guidance**: Make sure required columns are not empty on every non-empty rows.

### E_TimeSetOverride_03
- **FullCode**: EBA201003
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Argument
- **Code**: 3
- **Level**: Error
- **#Args**: 1
- **Template**: Frequency variable name [{0}] is invalid in TimeSetClockOverride sheet
- **Guidance**: Variable names must start with a letter and may only contain letters, numbers, and underscores

### E_TimeSetOverride_04
- **FullCode**: EBA229030
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 30
- **Level**: Error
- **#Args**: 1
- **Template**: Incorrect Frequency format [{0}]
- **Guidance**: Make sure Frequency format looks like this "SOME_VAR" or "SOME_VAR = VALUE"

### E_TimeSetOverride_05
- **FullCode**: EBA201005
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Argument
- **Code**: 5
- **Level**: Error
- **#Args**: 1
- **Template**: Frequency value [{0}] is invalid in TimeSetClockOverride sheet
- **Guidance**: Frequency value must be a number with optional frequency unit MHz or KHz

### E_TimeSetOverride_06
- **FullCode**: EBA201006
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Argument
- **Code**: 6
- **Level**: Error
- **#Args**: 2
- **Template**: Frequency variable `{0}` is assigned multiple conflicting values ({1}) in TimeSetClockOverride sheet
- **Guidance**: Ensure that each variable is assigned a single consistent value throughout the sheet. If the variable is defined multiple times, all assignments must use the same value.

### E_TimeSetOverride_07
- **FullCode**: EBA201007
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Argument
- **Code**: 7
- **Level**: Error
- **#Args**: 0
- **Template**: Pin/Group Setup does not allow empty cell in TimeSetClockOverride sheet
- **Guidance**: Pin/Group Setup cannot be empty. Only support `io`, `clock` and `clock_2x`

### E_TimeSetOverride_08
- **FullCode**: EBA201008
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Argument
- **Code**: 8
- **Level**: Error
- **#Args**: 1
- **Template**: Pin/Group Setup value {0} is invalid in TimeSetClockOverride sheet
- **Guidance**: Pin/Group Setup only support `io`, `clock` and `clock_2x`

### E_TimeSetOverride_09
- **FullCode**: EBA201009
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Argument
- **Code**: 9
- **Level**: Error
- **#Args**: 1
- **Template**: Data Src value {0} is invalid
- **Guidance**: Data Src only support `ALLHI`, `ALLLO`, `PA`, `PAT`, `PATHI`, `PATLO`, `PATNOT`

### E_TimeSetOverride_10
- **FullCode**: EBA201010
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Argument
- **Code**: 10
- **Level**: Error
- **#Args**: 1
- **Template**: Data Fmt value {0} is invalid
- **Guidance**: Data Fmt only support `NR`, `RH`, `RL`, `STAY`. And suffix `-2X` is only allowed when Setup is `clock_2x`

### E_TimeSetOverride_11
- **FullCode**: EBA201011
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Argument
- **Code**: 11
- **Level**: Error
- **#Args**: 0
- **Template**: Data Src ALLHI must return low `RL` in TimeSetClockOverride sheet
- **Guidance**: Make sure TimeSetClockOverride sheet has correct Data Src and Fmt

### E_TimeSetOverride_12
- **FullCode**: EBA201012
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Argument
- **Code**: 12
- **Level**: Error
- **#Args**: 0
- **Template**: Data Src ALLLO must return high `RH` in TimeSetClockOverride sheet
- **Guidance**: Make sure TimeSetClockOverride sheet has correct Data Src and Fmt

### E_TimeSetOverride_13
- **FullCode**: EBA425013
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: File
- **Code**: 13
- **Level**: Error
- **#Args**: 1
- **Template**: Cannot override TimeSet `{0}` because it is not used in the Pattern Dashboard.
- **Guidance**: Make sure TimeSetClockOverride sheet only override TimeSets inside Pattern Dashboard

### E_TimeSetOverride_14
- **FullCode**: EBA161014
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Duplicate
- **EnumErrorTarget**: Row
- **Code**: 14
- **Level**: Error
- **#Args**: 3
- **Template**: TimeSet: `{0}` Pin/Group: `{1}` is already defined at row {2}.
- **Guidance**: Make sure TimeSetClockOverride sheet has no duplicate `TimeSet` and `Pin/Group` pair under the same `TimeSet File`

### E_TimeSetOverride_15
- **FullCode**: EBA476015
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Timing
- **Code**: 15
- **Level**: Error
- **#Args**: 3
- **Template**: Cannot find `Pin/Group` `{1}` under `TimeSet` `{0}` in TimeSet File `{2}`.
- **Guidance**: Verify that the specified `TimeSet` and `Pin/Group` exist in the target TimeSet File and match the values defined in the `TimeSetClockOverride` sheet.

### E_TimeSetOverride_16
- **FullCode**: EBA276016
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Timing
- **Code**: 16
- **Level**: Error
- **#Args**: 3
- **Template**: The `Pin/Group Clock Period` formula: `{0}` at row {1} of TimeSet File `{2}` does not contain any variable.
- **Guidance**: This might be the TimeSet File error. Specify exactly one variable in the formula. For example: `=(1/_TCK_Freq_Var)`.

### E_TimeSetOverride_17
- **FullCode**: EBA276017
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Timing
- **Code**: 17
- **Level**: Error
- **#Args**: 3
- **Template**: Multiple variables were found in the `Pin/Group Clock Period` formula: `{0}` at row {1} of TimeSet File `{2}`
- **Guidance**: Ensure the formula contains exactly one variable reference.

### E_TimeSetOverride_18
- **FullCode**: EBA201018
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Argument
- **Code**: 18
- **Level**: Error
- **#Args**: 3
- **Template**: An unexpected error occurred while applying invalid Pin/Group Setup `{0}` to TimeSet File `{1}` at row {2}. This value should have been rejected during validation.
- **Guidance**: This error indicates an internal consistency issue in Autogen. Please contact the HardIp owner or support team and provide the generated error report for further investigation.

### E_TimeSetOverride_19
- **FullCode**: EBA201019
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Argument
- **Code**: 19
- **Level**: Error
- **#Args**: 2
- **Template**: While modifying TimeSet File: `{0}` at row {1}, Data Src is set to `ALLHI` but Data Fmt did not set to `RL`
- **Guidance**: Verify the final Pin/Group configuration in both the TimeSetClockOverride sheet and the target TimeSet File. Ensure that any Pin/Group using Data Src `ALLHI` has Data Fmt set to `RL`.

### E_TimeSetOverride_20
- **FullCode**: EBA201020
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Argument
- **Code**: 20
- **Level**: Error
- **#Args**: 2
- **Template**: While modifying TimeSet File: `{0}` at row {1}, Data Src is set to `ALLLO` but Data Fmt did not set to `RH`
- **Guidance**: Verify the final Pin/Group configuration in both the TimeSetClockOverride sheet and the target TimeSet File. Ensure that any Pin/Group using Data Src `ALLLO` has Data Fmt set to `RH`.

### E_TimeSetOverride_21
- **FullCode**: EBA476021
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Timing
- **Code**: 21
- **Level**: Error
- **#Args**: 2
- **Template**: Cannot find `TimeSet` `{0}` in TimeSet File `{1}`.
- **Guidance**: Verify that the specified `TimeSet` exist in the target TimeSet File and match the values defined in the `TimeSetClockOverride` sheet.

### E_TimeSetOverride_22
- **FullCode**: EBA201022
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Argument
- **Code**: 22
- **Level**: Error
- **#Args**: 1
- **Template**: Brand new Frequency variable {0} must assign a value
- **Guidance**: If a variable does not exist in TimeSet VAR Definition and is assigned in Frequency cell, you must assign a value to it

### E_InvalidTiming_01
- **FullCode**: EBA276001
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Timing
- **Code**: 1
- **Level**: Error
- **#Args**: 2
- **Template**: Parsing error in TimeSet: {0} Row: {1}
- **Guidance**: Review the TimeSet sheet for structural issues and correct the flagged entry.

### E_InvalidTiming_02
- **FullCode**: EBA276002
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Timing
- **Code**: 2
- **Level**: Error
- **#Args**: 2
- **Template**: Equation base variable '{0}' used in Time Set file {1} is not assigned an initial value
- **Guidance**: Review the TimeSet sheet for structural issues and correct the flagged entry.

### E_InvalidTiming_03
- **FullCode**: EBA276003
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Timing
- **Code**: 3
- **Level**: Error
- **#Args**: 2
- **Template**: Equation base variable '{0}' used in the context of Time Set file {1} is not assigned value in comment
- **Guidance**: Review the TimeSet sheet for structural issues and correct the flagged entry.

## BinCutErrorType (`BinCutErrorType.cs`)

### E_AllowEqual_01
- **FullCode**: EBC654001
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: PerformanceMode
- **Code**: 1
- **Level**: Error
- **#Args**: 2
- **Template**: The format of Allow Equal " {0} " was worng (Should be {1})!!!
- **Guidance**: Check the AllowEqual mode column in the binning table. AllowEqual mode can only be the previous performance mode

### E_AllowEqual_02
- **FullCode**: EBC654002
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: PerformanceMode
- **Code**: 2
- **Level**: Error
- **#Args**: 0
- **Template**: The Allow Equal should not be empty !!!
- **Guidance**: Fill in the Allow Equal column for the flagged mode row in the binning table. Allow Equal column should have same content for same performance mode

### E_AllowEqual_03
- **FullCode**: EBC654003
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: PerformanceMode
- **Code**: 3
- **Level**: Error
- **#Args**: 0
- **Template**: The Comment should not be empty !!!
- **Guidance**: Fill in the Comment column for the flagged mode row in the binning table. Comment column should have same content for same performance mode if it is start with "Max PV"

### E_AllowEqual_04
- **FullCode**: EBC654004
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: PerformanceMode
- **Code**: 4
- **Level**: Error
- **#Args**: 1
- **Template**: The Comment should contain mode {0} !!!
- **Guidance**: Update the Comment value to include all modes from the allow-equal group.

### E_BistInitPattern_01
- **FullCode**: EBC652005
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Pattern
- **Code**: 5
- **Level**: Error
- **#Args**: 2
- **Template**: Row {0} and row {1} of the init patterns are the same!!!
- **Guidance**: Check that the BIST initialization pattern, they need to be different between two neighbor rows 

### E_Business_01
- **FullCode**: EBC634006
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Ids
- **Code**: 6
- **Level**: Error
- **#Args**: 5
- **Template**: The Ids {0} of {1} was larger than the spec {2} in {3} @ Bin {4}
- **Guidance**: Ids value is larger the maximum Ids value. 

### E_Business_02
- **FullCode**: EBC421007
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Domain
- **Code**: 7
- **Level**: Error
- **#Args**: 1
- **Template**: The domain : {0} in Other Rails can't be found.
- **Guidance**: Check if the domain for other rail is defined in Binning table

### E_C_01
- **FullCode**: EBC609008
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: C
- **Code**: 8
- **Level**: Error
- **#Args**: 3
- **Template**: C value {0} is different from {1} in {2} when setting allow equal mode!!!
- **Guidance**: Check the BinCut C (constant) value definition in Binning table, should be the same when setting to allow equal mode 

### E_Cp2Gb_01
- **FullCode**: EBC631009
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: GB
- **Code**: 9
- **Level**: Error
- **#Args**: 3
- **Template**: CP2GB value {0} is different from {1} in {2} when setting allow equal mode!!!
- **Guidance**: Check the Cp2Gb (CP2 guardband) value in the Binning table, should be the same with its allow equal mode when setting allow equal mode. 

### E_Cpgb_01
- **FullCode**: EBC631010
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: GB
- **Code**: 10
- **Level**: Error
- **#Args**: 3
- **Template**: CPGB value {0} is different from {1} in {2} when setting allow equal mode!!!
- **Guidance**: Check the Cpgb (CP guardband) value in the Binning table, should be the same with its allow equal mode when setting allow equal mode. 

### E_CpMax_01
- **FullCode**: EBC615011
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Cpvmax
- **Code**: 11
- **Level**: Error
- **#Args**: 6
- **Template**: The cpVMax {0} + {1} - {2} is larger than the limit in efuseBitDef {3} - ( 2 ^ {4} - 1 ) * {5} !!!
- **Guidance**: Check the CP maximum value against the specification. Ensure the value does not exceed the allowed limit for this bin.

### E_Domain_01
- **FullCode**: WBC421012
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Domain
- **Code**: 12
- **Level**: Warning
- **#Args**: 1
- **Template**: The domain {0} is not existed !!!
- **Guidance**: Check the domain name in the Binning table against the supported domain list. Update the configuration to use a valid domain name.

### E_Domain_02
- **FullCode**: EBC321013
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Domain
- **Code**: 13
- **Level**: Error
- **#Args**: 3
- **Template**: The domain {0} of {1} is incorrect (Should be {2}) !!!
- **Guidance**: Open the BinCut binning table and correct the domain assignment for the flagged row. All rows with the same performance mode must reference the same domain.

### E_efuseBitDef_01
- **FullCode**: WBC424014
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Field
- **Code**: 14
- **Level**: Warning
- **#Args**: 2
- **Template**: The IDS_VDD_{0} can not be found in efuseBitDef, and only found relative name {1} !!!
- **Guidance**: Verify the eFuse bit definition name against IDS name. Check for typos or missing bit definitions.

### E_efuseBitDef_02
- **FullCode**: EBC424015
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Field
- **Code**: 15
- **Level**: Error
- **#Args**: 1
- **Template**: The IDS_VDD_{0} can not be found in efuseBitDef !!!
- **Guidance**: Verify the eFuse bit definition name against IDS name. Check for typos or missing bit definitions.

### E_Equation_01
- **FullCode**: EBC323016
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Equation
- **Code**: 16
- **Level**: Error
- **#Args**: 2
- **Template**: {0} and {1} have different number of equation
- **Guidance**: Check the equation count between flagged mode and its allow equal mode in Binning table. Ensure all referenced variables are defined and the expression is well-formed.

### E_Equation_02
- **FullCode**: EBC423017
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Equation
- **Code**: 17
- **Level**: Error
- **#Args**: 2
- **Template**: Equation {0} not be found in {1} when setting allow equal mode!!!
- **Guidance**: Open the BinCut binning table and verify that the flagged equation exists in the allow-equal mode rows. 

### E_Flow_01
- **FullCode**: EBC355018
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Pin
- **Code**: 18
- **Level**: Error
- **#Args**: 3
- **Template**: Please check core power pin in different jobs {0} {1} vs {2} !!!
- **Guidance**: Compare the core power pin definitions across all jobs in the BinCut flow sheet. Ensure the pin names and order are consistent between all job definitions.

### E_Flow_02
- **FullCode**: EBC355019
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Pin
- **Code**: 19
- **Level**: Error
- **#Args**: 2
- **Template**: The core power count in {0} does not match with {1} !!!
- **Guidance**: Open the BinCut flow sheet and verify that all jobs define the same number of core power pins. Align the pin counts across all job definitions.

### E_Flow_03
- **FullCode**: EBC454020
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: PerformanceMode
- **Code**: 20
- **Level**: Error
- **#Args**: 1
- **Template**: The mode {0} in flow can not be found in the column performance mode !!!
- **Guidance**: Open the BinCut flow sheet and verify that the flagged mode appears in the performance mode column. Add the missing mode or correct the reference.

### E_FlowName_01
- **FullCode**: EBC427021
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Flow
- **Code**: 21
- **Level**: Error
- **#Args**: 1
- **Template**: The flowName "{0}" is missing in the all BinCut_Instance sheet !!!
- **Guidance**: Verify the flow name against the defined flow list. Check for typos or missing flow definitions.

### E_FlowName_02
- **FullCode**: EBC652022
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Pattern
- **Code**: 22
- **Level**: Error
- **#Args**: 0
- **Template**: The first init can not be empty in one TTR instance !!!
- **Guidance**: Open the BinCut instance sheet and assign a valid init pattern to the first row of the TTR instance group. The first init pattern is required for TTR instance initialization.

### E_FlowName_03
- **FullCode**: EBC229023
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 23
- **Level**: Error
- **#Args**: 1
- **Template**: The flowName "{0}" with extra spaces !!!
- **Guidance**: Open the BinCut instance sheet and remove the extra spaces from the flagged flow name. Flow names must not contain leading, trailing, or embedded spaces.

### E_FormatError_01
- **FullCode**: EBC229024
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 24
- **Level**: Error
- **#Args**: 1
- **Template**: The subflow {0} had not start with Performance Mode !!!
- **Guidance**: Open the BinCut sheet and inspect the flagged field for formatting issues. Ensure all values conform to the expected format and data type.

### E_FormatError_02
- **FullCode**: EBC429025
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Format
- **Code**: 25
- **Level**: Error
- **#Args**: 4
- **Template**: Can not find voltage condition in flow sheet (Job: {0}, TableType: {1} ,TableBinType: {2}, mode: {3})!!!
- **Guidance**: Open the BinCut flow sheet and verify the voltage condition entry exists for the specified job, table type, and performance mode. Add the missing row or correct the reference.

### E_FormatError_03
- **FullCode**: EBC454026
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: PerformanceMode
- **Code**: 26
- **Level**: Error
- **#Args**: 1
- **Template**: The performance mode {0} can not be found in flow sheet!
- **Guidance**: Open the BinCut order sheet and inspect the flagged row for formatting issues. Ensure the performance mode and related fields conform to the expected format.

### E_FormatError_04
- **FullCode**: EBC229027
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 27
- **Level**: Error
- **#Args**: 0
- **Template**: The columns TD, BIST or FUNC can not be null or empty!
- **Guidance**: Open the BinCut order sheet and inspect the flagged row for formatting issues. Verify all required columns are populated with valid values.

### E_FormatError_05
- **FullCode**: EBC429028
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Format
- **Code**: 28
- **Level**: Error
- **#Args**: 2
- **Template**: The missing column {0} in sheet {1} !!!
- **Guidance**: Compare the header and row definitions across all binning table sheets. Ensure all sheets use consistent column names and value formats.

### E_FtRoomGb_01
- **FullCode**: EBC631029
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: GB
- **Code**: 29
- **Level**: Error
- **#Args**: 3
- **Template**: FT1GB value {0} is different from {1} in {2} when setting allow equal mode!!!
- **Guidance**: Open the BinCut binning table and compare the FT1GB value for the flagged row against the allow-equal mode row. The FT1 GB values must be identical when allow-equal mode is set.

### E_FtHotGb_01
- **FullCode**: EBC631030
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: GB
- **Code**: 30
- **Level**: Error
- **#Args**: 3
- **Template**: FT2GB value {0} is different from {1} in {2} when setting allow equal mode!!!
- **Guidance**: Open the BinCut binning table and compare the FT2GB value for the flagged row against the allow-equal mode row. The FT2 GB values must be identical when allow-equal mode is set.

### E_HTOLGb_01
- **FullCode**: EBC631031
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: GB
- **Code**: 31
- **Level**: Error
- **#Args**: 2
- **Template**: The GB of HTOL {0} is bigger than FT {1}
- **Guidance**: Check the HTOLGb (HTOL guardband) value in the BinCut definition. Ensure it is positive and within the specification limits for HTOL testing.

### E_ID_01
- **FullCode**: EBC429032
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Format
- **Code**: 32
- **Level**: Error
- **#Args**: 1
- **Template**: The ID column does not exist in sheet {0}, please add it!!!
- **Guidance**: Check if ID column is existed in Bincut binning table.

### E_ID_02
- **FullCode**: EBC229033
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 33
- **Level**: Error
- **#Args**: 0
- **Template**: The ID value invalid
- **Guidance**: Open the BinCut binning table and verify the ID value for the first row of the flagged performance mode group. The ID must be a valid numeric value.

### E_ID_04
- **FullCode**: EBC629034
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Format
- **Code**: 34
- **Level**: Error
- **#Args**: 4
- **Template**: The larger equation {0} ({1}) needs to set a larger ID value than {2} ({3})
- **Guidance**: Open the BinCut binning table and correct the ID values so that larger equations have larger IDs. The ID values must be monotonically increasing with equation number.

### E_ID_05
- **FullCode**: EBC629035
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Format
- **Code**: 35
- **Level**: Error
- **#Args**: 4
- **Template**: The smaller equation {0} ({1}) needs to set a smaller ID value than {2} ({3})
- **Guidance**: Open the BinCut binning table and correct the ID values so that smaller equations have smaller IDs. The ID values must be monotonically decreasing with equation number.

### E_ID_06
- **FullCode**: EBC629036
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Format
- **Code**: 36
- **Level**: Error
- **#Args**: 6
- **Template**: The {0}(ID:{1}) of the performance mode {2} needs to be set to be greater than the {3}(ID:{4}) of the {5}
- **Guidance**: Open the BinCut binning table and ensure that the last equation's ID in each performance mode is greater than in the preceding mode. Reorder or reassign ID values to satisfy this constraint.

### E_ID_07
- **FullCode**: EBC129037
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Duplicate
- **EnumErrorTarget**: Format
- **Code**: 37
- **Level**: Error
- **#Args**: 1
- **Template**: The ID value is duplicated in the {0} domain, please change!!!
- **Guidance**: Open the BinCut binning table and assign unique ID values within the flagged domain. No two rows within the same domain may share the same ID value.

### E_IdsMax_01
- **FullCode**: EBC634038
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Ids
- **Code**: 38
- **Level**: Error
- **#Args**: 3
- **Template**: The IdsMax {0} is larger than the limit in efuseBitDef {1} - ( 2 ^ {2} - 1 )  !!!
- **Guidance**: Check the IDS maximum value against the device specification. Ensure the value is within the allowable range for this bin category.

### E_InstanceName_01
- **FullCode**: WBC629039
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Format
- **Code**: 39
- **Level**: Warning
- **#Args**: 0
- **Template**: Instance name too long(over 150), please modify instance name by PatSetName(Orange)!!!
- **Guidance**: Instance name generated by autogen wil be too long and cause error in run time, use PatSetName(Orange) in test plan to shorten it. 

### E_Interpolation_01
- **FullCode**: EBC638040
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Interpolation
- **Code**: 40
- **Level**: Error
- **#Args**: 1
- **Template**: The value of performance {0} is different with other equations!!!
- **Guidance**: Ensure all equations for the same mode have the same Interpolation ModeH, ModeL, Offset and SkipTest value.

### E_Interpolation_03
- **FullCode**: EBC238041
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Interpolation
- **Code**: 41
- **Level**: Error
- **#Args**: 2
- **Template**: The value of Int_Mode_L/Int_Mode_H {0} is not a legal performance mode for {1}!!!
- **Guidance**: Ensure Int_Mode_L/Int_Mode_H values are performance modes.

### E_JobMisMatch_01
- **FullCode**: WBC339042
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Job
- **Code**: 42
- **Level**: Warning
- **#Args**: 4
- **Template**: Jobs in the flow sheet (Performance Mode / FlowName : {0} / {1}) is {2} , but jobs in BinCut_instance is {3} !!!
- **Guidance**: Compare the job assignment in the BinCut sheet against the expected job definition. Ensure the job name is consistent between all references.

### E_M_01
- **FullCode**: EBC645043
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: M
- **Code**: 43
- **Level**: Error
- **#Args**: 3
- **Template**: M value {0} is different from {1} in {2} when setting allow equal mode!!!
- **Guidance**: Check the BinCut M (multiplier) value definition. Ensure the value is within the allowed range and correctly specified.

### E_Missing_01
- **FullCode**: EBC527044
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Redundant
- **EnumErrorTarget**: Flow
- **Code**: 44
- **Level**: Error
- **#Args**: 2
- **Template**: The instance row {0} not be used in the flow sheet ({1})!!!
- **Guidance**: Identify the missing item in the BinCut configuration and add it. Check the BinCut specification to confirm which items are required.

### E_Missing_02
- **FullCode**: EBC427045
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Flow
- **Code**: 45
- **Level**: Error
- **#Args**: 2
- **Template**: The IPL/EPL instance {0} does not have the corresponding sub flow {1} in the flow_vddbinning sheet!!!
- **Guidance**: Identify the missing sub flows in the flow_vddbinning sheet. Check the Binning sheet to confirm the performance items are not required to generate IPL instances.

### E_MissingParameter_01
- **FullCode**: EBC451046
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Parameter
- **Code**: 46
- **Level**: Error
- **#Args**: 3
- **Template**: Missing Parameter in {0}({1}) : {2}
- **Guidance**: Missing argument in C# library function, check C# library version.

### E_MissVbtModule_01
- **FullCode**: EBC447047
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Module
- **Code**: 47
- **Level**: Error
- **#Args**: 2
- **Template**: The VBT function: {0} can not find in {1} library!
- **Guidance**: Missing function in C# library, check C# library version. 

### E_Pattern_01
- **FullCode**: EBC252048
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Pattern
- **Code**: 48
- **Level**: Error
- **#Args**: 1
- **Template**: The pattern : {0} format illegal
- **Guidance**: Verify the pattern name follows the expected naming convention (minimum 11 underscore-separated segments). Check that the pattern exists and is correctly referenced in the BinCut sheet.

### E_Pattern_02
- **FullCode**: EBC623049
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Equation
- **Code**: 49
- **Level**: Error
- **#Args**: 2
- **Template**: The eqn start {0} can not exceed the max eqn {1} !!!
- **Guidance**: Open the BinCut IDS distribution configuration and reduce the equation start value. The start value must not exceed the maximum equation number defined for this pattern.

### E_Pattern_03
- **FullCode**: EBC252050
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Pattern
- **Code**: 50
- **Level**: Error
- **#Args**: 1
- **Template**: The type of patten {0} is unknown !!!
- **Guidance**: Check the pattern type classification in the instance sheet reader. Ensure the pattern name follows the expected naming convention for RTOS pattern types.

### E_Pattern_04
- **FullCode**: EBC252051
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Pattern
- **Code**: 51
- **Level**: Error
- **#Args**: 1
- **Template**: The type of patten {0} is unknown !!!
- **Guidance**: Check the pattern type classification in the secondary parsing path. Ensure the pattern name follows the expected naming convention for all supported pattern types.

### E_Pattern_05
- **FullCode**: EBC252052
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Pattern
- **Code**: 52
- **Level**: Error
- **#Args**: 1
- **Template**: An empty string exists in the pattern: {0} !!!
- **Guidance**: Open the BinCut instance sheet and remove the empty string from the pattern cell. Pattern cells must not contain whitespace-only values.

### E_Pattern_06
- **FullCode**: EBC252053
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Pattern
- **Code**: 53
- **Level**: Error
- **#Args**: 0
- **Template**: The fisrt pattern can not be empty !!!
- **Guidance**: Open the BinCut instance sheet and provide a valid pattern name in the first pattern column. The first pattern in every row is required and must not be empty.

### E_PatternTimeSet_01
- **FullCode**: EBC452054
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pattern
- **Code**: 54
- **Level**: Error
- **#Args**: 0
- **Template**: This test can not find any time set by pattern dashboard !!!
- **Guidance**: Verify the time set assigned to the BinCut pattern against the defined time set list. Ensure the time set is compatible with the pattern's requirements.

### E_PatternTimeSet_02
- **FullCode**: EBC252055
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Pattern
- **Code**: 55
- **Level**: Error
- **#Args**: 1
- **Template**: There are multi time set {0} by patterns !!!
- **Guidance**: Multiple patterns reference different time sets and no override is provided. Set the TimeSet column in the BinCut instance sheet to explicitly select the intended time set.

### E_PatternTimeSet_03
- **FullCode**: EBC252056
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Pattern
- **Code**: 56
- **Level**: Error
- **#Args**: 1
- **Template**: There are multi time set {0} by patterns !!!
- **Guidance**: Multiple patterns reference different time sets in the secondary check path and no override is provided. Set the TimeSet column in the BinCut instance sheet to explicitly select the intended time set.

### E_PerformanceMode_01
- **FullCode**: EBC154057
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Duplicate
- **EnumErrorTarget**: PerformanceMode
- **Code**: 57
- **Level**: Error
- **#Args**: 1
- **Template**: The performance {0} is duplicated in other post flow sheets !!!
- **Guidance**: Verify the performance mode name against the supported mode list. Update the configuration to use a valid performance mode.

### E_Resolution_01
- **FullCode**: EBC359058
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Resolution
- **Code**: 58
- **Level**: Error
- **#Args**: 3
- **Template**: The efuse {0} resolution {1} is different with StepSize {2} from the Notes sheet !!!
- **Guidance**: Check the BinCut resolution value against the specification. Ensure the resolution step size is valid for the bin type.

### E_SRAMthresh_01
- **FullCode**: EBC667059
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: SRAM
- **Code**: 59
- **Level**: Error
- **#Args**: 0
- **Template**: The SRAMthresh_Product value is not allowed <= 0 !!!
- **Guidance**: Check the SRAM threshold value in the BinCut definition. Ensure the threshold is within the valid range for the specified SRAM type.

### E_SRAMthresh_02
- **FullCode**: EBC667060
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: SRAM
- **Code**: 60
- **Level**: Error
- **#Args**: 0
- **Template**: The SRAMthresh_binSearch (CP1) value is not allowed to be <= 0.
- **Guidance**: Set the SRAMthresh_binSearch column value to a positive number greater than zero.

### E_SRAMthresh_03
- **FullCode**: EBC667061
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: SRAM
- **Code**: 61
- **Level**: Error
- **#Args**: 0
- **Template**: The SRAMthresh_binSearch (CP2) value is not allowed to be <= 0.
- **Guidance**: Set the SRAMthresh_binSearch CP2 column value to a positive number greater than zero.

### E_TestingStage_01
- **FullCode**: EBC239062
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Job
- **Code**: 62
- **Level**: Error
- **#Args**: 1
- **Template**: The job name of Testing Stage "{0}" is unknown !!!
- **Guidance**: Verify the testing stage name against the defined stage list. Check for typos or missing stage definitions in the configuration.

### E_TimeSet_01
- **FullCode**: EBC476063
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Timing
- **Code**: 63
- **Level**: Error
- **#Args**: 1
- **Template**: TimeSet {0} does not exist is K folder
- **Guidance**: Verify the time set name against the defined time set list. Add the missing time set or correct the reference.

### E_TimeSet_02
- **FullCode**: EBC476064
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Timing
- **Code**: 64
- **Level**: Error
- **#Args**: 0
- **Template**: Can't get any SCAN Tset in all payload
- **Guidance**: Check the pattern dashboard for all payload patterns in this BinCut instance row. Ensure at least one payload pattern has a valid SCAN time set (ScanTset) defined.

### E_TTR_01
- **FullCode**: WBC629065
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Format
- **Code**: 65
- **Level**: Warning
- **#Args**: 1
- **Template**: Can not use multi enable {0} in one TTR instance !!!
- **Guidance**: Check the Enable value for TTR group in the BinCut_Instance. Enable value should be the same in one TTR group.

### E_UserFunction_01
- **FullCode**: EBC481066
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Userfunction
- **Code**: 66
- **Level**: Error
- **#Args**: 1
- **Template**: DigSrc {0} in UserFunction is not defined in UF_DigSrc sheet
- **Guidance**: Check if the Userfunction key is defined in Userfunction sheet.

### E_UserFunction_02
- **FullCode**: EBC481067
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Userfunction
- **Code**: 67
- **Level**: Error
- **#Args**: 1
- **Template**: Need to assign digital source group for digital source pattern {0}
- **Guidance**: Assign a DigSrc group in the UserFunction column for the row containing the digital source pattern.

### E_UserFunction_03
- **FullCode**: EBC281068
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Userfunction
- **Code**: 68
- **Level**: Error
- **#Args**: 0
- **Template**: UserFunction format error, should be "DigSrc:[DigSrcGroup]"
- **Guidance**: Correct the UserFunction value to follow the format 'DigSrc:[DigSrcGroup]'.

### E_VoltageType_01
- **FullCode**: EBC684069
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Voltage
- **Code**: 69
- **Level**: Error
- **#Args**: 0
- **Template**: Voltage type is different in the Sub Flow column and Voltage Category column !!!
- **Guidance**: Verify the voltage type against the list of supported BinCut voltage types. Update the configuration to use a valid voltage type.

### W_Domain_01
- **FullCode**: WBC221070
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Domain
- **Code**: 70
- **Level**: Warning
- **#Args**: 0
- **Template**: Unable to get domain from flow name and pattern !!!
- **Guidance**: Open the BinCut instance sheet and verify that the flow name or at least one pattern can be resolved to a known domain. Ensure the flow name follows the naming convention that encodes the domain.

### W_Domain_02
- **FullCode**: WBC221071
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Domain
- **Code**: 71
- **Level**: Warning
- **#Args**: 0
- **Template**: Unable to get domain from flow name !!!
- **Guidance**: Open the BinCut instance sheet and verify the flow name contains a recognizable domain syntax. The pattern list domain will be used as a fallback.

### W_Domain_03
- **FullCode**: WBC221072
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Domain
- **Code**: 72
- **Level**: Warning
- **#Args**: 0
- **Template**: Unable to get domain from pattern !!!
- **Guidance**: Open the BinCut instance sheet and verify the pattern names contain a recognizable domain syntax. The flow name domain will be used as a fallback.

### W_Domain_04
- **FullCode**: WBC321073
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Domain
- **Code**: 73
- **Level**: Warning
- **#Args**: 2
- **Template**: This domain of flowName {0} is different from the pattern {1}!!!
- **Guidance**: Open the BinCut instance sheet and verify the domain consistency between the flow name and the assigned patterns. All patterns in a row should belong to the same domain as the flow name.

### W_Domain_05
- **FullCode**: WBC321074
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Domain
- **Code**: 74
- **Level**: Warning
- **#Args**: 0
- **Template**: This domain is different from other patterns!!!
- **Guidance**: Open the BinCut instance sheet and verify that all patterns in the row belong to the same domain. Mixed-domain patterns within a single row are not supported.

### W_efuseBitDef_01
- **FullCode**: WBC424075
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Field
- **Code**: 75
- **Level**: Warning
- **#Args**: 3
- **Template**: The VDD_{0}_{1} can not be found in efuseBitDef, and only found relative name {2} !!!
- **Guidance**: Open the eFuse bit definition table and add an exact entry for the flagged VDD domain+mode combination. The relative match may not produce correct voltage mapping.

### W_efuseBitDef_02
- **FullCode**: WBC424076
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Field
- **Code**: 76
- **Level**: Warning
- **#Args**: 2
- **Template**: The ids name of {0} in flow can not be found in efuseBitDef, and only found relative name {1} !!!
- **Guidance**: Open the eFuse bit definition table and add an exact entry for the flagged IDS pin. The relative match may not produce correct IDS value mapping in the flow.

### W_efuseBitDef_03
- **FullCode**: WBC424077
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Field
- **Code**: 77
- **Level**: Warning
- **#Args**: 1
- **Template**: The ids name of {0} in flow can not be found in efuseBitDef !!!
- **Guidance**: Open the eFuse bit definition table and add an entry for the flagged IDS pin. Verify the IDS pin name follows the naming convention used in the efuseBitDef configuration.

### W_Flow_01
- **FullCode**: WBC227078
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Flow
- **Code**: 78
- **Level**: Warning
- **#Args**: 1
- **Template**: The mode {0} has multiple types in FUNC block !!!
- **Guidance**: Verify the flow reference in the BinCut sheet against the defined flow list. Check for typos or missing flow definitions.

### W_Flow_02
- **FullCode**: WBC427079
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Flow
- **Code**: 79
- **Level**: Warning
- **#Args**: 1
- **Template**: Please check the {0} syntax of performance, that should be existed in binning sheeet !!!
- **Guidance**: Open the BinCut flow sheet and verify the performance mode prefix against the binning sheet. Ensure the performance mode value begins with a recognized mode name.

### W_Flow_03
- **FullCode**: WBC284080
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Voltage
- **Code**: 80
- **Level**: Warning
- **#Args**: 1
- **Template**: AllOther in row {0} is empty !!!
- **Guidance**: Open the BinCut flow sheet and fill in the AllOther voltage value for the flagged row. Every row must have a non-empty AllOther voltage assignment.

### W_Flow_04
- **FullCode**: WBC484081
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Voltage
- **Code**: 81
- **Level**: Warning
- **#Args**: 0
- **Template**: 0 in row 1 is empty !!!
- **Guidance**: Open the BinCut flow sheet and fill in the missing voltage value for the flagged pin column. Every power pin must have a valid voltage assignment in each row.

### W_Flow_05
- **FullCode**: WBC454082
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: PerformanceMode
- **Code**: 82
- **Level**: Warning
- **#Args**: 1
- **Template**: The mode {0} in flow can not be found in binning, binning_binX ro binning_binY!!!
- **Guidance**: Open the BinCut binning sheets and verify the performance mode definition. Add the missing mode to the appropriate binning sheet.

### W_Flow_06
- **FullCode**: WBC454083
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: PerformanceMode
- **Code**: 83
- **Level**: Warning
- **#Args**: 1
- **Template**: Please check the {0} syntax of performance, that should be existed in binning sheeet !!!
- **Guidance**: Open the new BinCut flow sheet and verify the performance mode prefix against the binning sheet. Ensure the performance mode value begins with a recognized mode name.

### W_FlowName_01
- **FullCode**: WBC229084
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 84
- **Level**: Warning
- **#Args**: 1
- **Template**: Can not use multi flow name {0} in one TTR instance !!!
- **Guidance**: Open the BinCut instance sheet and assign a single flow name per TTR instance row. Merge or split TTR rows as needed to satisfy the one-flow-name-per-instance constraint.

### W_FlowName_02
- **FullCode**: WBC427085
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Flow
- **Code**: 85
- **Level**: Warning
- **#Args**: 1
- **Template**: The flowName "{0}" not found in the flow sheet !!!
- **Guidance**: Open the BinCut flow sheet and verify the flow name is defined. Add the missing flow definition or correct the flow name reference in the instance sheet.

### W_HTOLGb_01
- **FullCode**: WBC631086
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: GB
- **Code**: 86
- **Level**: Warning
- **#Args**: 2
- **Template**: The GB of HTOL {0} is bigger than FT {1}
- **Guidance**: Open the BinCut binning table and review the HTOL and FT hot guardband values. The HTOL RO guardband should not exceed the FT hot guardband.

### W_HTOLGb_02
- **FullCode**: WBC631087
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: GB
- **Code**: 87
- **Level**: Warning
- **#Args**: 2
- **Template**: The GB of HTOL {0} is bigger than FT {1}
- **Guidance**: Open the BinCut binning table and review the HTOL room and FT room guardband values. The HTOL room guardband should not exceed the FT room guardband.

### W_HTOLGb_03
- **FullCode**: WBC631088
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: GB
- **Code**: 88
- **Level**: Warning
- **#Args**: 2
- **Template**: The GB of HTOL {0} is bigger than FT {1}
- **Guidance**: Open the BinCut binning table and review the HTOL hot and FT room guardband values. The HTOL hot guardband should not exceed the FT room guardband.

### W_Pattern_01
- **FullCode**: WBC452089
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pattern
- **Code**: 89
- **Level**: Warning
- **#Args**: 1
- **Template**: {0} can't be found the pattern in CSV !!!
- **Guidance**: Verify the pattern name against the pattern CSV dashboard. Ensure the pattern is listed and the name spelling is correct.

### W_Pattern_02
- **FullCode**: WBC752090
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Info
- **EnumErrorTarget**: Pattern
- **Code**: 90
- **Level**: Warning
- **#Args**: 1
- **Template**: This pattern {0} is "Dont_useInCsv" !!!
- **Guidance**: The pattern is intentionally excluded via the 'dont_use' flag in the CSV. If the pattern is needed, update its use flag in the pattern CSV dashboard.

### W_Pattern_03
- **FullCode**: WBC452091
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pattern
- **Code**: 91
- **Level**: Warning
- **#Args**: 1
- **Template**: There are no FileVersion for {0} !!!
- **Guidance**: Check the pattern CSV dashboard for the flagged pattern. Ensure a valid FileVersion value is assigned to the pattern entry.

### W_Pattern_04
- **FullCode**: WBC452092
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pattern
- **Code**: 92
- **Level**: Warning
- **#Args**: 1
- **Template**: {0} can't be found the pattern in CSV !!!
- **Guidance**: Verify the pattern name against the pattern CSV dashboard in the secondary check path. Ensure the pattern is listed and the name spelling is correct.

### W_Pattern_05
- **FullCode**: WBC752093
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Info
- **EnumErrorTarget**: Pattern
- **Code**: 93
- **Level**: Warning
- **#Args**: 1
- **Template**: This pattern {0} is "Dont_useInCsv" !!!
- **Guidance**: The pattern is intentionally excluded via the 'dont_use' flag in the CSV. If the pattern is needed, update its use flag in the pattern CSV dashboard.

### W_Pattern_06
- **FullCode**: WBC252094
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Pattern
- **Code**: 94
- **Level**: Warning
- **#Args**: 0
- **Template**: The pattern can not be empty !!!
- **Guidance**: Open the BinCut instance sheet and fill in the empty pattern cell, or remove the extra column. Non-first pattern cells should either contain a valid pattern name or be omitted.

### W_PatternTimeSet_01
- **FullCode**: WBC176095
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Duplicate
- **EnumErrorTarget**: Timing
- **Code**: 95
- **Level**: Warning
- **#Args**: 2
- **Template**: There are multi time set {0} by patterns overwrite to {1}!!!
- **Guidance**: Multiple patterns reference different time sets. The specified time set value was used as an override. Verify that the override time set is correct for all patterns in this row.

### W_PatternTimeSet_02
- **FullCode**: WBC452096
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pattern
- **Code**: 96
- **Level**: Warning
- **#Args**: 0
- **Template**: This test can not find any time set by pattern dashboard !!!
- **Guidance**: Check the pattern dashboard for all payload patterns in this row. Ensure at least one pattern has a valid time set (TimeSetVersion) defined in the CSV.

### W_PatternTimeSet_03
- **FullCode**: WBC152097
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Duplicate
- **EnumErrorTarget**: Pattern
- **Code**: 97
- **Level**: Warning
- **#Args**: 2
- **Template**: There are multi time set {0} by patterns overwrite to {1}!!!
- **Guidance**: Multiple patterns reference different time sets in the secondary check path. The specified time set value was used as an override. Verify that the override time set is correct for all patterns in this row.

### W_TTR_01
- **FullCode**: WBC629098
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Format
- **Code**: 98
- **Level**: Warning
- **#Args**: 1
- **Template**: Can not use multi enable {0} in one TTR instance !!!
- **Guidance**: Open the BinCut instance sheet and reduce the TTR enable values to a single entry per instance. Merge or split rows as needed to satisfy the one-enable-per-instance constraint.

### W_TTR_02
- **FullCode**: WBC629099
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Format
- **Code**: 99
- **Level**: Warning
- **#Args**: 1
- **Template**: Can not use multi JobTestStage {0} in one TTR instance !!!
- **Guidance**: Open the BinCut instance sheet and reduce the TTR JobTestStage values to a single entry per instance. Each TTR instance must specify exactly one JobTestStage.

### W_TTR_03
- **FullCode**: WBC629100
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Format
- **Code**: 100
- **Level**: Warning
- **#Args**: 1
- **Template**: Can not use multi SiteVar {0} in one TTR instance !!!
- **Guidance**: Open the BinCut instance sheet and reduce the TTR SiteVar values to a single entry per instance. Each TTR instance must reference exactly one SiteVar.

### W_TTR_04
- **FullCode**: WBC629101
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Format
- **Code**: 101
- **Level**: Warning
- **#Args**: 1
- **Template**: Can not use multi ShiftSpeed {0} in one TTR instance !!!
- **Guidance**: Open the BinCut instance sheet and reduce the TTR ShiftSpeed values to a single entry per instance. Each TTR instance must have exactly one ShiftSpeed assignment.

### W_TTR_05
- **FullCode**: WBC629102
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Format
- **Code**: 102
- **Level**: Warning
- **#Args**: 1
- **Template**: Can not use multi Voltage Category {0} in one TTR instance !!!
- **Guidance**: Open the BinCut instance sheet and reduce the TTR Voltage Category values to a single entry per instance. Each TTR instance must reference exactly one voltage category.

### W_TTR_06
- **FullCode**: WBC629103
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Format
- **Code**: 103
- **Level**: Warning
- **#Args**: 1
- **Template**: Can not use multi timeSet {0} in one TTR instance !!!
- **Guidance**: Open the BinCut instance sheet and reduce the TTR timeSet values to a single entry per instance. Each TTR instance must specify exactly one time set.

### W_TTR_07
- **FullCode**: WBC629104
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Format
- **Code**: 104
- **Level**: Warning
- **#Args**: 1
- **Template**: Can not use multi FailFlag {0} in one TTR instance !!!
- **Guidance**: Open the BinCut instance sheet and reduce the TTR FailFlag values to a single entry per instance. Each TTR instance must reference exactly one FailFlag.

### W_TTR_08
- **FullCode**: WBC629105
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Format
- **Code**: 105
- **Level**: Warning
- **#Args**: 1
- **Template**: Can not use multi PatSetName(Orange) {0} in one TTR instance !!!
- **Guidance**: Open the BinCut instance sheet and reduce the TTR PatSetName (Orange) values to a single entry per instance. Each TTR instance must reference exactly one pattern set name.

### E_FormatError_06
- **FullCode**: EBC229106
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 106
- **Level**: Error
- **#Args**: 1
- **Template**: Please check syntax for"{0}" ,the count of "(" and ")" are mismatch !!!
- **Guidance**: Check the voltage pin entry in the BinCut flow table for format issues.

### E_FormatError_07
- **FullCode**: EBC229107
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 107
- **Level**: Error
- **#Args**: 2
- **Template**: The syntax "{0}"  of {1} was unknown !!!
- **Guidance**: Check the voltage pin entry in the BinCut flow table for format issues.

### E_FormatError_08
- **FullCode**: EBC629108
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Format
- **Code**: 108
- **Level**: Error
- **#Args**: 1
- **Template**: {0} has not evaluate bin before Bin Result !!!
- **Guidance**: Check the performance mode column entry in the BinCut flow table for format issues.

### E_FormatError_09
- **FullCode**: EBC229109
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 109
- **Level**: Error
- **#Args**: 0
- **Template**: The 0 is incorrect DC category in the Bincut !!!
- **Guidance**: Check the AllOther column entry in the BinCut flow table for format issues.

### E_FormatError_10
- **FullCode**: EBC229110
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 110
- **Level**: Error
- **#Args**: 0
- **Template**: There are extra spaces !!!
- **Guidance**: Check the ATPG column entry in the BinCut flow table for format issues.

### E_FormatError_11
- **FullCode**: EBC229111
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 111
- **Level**: Error
- **#Args**: 0
- **Template**: There are extra spaces !!!
- **Guidance**: Check the MBIST column entry in the BinCut flow table for format issues.

### E_FormatError_12
- **FullCode**: EBC229112
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 112
- **Level**: Error
- **#Args**: 0
- **Template**: There are extra spaces !!!
- **Guidance**: Check the SPI/RTOS column entry in the BinCut flow table for format issues.

### W_DuplicateProgramSheet_01
- **FullCode**: WBC166113
- **EnumErrorCategory**: BinCut
- **EnumErrorBehavior**: Duplicate
- **EnumErrorTarget**: Source
- **Code**: 113
- **Level**: Warning
- **#Args**: 1
- **Template**: Duplicate sheet from {0} in PowerBinning Workbook!
- **Guidance**: Check for duplicate sheet names in the program workbook. Remove or rename the duplicate sheet so each sheet has a unique name.

## BinOutReportErrorType (`BinOutReportErrorType.cs`)

### E_MismatchCount_01
- **FullCode**: EBO314001
- **EnumErrorCategory**: BinOutReport
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Count
- **Code**: 1
- **Level**: Error
- **#Args**: 4
- **Template**: Count mismatch detected for job "{0}": binout report count ({1}) does not match the count in program sheet "{2}" ({3}). (The updated value of Binout status and limit value is not exactly the same in each row)
- **Guidance**: Please check for any missing or extra rows in the flow sheet that may cause the count mismatch.

### E_MismatchCount_02
- **FullCode**: EBO314002
- **EnumErrorCategory**: BinOutReport
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Count
- **Code**: 2
- **Level**: Error
- **#Args**: 4
- **Template**: Count mismatch detected for job "{0}": binout report count ({1}) does not match the count in program sheet "{2}" ({3}). (The updated value of Binout status is not exactly the same in each row)
- **Guidance**: Please check for any missing or extra rows in the flow sheet that may cause the count mismatch.

### E_MismatchCount_03
- **FullCode**: EBO314003
- **EnumErrorCategory**: BinOutReport
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Count
- **Code**: 3
- **Level**: Error
- **#Args**: 4
- **Template**: Count mismatch detected for job "{0}": binout report count ({1}) does not match the count in program sheet "{2}" ({3}). (The updated value of limit value is not exactly the same in each row)
- **Guidance**: Please check for any missing or extra rows in the flow sheet that may cause the count mismatch.

### E_MismatchCount_04
- **FullCode**: EBO314004
- **EnumErrorCategory**: BinOutReport
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Count
- **Code**: 4
- **Level**: Error
- **#Args**: 2
- **Template**: Count mismatch for functional item "{0}" between TP and datalog. The value cannot be updated in "{1}".
- **Guidance**: Please check for any missing or extra rows in the flow sheet that may cause the count mismatch.

### E_MismatchCount_05
- **FullCode**: EBO314005
- **EnumErrorCategory**: BinOutReport
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Count
- **Code**: 5
- **Level**: Error
- **#Args**: 3
- **Template**: "{0}" with TName "{1}" was not found in the binout report for "{2}", so it cannot be updated.
- **Guidance**: Please verify that the item and TName exist in the report for the specified job and part.

### E_MismatchCount_06
- **FullCode**: EBO314006
- **EnumErrorCategory**: BinOutReport
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Count
- **Code**: 6
- **Level**: Error
- **#Args**: 3
- **Template**: TP/binout report count mismatch or inconsistent status/limit values found for "{0}" with TName "{1}" in "{2}". It cannot be updated.
- **Guidance**: Please verify that the TP and binout report counts match and that the status and limit values are consistent for the same item and TName within the job and part.

### E_MissingBinTable_01
- **FullCode**: EBO405001
- **EnumErrorCategory**: BinOutReport
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: BinTable
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Fail flag(s) "{0}" are not referenced by any bintable row in Bin_Table_HardIP.
- **Guidance**: Please ensure the fail flag(s) are referenced in at least one row of the Bin_Table_HardIP sheet.

### E_MissingBinTable_02
- **FullCode**: EBO405002
- **EnumErrorCategory**: BinOutReport
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: BinTable
- **Code**: 2
- **Level**: Error
- **#Args**: 2
- **Template**: Bin table "{0}" is not referenced in flow sheet "{1}".
- **Guidance**: Please ensure the bin table is referenced in the flow sheet.

### E_MissingFlag_01
- **FullCode**: EBO426001
- **EnumErrorCategory**: BinOutReport
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Flag
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: No fail flags found for instance "{0}".
- **Guidance**: Binout update requires at least one fail flag. Please verify that the correct flags are assigned to the instance.

### E_MissingInstance_01
- **FullCode**: EBO437001
- **EnumErrorCategory**: BinOutReport
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Instance
- **Code**: 1
- **Level**: Error
- **#Args**: 2
- **Template**: Instance {0} not be found in the flow sheet {1}.
- **Guidance**: Please ensure the instance exists in the flow sheet before performing the binout status update.

### E_MissingInstance_02
- **FullCode**: EBO437002
- **EnumErrorCategory**: BinOutReport
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Instance
- **Code**: 2
- **Level**: Error
- **#Args**: 1
- **Template**: Instance "{0}" is listed in the report, but no corresponding flow row was found in any flow sheet of the test program.
- **Guidance**: Please verify that the instance exists and is referenced in the correct flow sheet

### W_InvalidLimit_01
- **FullCode**: WBO242001
- **EnumErrorCategory**: BinOutReport
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Limit
- **Code**: 1
- **Level**: Warning
- **#Args**: 3
- **Template**: The {0} limit "{1}" in flow sheet "{2}" is not a numeric value.
- **Guidance**: Since the value is not numeric, please verify that it is the intended value.

### W_MismatchJob_01
- **FullCode**: WBO339001
- **EnumErrorCategory**: BinOutReport
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Job
- **Code**: 1
- **Level**: Warning
- **#Args**: 2
- **Template**: Instance "{0}" is not tested in job "{1}", so the tool will not update the value in the test program.
- **Guidance**: Please ensure the instance is included in the test flow for the specified job before updating the test program.

### W_MissingFlag_01
- **FullCode**: WBO426001
- **EnumErrorCategory**: BinOutReport
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Flag
- **Code**: 1
- **Level**: Warning
- **#Args**: 2
- **Template**: Original flag for instance "{0}" was not found. The tool will generate a new flag "{1}".
- **Guidance**: Please verify the original flag assignment in the flow sheet and confirm that the generated flag is correct.

## CharErrorType (`CharErrorType.cs`)

### E_IllegalChar_01
- **FullCode**: ECZ282001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Value
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: There is illegal char in {0}
- **Guidance**: Remove all the underline for each USERDEF.

### E_DuplicateTpName_01
- **FullCode**: ECZ137001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Duplicate
- **EnumErrorTarget**: Instance
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: Duplicate TP Name
- **Guidance**: Please remove one of duplicate case from char plan. Or re-name to another TP name (if the pattern set is different).

### E_DuplicatePattern_01
- **FullCode**: ECZ152001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Duplicate
- **EnumErrorTarget**: Pattern
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: duplicate patterns
- **Guidance**: There is duplicate pattern in a char item. Please check the correct pattern is fill in char plan

### E_ErrorShmooRange_01
- **FullCode**: ECZ357001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Range
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: Step is not empty besides primary shmoo
- **Guidance**: Only the primary shmoo can define a Step value. The Step field of all other shmoo settings must be empty.

### E_ErrorShmooRange_02
- **FullCode**: ECZ357002
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Range
- **Code**: 2
- **Level**: Error
- **#Args**: 0
- **Template**: NumStepNotMatch
- **Guidance**: Shmoo name not match, need to check.

### E_ErrorMethod_01
- **FullCode**: ECZ273001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Sweep
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: Error Search Method
- **Guidance**: Pre check search method should be liner for sweep on freq char row.

### E_OppositeUsl_01
- **FullCode**: ECZ242001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Limit
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: Opposite USL/LSL
- **Guidance**: USL must be greater than LSL.

### E_MissingPinName_01
- **FullCode**: ECZ455001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: Missing (pin name)/(pin group) in PinMap.txt/pinList sheet
- **Guidance**: Check if pin name exists in the PinList from PinMap.

### E_MissingPinName_02
- **FullCode**: ECZ455002
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 2
- **Level**: Error
- **#Args**: 2
- **Template**: Missing pin in pinmap file in {0}, Row {1}
- **Guidance**: Pin in char plan but not in pin map.

### E_MissingPinName_03
- **FullCode**: ECZ455003
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 3
- **Level**: Error
- **#Args**: 0
- **Template**: Missing VDD name in PinMap.txt
- **Guidance**: Pin in char plan but not in pin map.

### W_WrongGroup_01
- **FullCode**: WCZ254001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: PerformanceMode
- **Code**: 1
- **Level**: Warning
- **#Args**: 1
- **Template**: Wrong group for {0}
- **Guidance**: P-mode and Category mismatch. The p-mode and category should be the same block. You can ignore this error, if you want to use CPU P-mode to test SOC pattern.

### E_MissingHeader_01
- **FullCode**: ECZ433001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Header
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Missing Header: {0}
- **Guidance**: Wrong header in input sheet, please check!

### W_WrongRetention_01
- **FullCode**: WCZ229001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 1
- **Level**: Warning
- **#Args**: 0
- **Template**: Retention time format is wrong, the commaon count must be  in [0, 4, 14]
- **Guidance**: If you want the wait time after payload1. You should fill the comma in the corresponding position.

### W_WrongRetention_02
- **FullCode**: WCZ482001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Value
- **Code**: 1
- **Level**: Warning
- **#Args**: 0
- **Template**: Missing Retention time or PowerRunScenario for Category 'EXTRET' or 'NAPRET'
- **Guidance**: For category "EXTRET", must specifiy retention and power-run-scenario.

### W_WrongRetention_03
- **FullCode**: WCZ482002
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Value
- **Code**: 2
- **Level**: Warning
- **#Args**: 0
- **Template**: Missing PowerRunScenario for Category 'INTRET'
- **Guidance**: For category "INTRET", must specify power-run-scenario.

### W_WrongRetention_04
- **FullCode**: WCZ482003
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Value
- **Code**: 3
- **Level**: Warning
- **#Args**: 0
- **Template**: Missing PowerRunScenario for Category 'DISTURB'
- **Guidance**: For category "DISTURB", must specify power-run-scenario.

### W_WrongRetention_05
- **FullCode**: WCZ382001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Value
- **Code**: 1
- **Level**: Warning
- **#Args**: 0
- **Template**: Wait time exists but Category is not 'EXTRET'
- **Guidance**: If specified retention, the category must be EXTRET.

### E_WrongRetention_06
- **FullCode**: ECZ352001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Pattern
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: No matching write pattern for read pattern: {0}
- **Guidance**: No matching write pattern for read pattern.

### E_WrongRetention_07
- **FullCode**: ECZ352002
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Pattern
- **Code**: 2
- **Level**: Error
- **#Args**: 1
- **Template**: Unmatched write pattern {0} no found with read pattern
- **Guidance**: No matching write pattern for read pattern.

### W_WrongRetention_08
- **FullCode**: WCZ382002
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Value
- **Code**: 2
- **Level**: Warning
- **#Args**: 2
- **Template**: Retention time " {0} " is not same to tool expect value " {1} "
- **Guidance**: Retention time is not same to tool expect value.

### W_WrongRetention_09
- **FullCode**: WCZ382003
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Value
- **Code**: 3
- **Level**: Warning
- **#Args**: 0
- **Template**: Retention time due to exist extra inits(10) or payloads(5)
- **Guidance**: Retention time due to exist extra inits(10) or payloads(5).

### E_MissingShmooValue_01
- **FullCode**: ECZ482004
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Value
- **Code**: 4
- **Level**: Error
- **#Args**: 1
- **Template**: {0}
- **Guidance**: Missing start or stop value.

### E_MissingShmooCondition_01
- **FullCode**: ECZ482005
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Value
- **Code**: 5
- **Level**: Error
- **#Args**: 2
- **Template**: Missing shmoo condition in {0}, Row: {1}
- **Guidance**: Shmoo Condition must to define both start and stop value.

### E_MissingShmooCondition_02
- **FullCode**: ECZ482006
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Value
- **Code**: 6
- **Level**: Error
- **#Args**: 2
- **Template**: Missing VIH/VIL shmoo for HIO in {0}, Row: {1}
- **Guidance**: For HIO, should have some pin sweep setting.

### W_DummyShmooCondition_01
- **FullCode**: WCZ282002
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Value
- **Code**: 2
- **Level**: Warning
- **#Args**: 2
- **Template**: Dummy shmoo condition for HAC in {0}, Row: {1}
- **Guidance**: For HAC, start and stop values of power/IO pins should be the same under normal conditions.

### E_WrongTpNameEquation_01
- **FullCode**: ECZ223001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Equation
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: Wrong equation of TP Name
- **Guidance**: Please check the correct TP name in char plan should be equation base, and the any space in userdef1-userdef9, group, category should be removed.

### E_WrongUserdef1_01
- **FullCode**: ECZ282003
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Value
- **Code**: 3
- **Level**: Error
- **#Args**: 3
- **Template**: Wrong USERDEF1 in {0} sheets in {1}, Row: {2}
- **Guidance**: The value specified in UserDef1 must be defined in the corresponding sheet configuration. (HAC\|HFL\|HFH\|HFLH\|HIO) or (DFTL\|DFTH\|DFTLH\|MCL\|MCH\|MCLH)

### E_WrongMeasOfHac_01
- **FullCode**: ECZ246001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Measurement
- **Code**: 1
- **Level**: Error
- **#Args**: 3
- **Template**: Wrong MeasType for {0} in {1}, Row: {2}
- **Guidance**: Should have correct meas type in user_def2 column.

### E_WrongUserdef3OfHac_01
- **FullCode**: ECZ229002
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 2
- **Level**: Error
- **#Args**: 2
- **Template**: Wrong USERDEF3 for HAC in {0}, Row: {1}
- **Guidance**: Should have correct format in USERDEF3 column.

### E_WrongMeasCount_01
- **FullCode**: ECZ246002
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Measurement
- **Code**: 2
- **Level**: Error
- **#Args**: 3
- **Template**: {0} for HAC in {1}, Row: {2}
- **Guidance**: Wrong MeasCount.

### E_WrongMeasCount_02
- **FullCode**: ECZ246003
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Measurement
- **Code**: 3
- **Level**: Error
- **#Args**: 2
- **Template**: {0} total MeasCount in Char_Plan for HAC in {1} than in HardIP_info
- **Guidance**: Measure counts are not equal.

### E_WrongMeasPin_01
- **FullCode**: ECZ246004
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Measurement
- **Code**: 4
- **Level**: Error
- **#Args**: 3
- **Template**: {0} for HAC in {1}, Row: {2}
- **Guidance**: Wrong MeasPin.

### E_MissingPinSeq_01
- **FullCode**: ECZ455004
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 4
- **Level**: Error
- **#Args**: 0
- **Template**: Missing pin seq or specified of pattern in PatInfo file
- **Guidance**: Measure count are different and pat info meas count == 0.

### E_MissingPatternInPatInfo_01
- **FullCode**: ECZ453001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: PatternInfo
- **Code**: 1
- **Level**: Error
- **#Args**: 2
- **Template**: Missing patten in PatInfo file in {0}, Row: {1}
- **Guidance**: Cannot find patternName in PatInfoDict.

### E_WrongUserdef1OfVih_01
- **FullCode**: ECZ282004
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Value
- **Code**: 4
- **Level**: Error
- **#Args**: 2
- **Template**: TpName Contains VIH/VIL but USERDEF1 is not HFL\|HFH\|HIO in {0}, Row: {1}
- **Guidance**: For item in hard ip sheets and tp name contains vih \| vil, their user def 1 should be HFH \| HFL \| HIO.

### E_WrongShmooNv_01
- **FullCode**: ECZ282005
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Value
- **Code**: 5
- **Level**: Error
- **#Args**: 0
- **Template**: NV GlobalSpecs not within Start/Stop
- **Guidance**: The NV value defined in GlobalSpecs must be within the Start and Stop range of the shmoo setting.

### E_WrongShmooNv_02
- **FullCode**: ECZ282006
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Value
- **Code**: 6
- **Level**: Error
- **#Args**: 0
- **Template**: VDD power setting over than 1.7 of  nominal voltage 
- **Guidance**: The Start and Stop values of a VDD power shmoo must not exceed 1.7 times the nominal voltage (NV) defined in GlobalSpecs.

### E_MissingPinInGlobalSpecs_01
- **FullCode**: ECZ455005
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 5
- **Level**: Error
- **#Args**: 0
- **Template**: Shmoo pin does not exist in GlobalSpecs
- **Guidance**: Shmoo pin does not exist in GlobalSpecs.

### E_NoOtherSupplies_01
- **FullCode**: ECZ229003
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 3
- **Level**: Error
- **#Args**: 2
- **Template**: missing or wrong other supplies in {0}, Row: {1}
- **Guidance**: Please must define the other supply power condition. The other supply power should be NV or HV or LV.

### E_WrongVddInPinColumn_01
- **FullCode**: ECZ255001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Pin
- **Code**: 1
- **Level**: Error
- **#Args**: 2
- **Template**: Error VDD pins in Pin column in {0}, Row {1}
- **Guidance**: The VDD pin specified in the Pin column is invalid. Please verify that the pin name and configuration are defined correctly.

### E_PatternOtherThanInPayload1_01
- **FullCode**: ECZ252001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Pattern
- **Code**: 1
- **Level**: Error
- **#Args**: 2
- **Template**: There is pattern other than in payload1 in {0}, With: {1}
- **Guidance**: Only patterns defined in Payload1 are supported. Please remove any patterns assigned to other payloads.

### E_PatternsWithoutPayload_01
- **FullCode**: ECZ452001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pattern
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: There is no payload in these patterns: {0}
- **Guidance**: For non-Rtos cmd char row, there should be at least one payload in the patterns

### W_MissingHlnCondition_01
- **FullCode**: WCZ463001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Setting
- **Code**: 1
- **Level**: Warning
- **#Args**: 1
- **Template**: Missing HLN condition in {0}
- **Guidance**: Each H, L, N, and N1 condition must have a corresponding counter row. Please verify that all required HLN conditions are defined correctly.

### W_TooManyInstanceName_01
- **FullCode**: WCZ237001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Instance
- **Code**: 1
- **Level**: Warning
- **#Args**: 0
- **Template**: Too many instance names
- **Guidance**: The number of instance names associated with the same Payload1 must not exceed 200.

### E_MixedSIandDm_01
- **FullCode**: ECZ352003
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Pattern
- **Code**: 3
- **Level**: Error
- **#Args**: 2
- **Template**: Init/Payload patterns is mixed with SI/DM in {0}, Row {1}
- **Guidance**: SI/DM patterns must not be mixed with init or payload patterns. Please separate SI/DM patterns from the init/payload configuration.

### E_PinGroupNotMatch_01
- **FullCode**: ECZ356001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: PinGroup
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: Pin group in PinList sheet not match pinmap file
- **Guidance**: The pin group definition in the PinList sheet must match the corresponding pin group definition in the pinmap file. Please verify that both the pin count and pin members are consistent.

### E_WrongUsllslRange_01
- **FullCode**: ECZ257001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Range
- **Code**: 1
- **Level**: Error
- **#Args**: 2
- **Template**: {0} value is invalid or Wrong {0} for shmoo range in " {1} "
- **Guidance**: Invalid USL/LSL: the value is not within the shmoo range

### E_MissingForceCondition_01
- **FullCode**: ECZ455006
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 6
- **Level**: Error
- **#Args**: 1
- **Template**: USERDEF6 contains Vih\|Vil\|Vicm\|Vid, But do not has pin sweep {0}
- **Guidance**: For USERDEF6 contains Vih\|Vil\|Vicm\|Vid, should also specified the pin sweep with the corresponding type.

### E_WrongCapture_01
- **FullCode**: ECZ214001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Count
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: Wrong Capture count in USERDEF6
- **Guidance**: Need to check plan measC count matches with pattern info.

### E_WrongCapture_02
- **FullCode**: ECZ262001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Sequence
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: Wrong Capture sequence in USERDEF6 for MeasC
- **Guidance**: Need to check the measC sequence matches with pattern info.

### E_Userdef3MismatchShmoo_01
- **FullCode**: ECZ382004
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Value
- **Code**: 4
- **Level**: Error
- **#Args**: 0
- **Template**: There is an issue with USERDEF3 related Vmain or Valt pin shmoo setting
- **Guidance**: PE will base on the TP name to analysis the char data, so need to check if the shmoo setting is sweep.

### E_Userdef3MismatchShmoo_02
- **FullCode**: ECZ382005
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Value
- **Code**: 5
- **Level**: Error
- **#Args**: 0
- **Template**: USERDEF3 power does not match with shmoo setting
- **Guidance**: PE will base on the TP name to analysis the char data, so the information of shmoo sweep pins and userdef3 must be the same.

### E_WrongShmooSteps_01
- **FullCode**: ECZ282007
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Value
- **Code**: 7
- **Level**: Error
- **#Args**: 0
- **Template**: Start didn't equal Stop but step is 0
- **Guidance**: The shmoo steps from this char looks not reasonable. Start and Stop must be the same when Step is 0.

### W_WrongShmooSteps_02
- **FullCode**: WCZ282008
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Value
- **Code**: 8
- **Level**: Warning
- **#Args**: 0
- **Template**: Shmoo step > 0.005
- **Guidance**: The shmoo steps from this char looks not reasonable, Step value must be less than 0.005.

### W_WrongShmooSteps_03
- **FullCode**: WCZ282009
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Value
- **Code**: 9
- **Level**: Warning
- **#Args**: 0
- **Template**: Shmoo point is not integer
- **Guidance**: There is a shmoo point which is not integer in a shmoo setting.

### W_HeatAlarm_01
- **FullCode**: WCZ282010
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Value
- **Code**: 10
- **Level**: Warning
- **#Args**: 0
- **Template**: Stop voltage higher than start voltage
- **Guidance**: Stop voltage higher than start voltage. It might cause heat alarm, please check.

### E_LessThan6ShmooPoints_01
- **FullCode**: ECZ282011
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Value
- **Code**: 11
- **Level**: Error
- **#Args**: 0
- **Template**: Less than 6sShmoo points
- **Guidance**: A shmoo sweep must contain at least 6 points. Please adjust the Start, Stop, or Step value to increase the number of shmoo points.

### E_WrongNetName_01
- **FullCode**: ECZ233001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Header
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Header power name " {0} " is not a Net Name
- **Guidance**: VDD power on the header / USERDEF3 is not net name.

### E_WrongLimitFormat_01
- **FullCode**: ECZ229004
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 4
- **Level**: Error
- **#Args**: 0
- **Template**: Wrong format of USL/LSL
- **Guidance**: Invalid USL/LSL format. The value must be a valid double number.

### E_WrongDigSrc_01
- **FullCode**: ECZ402001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Assignment
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: Missing digsrc assignment in Userdef6/7/8
- **Guidance**: There is one pattern specified with ':DigSrc' but no digsrc assignment in Userdef6/7/8.

### E_WrongDigSrc_02
- **FullCode**: ECZ402002
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Assignment
- **Code**: 2
- **Level**: Error
- **#Args**: 0
- **Template**: Missing digsrc assignment in Userdef6/7
- **Guidance**: There are two patterns specified with ':DigSrc' but no digsrc assignment in Userdef6/7.

### E_WrongDigSrc_03
- **FullCode**: ECZ202001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Assignment
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: There are more than two patterns specified with ':DigSrc'
- **Guidance**: More than two patterns are specified with ':DigSrc'. Please limit the configuration to two patterns.

### E_WrongDigSrc_04
- **FullCode**: ECZ402003
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Assignment
- **Code**: 3
- **Level**: Error
- **#Args**: 0
- **Template**: Missing digsrc assignment in Userdef6
- **Guidance**: There are two patterns specified with ':DigSrc' but no digsrc assignment in Userdef6.

### E_WrongDigSrc_05
- **FullCode**: ECZ402004
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Assignment
- **Code**: 4
- **Level**: Error
- **#Args**: 0
- **Template**: Missing digsrc assignment in Userdef7
- **Guidance**: There are two patterns specified with ':DigSrc' but no digsrc assignment in Userdef7.

### E_WrongDigSrc_06
- **FullCode**: ECZ202002
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Assignment
- **Code**: 2
- **Level**: Error
- **#Args**: 0
- **Template**: There are more than two MTD patterns specified with ':DigSrc'
- **Guidance**: More than two MTD patterns are specified with ':DigSrc'. Please limit the configuration to two patterns.

### E_WrongDigSrc_07
- **FullCode**: ECZ282012
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Value
- **Code**: 12
- **Level**: Error
- **#Args**: 1
- **Template**: S and binary can't be mixed in "{0}"}
- **Guidance**: Please remove the mixed usage of 'S' and binary digits in the same segment.

### E_WrongDigSrc_08
- **FullCode**: ECZ482007
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Value
- **Code**: 7
- **Level**: Error
- **#Args**: 1
- **Template**: "{0}" doesn't exist in Send Bit Str of pattern info
- **Guidance**: All DigSrc segments must have a corresponding definition in the Send Bit Str of pattern info. Please make sure the segment name is defined correctly.

### E_WrongDigSrc_09
- **FullCode**: ECZ282013
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Value
- **Code**: 13
- **Level**: Error
- **#Args**: 2
- **Template**: "{0}" has S that exceed the length of " {1} "
- **Guidance**: The total number of 'S' bits assigned to DigSrc segments cannot exceed the available S length defined in UserDef9. Please reduce the number of 'S' bits or update UserDef9 accordingly.

### E_WrongDigSrc_10
- **FullCode**: ECZ282014
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Value
- **Code**: 14
- **Level**: Error
- **#Args**: 3
- **Template**: S Length of {0} ({1}) isn't the same as segment length in pattern info ({2})
- **Guidance**: The number of 'S' bits defined in the DigSrc segment must match the segment length specified in the pattern information. Please ensure both definitions use the same length.

### E_WrongDigSrc_11
- **FullCode**: ECZ282015
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Value
- **Code**: 15
- **Level**: Error
- **#Args**: 2
- **Template**: {0} over the S length of {1}
- **Guidance**: The combined length of all DigSrc segments using 'S' exceeds the S length defined in UserDef9. Please adjust the segment assignments or update UserDef9 to provide sufficient S bits.

### E_WrongDigSrc_12
- **FullCode**: ECZ202003
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Assignment
- **Code**: 3
- **Level**: Error
- **#Args**: 1
- **Template**: {0}: segments for S are not continuous
- **Guidance**: All DigSrc segments assigned with 'S' must be placed in consecutive order. Please ensure the S-related segments are continuous without interruption from other segment assignments.

### E_WrongDigSrc_13
- **FullCode**: ECZ229005
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 5
- **Level**: Error
- **#Args**: 1
- **Template**: Data in {0} must be 0, 1, S
- **Guidance**: For segments with the 'F' format, only '0', '1', and 'S' are allowed. Please remove any unsupported characters from the segment definition.

### E_WrongDigSrc_14
- **FullCode**: ECZ229006
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 6
- **Level**: Error
- **#Args**: 1
- **Template**: Data in {0} must be 0-9 and A-F
- **Guidance**: For segments with the 'G' format, only hexadecimal characters (0-9 and A-F) are allowed. Please update the segment definition to use valid hexadecimal values.

### E_WrongDigSrc_15
- **FullCode**: ECZ229007
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 7
- **Level**: Error
- **#Args**: 1
- **Template**: " {0} " didn't be recognized
- **Guidance**: The segment format is not recognized. Please use a supported segment type and ensure the segment definition follows the expected syntax.

### W_WrongDigSrc_16
- **FullCode**: WCZ229008
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 8
- **Level**: Warning
- **#Args**: 0
- **Template**: Sgmt default value is not found, use '0' as sgmt default value
- **Guidance**: Sgmt default value is not found, use '0' as sgmt default value. If wish to define the sgmt default value, just add 'sgmt_default=0', or 'sgmt_default=1' in the Revision sheet.

### W_MissingRevisonSheet_01
- **FullCode**: WCZ425001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: File
- **Code**: 1
- **Level**: Warning
- **#Args**: 0
- **Template**: Missing revision sheet in CharPlan
- **Guidance**: Need to add Revision sheet in CharPlan to specify default digsrc value.

### E_MisMatchDigSrc_01
- **FullCode**: ECZ382006
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Value
- **Code**: 6
- **Level**: Error
- **#Args**: 0
- **Template**: Pattern ends with DigSrc without SendBitStr in pat info
- **Guidance**: pattern:DigSrc need to have sendBitStr in pattern info.

### E_TestNameOverLength_01
- **FullCode**: ECZ229009
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 9
- **Level**: Error
- **#Args**: 0
- **Template**: Test Name length over 255 bytes
- **Guidance**: Test Name length need less than 255 bytes, please check it.

### E_TestNameNotEndswithSingleUnderline_01
- **FullCode**: ECZ229010
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 10
- **Level**: Error
- **#Args**: 0
- **Template**: Test Name not ends with single underline
- **Guidance**: Please check userdef9 and TP name cell if existed only one underline at the last.

### W_EmptySheet_01
- **FullCode**: WCZ225001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: File
- **Code**: 1
- **Level**: Warning
- **#Args**: 1
- **Template**: sheet {0} is empty, all items are unused!
- **Guidance**: Sheet is empty, all items are unused!

### W_EmptyUse_01
- **FullCode**: WCZ229011
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 11
- **Level**: Warning
- **#Args**: 0
- **Template**: Use/Not Use of this Char Item is empty, please check
- **Guidance**: Use/Not Use of this Char Item is empty, please check.

### E_WrongPowerRunScenario_01
- **FullCode**: ECZ282016
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Value
- **Code**: 16
- **Level**: Error
- **#Args**: 2
- **Template**: Wrong PowerRunScenario ({0}): {1}
- **Guidance**: PowerRunScenario format check fail, need to modify.

### W_WrongPowerRunScenario_02
- **FullCode**: WCZ282017
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Value
- **Code**: 17
- **Level**: Warning
- **#Args**: 1
- **Template**: Wrong PowerRunScenario : {0}
- **Guidance**: PowerRunScenario should be init_(NV\|Sweep\|VRS)_PL_(NV\|Sweep\|VRS) due to exist extra inits or payloads.

### E_WrongPowerRunScenario_03
- **FullCode**: ECZ229012
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 12
- **Level**: Error
- **#Args**: 0
- **Template**: Spell wrong PowerRunScenario
- **Guidance**: The PowerRunScenario header is misspelled. Please use the exact header name 'PowerRunScenario'.

### E_EmptyHeader_01
- **FullCode**: ECZ229013
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 13
- **Level**: Error
- **#Args**: 1
- **Template**: EmptyColumn Need Check: {0}
- **Guidance**: An empty column header was detected. Please enter a valid header name or remove the unused column.

### W_PowerGroupApply_01
- **FullCode**: WCZ282018
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Value
- **Code**: 18
- **Level**: Warning
- **#Args**: 1
- **Template**: Suggest tie all {0} related power together
- **Guidance**: Sweep condition should apply to pin group for vdd_low and vdd_fixed, instead of indivisual of them.

### W_RtosUserdef6Syntax_01
- **FullCode**: WCZ402005
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Assignment
- **Code**: 5
- **Level**: Warning
- **#Args**: 0
- **Template**: There is no RTOS_Userdef6_Syntax in Revision sheet
- **Guidance**: RTOS Userdef6 Syntax is not assigned, please check it.

### E_RtosUserdef6Syntax_02
- **FullCode**: ECZ282019
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Value
- **Code**: 19
- **Level**: Error
- **#Args**: 1
- **Template**: RTOS_Userdef6_Syntax: {0} is wrong
- **Guidance**: The RTOS command format in UserDef6 is invalid. Please verify that the syntax follows the supported RTOS command format.

### E_TrackingPinMismatchPrimaryStep_01
- **FullCode**: ECZ355001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Pin
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Mismatch tracking pins: {0}
- **Guidance**: Tracking pins with the same Order must have the same number of shmoo points.  Please verify that the Start, Stop, and Step settings generate consistent shmoo point counts.

### E_PrimaryTrackingPinValtInconsist_01
- **FullCode**: ECZ263001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Setting
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: Shmoo Setting Contains Valt and Vmain
- **Guidance**: A shmoo tracking group must use either Valt or Vmain consistently.  Please verify the voltage type setting for all shmoos with the same Order.

### W_SramPinNotTieLogic_01
- **FullCode**: WCZ255002
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Pin
- **Code**: 2
- **Level**: Warning
- **#Args**: 3
- **Template**: Shmoo pin only define: {0}; suggest to tie: {1}, {2}
- **Guidance**: SELSRAM-related shmoo pins must be defined together. Please include all associated pins in the same shmoo setting.

### W_InitPatternbehindSelsram_01
- **FullCode**: WCZ252002
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Pattern
- **Code**: 2
- **Level**: Warning
- **#Args**: 0
- **Template**: Exist init pattern behind selsram pattern
- **Guidance**: Check if selsram pattern is always the last init pattern in a single char row.

### E_SelsramPatternNotFoundInMappingTable_01
- **FullCode**: ECZ252003
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Pattern
- **Code**: 3
- **Level**: Error
- **#Args**: 0
- **Template**: Cannot find any init patterns in mapping table
- **Guidance**: Confirm the selsram pattern is exist in mapping table.

### E_SelsramBitMismatch_01
- **FullCode**: ECZ382007
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Value
- **Code**: 7
- **Level**: Error
- **#Args**: 0
- **Template**: UserDef9 data mismatch with mapping table
- **Guidance**: Confirm the selsram refer length in char plan and selsram pattern bits in mapping table is match.

### W_MissingManualAcInProgram_01
- **FullCode**: WCZ400001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: AcCategory
- **Code**: 1
- **Level**: Warning
- **#Args**: 2
- **Template**: Can't serach category: {0} in program ac specs, use {1} as base to generate
- **Guidance**: The specified AC category was not found in the program AC specs. Please verify the category name or add the corresponding AC category definition.

### W_MissingManualAcInTimeSettingsSheet_01
- **FullCode**: WCZ400002
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: AcCategory
- **Code**: 2
- **Level**: Warning
- **#Args**: 1
- **Template**: Char manual ac: {0} is not in timesettings sheet, directly use program ac spec
- **Guidance**: The Manual AC setting is not defined in the TimeSettings sheet. The program AC spec will be used as the fallback source.

### E_MissingManualAcInProgAndTsSheet_01
- **FullCode**: ECZ400003
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: AcCategory
- **Code**: 3
- **Level**: Error
- **#Args**: 1
- **Template**: Char manual ac: {0} is not in timesettings sheet and program ac spec
- **Guidance**: The Manual AC setting must exist in the TimeSettings sheet or program AC specs.

### E_IllegalShmooValue_01
- **FullCode**: ECZ282020
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Value
- **Code**: 20
- **Level**: Error
- **#Args**: 0
- **Template**: Shmoo value is non-numeric string
- **Guidance**: Shmoo Start, Stop, and Step values must be valid numeric values. Please verify that all shmoo settings contain numeric data only.

### W_MultiUsedAcForPayload_01
- **FullCode**: WCZ300001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: AcCategory
- **Code**: 1
- **Level**: Warning
- **#Args**: 2
- **Template**: Multi ac category for payload({0}) in base program: {1}
- **Guidance**: Each payload in the base program must be associated with only one AC category. Please verify the payload configuration and remove any unexpected AC category assignments.

### W_MultiUsedTimesetForPayload_01
- **FullCode**: WCZ376001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Timing
- **Code**: 1
- **Level**: Warning
- **#Args**: 2
- **Template**: Multi timeset for payload({0}) in base program: {1}
- **Guidance**: Each payload in the base program must be associated with only one timeset. Please verify the payload configuration and remove any unexpected timeset assignments.

### W_MultiUsedDcForPayload_01
- **FullCode**: WCZ318001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: DcCategory
- **Code**: 1
- **Level**: Warning
- **#Args**: 2
- **Template**: Multi dc category for payload({0}) in base program: {1}
- **Guidance**: Each payload in the base program must be associated with only one DC category. Please verify the payload configuration and remove any unexpected DC category assignments.

### E_MissingUsedAcForPayload_01
- **FullCode**: ECZ400004
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: AcCategory
- **Code**: 4
- **Level**: Error
- **#Args**: 1
- **Template**: None of ac category for payload({0}) in base program.
- **Guidance**: Each payload in the base program must have a corresponding AC category. Please verify the payload configuration and ensure an AC category is defined.

### E_MissingUsedTimesetForPayload_01
- **FullCode**: ECZ476001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Timing
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: None of timeset for payload({0}) in base program
- **Guidance**: Each payload in the base program must have a corresponding timeset. Please verify the payload configuration and ensure an timeset is defined.

### E_MissingUsedDcForPayload_01
- **FullCode**: ECZ418001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: DcCategory
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: None of dc category for payload({0}) in base program
- **Guidance**: Each payload in the base program must have a corresponding DC category. Please verify the payload configuration and ensure an DC category is defined.

### W_MissingDigSrcSgmtInPattern_01
- **FullCode**: WCZ466001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Source
- **Code**: 1
- **Level**: Warning
- **#Args**: 2
- **Template**: Pattern: "{0}" has not "{1}"
- **Guidance**: The specified segment was not found in the pattern's Send Bit String definition. Please verify the segment name and ensure it is defined in the corresponding pattern information.

### W_VoltageHigherThan1P3_01
- **FullCode**: WCZ284001
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Voltage
- **Code**: 1
- **Level**: Warning
- **#Args**: 1
- **Template**: Shmoo {0} value >= 1.3
- **Guidance**: The absolute value of shmoo value should be less than 1.3. Please verify the Start/STOP setting and adjust it to a valid range.

### E_VddShmooLowToHigh_01
- **FullCode**: WCZ263002
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Setting
- **Code**: 2
- **Level**: Warning
- **#Args**: 0
- **Template**: Shmoo range is from low to high
- **Guidance**: The Shmoo Start value must be greater than the Stop value. Please define the shmoo range from high to low.

### E_MissingPinMap_01
- **FullCode**: ECZ425002
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: File
- **Code**: 2
- **Level**: Error
- **#Args**: 0
- **Template**: Missing PinMap.txt
- **Guidance**: Missing PinMap.txt. Please Reload PinMap or Test Program.

### W_MissingPatInfo_01
- **FullCode**: WCZ425003
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: File
- **Code**: 3
- **Level**: Warning
- **#Args**: 0
- **Template**: Pattern info is empty, please check!
- **Guidance**: Pattern info is empty, please check!

### E_IllegalForNewTChar_01
- **FullCode**: ECZ229014
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 14
- **Level**: Error
- **#Args**: 0
- **Template**: Define multi patterns
- **Guidance**: Can't define multi patterns here, please define one cell by one pattern.

### E_IllegalForNewTChar_02
- **FullCode**: ECZ229015
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 15
- **Level**: Error
- **#Args**: 0
- **Template**: Retention format error
- **Guidance**: Please follow {INIT/PL}{Index}:{RetentionTime}:{GuardBand}mV, e.g INIT2:0.02:+25mV.

### E_IllegalForNewTChar_03
- **FullCode**: ECZ229016
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 16
- **Level**: Error
- **#Args**: 1
- **Template**: {0} pattern didn't be defined, please check.
- **Guidance**: INIT/Payload pattern in retention cannot find corresponding INIT/Payload pattern in char plan.

### E_IllegalCharUslLsl_01
- **FullCode**: ECZ229017
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 17
- **Level**: Error
- **#Args**: 1
- **Template**: Can not convert {0} to a value for limit!
- **Guidance**: Invalid USL/LSL format. The value must be a valid double number.

### W_NotInHarvTable_01
- **FullCode**: WCZ482008
- **EnumErrorCategory**: Char
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Value
- **Code**: 8
- **Level**: Warning
- **#Args**: 3
- **Template**: {0} in {1}, Row: {2} is not defined in Harv_Mapping_Table
- **Guidance**: The specified HARV FSTP was not found in the Harv_Mapping_Table. Please verify the HARV FSTP name and ensure a corresponding mapping entry exists.

## ClockCheckErrorType (`ClockCheckErrorType.cs`)

### E_MissingFlow_01
- **FullCode**: ECL427001
- **EnumErrorCategory**: ClockCheck
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Flow
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Type/INIT sequence: "{0}" does not have a matching entry in "Sub Flow" column of Instance_Clock_Check sheet.
- **Guidance**: Check the Type/INIT sequence value in Clock_Check sheet has a corresponding entry in the "Sub Flow" column of of Instance_Clock_Check sheet.

### E_MissingLibrary_01
- **FullCode**: ECL441001
- **EnumErrorCategory**: ClockCheck
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Library
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: No library setting was found based on MiscInfo column value "{0}".
- **Guidance**: Verify that the MiscInfo column value is correct and that a corresponding library setting is defined.

### E_MissingSetting_01
- **FullCode**: ECL463001
- **EnumErrorCategory**: ClockCheck
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Setting
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: This instance does not have a matching entry in "{0}" sheet.
- **Guidance**: Verify that each Sub Flow in the Instance_Clock_Check sheet exists as a Type/INIT sequence in Clock_Check sheet.

### I_BurstPatJudgeFlag_01
- **FullCode**: ICL726001
- **EnumErrorCategory**: ClockCheck
- **EnumErrorBehavior**: Info
- **EnumErrorTarget**: Flag
- **Code**: 1
- **Level**: Info
- **#Args**: 1
- **Template**: Burst pattern instance uses fail flag "{0}".
- **Guidance**: Verify the burst pattern judge flag definition. Ensure the flag name and condition are valid and match the burst pattern flow.

## EFuseErrorType (`EFuseErrorType.cs`)

### E_RuleViolationBit_01
- **FullCode**: EEF606001
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Bit
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: LSB BIT isn't a Number. LSB BIT : {0}
- **Guidance**: Check the EFuse bit definition for the specified bit. Verify the bit address, width, and value are all within valid ranges.

### E_RuleViolationBit_02
- **FullCode**: EEF606002
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Bit
- **Code**: 2
- **Level**: Error
- **#Args**: 1
- **Template**: MSB BIT isn't a Number. MSB BIT : {0}
- **Guidance**: Check the EFuse bit definition for the specified bit. Verify the bit address, width, and value are all within valid ranges.

### E_RuleViolationBit_03
- **FullCode**: EEF606003
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Bit
- **Code**: 3
- **Level**: Error
- **#Args**: 1
- **Template**: Bit Width isn't a Number. Bit Width : {0}
- **Guidance**: Check the EFuse bit definition for the specified bit. Verify the bit address, width, and value are all within valid ranges.

### E_RuleViolationBit_04
- **FullCode**: EEF606004
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Bit
- **Code**: 4
- **Level**: Error
- **#Args**: 1
- **Template**: Bit Width doesn't Equals with End Bit - Start Bit + 1. Result is {0}
- **Guidance**: Check the EFuse bit definition for the specified bit. Verify the bit address, width, and value are all within valid ranges.

### E_RuleViolationBit_05
- **FullCode**: EEF606005
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Bit
- **Code**: 5
- **Level**: Error
- **#Args**: 3
- **Template**: Incorrect bit width in the column Default Value. Bit Width (Math.Ceiling({0}/4) = {1}) < Default value ({2})
- **Guidance**: Check the EFuse bit definition for the specified bit. Verify the bit address, width, and value are all within valid ranges.

### E_RuleViolationBit_06
- **FullCode**: EEF606006
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Bit
- **Code**: 6
- **Level**: Error
- **#Args**: 2
- **Template**: Incorrect bit width in the column Default Value. Bit Width ({0}) < Default value ({1})
- **Guidance**: Check the EFuse bit definition for the specified bit. Verify the bit address, width, and value are all within valid ranges.

### E_RuleViolationBit_07
- **FullCode**: EEF606007
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Bit
- **Code**: 7
- **Level**: Error
- **#Args**: 1
- **Template**: The content of this column must be [MSB:LSB] or [LSB:MSB]. Content {0}
- **Guidance**: Check the EFuse bit definition for the specified bit. Verify the bit address, width, and value are all within valid ranges.

### E_RuleViolationBit_08
- **FullCode**: EEF206008
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Bit
- **Code**: 8
- **Level**: Error
- **#Args**: 1
- **Template**: Non-numeric string exists in Fuse column. Content {0}
- **Guidance**: Check the EFuse bit definition for the specified bit. Verify the bit address, width, and value are all within valid ranges.

### E_RuleViolationBit_09
- **FullCode**: EEF606009
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Bit
- **Code**: 9
- **Level**: Error
- **#Args**: 1
- **Template**: This row format is MSB:LSB is inconsistent with the LSB:MSB used by other rows. Content {0}
- **Guidance**: Check the EFuse bit definition for the specified bit. Verify the bit address, width, and value are all within valid ranges.

### E_RuleViolationBit_10
- **FullCode**: EEF606010
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Bit
- **Code**: 10
- **Level**: Error
- **#Args**: 1
- **Template**: This row format is LSB:MSB is inconsistent with the MSB:LSB used by other rows. Content {0}
- **Guidance**: Check the EFuse bit definition for the specified bit. Verify the bit address, width, and value are all within valid ranges.

### E_RuleViolationCMP_01
- **FullCode**: EEF607001
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Block
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: The item {0}'s Bitwidth of the CMP block is not same as UDR Block
- **Guidance**: Check the EFuse comparator configuration and value. Ensure the comparator threshold is valid for the specified fuse field.

### E_InvalidCRC_01
- **FullCode**: EEF229001
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: Crc bit format needs to use [number:number]
- **Guidance**: Recalculate the EFuse CRC for the specified block. Ensure the input data matches the programmed fuse values before recomputing.

### E_InvalidCRC_02
- **FullCode**: EEF229002
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 2
- **Level**: Error
- **#Args**: 3
- **Template**: Item {0} bit range {1}:{2} exist itself
- **Guidance**: Recalculate the EFuse CRC for the specified block. Ensure the input data matches the programmed fuse values before recomputing.

### E_InvalidCRC_03
- **FullCode**: EEF229003
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 3
- **Level**: Error
- **#Args**: 3
- **Template**: Item {0} bit range {1}:{2} does not exist in ignore bit
- **Guidance**: Recalculate the EFuse CRC for the specified block. Ensure the input data matches the programmed fuse values before recomputing.

### E_InvalidCRC_04
- **FullCode**: EEF229004
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 4
- **Level**: Error
- **#Args**: 5
- **Template**: The CRC bit range {0}:{1} exist in itself and the programming stage {2} is untested job. CRC item: {3}, programming stage: {4}
- **Guidance**: Recalculate the EFuse CRC for the specified block. Ensure the input data matches the programmed fuse values before recomputing.

### E_InvalidCRC_05
- **FullCode**: EEF229005
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 5
- **Level**: Error
- **#Args**: 5
- **Template**: The bit range {0}:{1} does not exist in the ignore bit and the programming stage {2} is untested job. CRC item: {3}, programming stage: {4}
- **Guidance**: Recalculate the EFuse CRC for the specified block. Ensure the input data matches the programmed fuse values before recomputing.

### E_InvalidCRC_06
- **FullCode**: EEF229006
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 6
- **Level**: Error
- **#Args**: 6
- **Template**: The CRC bit range {0}:{1} exist in itself and the programming stage {3} is untested job. CRC item: {4}, programming stage: {5}
- **Guidance**: Recalculate the EFuse CRC for the specified block. Ensure the input data matches the programmed fuse values before recomputing.

### E_InvalidCRC_07
- **FullCode**: EEF229007
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 7
- **Level**: Error
- **#Args**: 6
- **Template**: The {0} bit range {1}:{2} does not exist in ignore bit and the programming stage {3} is untested job. CRC item: {4}, programming stage: {5}
- **Guidance**: Recalculate the EFuse CRC for the specified block. Ensure the input data matches the programmed fuse values before recomputing.

### E_MissingField_01
- **FullCode**: EEF424001
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Field
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Missing {0} name in the config table
- **Guidance**: Check the EFuse default value against the allowed bit-width for the field. Reduce the value or increase the allocated width to accommodate it.

### E_RuleViolationStage_01
- **FullCode**: EEF669001
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Stage
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: ECID Programming Stage is {0}
- **Guidance**: Verify the ECID non-CP programming stage against the expected stage definition. Ensure the stage is correctly configured for non-CP (non-circuit probe) testing.

### E_InvalidValue_01
- **FullCode**: EEF282001
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Value
- **Code**: 1
- **Level**: Error
- **#Args**: 2
- **Template**: {0} is integer, tool change to {1}
- **Guidance**: Verify that the EFuse sheet or referenced resource is present in the workbook. Check the sheet name spelling and ensure it was not accidentally deleted.

### W_RuleViolationFuseBlowLocation_01
- **FullCode**: WEF682001
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Value
- **Code**: 1
- **Level**: Warning
- **#Args**: 0
- **Template**: Fuse Blow Location is FTF
- **Guidance**: Verify the EFuse blow location address against the device fuse map. Ensure the location is within the valid fuse address range.

### E_InvalidLimit_01
- **FullCode**: EEF242001
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Limit
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Low Limit isn't a Number. Low Limit : {0}
- **Guidance**: Check the EFuse limit value against the specification. Ensure the limit is within the allowed range for the fuse type.

### E_InvalidLimit_02
- **FullCode**: EEF242002
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Limit
- **Code**: 2
- **Level**: Error
- **#Args**: 1
- **Template**: High Limit isn't a Number. High Limit : {0}
- **Guidance**: Check the EFuse limit value against the specification. Ensure the limit is within the allowed range for the fuse type.

### E_InvalidLimit_03
- **FullCode**: EEF242003
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Limit
- **Code**: 3
- **Level**: Error
- **#Args**: 0
- **Template**: When DefalutOrReal is Real, High Limit must large than Low Limit.
- **Guidance**: Check the EFuse limit value against the specification. Ensure the limit is within the allowed range for the fuse type.

### E_InvalidLimit_04
- **FullCode**: EEF242004
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Limit
- **Code**: 4
- **Level**: Error
- **#Args**: 0
- **Template**: IDS resolution * 2^(Bit Range) < High Limit.
- **Guidance**: Check the EFuse limit value against the specification. Ensure the limit is within the allowed range for the fuse type.

### E_InvalidLsbMsb_01
- **FullCode**: EEF279001
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Type
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: The lsb bit can not be recognized
- **Guidance**: Check the bit ordering (LSB/MSB) configuration for the specified EFuse field. Ensure the ordering matches the device specification and programming tool expectation.

### E_InvalidLsbMsb_02
- **FullCode**: EEF279002
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Type
- **Code**: 2
- **Level**: Error
- **#Args**: 0
- **Template**: The msb bit can not be recognized
- **Guidance**: Check the bit ordering (LSB/MSB) configuration for the specified EFuse field. Ensure the ordering matches the device specification and programming tool expectation.

### E_InvalidMaximumBits_01
- **FullCode**: EEF242005
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Limit
- **Code**: 5
- **Level**: Error
- **#Args**: 2
- **Template**: The last config condition MSB ({0}) < Maximum Bit ({1})
- **Guidance**: Verify the total EFuse bit allocation does not exceed the device maximum. Reduce the number of fuse bits or reorganize the fuse map.

### E_MismatchRevision_01
- **FullCode**: EEF360001
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Revision
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: The database revision is not same with the efuse database revision sheet content, please check it!
- **Guidance**: Compare the EFuse value against the expected definition. Check that the programmed value matches the test plan specification.

### E_MismatchDefaultType_01
- **FullCode**: EEF319001
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Default
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: The database revision should be defined as default value, please check it!
- **Guidance**: Check the EFuse default type definition against the expected type. Ensure the type is consistent across all references to this field.

### E_MissingParameter_01
- **FullCode**: EEF479001
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Type
- **Code**: 1
- **Level**: Error
- **#Args**: 3
- **Template**: Missing Parameter in {0}({1}) : {2}
- **Guidance**: Open the EFuse sheet and add the required parameter. Check the EFuse specification to confirm which parameters are mandatory.

### E_MissVbtModule_01
- **FullCode**: EEF447001
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Module
- **Code**: 1
- **Level**: Error
- **#Args**: 2
- **Template**: The VBT function: {0} can not find in {1} library!
- **Guidance**: Verify the VBT module name against the expected module list. Add the missing module definition or correct the reference.

### E_NotMatchToBDF_01
- **FullCode**: EEF379001
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Type
- **Code**: 1
- **Level**: Error
- **#Args**: 2
- **Template**: The pattern type is match to the BDF, pattern is {0}, but BDF is {1}
- **Guidance**: Compare the EFuse definition against the BDF. Update either the EFuse configuration or the BDF so the definitions are consistent.

### E_MismatchBit_01
- **FullCode**: EEF306001
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Bit
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: This bank count should be mod by 16 due to the ECC
- **Guidance**: Verify the EFuse bit count against the device specification. Ensure the total bits allocated do not exceed the physical fuse capacity.

### E_MismatchBit_02
- **FullCode**: EEF306002
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Bit
- **Code**: 2
- **Level**: Error
- **#Args**: 0
- **Template**: This column cannot be empty
- **Guidance**: Verify the EFuse bit count against the device specification. Ensure the total bits allocated do not exceed the physical fuse capacity.

### E_MismatchBit_03
- **FullCode**: EEF306003
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Bit
- **Code**: 3
- **Level**: Error
- **#Args**: 0
- **Template**: This column must be a number
- **Guidance**: Verify the EFuse bit count against the device specification. Ensure the total bits allocated do not exceed the physical fuse capacity.

### E_MissingPattern_01
- **FullCode**: EEF452001
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pattern
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: The pattern {0} is not found in pattern directory!!!
- **Guidance**: Verify the EFuse pattern name and its definition. Check that the pattern exists in the pattern folder and is correctly referenced.

### E_MissingPattern_02
- **FullCode**: EEF452002
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pattern
- **Code**: 2
- **Level**: Error
- **#Args**: 1
- **Template**: The pattern {0} can't be found in the CSV !!!
- **Guidance**: Verify the EFuse pattern name and its definition. Check that the pattern exists in the pattern folder and is correctly referenced.

### E_MissingPattern_03
- **FullCode**: EEF452003
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pattern
- **Code**: 3
- **Level**: Error
- **#Args**: 1
- **Template**: This pattern {0} is "Dont_useInCsv" !!!
- **Guidance**: Verify the EFuse pattern name. Check that the pattern exists in the pattern folder and is correctly referenced.

### E_MissingPattern_04
- **FullCode**: EEF452004
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pattern
- **Code**: 4
- **Level**: Error
- **#Args**: 1
- **Template**: This pattern {0} of the FileVersion column is "n/a" !!!
- **Guidance**: Verify the EFuse pattern name and its definition. Check that the pattern exists in the pattern folder and is correctly referenced.

### E_MissingPattern_05
- **FullCode**: EEF452005
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pattern
- **Code**: 5
- **Level**: Error
- **#Args**: 1
- **Template**: This pattern {0} can't get Read/Write pin !!!
- **Guidance**: Verify the EFuse pattern name and its definition. Check that the pattern exists in the pattern folder and is correctly referenced.

### E_MissingPattern_06
- **FullCode**: EEF452006
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pattern
- **Code**: 6
- **Level**: Error
- **#Args**: 1
- **Template**: The pattern {0} can't be found in the CSV !!!
- **Guidance**: Verify the EFuse pattern name and its definition. Check that the pattern exists in the pattern folder and is correctly referenced.

### E_MismatchPattern_01
- **FullCode**: EEF352001
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Pattern
- **Code**: 1
- **Level**: Error
- **#Args**: 6
- **Template**: The pattern {0} send count {1} {1} is not equal to {2} in the bank {3}, {4}*{5} !!!!!
- **Guidance**: Verify the EFuse pattern name and its definition. Check that the pattern exists in the pattern folder and is correctly referenced.

### E_MismatchPattern_02
- **FullCode**: EEF352002
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Pattern
- **Code**: 2
- **Level**: Error
- **#Args**: 2
- **Template**: The pattern {0} access mode (DAA) is not equal to BDF file {1}
- **Guidance**: Verify the EFuse pattern name and its definition. Check that the pattern exists in the pattern folder and is correctly referenced.

### E_MismatchPattern_03
- **FullCode**: EEF352003
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Pattern
- **Code**: 3
- **Level**: Error
- **#Args**: 2
- **Template**: The pattern {0} access mode (JTG) is not equal to BDF file {1}
- **Guidance**: Verify the EFuse pattern name and its definition. Check that the pattern exists in the pattern folder and is correctly referenced.

### E_MismatchPattern_04
- **FullCode**: EEF352004
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Pattern
- **Code**: 4
- **Level**: Error
- **#Args**: 8
- **Template**: The pattern {0} ({1}) store count {2} is not equal to {3} in the bank {4} ({5}), {6}*{7} !!!
- **Guidance**: Verify the EFuse pattern name and its definition. Check that the pattern exists in the pattern folder and is correctly referenced.

### E_MismatchPattern_05
- **FullCode**: EEF352005
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Pattern
- **Code**: 5
- **Level**: Error
- **#Args**: 1
- **Template**: This pattern {0} can't be discerned in what mode !!!
- **Guidance**: Check the EFuse pattern test mode against the supported mode list. Update the configuration to use a valid test mode.

### E_RuleViolationStage_02
- **FullCode**: EEF669002
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Stage
- **Code**: 2
- **Level**: Error
- **#Args**: 0
- **Template**: Programming Stage is FTF
- **Guidance**: Verify the EFuse programming stage definition against the expected stage list. Ensure the stage name and sequence are consistent with the test plan.

### W_RuleViolationStage_03
- **FullCode**: WEF669003
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Stage
- **Code**: 3
- **Level**: Warning
- **#Args**: 2
- **Template**: The programming stage is same between {0} & {1}
- **Guidance**: Verify the shadow programming stage definition against the expected stage list. Ensure the shadow stage is correctly configured for the fuse type.

### E_MismatchRow_01
- **FullCode**: EEF361001
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Row
- **Code**: 1
- **Level**: Error
- **#Args**: 4
- **Template**: The count of rows in sheet {0} ({1}) is inconsistent with sheet {2} ({3})
- **Guidance**: Check the EFuse row information against the device fuse map layout. Ensure row addresses and counts are consistent with the specification.

### I_RuleViolationBank_01
- **FullCode**: IEF603001
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Bank
- **Code**: 1
- **Level**: Info
- **#Args**: 1
- **Template**: Tool filter out the bank name {0} which contains the keyword bira!!!
- **Guidance**: Open the EFuse bank definition and inspect it for invalid fields. Verify bank address, width, and bit assignments against the device specification.

### I_RuleViolationBank_02
- **FullCode**: IEF603002
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Bank
- **Code**: 2
- **Level**: Info
- **#Args**: 1
- **Template**: Tool filter out the bank name {0} which job is SLT!!!
- **Guidance**: Open the EFuse bank definition and inspect it for invalid fields. Verify bank address, width, and bit assignments against the device specification.

### W_RuleViolationBank_01
- **FullCode**: WEF603001
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Bank
- **Code**: 1
- **Level**: Warning
- **#Args**: 1
- **Template**: Bank Name ({0}) is not recognized
- **Guidance**: Check the EFuse bank identifier against the defined bank list. Add the missing bank definition or correct the reference to an existing bank.

### W_MismatchRow_01
- **FullCode**: WEF361001
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Row
- **Code**: 1
- **Level**: Warning
- **#Args**: 0
- **Template**: The base row has different value from other base rows
- **Guidance**: Verify the EFuse base voltage against the device electrical specifications. Ensure the voltage level is within the allowed operating range.

### W_InvalidBit_01
- **FullCode**: WEF206001
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Bit
- **Code**: 1
- **Level**: Warning
- **#Args**: 0
- **Template**: Bit width over 1024 is not allowed.
- **Guidance**: Open the EFuse bit definition table and verify the flagged bit entry. Check the bit address, width, and value are within valid ranges.

### W_InvalidValue_01
- **FullCode**: WEF282001
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Value
- **Code**: 1
- **Level**: Warning
- **#Args**: 1
- **Template**: Default Value isn't a Number. Default Value : {0}
- **Guidance**: Open the EFuse definition and verify the default value for the second occurrence. Ensure the value is within the bit-width range for the specified field.

### W_RuleViolationRevision_01
- **FullCode**: WEF660001
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Revision
- **Code**: 1
- **Level**: Warning
- **#Args**: 0
- **Template**: The database revision is judge by tool, please check it!
- **Guidance**: Check the EFuse data revision logic in the tool configuration. Ensure the revision comparison rules and input data are correct.

### W_RuleViolationLimit_01
- **FullCode**: WEF642001
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Limit
- **Code**: 1
- **Level**: Warning
- **#Args**: 0
- **Template**: If Algorithm is base, Low Limit and High Limit and Default Value must be same.
- **Guidance**: Check the EFuse limit value against the specification. Ensure the limit is within the allowed range for the fuse type.

### W_RedundantPattern_01
- **FullCode**: WEF552001
- **EnumErrorCategory**: EFuse
- **EnumErrorBehavior**: Redundant
- **EnumErrorTarget**: Pattern
- **Code**: 1
- **Level**: Warning
- **#Args**: 3
- **Template**: {0} is unused pattern (bank name: {1}, test mode: {2}) !!!
- **Guidance**: Review the EFuse pattern list and remove unused patterns, or add the missing test references to consume them.

## EvsErrorType (`EvsErrorType.cs`)

### E_MissingParameter_01
- **FullCode**: EEV482001
- **EnumErrorCategory**: Evs
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Value
- **Code**: 1
- **Level**: Error
- **#Args**: 3
- **Template**: Missing Parameter in {0}({1}) : {2}
- **Guidance**: Open the EVS sheet and add the required parameter. Check the EVS specification to confirm which parameters are mandatory.

### E_MissVbtModule_01
- **FullCode**: EEV447001
- **EnumErrorCategory**: Evs
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Module
- **Code**: 1
- **Level**: Error
- **#Args**: 2
- **Template**: The VBT function: {0} can not find in {1} library!
- **Guidance**: Verify the VBT module name against the expected module list. Add the missing module or correct the reference.

### E_MissVbtModule_02
- **FullCode**: EEV447002
- **EnumErrorCategory**: Evs
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Module
- **Code**: 2
- **Level**: Error
- **#Args**: 2
- **Template**: The VBT function: {0} can not find in {1} library!
- **Guidance**: Verify the VBT module name against the expected module list for the basic block. Add the missing module definition or correct the reference.

### E_Pattern_01
- **FullCode**: EEV252001
- **EnumErrorCategory**: Evs
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Pattern
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: The pattern : {0} format illegal
- **Guidance**: Verify the pattern name and its definition in the EVS sheet. Check that the pattern exists and is correctly referenced.

## FlowMainErrorType (`FlowMainErrorType.cs`)

### E_InvalidFormat_01
- **FullCode**: EFM229001
- **EnumErrorCategory**: FlowMain
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: No data found in range: A1:J10.
- **Guidance**: Verify that the correct input document was provided and that the Flow Main sheet is not empty.

### E_MissingDocument_01
- **FullCode**: EFM420001
- **EnumErrorCategory**: FlowMain
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Document
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Referenced Flow Main Sheet '{0}' does not exist in the test plan.
- **Guidance**: Verify that the specified Flow Main Sheet exists in the test plan.Check for any sheet name mismatches or missing sheets.

### W_MismatchFlow_01
- **FullCode**: WFM327001
- **EnumErrorCategory**: FlowMain
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Flow
- **Code**: 1
- **Level**: Warning
- **#Args**: 3
- **Template**: The Sheet Name: "{0}"{1} defined for Job "{2}" does not match any existing sheet or source.
- **Guidance**: Verify that the specified mapping matches an existing sheet or source in the test plan.

## HardIpErrorType (`HardIpErrorType.cs`)

### E_ADC_Convertor_01
- **FullCode**: EHP342001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Limit
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: Limit count not enough for ADC convertor use
- **Guidance**: Check the ADC converter definition for the specified pin or instance. Verify the conversion formula, reference voltage, and scaling factors.

### E_ADC_Convertor_02
- **FullCode**: EHP342002
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Limit
- **Code**: 2
- **Level**: Error
- **#Args**: 1
- **Template**: Previous use-limits count:{0} contains simple calc equation usage, need to check
- **Guidance**: Check the ADC converter definition for the specified pin or instance. Verify the conversion formula, reference voltage, and scaling factors.

### E_CalcEqnCheck_01
- **FullCode**: EHP623001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Equation
- **Code**: 1
- **Level**: Error
- **#Args**: 2
- **Template**: CalcEquation :{0} ,convert to Cus_Str_MainProgram:{1} for TTR usage
- **Guidance**: Review the calculation equation for syntax and semantic errors. Ensure all referenced variables are defined and the result is within expected bounds.

### E_CalcEqnCheck_02
- **FullCode**: EHP623002
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Equation
- **Code**: 2
- **Level**: Error
- **#Args**: 2
- **Template**: CalcEquation :{0} , can not convert to Cus_Str_MainProgram:{1} for TTR usage, need check
- **Guidance**: Review the calculation equation for syntax and semantic errors. Ensure all referenced variables are defined and the result is within expected bounds.

### E_CanNotGetSelsramSetting_01
- **FullCode**: EHP463001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Setting
- **Code**: 1
- **Level**: Error
- **#Args**: 3
- **Template**: Can not get selsrm setting by Block: {0}; pattern: {1} in {2}.
- **Guidance**: Verify the SELSRAM instance name and configuration. Ensure the SELSRAM setting is properly defined and accessible for the specified instance.

### E_DuplicatePatternInPatInfo_01
- **FullCode**: EHP153001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Duplicate
- **EnumErrorTarget**: PatternInfo
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Error: {0} is Duplicated in PatInfoFile.txt
- **Guidance**: Check the PatInfo sheet for duplicate pattern entries. Remove the duplicate or rename one entry so each pattern name is unique.

### E_DuplicateStoreName_01
- **FullCode**: EHP170001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Duplicate
- **EnumErrorTarget**: StoreName
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: We found duplicate Store name "{0}" in HardIP multiple pattern.
- **Guidance**: Check the Hard IP store name list for duplicate entries. Rename one of the duplicate store names so each name is unique.

### E_HIPeFuseCatename_01
- **FullCode**: EHP401001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Argument
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Cannot find "{0}" in EFUSE_BitDef_Table
- **Guidance**: Verify the Hard IP eFuse category name against the EFuse category list. Check for typos or missing category definitions.

### E_HIPeFuseDSPWAVE_01
- **FullCode**: EHP622001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: DSPWave
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Not allow syntax:{0}
- **Guidance**: Verify the Hard IP eFuse DSP wave configuration. Ensure the DSP wave name and parameters are consistent with the eFuse specification.

### E_HIPeFuseDSPWAVE_02
- **FullCode**: EHP322002
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: DSPWave
- **Code**: 2
- **Level**: Error
- **#Args**: 3
- **Template**: "{0}" dspwave size :{1} is not same to efuse_bitdef: {2}
- **Guidance**: Verify the Hard IP eFuse DSP wave configuration. Ensure the DSP wave name and parameters are consistent with the eFuse specification.

### E_HIPeFuseInput_01
- **FullCode**: EHP420001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Document
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: Please provide EFUSE_BitDef_Table to do the cross-checking of efuse and hardip.
- **Guidance**: Verify the Hard IP eFuse input definition against the eFuse specification. Check the input name, bit address, and allowed values.

### E_HIPeFuseInput_02
- **FullCode**: EHP420002
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Document
- **Code**: 2
- **Level**: Error
- **#Args**: 0
- **Template**: Please provide eFuse_HardIP_Table to do the cross-checking of efuse and hardip.
- **Guidance**: Verify the Hard IP eFuse input definition against the eFuse specification. Check the input name, bit address, and allowed values.

### E_HIPeFuseInput_03
- **FullCode**: EHP420003
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Document
- **Code**: 3
- **Level**: Error
- **#Args**: 1
- **Template**: {0} func but not found "EFUSE_BitDef_Table"
- **Guidance**: Verify the Hard IP eFuse input definition against the eFuse specification. Check the input name, bit address, and allowed values.

### E_HIPeFuseInput_04
- **FullCode**: EHP420004
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Document
- **Code**: 4
- **Level**: Error
- **#Args**: 1
- **Template**: {0} func cannot find "m_catename" or "fieName"(.NET)
- **Guidance**: Verify the Hard IP eFuse input definition against the eFuse specification. Check the input name, bit address, and allowed values.

### E_HIPeFuseInput_05
- **FullCode**: EHP420005
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Document
- **Code**: 5
- **Level**: Error
- **#Args**: 1
- **Template**: {0} func cannot find "DSPWAVESIZE" or "sampleSize"(.NET)
- **Guidance**: Verify the Hard IP eFuse input definition against the eFuse specification. Check the input name, bit address, and allowed values.

### E_HIPeFuseInput_06
- **FullCode**: EHP420006
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Document
- **Code**: 6
- **Level**: Error
- **#Args**: 1
- **Template**: {0} func but not found "EFUSE_BitDef_Table"
- **Guidance**: Verify the Hard IP eFuse input definition against the eFuse specification. Check the input name, bit address, and allowed values.

### E_IgxlLimitation_01
- **FullCode**: EHP636001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Igxl
- **Code**: 1
- **Level**: Error
- **#Args**: 2
- **Template**: IG-XL 9.0 need to keep the String length < 8000 characters ;Arg:{0};Value:{1}
- **Guidance**: Review the flagged IGXL limitation and restructure the test to stay within the allowed bounds. Consult the IGXL platform documentation for the specific constraint.

### E_MisCalculationParaDefine_01
- **FullCode**: EHP470001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: StoreName
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Missing Dictionary key in calculation does not exist above test items : {0}
- **Guidance**: Add the missing calculation parameter definition. Please make sure either pin or storename exist in calculation

### E_MismatchFieldJob_01
- **FullCode**: EHP339001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Job
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: The job of the {0} is not match to the Efuse_Bit_Def definition.
- **Guidance**: Compare the field job assignment against the expected job definition. Update the Hard IP sheet so the field job is consistent with the specification.

### E_MismatchFieldWidth_01
- **FullCode**: EHP306001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Bit
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: The bit width of the {0} is not match to the Efuse_Bit_Def definition.
- **Guidance**: Compare the field width in the Hard IP sheet against the EFuse bit definition. Correct the width so both definitions are consistent.

### E_MismatchLimitUnit_01
- **FullCode**: EHP380001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Unit
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: Measure type is different from limit unit
- **Guidance**: Check the unit specification in the limit definition. Ensure the unit matches the measurement type (e.g. V, A, Hz).

### E_MismatchRealValue_01
- **FullCode**: EHP324001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Field
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: The field of the {0} is not real, please check with the Efuse_Bit_Def definition.
- **Guidance**: Compare the real value in the Hard IP sheet against the expected definition. Correct the value in the Hard IP sheet to match the specification.

### E_MismatchSequenceAndCallSubrsCnt_01
- **FullCode**: EHP352001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Pattern
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Please Check the pattern "{0}" , the sequence count and call subrs cnt is not match.
- **Guidance**: Check that the number of measurement sequence entries equals the number of call subroutine entries. Add or remove entries to restore the count balance.

### E_MisPatternForMeasument_01
- **FullCode**: EHP452001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pattern
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: Missing pattern for measuments
- **Guidance**: Check whether the pattern is configured for measurement operations. Use a pattern that includes measurement sequences for this test.

### E_MissingBinNum_01
- **FullCode**: EHP404001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Bin
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: Missing bin number setting
- **Guidance**: Open the bin definition and assign a bin number for the specified test. Ensure the bin number is unique and within the valid bin range.

### E_MissingCalcFunction_01
- **FullCode**: EHP270001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: StoreName
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Missing Dictionary key in calculation does not exist above calculation funtion : {0}
- **Guidance**: Add the missing calculation function definition to the Hard IP configuration. Verify the function name matches the reference in the sheet.

### E_MissingCallSubroutine_01
- **FullCode**: EHP252001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Pattern
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Please Check the pattern "{0}" have no call_subr but have meas seq info.
- **Guidance**: Add the required call subroutine entry for the pattern in PatInfo. Ensure the subroutine name matches the one used in the pattern file.

### E_MissingFieldName_01
- **FullCode**: EHP401002
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Argument
- **Code**: 2
- **Level**: Error
- **#Args**: 1
- **Template**: Can not find {0} in Efuse_Bit_Def definition.
- **Guidance**: Open the Hard IP eFuse sheet and add the required field name. Verify the field name matches the EFuse bit definition.

### E_MissingHardIpCategory_01
- **FullCode**: EHP418001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: DcCategory
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Can not find HardIp category: {0} from TestSettting sheet
- **Guidance**: Add the missing Hard IP category definition. Verify the category name against the expected Hard IP configuration.

### E_MissingHardIpDcPin_01
- **FullCode**: EHP455001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: Missing HardIpDc pin in Pinmap
- **Guidance**: Add the missing DC pin to the Hard IP pin list. Verify the pin name against the device pin map.

### E_MissingHardIpSheet_01
- **FullCode**: EHP282001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Value
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Here is one HardIP sheet :"{0}" but sheet name is not start with "HardIP_"
- **Guidance**: Add "HardIP_" prefix to the sheet to solve the problem

### E_MissingHeader_01
- **FullCode**: EHP433001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Header
- **Code**: 1
- **Level**: Error
- **#Args**: 2
- **Template**: Missing header: {0} in sheet {1}
- **Guidance**: Open the sheet and add the required header column. Compare the column headers against the specification to identify which ones are missing.

### E_MissingInPreWrite_01
- **FullCode**: EHP252002
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Pattern
- **Code**: 2
- **Level**: Error
- **#Args**: 1
- **Template**: The field {0} is not defined in the prewrite HardIP testplan.
- **Guidance**: Add the missing entry to the pre-write definition in the Hard IP sheet. Ensure all required fields are present before the write sequence.

### E_MissingMeasureSequence_01
- **FullCode**: EHP252003
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Pattern
- **Code**: 3
- **Level**: Error
- **#Args**: 1
- **Template**: Please Check the pattern "{0}" have no meas seq info but have call_subr.
- **Guidance**: Open the PatInfo sheet and add the required measurement sequence for the pattern. Verify the sequence is needed and add all required measurement entries.

### E_MissingNeededSheets_02
- **FullCode**: EHP420007
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Document
- **Code**: 7
- **Level**: Error
- **#Args**: 1
- **Template**: Missing {0} sheet in TestPlan
- **Guidance**: Add the missing sheet to the workbook or check whether it was renamed. Compare the workbook contents against the expected sheet list.

### E_MissingParameter_01
- **FullCode**: EHP401003
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Argument
- **Code**: 3
- **Level**: Error
- **#Args**: 2
- **Template**: Missing Parameter in {0} : {1} Or wrong key word in misc-info
- **Guidance**: Open the Hard IP sheet and add the required parameter. Check the Hard IP specification to confirm which parameters are mandatory.

### E_MissingParameter_03
- **FullCode**: EHP401007
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Argument
- **Code**: 7
- **Level**: Error
- **#Args**: 3
- **Template**: Missing Parameter in {0}({1}) : {2}
- **Guidance**: Open the Hard IP sheet and add the required parameter. Check the Hard IP specification to confirm which parameters are mandatory.

### E_MissingPatInfoPinInPinMap_01
- **FullCode**: EHP455002
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 2
- **Level**: Error
- **#Args**: 1
- **Template**: Missing MeasC pin in Pinmap:{0}
- **Guidance**: Check the pin map and add the missing PatInfo pin entry. Ensure the pin name in PatInfo exactly matches the pin map definition.

### E_MissingPatternInPatInfo_01
- **FullCode**: EHP253001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: PatternInfo
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Missing patten in PatInfo file: {0}
- **Guidance**: Add the missing pattern definition to the PatInfo sheet. Ensure all patterns used in the test plan have a corresponding PatInfo entry.

### E_MissingPatternInTestPlan_01
- **FullCode**: EHP252004
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Pattern
- **Code**: 4
- **Level**: Error
- **#Args**: 0
- **Template**: Missing pattern in "Pattern" Column
- **Guidance**: Add the missing pattern entry to the test plan. Verify the pattern file exists and is included in the pattern compilation step.

### E_MissingPinName_01
- **FullCode**: EHP455003
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 3
- **Level**: Error
- **#Args**: 1
- **Template**: Missing Hi limit pin in Pinmap : {0}
- **Guidance**: Open the Hard IP sheet and fill in the pin name for the flagged row. All pin rows must have a non-empty pin name.

### E_MissingPinName_02
- **FullCode**: EHP455004
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 4
- **Level**: Error
- **#Args**: 1
- **Template**: Missing Lo limit pin in Pinmap : {0}
- **Guidance**: Open the Hard IP sheet and fill in the pin name for the flagged row. All pin rows must have a non-empty pin name.

### E_MissingPinName_03
- **FullCode**: EHP455005
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 5
- **Level**: Error
- **#Args**: 1
- **Template**: Missing force pin in Pinmap : {0}
- **Guidance**: Open the Hard IP sheet and fill in the pin name for the flagged row. All pin rows must have a non-empty pin name.

### E_MissingPinName_04
- **FullCode**: EHP455006
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 6
- **Level**: Error
- **#Args**: 1
- **Template**: Missing sweep pin in Pinmap : {0}
- **Guidance**: Open the Hard IP sheet and fill in the pin name for the flagged row. All pin rows must have a non-empty pin name.

### E_MissingPinName_05
- **FullCode**: EHP455007
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 7
- **Level**: Error
- **#Args**: 1
- **Template**: Missing meas pins in Pinmap : {0}
- **Guidance**: Open the Hard IP sheet and fill in the pin name for the flagged row. All pin rows must have a non-empty pin name.

### E_MissingPinName_06
- **FullCode**: EHP455008
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 8
- **Level**: Error
- **#Args**: 1
- **Template**: Missing force pins of MeasPin in Pin map : {0}
- **Guidance**: Open the Hard IP sheet and fill in the pin name for the flagged row. All pin rows must have a non-empty pin name.

### E_MissingPinName_07
- **FullCode**: EHP455009
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 9
- **Level**: Error
- **#Args**: 0
- **Template**: Empty Meas PinName but contains Limits
- **Guidance**: Open the Hard IP sheet and fill in the pin name for the flagged row. All pin rows must have a non-empty pin name.

### E_MissingPinName_08
- **FullCode**: EHP455010
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pin
- **Code**: 10
- **Level**: Error
- **#Args**: 1
- **Template**: Missing "Ignore_Flow_Limit" in pin's misc info to ingore limit for trim instance : {0}
- **Guidance**: Open the Hard IP sheet and fill in the pin name for the flagged row. All pin rows must have a non-empty pin name.

### E_MissingSweepSingleloopTable_01
- **FullCode**: EHP420008
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Document
- **Code**: 8
- **Level**: Error
- **#Args**: 1
- **Template**: Missing sweep or single-loop table definition sheet {0} in TestPlan
- **Guidance**: Add the missing sweep or single-loop table definition. Verify the table name against the expected Hard IP register configuration.

### E_MissVbtModule_02
- **FullCode**: EHP447002
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Module
- **Code**: 2
- **Level**: Error
- **#Args**: 1
- **Template**: The function: {0} can not find in library!
- **Guidance**: Verify the VBT module name against the expected module list. Add the missing module definition or correct the reference.

### E_OppositeLimit_01
- **FullCode**: EHP242001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Limit
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: Opposite limit value
- **Guidance**: Check that the lower limit is less than the upper limit. Swap the values if they were accidentally reversed.

### E_PatternExistMeasument_01
- **FullCode**: EHP661001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Row
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: pattern row exist measuments
- **Guidance**: Check the pattern list for an existing measurement pattern with the same name. Remove the duplicate reference or rename the new pattern.

### E_RegAssignError_01
- **FullCode**: EHP602001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Assignment
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Register Assign Type {0} is not support to split
- **Guidance**: Open the register assignment definition and correct the error. Verify the assignment value is within the allowed range and format.

### E_RegAssignError_02
- **FullCode**: EHP302002
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Assignment
- **Code**: 2
- **Level**: Error
- **#Args**: 1
- **Template**: The count of Register Assign {0} differs from measureSequence
- **Guidance**: Open the register assignment definition and correct the error. Verify the assignment value is within the allowed range and format.

### E_RegAssignError_03
- **FullCode**: EHP202003
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Assignment
- **Code**: 3
- **Level**: Error
- **#Args**: 1
- **Template**: Register Assign Type {0} is not support to split
- **Guidance**: Open the register assignment definition and correct the error. Verify the assignment value is within the allowed range and format.

### E_RepeatSubBlock_01
- **FullCode**: EHP171001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Duplicate
- **EnumErrorTarget**: Subblock
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Repeat SubBlock Name:{0}
- **Guidance**: Check the Hard IP sub-block list for duplicate entries. Remove the repeated sub-block reference or consolidate the duplicates.

### E_SelsramDigSrcAssignmentNotDefineInInstance_01
- **FullCode**: EHP202001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Assignment
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: SELSRAM digital source assignment '{0}' is not defined in the instance.
- **Guidance**: Add the missing SELSRAM digital source assignment to the instance definition. Ensure the assignment matches the one referenced in the mapping table.

### E_SelsramDigSrcAssignmentNotDefineInTable_01
- **FullCode**: EHP202002
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Assignment
- **Code**: 2
- **Level**: Error
- **#Args**: 4
- **Template**: Block: {0} Pattern: {1} Alpha: {2} of DigSrc_Assignment is not defined in {3}.
- **Guidance**: Add the missing SELSRAM digital source assignment to the mapping table. Verify the assignment name against the table definition.

### E_SelsramMappingTableError_01
- **FullCode**: EHP420009
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Document
- **Code**: 9
- **Level**: Error
- **#Args**: 0
- **Template**: Can't get SELSRAM_Mapping_Table in the test plan.
- **Guidance**: Check the SELSRAM mapping table for invalid or inconsistent entries. Verify the mapping conforms to the SELSRAM specification.

### E_SelsramMappingTableError_02
- **FullCode**: EHP411002
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Column
- **Code**: 2
- **Level**: Error
- **#Args**: 0
- **Template**: DigSrc_Assignment column didn't exist in SELSRAM_Mapping_Table.
- **Guidance**: Check the SELSRAM mapping table for invalid or inconsistent entries. Verify the mapping conforms to the SELSRAM specification.

### E_SubBlockNotFound_01
- **FullCode**: EHP471001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Subblock
- **Code**: 1
- **Level**: Error
- **#Args**: 2
- **Template**: SubBlock :{0}, not found in {1} sheet.
- **Guidance**: Verify that the sub-block name is spelled correctly and defined in the Hard IP sheet. Add the missing sub-block or correct the reference.

### E_SubBlockUsage_01
- **FullCode**: EHP271001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Subblock
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Invalid SubBlock Usage:{0}
- **Guidance**: Check the sub-block usage in the Hard IP configuration. Ensure the sub-block is referenced correctly and all required usage rules are satisfied.

### E_WrongDigSrcSignalName_01
- **FullCode**: EHP353001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: PatternInfo
- **Code**: 1
- **Level**: Error
- **#Args**: 2
- **Template**: DigSrc Signal Name "{0}" is different from VM_Vector "{1} " 
- **Guidance**: Verify the digital source signal name against the defined signal list. Check for typos or renamed signals in the source definition.

### E_WrongForceCondition_01
- **FullCode**: EHP228001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Force
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Wrong Force Type for InterPose_PreMeas:{0}
- **Guidance**: Open the Hard IP force condition definition and verify the values. Ensure the condition is compatible with the instance type and test requirements.

### E_WrongForceCondition_02
- **FullCode**: EHP228002
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Force
- **Code**: 2
- **Level**: Error
- **#Args**: 1
- **Template**: Wrong Force Type for others force condition:{0}
- **Guidance**: Open the Hard IP force condition definition and verify the values. Ensure the condition is compatible with the instance type and test requirements.

### E_WrongForceCondition_03
- **FullCode**: EHP628003
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Force
- **Code**: 3
- **Level**: Error
- **#Args**: 1
- **Template**: {0} could not support
- **Guidance**: Open the Hard IP force condition definition and verify the values. Ensure the condition is compatible with the instance type and test requirements.

### E_WrongForceCondition_04
- **FullCode**: EHP228004
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Force
- **Code**: 4
- **Level**: Error
- **#Args**: 1
- **Template**: Wrong ForceCondition for {0}
- **Guidance**: Open the Hard IP force condition definition and verify the values. Ensure the condition is compatible with the instance type and test requirements.

### E_WrongForceCondition_05
- **FullCode**: EHP355005
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Pin
- **Code**: 5
- **Level**: Error
- **#Args**: 1
- **Template**: Wrong pins of AC special setting: {0}
- **Guidance**: Open the Hard IP force condition definition and verify the values. Ensure the condition is compatible with the instance type and test requirements.

### E_WrongForceCondition_06
- **FullCode**: EHP229006
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 6
- **Level**: Error
- **#Args**: 1
- **Template**: Wrong voltage format of AcSelector setting: {0}
- **Guidance**: Open the Hard IP force condition definition and verify the values. Ensure the condition is compatible with the instance type and test requirements.

### E_WrongForceCondition_07
- **FullCode**: EHP229007
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 7
- **Level**: Error
- **#Args**: 1
- **Template**: Wrong type format of AcSelector setting: {0}
- **Guidance**: Open the Hard IP force condition definition and verify the values. Ensure the condition is compatible with the instance type and test requirements.

### E_WrongForceCondition_08
- **FullCode**: EHP229008
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 8
- **Level**: Error
- **#Args**: 1
- **Template**: Wrong format of AcSelector setting: {0}
- **Guidance**: Open the Hard IP force condition definition and verify the values. Ensure the condition is compatible with the instance type and test requirements.

### E_WrongForceCondition_09
- **FullCode**: EHP228009
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Force
- **Code**: 9
- **Level**: Error
- **#Args**: 1
- **Template**: Wrong ForceCondition for {0}
- **Guidance**: Open the Hard IP force condition definition and verify the values. Ensure the condition is compatible with the instance type and test requirements.

### E_WrongForceCondition_10
- **FullCode**: EHP228010
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Force
- **Code**: 10
- **Level**: Error
- **#Args**: 1
- **Template**: Wrong ForceCondition for {0}
- **Guidance**: Open the Hard IP force condition definition and verify the values. Ensure the condition is compatible with the instance type and test requirements.

### E_WrongForceCondition_11
- **FullCode**: EHP229011
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 11
- **Level**: Error
- **#Args**: 0
- **Template**: Wrong format of force condition in Test Plan
- **Guidance**: Open the Hard IP force condition definition and verify the values. Ensure the condition is compatible with the instance type and test requirements.

### E_WrongLimitValue_01
- **FullCode**: EHP242002
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Limit
- **Code**: 2
- **Level**: Error
- **#Args**: 0
- **Template**: Unrecognied limit value
- **Guidance**: Check the limit value against the device specification. Ensure the value does not exceed the allowed min/max bounds.

### E_WrongMeasC_01
- **FullCode**: EHP253002
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: PatternInfo
- **Code**: 2
- **Level**: Error
- **#Args**: 1
- **Template**: Can't get cap pin from pattern info, pattern : {0}
- **Guidance**: Compare the MeasC pin assignment with the cap pin defined in PatInfo. Update one of the definitions so both references point to the same physical pin.

### E_WrongMeasC_02
- **FullCode**: EHP252005
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Pattern
- **Code**: 5
- **Level**: Error
- **#Args**: 0
- **Template**: Define MeasC, but not define number of capture bits
- **Guidance**: Compare the MeasC pin assignment with the cap pin defined in PatInfo. Update one of the definitions so both references point to the same physical pin.

### E_WrongMeasC_03
- **FullCode**: EHP346003
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Measurement
- **Code**: 3
- **Level**: Error
- **#Args**: 2
- **Template**: MeasC capture bit mismatch between test plan {0} and patInfo {1}
- **Guidance**: Compare the MeasC pin assignment with the cap pin defined in PatInfo. Update one of the definitions so both references point to the same physical pin.

### E_WrongMeasContent_01
- **FullCode**: EHP246001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Measurement
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: Wrong measure data in 'Meas' column
- **Guidance**: Check the measurement content field in PatInfo for the specified entry. Ensure the content value is within the allowed set for this measurement type.

### E_WrongMeasContent_02
- **FullCode**: EHP246002
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Measurement
- **Code**: 2
- **Level**: Error
- **#Args**: 1
- **Template**: CusStr of"{0}" is not defined in this test item, and default calc type would be generated with C
- **Guidance**: Check the measurement content field in PatInfo for the specified entry. Ensure the content value is within the allowed set for this measurement type.

### E_WrongMeasPinInPatInfo_01
- **FullCode**: EHP355001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Pin
- **Code**: 1
- **Level**: Error
- **#Args**: 2
- **Template**: MeasC pin:{0} is different from pat info cap pin :{1}
- **Guidance**: Verify the measurement pin name against the pin map and PatInfo definition. Ensure the pin is a valid measurement pin for this Hard IP type.

### E_WrongMeasPinInPatInfo_02
- **FullCode**: EHP355002
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Pin
- **Code**: 2
- **Level**: Error
- **#Args**: 1
- **Template**: pattern info pin:{0}, can't match to test plan
- **Guidance**: Verify the measurement pin name against the pin map and PatInfo definition. Ensure the pin is a valid measurement pin for this Hard IP type.

### E_WrongMeasType_01
- **FullCode**: EHP246003
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Measurement
- **Code**: 3
- **Level**: Error
- **#Args**: 1
- **Template**: WrongMeasType of '{0}'
- **Guidance**: Check the measurement type assignment in the PatInfo sheet. Ensure the type is supported for the specified pin and Hard IP category.

### E_WrongRegisterAssignment_01
- **FullCode**: EHP302001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Assignment
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: One of RegisterAssignment or digsrc of pattern does not have information, please check
- **Guidance**: Open the register assignment definition and verify the value against the register specification. Ensure the assigned value is within the bit-width and allowed range.

### E_WrongRegisterAssignment_02
- **FullCode**: EHP306002
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Bit
- **Code**: 2
- **Level**: Error
- **#Args**: 3
- **Template**: Send bit: "{0}", Mismatch bit width between pattern send bit: {1} and Register Assignment in test plan: {2} .
- **Guidance**: Open the register assignment definition and verify the value against the register specification. Ensure the assigned value is within the bit-width and allowed range.

### E_WrongRegisterAssignment_04
- **FullCode**: EHP670004
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: StoreName
- **Code**: 4
- **Level**: Error
- **#Args**: 1
- **Template**: "{0}" is not defined before use it in register assignment
- **Guidance**: Open the register assignment definition and verify the value against the register specification. Ensure the assigned value is within the bit-width and allowed range.

### E_WrongRegisterAssignment_05
- **FullCode**: EHP302005
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Assignment
- **Code**: 5
- **Level**: Error
- **#Args**: 1
- **Template**: {0}
- **Guidance**: Open the register assignment definition and verify the value against the register specification. Ensure the assigned value is within the bit-width and allowed range.

### E_WrongRegisterAssignment_06
- **FullCode**: EHP202006
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Assignment
- **Code**: 6
- **Level**: Error
- **#Args**: 1
- **Template**: There is a Sweep format issue : {0}
- **Guidance**: Open the register assignment definition and verify the value against the register specification. Ensure the assigned value is within the bit-width and allowed range.

### E_WrongRegisterAssignment_07
- **FullCode**: EHP302007
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Assignment
- **Code**: 7
- **Level**: Error
- **#Args**: 2
- **Template**: Burst register assignment length {0} is different from burst pattern count {1}
- **Guidance**: Open the register assignment definition and verify the value against the register specification. Ensure the assigned value is within the bit-width and allowed range.

### E_WrongRegisterAssignment_08
- **FullCode**: EHP253008
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: PatternInfo
- **Code**: 8
- **Level**: Error
- **#Args**: 2
- **Template**: Can't get assignment: {0} from pattern info: {1}
- **Guidance**: Open the register assignment definition and verify the value against the register specification. Ensure the assigned value is within the bit-width and allowed range.

### E_WrongRegisterAssignment_09
- **FullCode**: EHP302009
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Assignment
- **Code**: 9
- **Level**: Error
- **#Args**: 2
- **Template**: Lose pattern info assignment in plan: {0} from pattern info: {1}
- **Guidance**: Open the register assignment definition and verify the value against the register specification. Ensure the assigned value is within the bit-width and allowed range.

### E_WrongRegisterAssignment_10
- **FullCode**: EHP302010
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Assignment
- **Code**: 10
- **Level**: Error
- **#Args**: 2
- **Template**: # of registers between patinfo {0} and test plan {1} not match
- **Guidance**: Open the register assignment definition and verify the value against the register specification. Ensure the assigned value is within the bit-width and allowed range.

### E_WrongRegisterAssignment_11
- **FullCode**: EHP302011
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Assignment
- **Code**: 11
- **Level**: Error
- **#Args**: 1
- **Template**: patinfo contains registers: {0}, not exist in test plan
- **Guidance**: Open the register assignment definition and verify the value against the register specification. Ensure the assigned value is within the bit-width and allowed range.

### E_WrongRegisterAssignment_12
- **FullCode**: EHP302012
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Assignment
- **Code**: 12
- **Level**: Error
- **#Args**: 1
- **Template**: test plan contains registers: {0}, not exist in patinfo
- **Guidance**: Open the register assignment definition and verify the value against the register specification. Ensure the assigned value is within the bit-width and allowed range.

### E_WrongSendInformation_01
- **FullCode**: EHP306003
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Bit
- **Code**: 3
- **Level**: Error
- **#Args**: 1
- **Template**: Same name in "Send Bit Name" but with different num of bit in "Send Bit Str" -- {0}
- **Guidance**: Check the send information fields in PatInfo for the specified pattern. Verify all values conform to the expected format and bit-width constraints.

### E_WrongSweepStep_01
- **FullCode**: EHP273001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Sweep
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: The step for sweep voltage is incorrect due to {0}
- **Guidance**: Check the sweep step value against the specification. Ensure the step size is positive, within range, and compatible with the sweep range.

### E_WrongTimeSet_01
- **FullCode**: EHP378001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Tset
- **Code**: 1
- **Level**: Error
- **#Args**: 2
- **Template**: The Payload of multiple init : {0} is not the same timeSet with pattern : {1}, please check
- **Guidance**: Verify the time set name against the defined time set list. Ensure the time set is compatible with the Hard IP instance requirements.

### I_DuplicateStoreName_01
- **FullCode**: IHP170001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Duplicate
- **EnumErrorTarget**: StoreName
- **Code**: 1
- **Level**: Info
- **#Args**: 1
- **Template**: We found duplicate Store name "{0}" in HardIP single pattern.
- **Guidance**: Check the Hard IP store name list for duplicate entries. Rename one of the duplicate store names so each name is unique.

### I_UnsupportedOpcode_01
- **FullCode**: IHP249001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Opcode
- **Code**: 1
- **Level**: Info
- **#Args**: 1
- **Template**: We have found an invalid Opcode:{0}.
- **Guidance**: Check the IGXL version for supported opcodes. Replace the unsupported opcode with a supported equivalent.

### W_BurstPatJudgeFlag_02
- **FullCode**: WHP652002
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Pattern
- **Code**: 2
- **Level**: Warning
- **#Args**: 2
- **Template**: {0}: Burst pattern judged by {1}
- **Guidance**: Verify the burst pattern judge flag definition. Ensure the flag name and condition are valid and match the burst pattern flow.

### W_ClearLimit_01
- **FullCode**: WHP642001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Limit
- **Code**: 1
- **Level**: Warning
- **#Args**: 0
- **Template**: Limit clear by tool to let VBT apply limit directly
- **Guidance**: Verify whether the limit clearing is intentional. If not, restore the limit definition from the specification.

### W_DuplicateInstance_01
- **FullCode**: WHP137001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Duplicate
- **EnumErrorTarget**: Instance
- **Code**: 1
- **Level**: Warning
- **#Args**: 0
- **Template**: We found duplicate instance but we add index to avoid validate fail, please check
- **Guidance**: Check the instance list for duplicate entries. Remove or rename the duplicate so each instance name is unique.

### W_ExistedSubBlock_01
- **FullCode**: WHP171001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Duplicate
- **EnumErrorTarget**: Subblock
- **Code**: 1
- **Level**: Warning
- **#Args**: 1
- **Template**: SubBlock:{0}, existed on other sheet
- **Guidance**: Check whether the sub-block has already been defined. Remove the duplicate definition or update the existing one.

### W_FuseWriteJudgeFlag_01
- **FullCode**: WHP602002
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Assignment
- **Code**: 2
- **Level**: Warning
- **#Args**: 1
- **Template**: Need to perform eFuse pre-write, but flag is missing in MiscInfo, Autogen judged that by the latest fail flag: {0}
- **Guidance**: Verify the fuse write judge flag definition. Ensure the flag name and condition are valid and consistent with the fuse write flow.

### W_ManualItems_01
- **FullCode**: WHP652001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Pattern
- **Code**: 1
- **Level**: Warning
- **#Args**: 0
- **Template**: Manual items that Autogen cannot support yet!
- **Guidance**: Review the flagged manual item and take the described action. Manual items are reminders for configuration steps that cannot be automated.

### W_MissingLimit_01
- **FullCode**: WHP442001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Limit
- **Code**: 1
- **Level**: Warning
- **#Args**: 0
- **Template**: Measure pin had no limit
- **Guidance**: Add the required limit definition for the specified test. Check whether the test requires upper, lower, or both limits.

### W_MissingLimitUnit_01
- **FullCode**: WHP480001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Unit
- **Code**: 1
- **Level**: Warning
- **#Args**: 0
- **Template**: Measure pin had no limit unit
- **Guidance**: Open the limit definition and add the required unit. All limits must specify a unit consistent with the measurement type.

### W_MissingPinSeq_01
- **FullCode**: WHP346001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Measurement
- **Code**: 1
- **Level**: Warning
- **#Args**: 0
- **Template**: No measure sequence in PatInfo file, but specified measure pins in testPlan
- **Guidance**: Add the missing pin sequence definition. Ensure the sequence covers all required pins in the correct order.

### W_MissingRealFieldNameInEfuseHardIP_01
- **FullCode**: WHP324001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Field
- **Code**: 1
- **Level**: Warning
- **#Args**: 1
- **Template**: {0} is real value in Efuse_Bit_Def but can not find in eFuse_HardIP_Table.
- **Guidance**: Add the missing real field name to the eFuse Hard IP sheet. Verify the field name matches exactly the name in the EFuse bit definition.

### W_MissingTestplanPin_01
- **FullCode**: WHP355003
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Pin
- **Code**: 3
- **Level**: Warning
- **#Args**: 2
- **Template**: this pin :{0} can not find in pattern info {1}
- **Guidance**: Please compare the pin you assigned with pins in PatternInfo.

### W_NoBurstSubBlock_01
- **FullCode**: WHP471001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Subblock
- **Code**: 1
- **Level**: Warning
- **#Args**: 1
- **Template**: The SubBlock:{0}, not be referenced by burst item
- **Guidance**: Add a burst sub-block definition for the specified instance. Check whether the instance requires a burst configuration and add it accordingly.

### W_RegAssignWarning_01
- **FullCode**: WHP652003
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Pattern
- **Code**: 3
- **Level**: Warning
- **#Args**: 1
- **Template**: The legnth of {0} content is over 6000, move to reg assign table
- **Guidance**: Review the register assignment warning. Confirm the assignment is intentional; it may produce unexpected device behavior.

### W_RelayRestore_01
- **FullCode**: WHP652004
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Pattern
- **Code**: 4
- **Level**: Warning
- **#Args**: 1
- **Template**: RelayOn/Off would not restore in PrePat, suggest fill in Misc info with 'RelayOn:{0}' or 'RelayOff:{0}'
- **Guidance**: If use [PinName]:RelayOn in force condition, highly recommended to use RelayOff:[PinName] in MiscInfo to restore relay state. Vise versa

### W_RepeatSubBlock_01
- **FullCode**: WHP271001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Subblock
- **Code**: 1
- **Level**: Warning
- **#Args**: 1
- **Template**: Invalid SubBlock Name:{0}
- **Guidance**: Check the Hard IP sub-block list for duplicate entries. Remove the repeated sub-block reference or consolidate the duplicates.

### W_UnrecognisedHeader_01
- **FullCode**: WHP233001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Header
- **Code**: 1
- **Level**: Warning
- **#Args**: 1
- **Template**: Unrecognied header : {0}
- **Guidance**: Remove or rename the unrecognised column header. Check the specification for the list of accepted header names.

### W_WrongMeasC_01
- **FullCode**: WHP170001
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Duplicate
- **EnumErrorTarget**: StoreName
- **Code**: 1
- **Level**: Warning
- **#Args**: 2
- **Template**: The capture store name : {0} is repeated with {1} times in this item! 
- **Guidance**: Compare the MeasC pin assignment with the cap pin defined in PatInfo. Update one of the definitions so both references point to the same physical pin.

### W_WrongPatternName_01
- **FullCode**: WHP652005
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Pattern
- **Code**: 5
- **Level**: Warning
- **#Args**: 0
- **Template**: Patterns that not start with “dd_”, ”cz_” or “pp_” will be ignored.
- **Guidance**: Please make sure again. Do you really want to use pattern naming that will be ignored by default?

### W_WrongRegisterAssignment_01
- **FullCode**: WHP602003
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Assignment
- **Code**: 3
- **Level**: Warning
- **#Args**: 2
- **Template**: start:{0} is larger than stop:{1}
- **Guidance**: RegisterAssign syntax <KEY>:<START>:<END> or <KEY>[<START>:<END>]. <START> must not larger than <END>

### W_WrongTotalMeasCount_01
- **FullCode**: WHP355004
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Pin
- **Code**: 4
- **Level**: Warning
- **#Args**: 2
- **Template**: Total MeasCount different, TestPlan:{0}, PatInfo:{1}
- **Guidance**: Compare the expected and actual measurement counts in the PatInfo sheet. Add or remove measurement entries to match the expected count.

## HarvestErrorType (`HarvestErrorType.cs`)

### E_InvalidFormat_01
- **FullCode**: EHV229001
- **EnumErrorCategory**: Harvest
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Missing job name in harvesting truth table sheet name: "{0}".
- **Guidance**: Please ensure the sheet name contains the job name.

### E_InvalidFormat_02
- **FullCode**: EHV229002
- **EnumErrorCategory**: Harvest
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 2
- **Level**: Error
- **#Args**: 1
- **Template**: Invalid content "{0}". Only X, numeric values, and ranges are allowed.
- **Guidance**: Please enter only X, a number, or a valid range.

### E_InvalidFormat_03
- **FullCode**: EHV229003
- **EnumErrorCategory**: Harvest
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 3
- **Level**: Error
- **#Args**: 2
- **Template**: Cannot fill range {0} for single flag : {1}.
- **Guidance**: Unable to define the range. Ensure the flag defined in the top row is defined with a valid range.

### E_InvalidOrder_01
- **FullCode**: EHV250001
- **EnumErrorCategory**: Harvest
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Order
- **Code**: 1
- **Level**: Error
- **#Args**: 2
- **Template**: Header order is incorrect. '{0}' must appear after '{1}'.
- **Guidance**: Verify that headers in HarvestingTruthTable table follow the required order.

### E_MissingField_01
- **FullCode**: EHV424001
- **EnumErrorCategory**: Harvest
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Field
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Field {0} with "Real" type was not found in EFUSE_BitDef_Table.
- **Guidance**: Please verify that the field exists in EFUSE_BitDef_Table and is defined with type "Real".

### E_MissingFlag_01
- **FullCode**: EHV426001
- **EnumErrorCategory**: Harvest
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Flag
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Read Fuse flag "{0}" is defined in CP2 but is not defined in this truth table.
- **Guidance**: Please ensure that flag definitions are consistent across the both harvest truth tables.

### E_MissingHeader_01
- **FullCode**: EHV433001
- **EnumErrorCategory**: Harvest
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Header
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Cannot found header: {0}, please check.
- **Guidance**: Open HarvestingTruthTable and ensure the missing column header(s) are present and spelled correctly.

### E_RedundantFlag_01
- **FullCode**: EHV526001
- **EnumErrorCategory**: Harvest
- **EnumErrorBehavior**: Redundant
- **EnumErrorTarget**: Flag
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Read Fuse Flag "{0}" is defined in this truth table but is missing in CP2.
- **Guidance**: Please ensure that flag definitions are consistent across the both harvest truth tables.

### W_InvalidFormat_01
- **FullCode**: WHV229001
- **EnumErrorCategory**: Harvest
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 1
- **Level**: Warning
- **#Args**: 1
- **Template**: Special character(s) are not allowed: {0}.
- **Guidance**: Use only letters, numbers, and underscores (_).

### W_MismatchPattern_01
- **FullCode**: WHV352001
- **EnumErrorCategory**: Harvest
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Pattern
- **Code**: 1
- **Level**: Warning
- **#Args**: 1
- **Template**: Cannot found any matching pattern in Mapping_DigitalCores table. Pattern : "{0}"
- **Guidance**: Because PatternPinGroup is specified in this instance, verify that the corresponding Pattern Name or Pattern Name keyword exists in Mapping_DigitalCores table.

### W_MissingHeader_01
- **FullCode**: WHV433001
- **EnumErrorCategory**: Harvest
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Header
- **Code**: 1
- **Level**: Warning
- **#Args**: 1
- **Template**: Cannot found header: {0}, please check.
- **Guidance**: Open HarvestingTruthTable and ensure the missing column header(s) are present and spelled correctly.

### W_MissingHeader_02
- **FullCode**: WHV433002
- **EnumErrorCategory**: Harvest
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Header
- **Code**: 2
- **Level**: Warning
- **#Args**: 1
- **Template**: Cannot found header "FUSING({0})", please check.
- **Guidance**: Open HarvestingTruthTable and add the FUSING column for the flagged job. Ensure the column name matches the expected format 'FUSING(<job>)'.

## HtolErrorType (`HtolErrorType.cs`)

### E_MissingLibrary_01
- **FullCode**: EHT441001
- **EnumErrorCategory**: Htol
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Library
- **Code**: 1
- **Level**: Error
- **#Args**: 2
- **Template**: The function: {0} can not find in {1} library!
- **Guidance**: Please verify that the function exists in the specified library and that the library version and function reference are correct.

## MbistErrorType (`MbistErrorType.cs`)

### E_BurstInfoMismatch_01
- **FullCode**: EBI384001
- **EnumErrorCategory**: Mbist
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Voltage
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: Column[Voltage]: Same label with previous row, but Voltage mismatch.
- **Guidance**: Open the MBIST sheet and check the Voltage column for the flagged row. Rows sharing the same label must have identical Voltage values.

### E_BurstInfoMismatch_02
- **FullCode**: EBI343002
- **EnumErrorCategory**: Mbist
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Logic
- **Code**: 2
- **Level**: Error
- **#Args**: 0
- **Template**: Column[FailBranch]: Same label with previous row, but FailBranch mismatch.
- **Guidance**: Open the MBIST sheet and check the FailBranch column for the flagged row. Rows sharing the same label must have identical FailBranch values.

### E_BurstInfoMismatch_03
- **FullCode**: EBI376003
- **EnumErrorCategory**: Mbist
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Timing
- **Code**: 3
- **Level**: Error
- **#Args**: 0
- **Template**: Column[TimeSet]: Same label with previous row, but TimeSet mismatch.
- **Guidance**: Open the MBIST sheet and check the TimeSet column for the flagged row. Rows sharing the same label must have identical TimeSet values.

### E_BurstInfoMismatch_04
- **FullCode**: EBI451004
- **EnumErrorCategory**: Mbist
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Parameter
- **Code**: 4
- **Level**: Error
- **#Args**: 0
- **Template**: Argument "patternBeforeWait" for retention will be empty.
- **Guidance**: Check the retention pattern configuration and ensure 'patternBeforeWait' is assigned a valid pattern. Verify the burst info definition includes this argument.

### E_Business_01
- **FullCode**: EBI643005
- **EnumErrorCategory**: Mbist
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Logic
- **Code**: 5
- **Level**: Error
- **#Args**: 2
- **Template**: {0} should be last row in {1}!!!
- **Guidance**: The action column contest for the last row of Mbist sheet need to be "PASS" . 

### E_Business_02
- **FullCode**: EBI452006
- **EnumErrorCategory**: Mbist
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pattern
- **Code**: 6
- **Level**: Error
- **#Args**: 1
- **Template**: Cannot find any payload from pass branch : {0} , selsram will use init pattern category.
- **Guidance**: Review the MBIST business logic constraint that was violated. Check the non-logical data definition and correct the configuration.

### E_Business_03
- **FullCode**: EBI418007
- **EnumErrorCategory**: Mbist
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: DcCategory
- **Code**: 7
- **Level**: Error
- **#Args**: 2
- **Template**: Error! Can not found DC Spec by {0} in {1}
- **Guidance**: DC category defined in flagged row can not be find in voltage table, add DC category into voltage table

### E_MissingParameter_01
- **FullCode**: EBI451008
- **EnumErrorCategory**: Mbist
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Parameter
- **Code**: 8
- **Level**: Error
- **#Args**: 3
- **Template**: Missing Parameter in {0}({1}) : {2}
- **Guidance**: Missing argument in C# library function, check C# library version. 

### E_MissVbtModule_01
- **FullCode**: EBI447009
- **EnumErrorCategory**: Mbist
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Module
- **Code**: 9
- **Level**: Error
- **#Args**: 2
- **Template**: The VBT function: {0} can not find in {1} library!
- **Guidance**: Missing C# library function, check C# library version. 

## PatInfoType (`PatInfoType.cs`)

### E_MismatchGenericName_01
- **FullCode**: EHP353002
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: PatternInfo
- **Code**: 2
- **Level**: Error
- **#Args**: 1
- **Template**: Please Check the pattern "{0}" , the vm vector and generic name is not match.
- **Guidance**: Verify the generic name used in PatInfo matches the name defined in the source definition. Check for case differences or renamed entries.

### E_MismatchVmandSubr_01
- **FullCode**: EHP353003
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: PatternInfo
- **Code**: 3
- **Level**: Error
- **#Args**: 1
- **Template**: Please Check the pattern "{0}" , the vm vector and call subrs is not match.
- **Guidance**: Compare the VM definition against the call subroutine entries for the pattern. Ensure they reference the same signals and measurement targets.

### E_MissingVmvector_01
- **FullCode**: EHP253003
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: PatternInfo
- **Code**: 3
- **Level**: Error
- **#Args**: 1
- **Template**: Please Check the pattern "{0}" , the vm vector is missing
- **Guidance**: Open the PatInfo sheet and add the required VM vector definition for the pattern. Verify the VM vector name and ensure it is referenced correctly.

### E_WrongDigSrcSignalName_01
- **FullCode**: EHP353004
- **EnumErrorCategory**: HardIp
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: PatternInfo
- **Code**: 4
- **Level**: Error
- **#Args**: 1
- **Template**: Pattern : "{0}" , Sen bit is Empty while pattern info has DigSrc Signal Name
- **Guidance**: Verify the digital source signal name against the defined signal list. Check for typos or renamed signals in the source definition.

## PatternMissing (`PatternMissing.cs`)

### E_MissingPatternFile_01
- **FullCode**: EBA452001
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pattern
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: {0} doesn't exist in pattern folder.
- **Guidance**: Verify the pattern file path and name against the test-plan entry. Run the pattern compilation step if the file is generated. Check source-control to confirm the file was committed and not accidentally excluded.

### E_VersionChange_01
- **FullCode**: EBA252001
- **EnumErrorCategory**: Basic
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Pattern
- **Code**: 1
- **Level**: Error
- **#Args**: 2
- **Template**: Pattern's version changes from {0} to {1}
- **Guidance**: The pattern was regenerated with a different version than the test plan expects. Re-run the full pattern generation flow and update the test plan to reference the new version, or roll back the pattern source to the expected version.

## PostActionErrorType (`PostActionErrorType.cs`)

### W_DuplicateFile_01
- **FullCode**: WPO125001
- **EnumErrorCategory**: PostAction
- **EnumErrorBehavior**: Duplicate
- **EnumErrorTarget**: File
- **Code**: 1
- **Level**: Warning
- **#Args**: 1
- **Template**: Duplicate sheet name detected during IGXL composition: {0}. The previously loaded sheet was removed.
- **Guidance**: A duplicate sheet was found and has been removed from the output. Verify that the intended source file is the one retained.

### W_DuplicateInstance_01
- **FullCode**: WPO137001
- **EnumErrorCategory**: PostAction
- **EnumErrorBehavior**: Duplicate
- **EnumErrorTarget**: Instance
- **Code**: 1
- **Level**: Warning
- **#Args**: 3
- **Template**: Duplicate instance '{0}' was detected {1} times in the following sheet(s): {2}.
- **Guidance**: Please check whether the same instance has been defined multiple times in the input document, as duplicate definitions may lead to generation issues.

### W_DuplicateTestNumber_01
- **FullCode**: WPO175001
- **EnumErrorCategory**: PostAction
- **EnumErrorBehavior**: Duplicate
- **EnumErrorTarget**: TestNumber
- **Code**: 1
- **Level**: Warning
- **#Args**: 5
- **Template**: Duplicate TNum {0} detected between Flow Sheet '{1}' Item '{2}' and Flow Sheet '{3}' Item '{4}'.
- **Guidance**: Check the test number assignments for repeated values. Each test must have a unique number; assign the next available number to the duplicate.

## PreActionErrorType (`PreActionErrorType.cs`)

### E_DuplicateDcCategory_01
- **FullCode**: EPA118001
- **EnumErrorCategory**: PreAction
- **EnumErrorBehavior**: Duplicate
- **EnumErrorTarget**: DcCategory
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Duplicate DC category: {0}
- **Guidance**: Duplicate DC category definition found in the voltage table. Please ensure each DC category name is unique.

### E_DuplicateLibrary_01
- **FullCode**: EPA141001
- **EnumErrorCategory**: PreAction
- **EnumErrorBehavior**: Duplicate
- **EnumErrorTarget**: Library
- **Code**: 1
- **Level**: Error
- **#Args**: 3
- **Template**: Duplicate library: "{0}" in file {1}!={2}
- **Guidance**: Check the library list for duplicate entries. Remove the duplicate module or rename one entry so each module name is unique.

### E_InvalidDocument_01
- **FullCode**: EPA220001
- **EnumErrorCategory**: PreAction
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Document
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: Worksheet does not contain any data.
- **Guidance**: Verify that the document is not empty and contains the expected content.

### E_MismatchPinGroup_01
- **FullCode**: EPA356001
- **EnumErrorCategory**: PreAction
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: PinGroup
- **Code**: 1
- **Level**: Error
- **#Args**: 3
- **Template**: Pin group {0} from IO_PinMap/IO_PinGroup has inconsistent pin definitions between {1}. Please check these pins in IO_PinMap/IO_PinGroup : {2}
- **Guidance**: Please compare the pin definitions and resolve any inconsistencies.

### E_MismatchPinGroup_02
- **FullCode**: EPA356002
- **EnumErrorCategory**: PreAction
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: PinGroup
- **Code**: 2
- **Level**: Error
- **#Args**: 4
- **Template**: Pin group {0} type mismatch: {1} uses type "{2}", while the corresponding pin group in IO_PinMap/IO_PinGroup uses type "{3}",
- **Guidance**: Please verify the pin group type definitions in both sources and ensure they match.

### E_MissingDcCategory_01
- **FullCode**: EPA418001
- **EnumErrorCategory**: PreAction
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: DcCategory
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: Missing DC category name.
- **Guidance**: A category name is missing in a flagged column of the voltage table. Every category column must have a non-empty name in the header row.

### E_MissingDocument_01
- **FullCode**: EPA420001
- **EnumErrorCategory**: PreAction
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Document
- **Code**: 1
- **Level**: Error
- **#Args**: 0
- **Template**: Required sheet IO_PinMap or IO_Continuity was not found in the test plan.
- **Guidance**: Verify that the required sheets exist in the test plan and have not been renamed or removed.

### E_MissingDocument_02
- **FullCode**: EPA420002
- **EnumErrorCategory**: PreAction
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Document
- **Code**: 2
- **Level**: Error
- **#Args**: 0
- **Template**: Missing IO_ignore_list sheet in the test plan.
- **Guidance**: Verify that the required sheet exists in the test plan.

### E_MissingFile_01
- **FullCode**: EPA425001
- **EnumErrorCategory**: PreAction
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: File
- **Code**: 1
- **Level**: Error
- **#Args**: 2
- **Template**: Cannot find {0}: "{1}".
- **Guidance**: Verify that the specified path is correct and that the referenced file exists.

### E_MissingHeader_01
- **FullCode**: EPA433001
- **EnumErrorCategory**: PreAction
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Header
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Cannot find first header '{0}' in the sheet.
- **Guidance**: Check that the expected header row exists and the header name matches exactly.

### E_MissingLibrary_01
- **FullCode**: EPA441001
- **EnumErrorCategory**: PreAction
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Library
- **Code**: 1
- **Level**: Error
- **#Args**: 2
- **Template**: The function: {0} can not find in {1} library!
- **Guidance**: Please verify that the function exists in the specified library and that the library version and function reference are correct.

### E_MissingParameter_01
- **FullCode**: EPA451001
- **EnumErrorCategory**: PreAction
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Parameter
- **Code**: 1
- **Level**: Error
- **#Args**: 3
- **Template**: Missing parameter "{2}" in function "{0}" of library "({1})".
- **Guidance**: Please ensure the parameter exists in the library and that the argument name matches the documented definition.

### E_RuleViolationColumn_01
- **FullCode**: EPA611001
- **EnumErrorCategory**: PreAction
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Column
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Non-adjacent DC category: {0}
- **Guidance**: Columns belonging to the same DC category must be adjacent. Please check whether any previous column has the same DC category name.

### E_RuleViolationPin_01
- **FullCode**: EPA655001
- **EnumErrorCategory**: PreAction
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Pin
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: N pin need to behind to P pin of Pin group : {0} for library issue
- **Guidance**: Verify that the differential pair pins are correctly defined and paired. Ensure the positive and negative pins are assigned to the same differential pair group.

### E_RuleViolationVoltage_01
- **FullCode**: EPA684001
- **EnumErrorCategory**: PreAction
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Voltage
- **Code**: 1
- **Level**: Error
- **#Args**: 7
- **Template**: Pin: {0}, Category: {1}, {2}:{3} {4}:{5} {3}{6}{5}
- **Guidance**: Open the voltage table and check the voltage values for the flagged pin and category. Ensure LV <= NV <= HV and all values are self-consistent.

### W_InvalidFormat_01
- **FullCode**: WPA229001
- **EnumErrorCategory**: PreAction
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 1
- **Level**: Warning
- **#Args**: 1
- **Template**: Power pin name "{0}" contains spaces.
- **Guidance**: Remove any spaces from the power pin name and ensure it follows the expected naming convention.

### W_InvalidFormat_02
- **FullCode**: WPA229002
- **EnumErrorCategory**: PreAction
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Format
- **Code**: 2
- **Level**: Warning
- **#Args**: 1
- **Template**: DC category value "{0}" is not a valid number or percentage.
- **Guidance**: Verify that the DC category value in the Voltage Table is a valid number (e.g. 1.8) or percentage (e.g. 90%).

### W_InvalidVoltage_01
- **FullCode**: WPA284001
- **EnumErrorCategory**: PreAction
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Voltage
- **Code**: 1
- **Level**: Warning
- **#Args**: 5
- **Template**: DC Category "{0}", Pin "{1}" contains invalid NV/{2} content. NV: {3}, {2}: {4}.
- **Guidance**: Please check the values in the corresponding pin under this category and ensure they are valid numeric values.

### W_InvalidVoltage_02
- **FullCode**: WPA284002
- **EnumErrorCategory**: PreAction
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Voltage
- **Code**: 2
- **Level**: Warning
- **#Args**: 0
- **Template**: Voltage value cannot be less than 0.
- **Guidance**: Verify that the voltage value is a non-negative number.

### W_RuleViolationDcCategory_01
- **FullCode**: WPA618001
- **EnumErrorCategory**: PreAction
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: DcCategory
- **Code**: 1
- **Level**: Warning
- **#Args**: 2
- **Template**: DC category name "{0}" starts with an unrecognized test block: "{1}".
- **Guidance**: Verify that the specified DC Category is intended and matches the expected block naming conventions(evs ,ids ,mbist ,hardip ,conti ,nwire ,efuse ,rtos ,rto ,sa ,sachain ,td ,tdchain ,scan ,bincut ,htol).

### W_RuleViolationVoltage_01
- **FullCode**: WPA684001
- **EnumErrorCategory**: PreAction
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Voltage
- **Code**: 1
- **Level**: Warning
- **#Args**: 5
- **Template**: DC Category "{0}", Pin "{1}": {2} value ({3}) {4} NV.
- **Guidance**: Please review the corresponding pin for the specified category in the Voltage Table and ensure that LV <= NV <= HV.

### W_RuleViolationVoltage_02
- **FullCode**: WPA684002
- **EnumErrorCategory**: PreAction
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Voltage
- **Code**: 2
- **Level**: Warning
- **#Args**: 5
- **Template**: Pin "{0}" {1} vmain value ({2}) does not match the {1} valt value ({3}) defined for Mbist dc category "{4}".
- **Guidance**: Verify that the specified power pin value matches the VAlt definition for the corresponding MBist DC category.

## RtosErrorType (`RtosErrorType.cs`)

### E_MissingLibrary_01
- **FullCode**: ERT441001
- **EnumErrorCategory**: Rtos
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Library
- **Code**: 1
- **Level**: Error
- **#Args**: 2
- **Template**: The function: {0} can not find in {1} library!
- **Guidance**: Please verify that the function exists in the specified library and that the library version and function reference are correct.

### E_MissingParameter_01
- **FullCode**: ERT451001
- **EnumErrorCategory**: Rtos
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Parameter
- **Code**: 1
- **Level**: Error
- **#Args**: 3
- **Template**: Missing parameter "{2}" in function "{0}" of library "({1})".
- **Guidance**: Please ensure the parameter exists in the library and that the argument name matches the documented definition.

### E_MissingRtosCategory_01
- **FullCode**: ERT418001
- **EnumErrorCategory**: Rtos
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: DcCategory
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Can not find Rtos category: {0} from TestSettting sheet
- **Guidance**: Add the missing Rtos category definition. Verify the category name against the expected Rtos configuration.

## ScanErrorType (`ScanErrorType.cs`)

### E_FormatError_01
- **FullCode**: ESC329001
- **EnumErrorCategory**: Scan
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Format
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: The code mismatch {0} !!!
- **Guidance**: Open the scan sheet and inspect the flagged field for formatting issues. Ensure all values conform to the expected format and data type.

### E_InstanceName_01
- **FullCode**: WSC637002
- **EnumErrorCategory**: Scan
- **EnumErrorBehavior**: RuleViolation
- **EnumErrorTarget**: Instance
- **Code**: 2
- **Level**: Warning
- **#Args**: 0
- **Template**: Instance name too long(over 150), please modify instance name by PatSetName(Orange)!!!
- **Guidance**: Instance name generated by autogen wil be too long and cause error in run time, use PatSetName(Orange) in test plan to shorten it. 

### E_MissingParameter_01
- **FullCode**: ESC451003
- **EnumErrorCategory**: Scan
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Parameter
- **Code**: 3
- **Level**: Error
- **#Args**: 3
- **Template**: Missing Parameter in {0}({1}) : {2}
- **Guidance**: Missing argument in C# library function, check C# library version.

### E_MissVbtModule_01
- **FullCode**: ESC447004
- **EnumErrorCategory**: Scan
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Module
- **Code**: 4
- **Level**: Error
- **#Args**: 2
- **Template**: The VBT function: {0} can not find in {1} library!
- **Guidance**: Missing function in C# library, check C# library version. 

### E_Pattern_01
- **FullCode**: ESC252005
- **EnumErrorCategory**: Scan
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Pattern
- **Code**: 5
- **Level**: Error
- **#Args**: 1
- **Template**: The pattern : {0} format illegal
- **Guidance**: Verify the pattern name follows the expected naming convention (minimum 11 underscore-separated segments). Check that the pattern exists and is correctly referenced in the Scan sheet.

### E_PatternTimeSet_01
- **FullCode**: ESC452006
- **EnumErrorCategory**: Scan
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pattern
- **Code**: 6
- **Level**: Error
- **#Args**: 0
- **Template**: This test can not find any time set by pattern dashboard !!!
- **Guidance**: Verify the time set assigned to the scan pattern in pattern dashboard. 

### E_PatternTimeSet_02
- **FullCode**: ESC152007
- **EnumErrorCategory**: Scan
- **EnumErrorBehavior**: Duplicate
- **EnumErrorTarget**: Pattern
- **Code**: 7
- **Level**: Error
- **#Args**: 1
- **Template**: There are multi time set {0} by patterns !!!
- **Guidance**: Multiple time set is defined in pattern dashboard for this pattern. 

### E_TimeSet_01
- **FullCode**: ESC476008
- **EnumErrorCategory**: Scan
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Timing
- **Code**: 8
- **Level**: Error
- **#Args**: 1
- **Template**: TimeSet {0} does not exist is K folder
- **Guidance**: Time set missing in source folder. 

### E_TimeSet_02
- **FullCode**: ESC452009
- **EnumErrorCategory**: Scan
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pattern
- **Code**: 9
- **Level**: Error
- **#Args**: 0
- **Template**: Can't get any SCAN Tset in all payload
- **Guidance**: Check the pattern dashboard for all payload patterns in this row. Ensure at least one pattern has a valid SCAN time set (ScanTset) defined.

### E_UserFunction_01
- **FullCode**: ESC481010
- **EnumErrorCategory**: Scan
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Userfunction
- **Code**: 10
- **Level**: Error
- **#Args**: 1
- **Template**: DigSrc {0} in UserFunction is not defined in UF_DigSrc sheet
- **Guidance**: Check if the Userfunction key is defined in Userfunction sheet.

### E_UserFunction_02
- **FullCode**: ESC481011
- **EnumErrorCategory**: Scan
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Userfunction
- **Code**: 11
- **Level**: Error
- **#Args**: 1
- **Template**: Need to assign digital source group for digital source pattern {0}
- **Guidance**: Digital source pattern need digital source group in DigSrc sheet

### E_UserFunction_03
- **FullCode**: ESC281012
- **EnumErrorCategory**: Scan
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Userfunction
- **Code**: 12
- **Level**: Error
- **#Args**: 0
- **Template**: UserFunction format error, should be "DigSrc:[DigSrcGroup]"
- **Guidance**: Correct the UserFunction value in the scan instance sheet. The value must be two colon-separated parts: 'DigSrc' and the digital source group name.

### E_VoltageType_01
- **FullCode**: ESC384013
- **EnumErrorCategory**: Scan
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Voltage
- **Code**: 13
- **Level**: Error
- **#Args**: 0
- **Template**: Voltage type is different in the Sub Flow column and Voltage Category column !!!
- **Guidance**: Verify the voltage type in Sub Flow column and Voltage Category column in test planUpdate the configuration to use a valid voltage type.

### W_Pattern_01
- **FullCode**: WSC452014
- **EnumErrorCategory**: Scan
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pattern
- **Code**: 14
- **Level**: Warning
- **#Args**: 1
- **Template**: {0} can't be found the pattern in CSV !!!
- **Guidance**: Verify the pattern name against the pattern CSV dashboard. Ensure the pattern is listed and the name spelling is correct.

### W_Pattern_02
- **FullCode**: WSC752015
- **EnumErrorCategory**: Scan
- **EnumErrorBehavior**: Info
- **EnumErrorTarget**: Pattern
- **Code**: 15
- **Level**: Warning
- **#Args**: 1
- **Template**: This pattern {0} is "Dont_useInCsv" !!!
- **Guidance**: The pattern is intentionally excluded via the 'dont_use' flag in the CSV. If the pattern is needed, update its use flag in the pattern CSV dashboard.

### W_Pattern_03
- **FullCode**: WSC452016
- **EnumErrorCategory**: Scan
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: Pattern
- **Code**: 16
- **Level**: Warning
- **#Args**: 1
- **Template**: There are no FileVersion for {0} !!!
- **Guidance**: Check the pattern CSV dashboard for the flagged pattern. Ensure a valid FileVersion value is assigned to the pattern entry.

### W_PatternTimeSet_01
- **FullCode**: WSC176017
- **EnumErrorCategory**: Scan
- **EnumErrorBehavior**: Duplicate
- **EnumErrorTarget**: Timing
- **Code**: 17
- **Level**: Warning
- **#Args**: 2
- **Template**: There are multi time set {0} by patterns overwrite to {1}!!!
- **Guidance**: Patterns reference different time sets. The specified time set value was used as an override. Verify that the override time set is correct for all patterns in this row.

## SsnErrorType (`SsnErrorType.cs`)

### E_MismatchCore_01
- **FullCode**: ESN313001
- **EnumErrorCategory**: Ssn
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Core
- **Code**: 1
- **Level**: Error
- **#Args**: 3
- **Template**: Cannot found SsnCoreName "{0}" of pattern (HardipInfo) in the "{1}"(HarvestPinFlag_Table). Pattern : "{2}"
- **Guidance**: Compare the SSN core name against the defined core list. Update the SSN definition or the core reference to ensure they match.

### E_MismatchCore_02
- **FullCode**: ESN313002
- **EnumErrorCategory**: Ssn
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Core
- **Code**: 2
- **Level**: Error
- **#Args**: 3
- **Template**: Cannot found SsnCoreName "{0}" from "{1}" (HarvestPinFlag_Table) in the pattern (HardipInfo). Pattern : "{2}"
- **Guidance**: Compare the core names between the HarvestPinFlag_Table entry and the HardipInfo pattern definition. Add the missing core to HardipInfo or remove the extra entry from HarvestPinFlag_Table.

### E_DuplicatePattern_01
- **FullCode**: ESN152001
- **EnumErrorCategory**: Ssn
- **EnumErrorBehavior**: Duplicate
- **EnumErrorTarget**: Pattern
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Duplicate Pattern : {0}
- **Guidance**: Check the SSN pattern list for duplicate entries. Remove the duplicate or rename one entry so each pattern name is unique.

### E_DuplicatePinGroup_01
- **FullCode**: ESN156001
- **EnumErrorCategory**: Ssn
- **EnumErrorBehavior**: Duplicate
- **EnumErrorTarget**: PinGroup
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Duplicate Pin Group : {0}
- **Guidance**: Check the SSN pin group definitions for duplicate entries. Remove or rename the duplicate group so each name is unique.

### E_InvalidFlag_01
- **FullCode**: ESN226001
- **EnumErrorCategory**: Ssn
- **EnumErrorBehavior**: Invalid
- **EnumErrorTarget**: Flag
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Illegal Flag Name : {0}
- **Guidance**: Rename the SSN flag to avoid reserved keywords or illegal characters. Check the SSN naming rules for the list of disallowed names.

### E_MissingPatternInfo_01
- **FullCode**: ESN453001
- **EnumErrorCategory**: Ssn
- **EnumErrorBehavior**: Missing
- **EnumErrorTarget**: PatternInfo
- **Code**: 1
- **Level**: Error
- **#Args**: 1
- **Template**: Pattern : "{0}" , Missing SSN info in HardipInfo
- **Guidance**: Open the SSN definition and add the required SSN information. Verify all mandatory SSN fields are populated for the specified instance.

### E_MismatchFlag_01
- **FullCode**: ESN326001
- **EnumErrorCategory**: Ssn
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Flag
- **Code**: 1
- **Level**: Error
- **#Args**: 2
- **Template**: PinGroup Count : {0} , PinFlag Count : {1} Mismatch
- **Guidance**: Verify that the SSN pin group and the associated flag definition are aligned. Ensure the pin count and flag assignments are consistent between the two definitions.

### W_MismatchPattern_01
- **FullCode**: WSN352001
- **EnumErrorCategory**: Ssn
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Pattern
- **Code**: 1
- **Level**: Warning
- **#Args**: 1
- **Template**: Corresponding patterns > 1 in the HarvestPinFlag_Table. Pattern : "{0}"
- **Guidance**: Open the HarvestPinFlag_Table and verify that each pattern has a unique match. Remove or disambiguate the duplicate entries so only one row matches the pattern.

### W_MismatchPattern_02
- **FullCode**: WSN352002
- **EnumErrorCategory**: Ssn
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Pattern
- **Code**: 2
- **Level**: Warning
- **#Args**: 1
- **Template**: Cannot found any matching pattern in the HarvestPinFlag_Table. Pattern : "{0}"
- **Guidance**: Check the SSN field value against the expected specification. Ensure the value is within allowed bounds and uses the correct format.

### W_MismatchFlag_01
- **FullCode**: WSN326001
- **EnumErrorCategory**: Ssn
- **EnumErrorBehavior**: Mismatch
- **EnumErrorTarget**: Flag
- **Code**: 1
- **Level**: Warning
- **#Args**: 3
- **Template**: Pattern : "{0}" , SsnCoreName : "{1}" SsnFlag : "{2}" mismatch
- **Guidance**: Open the HarvestPinFlag_Table and check the flag name for the flagged SSN core. Ensure the flag follows the expected naming convention (F_<CoreName>).
