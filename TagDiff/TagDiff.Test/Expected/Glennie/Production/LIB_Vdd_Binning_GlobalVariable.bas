Attribute VB_Name = "LIB_Vdd_Binning_GlobalVariable"
Option Explicit

'=======================================================================================================
' Constant to control the flow condidion
'=======================================================================================================
Public Const Flag_VDD_Binning_Func_Pattern_Only = False                 'control running pattern only for Vddbinning especially for CP1.
Public Const Flag_VDD_Binning_Verbose = False
Public Const Flag_Tester_Offline As Boolean = False                     'set the vddbinning is tester offline.
Public Flag_VDD_Binning_Offline As Boolean
Public Const Flag_Print_Out_tables_enable = False                       'set the print out IDS ZONE and voltage enable.
Public Flag_Interpolation_enable As Boolean                             'If Vx calculation of interpolation exists, turn the flag as true.
'''20210420: C651 Si did internal syncup and confirmed that Montonicitiy Check should use product voltage(PV) only.
Public Const Flag_Only_Check_PV_for_VoltageHeritage As Boolean = True   'C651 Chris Vu requested that only check Efuse Product voltage (GradVDD, not Grade) for voltage heritage in "find_start_voltage" and "find_next_bin_eq_interpolation", 20201028.
'''20211130 Rhodes special case GPU0~GPUX
Public Const Flag_cross_domain_check_enable = False
'''For IDS distribution mode
Public Const Flag_IDS_Distribution_enable = False                       'The flag is to decide if BinCut search mode is IDS mode or not. True: IDS mode; False: Linear mode.
Public Pal_multifstp_pattern() As String
Public Pal_multifstp_header() As String

'''20211202: For IDS Mapping using
Public Flag_IDS_Mapping_enable As Boolean                               ''Added to support "BincutCalc_IDSrail" Column 20211202

'''20220624: For checking sort/hard bin in BinTable and BinCut Vdd_def sheet
Public Flag_BinTable_Parsed_ForBinCut As Boolean                        ''Added for ensuring hard bin for BinCut table

'''For projects with rail-switch
Public Const Flag_Enable_Rail_Switch = True                             'Added to enable VRS rail switch(Vmain and Valt).
Public Const Flag_SyncUp_DCVS_Output_enable = True                      'Added to control SyncUp_DCVS_Output.
Public Const Flag_Read_SafeVoltage_from_DCspecs = True                  'Added the flag to control reading safe voltages from DC Specs(True) or Global Specs(False).
Public Const Flag_Using_Payload_Voltage_for_Selsrm_Calc = False         'Added the flag to control BinCut Selsrm bit calculation by real BinCut Payoad voltage(True) or EQN-based voltage without dynamic_offset(False).

'''For Capture Memory(CMEM)
'''20210322: Modified to decide Flag_Enable_CMEM_Collection by checking TheExec.Flow.EnableWord("Vddbin_CMEM_Collection").
Public Flag_Enable_CMEM_Collection As Boolean                           'Added the flag to overwrite Enable_CMEM_Collection.

'''For COFInstance
Public Flag_Vddbin_COF_Instance As Boolean                              'Added the flag to overwrite COFInstance.
Public Flag_Vddbin_COF_Instance_with_PerEqnLog As Boolean               'Added the flag to print PerEqnLog for COFInstance in the datalog.
Public Flag_Vddbin_COF_StepInheritance As Boolean

'''For DoAll_DebugCollection
Public Flag_Vddbin_DoAll_DebugCollection As Boolean

'''For Pattern Burst Control
Public Flag_Vddbin_PatternBurst_PrintOut As Boolean                        '''20230323: Add enable word to print out pattern names under domain mode, requested by C651 Si

'''For BinOut flag
Public Const strGlb_Flag_Vddbinning_Fail_Stop = "F_Vddbinning_Fail_Stop"    'The flag controls BinOut. Please check "Vddbinning_Fail_Stop" in Bin_Table.
Public Const strGlb_Flag_Vddbinning_IDS_fail = "F_Vddbinning_IDS_fail"      'The flag controls BinOut. Please check "Bin_Vddbinning_IDS_fail" in Bin_Table.
Public Const strGlb_Flag_Vddbinning_Power_Binning_Fail_Stop = "F_Power_Binning_Fail" 'The flag controls Power_Binning BinOut. Please check "Vddbinning_Power_Binning_Fail" in Bin_Table.
Public Const strGlb_Flag_Vddbinning_Interpolation_fail = "F_Vddbinning_Interpolation_fail"
Public Const strGlb_Flag_HarvestBinningFlag_AllCorePass = "F_GFX_CORE_5_0_SUM_0"
Public Const strGlb_Flag_HarvestBinningFlag_SingleFSTP = "F_Gfx_SingleFSTP"

'''For PTE & TTR improvement
Public Const Flag_PrintDcvsShadowVoltage = False                        'Added to enable all print_alt related functions with DCVS shadow voltages.
Public Const Flag_noRestoreVoltageForPrepatt = False                    'Control vbt not to save and restore Payload voltages for Prepatt (only for projects with rail-switch).
Public Const Flag_Skip_ReApplyInitVolageToDCVS = False                  'Control vbt not to re-apply BinCut Init voltages to DCVS for PrePatt(Init Patt). If initial voltages and safe voltage(init voltage) use the same DC category, it can skip "set_core_power_vddbinning_VT" after initial voltages...
Public Const Flag_Skip_ReApplyPayloadVoltageToDCVS = False              'Control vbt not to re-apply BinCut Payload voltages to DCVS for FuncPatt(Payload Patt).
Public Const Flag_Remove_Printing_BV_voltages = False                   'Control vbt not to print BV strings of BinCut Initial/Safe/Payload voltages in the datalog.
Public Const Flag_Skip_Printing_Safe_Voltage = False                    'Control vbt not to print BV strings of BinCut Safe Voltages in the datalog.
Public Const Flag_Skip_Printing_SelSrm_DSSC_Info = False                'Control vbt not to print strings of BinCut SelSrm DSSC Info in the datalog.
Public Const Flag_PowerBinning_Harvest_Bin_FieldName = False             'Control vbt to select Power binning with harvest bin field name or not
Public Const Flag_Enable_Parsing_BinTable_ForBinCut = True              'For checking sort/hard bin in BinTable and BinCut Vdd_def sheet

'''Other flags
Public is_BinCutJob_for_StepSearch As Boolean
Public Flag_BinCut_Config_Printed As Boolean                            'Added to print the status of BinCut settings.
Public Flag_SelsrmMappingTable_Parsed As Boolean
Public Flag_PowerBinningTable_Parsed As Boolean
Public Flag_Enable_PowerBinning_Harvest As Boolean                      'PowerBinning for DUT with Harvest (new PwrSeq_Harvest) if column "Harvest_bin" in powerbinning table is not empty.
Public Flag_NonbinningrailOutsideBinCut_parsed As Boolean               'Must set it as "True" if sheet "Non_Binning_Rail_Outside_BinCut" is parsed.
'''20210510: Modified to parse HARV_Pmode_Table and HARVMappingTable for Harvest Core DSSC.
Public Flag_Harvest_Pmode_Table_Parsed As Boolean
Public Flag_Harvest_Mapping_Table_Parsed As Boolean
Public Flag_Harvest_Core_DSSC_Ready As Boolean
'''20210526: Modified to add "Flag_Get_column_Monotonicity_Offset" for Monotonicity_Offset check because C651 Si revised the check rules.
Public Flag_Get_column_Monotonicity_Offset As Boolean
Public Flag_T0TX_ROOM_format As Boolean 'WWWWW20220631 T0TX
Public Flag_T0TX_HOT_format As Boolean 'WWWWW20220631 T0TX
'''20220711: Added for optional use of HarvPinGrp in BinCut
Public Flag_HarvPinFlag_Mapping_Table_Parsed As Boolean
'''20220922: Coll FT2 uses shadow voltage, no binout
Public Flag_Vddbin_FT2_Shadow_NoBinout As Boolean

'=======================================================================================================
' Constant to define the size
'=======================================================================================================
Public MaxPassBinCut As Long                                        'How many BinCut tables we can read now. For setting the array size.
Public MaxSiteCount As Long                                          'How many sites tester can support now.
Public Max_IDS_Zone As Long  'add function                                     'Max IDS Zone number. For setting the array size.
Public Max_IDS_Step As Long  'Public Const TotalStepPerMode = 12   equal max eqn                                      'Max steps in one IDS Zone. For setting the array size.
Public MaxBincutPowerdomainCount As Long
Public MaxPerformanceModeCount As Long   'add function OK                          'Max numbers of performance mode. For setting the array size.
Public MaxAdditionalModeCount As Long     'add function                             'Max numbers of Additional Modes for performance modes in Flow Sheet ("Non_Binning_Rail"). ex: MS001_GPU, "GPU" is the additional mode of p_mode MS001.
Public MaxJobCountInVbt As Long                                     'Max job count in VBT library.
Public MaxBincutVoltageType As Long                                  'Max BinCut voltage type for Enum BincutVoltageType.
Public Const gC_StepVoltage = 5                                         'Step Size of efuse product voltage (refer to the sheet "Efuse_Bit_Def_Table").
Public Const VddbinIDSFailSortNumber = 9803                             'SortNumber of Flag_Vddbinning_IDS_fail.
Public Const VddbinIDSFailBinNumber = 20                                'BinNumber of Flag_Vddbinning_IDS_fail.
Public Const VddbinFailStopSortNumber = 9801                            'SortNumber of Flag_Vddbinning_Fail_Stop.
Public Const VddbinFailStopBinNumber = 20                               'BinNumber of Flag_Vddbinning_Fail_Stop & Flag_Vddbinning_IDS_fail.
'=======================================================================================================
' variable
'=======================================================================================================
Public PassBinCut_ary() As Long                                         'Store all BinCut numbers from header of the sheet "Vdd_Binning_Def_appA_1" into the array.
Public Total_Bincut_Num As Long                                         'Total enable bincut table in T/P
Public CurrentPassBinCutNum As New SiteLong                             'Passbin for this site right now.
Public BV_StepVoltage As Double                                         'Step Size shown in BinCut "Vdd_Binning_Def_appA".
Public FlagInitVddBinningTable As Boolean
Public Version_Vdd_Binning_Def As String                                'Version of BinCut table "Vdd_Binning_Def_appA".
Public Version_IDS_Distribution As String                               'Version of BinCut table "IDS_Distribution_Table".
Public VddbinningBaseVoltage As Double                                  'BaseVoltage is defined in the header of "Vdd_Binning_Def_appA".
Public BaseVoltageFromEfuseBDF As Double
Public IsLevelLoadedForApplyLevelsTiming As Boolean                     'Flag to check if ApplyLevelsTimingDone is done for the instance.
Public PreviousBinCutInstanceContext As String
Public CurrentBinCutInstanceContext As String
Public Flag_isPrinted_BincutVoltageOutput As Boolean                    '20221223: For print voltage output(Vmain/Valt) in debug print mode

'=======================================================================================================
' Define the address to store the data from BinCut Tables for different Performance Modes
'=======================================================================================================
'''BinCut power domains are enumerated first, then pmodes are enumerated.
'''Numbers of BinCut power domains should match the constant "MaxBincutPowerdomainCount".
'''cntVddbinPmode should match the maximum of pmode defined by the constant "MaxPerformanceModeCount".
Public VddbinPinDict As New Dictionary      '''Domain2enum
Public cntVddbinPin As Integer
Public VddbinPmodeDict As New Dictionary    '''pmode2enum
Public cntVddbinPmode As Integer
Public VddBinName() As String               '''enum2pmode
Public EquationVoltages_ARRAY() As String  '''20220321

'''ToDo: If no one uses type_bin2_flag/mode_bin2_flag, we will remove these variables later.
Public ExcludedPmode() As Boolean 'Public ExcludedPmode(MaxPerformanceModeCount) As Boolean

'''20191219: Dictionaries for Domain2Pin and Pin2Domain.
'''Bincut powerDomains are defined in the header of the sheet "Non_Binning_Rail".
Public domain2pinDict As New Dictionary
Public pin2domainDict As New Dictionary

'''20191219: Dictionary for storing DCVS type for each BinCut powerPin.
'''We will use this for UltraFlexPlus later!!!
Public VddbinPinDcvstypeDict As New Dictionary

'''20200423: Modified to move "Dim dictPin2Dcspec As New Dictionary" from initDomain2Pin into GlobalVariable.
Public dictPin2Dcspec As New Dictionary
Public dictDomain2DcSpecGrp As New Dictionary
'''20211130 Rhodes special case GPU0~GPUX
Public dict_cross_domain_compare As New Dictionary
'''20210703: Modified to use dict_strPmode2EfuseCategory as the dictionary of p_mode and array of the related Efuse category.
Public dict_strPmode2EfuseCategory As New Dictionary
'''20210703: Modified to use dict_EfuseCategory2BinCutTestJob as the dictionary of Efuse category and the matched programming state in Efuse.
Public dict_EfuseCategory2BinCutTestJob As New Dictionary
'''20220711: Modified to use dict_EfuseCategory2BDF
Public dict_EfuseCategory2BDF As New Dictionary

'=======================================================================================================
' Power_List_All comes in the sequence of MAX_ID and performance mode from Bincut voltage table(Vdd_Binning_Def).
'=======================================================================================================
'''ToDo: Maybe we can remove the globalVariable "Power_List_All" later...
Public Power_List_All As String

'''Replace the hard-code global variable "**_power_seq" with BinCut_Power_Seq() by index of powerDomain.
'''20190704: Added to replace each Power_Seq.
'''20210701: Modified to add the globalVariable gb_bincut_power_list(MaxBincutPowerdomainCount)
Public gb_bincut_power_list() As String 'Public gb_bincut_power_list(MaxBincutPowerdomainCount) As String

Public Type Power_Seq
    Power_Seq() As String
End Type

Public BinCut_Power_Seq() As Power_Seq 'Public BinCut_Power_Seq(MaxBincutPowerdomainCount) As Power_Seq

Public Bincut_Domain2BinX() As New SiteBoolean
'20211022: Added to support SRAM Voltage calculate
'Public gb_bincut_sram_power_list(MaxBincutPowerdomainCount) As String
Public BinCut_Sram_Power_Seq As New Dictionary '(MaxBincutPowerdomainCount) As Power_Seq
Public BinCut_Sram_Power_KeyMapping As New Dictionary

'''//Use dictionary to store CorePower/OtherRail, and columns of BinCut powerDomains in "sheet Non_Binning_Rail".
Public dict_IsCorePower As New Dictionary                   '''True: CorePower; False: OtherRail.
Public dict_BinCutFlow_Domain2Column As New Dictionary
Public dict_BinCutFlow_Column2Domain As New Dictionary
Public dict_IsCorePowerInBinCutFlowSheet As New Dictionary  '''True: CorePower; False: OtherRail.
Public dict_EquationVoltages As New Dictionary '20220321
Public dict_EquationVoltages_ALL As New Dictionary '20220321
Public dict_EquationVoltages_Domain2Column As New Dictionary '20220321
Public dict_EquationVoltages_Column2Domain As New Dictionary '20220321
'''Store BinCut powerDomain of CorePower and OtherRail in the strings "VDD_***,VDD_xxx".
Public FullBinCutPowerinFlowSheet As String
Public FullBinCutPowerinFlowSheet_outside As String
Public FullCorePowerinFlowSheet As String
Public FullOtherRailinFlowSheet As String
Public cntEqVoltagePmode As Integer
Public MaxEqnNum As Long
Public Max_IDS_Dis_Zone As Integer

'''Added to parse and store the strings "VDD_***,VDD_xxx" into string array.
Public pinGroup_BinCut() As String
Public pinGroup_CorePower() As String
Public pinGroup_OtherRail() As String

'=======================================================================================================
' Define the different Additional Modes (for Performance Mode with Additional Mode).
' ex: MS001_GPU, GPU is the additional mode of MS001.
'=======================================================================================================
'''Numbers of additional modes should match the constant "MaxAdditionalModeCount".
Public AdditionalModeDict As New Dictionary '''20200211: Modified to replace "FlowTestCondDict" with "AdditionalModeDict".
Public cntAdditionalMode As Integer         '''20200211: Modified to replace "cntFlowTestCond" with "cntAdditionalMode".
Public AdditionalModeName() As String       '''20200211: Modified to replace "FlowTestCondName" with "AdditionalModeName".
Public dict_OutsideBinCut_additionalMode As New Dictionary

'=======================================================================================================
' Define the address to store the data from BinCut Tables for Test Type
'=======================================================================================================
'''Please check the const "MaxTestType".
'''20190502: Modified to add "Func" for dynamic offset_Func.
Enum testType
    TD = 0
    Mbist = 1
    SPI = 2
    RTOS = 2
    TMPS = 3
    Func = 4
    ldcbfd = 5
    Total_Count = 6
End Enum

'''20191105: Added for Printing Testtype for BV strings in the datalog.
Public TestTypeName() As String 'TestTypeName(MaxTestType) As String 20230526

'=======================================================================================================
' Define the different search algorithm
'=======================================================================================================
Enum GradeSearchAlgorithm
    linear = 0
    BinAry = 1
    ids = 2
End Enum

'=======================================================================================================
' The data structure is for reading BinCut Tables, Other Rail Table and Flow Table
'=======================================================================================================
'''20200702: Modified to add OutsideBinCut_OTHER_VOLTAGE/OutsideBinCut_HVCC_OTHER_VOLTAGE/OutsideBinCut_Addtional_OTHER_VOLTAGE/OutsideBinCut_HVCC_Addtional_OTHER_VOLTAGE for Outside BinCut.
'''20210325: Modified to use the 1-dimension array to store SRAM_Vth.
'''20210422: Modified to remove IDS_CP_LIMIT_COUNT/IDS_FT_LIMIT_COUNT/IDS_QA_LIMIT_COUNT/IDS_FT2_LIMIT_COUNT/IDS_FT2_QA_LIMIT_COUNT from Public Type BINCUT_TYPE.
'''20210427: Modified to parse the column of "Monotonicity_Offset".
Public Type BINCUT_TYPE '''For read BinCut Table (vdd_binning_def), other rail and flow sheet(non_binning_rail).
    EQ_Num() As Long
    c() As Double ' down TotalStepPerMode
    m() As Double
    CP_Vmax() As Double
    CP_Vmin() As Double
    '''Montonicity_Offset
    Monotonicity_Offset() As Double
    '''GuardBand
    CP_GB() As Double
    CP2_GB() As Double
    FT1_GB() As Double
    FT2_GB() As Double
    SLT_GB() As Double
    FTQA_GB() As Double
    HTOL_RO_GB() As Double
    HTOL_RO_GB_ROOM() As Double
    HTOL_RO_GB_HOT() As Double
    SLT_FTQA_GB() As Double
    '''IDS_Limit
    IDS_CP_LIMIT() As Double
    IDS_FT_LIMIT() As Double
    IDS_QA_LIMIT() As Double
    IDS_FT2_LIMIT() As Double
    IDS_FT2_QA_LIMIT() As Double
    HVCC_CP() As Double
    HVCC_FT() As Double
    HVCC_QA() As Double ' up TotalStepPerMode
    SBIN_BINNING_FAIL() As Long 'SBIN_BINNING_FAIL(TotalStepPerMode, MaxTestType) As Long 20230526
    SBIN_LVCC_FAIL() As Long 'SBIN_LVCC_FAIL(TotalStepPerMode, MaxTestType) As Long 20230526
    HBIN_BINNING_FAIL() As Long 'HBIN_BINNING_FAIL(TotalStepPerMode, MaxTestType) As Long 20230526
    HBIN_LVCC_FAIL() As Long 'HBIN_LVCC_FAIL(TotalStepPerMode, MaxTestType) As Long 20230526
    Mode_Step As Long
    OTHER_CP_Vmax() As Double 'OTHER_CP_Vmax(MaxBincutPowerdomainCount) As Double
    OTHER_CP_Vmin() As Double 'OTHER_CP_Vmin(MaxBincutPowerdomainCount) As Double
    OTHER_VOLTAGE() As String 'OTHER_VOLTAGE(MaxBincutPowerdomainCount) As String
    HVCC_OTHER_VOLTAGE() As String ' HVCC_OTHER_VOLTAGE(MaxBincutPowerdomainCount) As String
    '''GuardBand for nonbinning CorePower and OtherRail
    OTHER_FT1_GB() As Double 'OTHER_FT1_GB(MaxBincutPowerdomainCount) As Double
    OTHER_FT2_GB() As Double 'OTHER_FT2_GB(MaxBincutPowerdomainCount) As Double
    OTHER_CP1_RAIL() As Double ' OTHER_CP1_RAIL(MaxBincutPowerdomainCount) As Double
    OTHER_CP1_GB() As Double ' OTHER_CP1_GB(MaxBincutPowerdomainCount) As Double
    OTHER_CP2_GB() As Double 'OTHER_CP2_GB(MaxBincutPowerdomainCount) As Double
    OTHER_PRODUCT_RAIL() As Double 'OTHER_PRODUCT_RAIL(MaxBincutPowerdomainCount) As Double
    OTHER_SLT_GB() As Double 'OTHER_SLT_GB(MaxBincutPowerdomainCount) As Double
    OTHER_ATE_FQA_GB() As Double 'OTHER_ATE_FQA_GB(MaxBincutPowerdomainCount) As Double
    OTHER_HTOL_RO_GB() As Double 'OTHER_HTOL_RO_GB(MaxBincutPowerdomainCount) As Double
    OTHER_HTOL_RO_GB_ROOM() As Double 'OTHER_HTOL_RO_GB_ROOM(MaxBincutPowerdomainCount) As Double
    OTHER_HTOL_RO_GB_HOT() As Double 'OTHER_HTOL_RO_GB_HOT(MaxBincutPowerdomainCount) As Double
    OTHER_SLT_FQA_GB() As Double 'OTHER_SLT_FQA_GB(MaxBincutPowerdomainCount) As Double
    HVCC_OTHER_CP_RAIL() As Double 'HVCC_OTHER_CP_RAIL(MaxBincutPowerdomainCount) As Double
    HVCC_OTHER_FT_RAIL() As Double 'HVCC_OTHER_FT_RAIL(MaxBincutPowerdomainCount) As Double
    HVCC_OTHER_QA_RAIL() As Double 'HVCC_OTHER_QA_RAIL(MaxBincutPowerdomainCount) As Double
    OTHER_CPIDS() As Double ' OTHER_CPIDS(MaxBincutPowerdomainCount) As Double
    OTHER_FTIDS() As Double 'OTHER_FTIDS(MaxBincutPowerdomainCount) As Double
    MAX_ID As Double                                                                            '''20151110: added this for each peformance mode in BinCut voltage table. 20190613: changed from long to double -by SY
    ExcludedPmode As Boolean                                                                    '''20150701: added to assign exist different bincut table.
    Allow_Equal() As Integer   'TotalStepPerMode                                                  '''20161223: added to assign the inheriting rule of allow equal for the performance mode.
    Addtional_OTHER_VOLTAGE() As String 'Addtional_OTHER_VOLTAGE(MaxBincutPowerdomainCount, MaxAdditionalModeCount) As String           '''store testCondition of each powerDomain for performance mode with additional mode in BV test instance.
    HVCC_Addtional_OTHER_VOLTAGE() As String 'HVCC_Addtional_OTHER_VOLTAGE(MaxBincutPowerdomainCount, MaxAdditionalModeCount) As String     '''store testCondition of each powerDomain for performance mode with additional mode in HBV test instance.
    INTP_MODE_L() As Integer   'TotalStepPerMode down                                                  '''start p_mode of interpolation.
    INTP_MODE_H() As Integer                                                    '''end p_mode of interpolation.
    INTP_MFACTOR() As Double                                                    '''factor of interpolation.
    INTP_OFFSET() As Double                                                     '''offset of interpolation.
    INTP_SKIPTEST() As Boolean   'TotalStepPerMode up                                               '''flag to skip interpolation tests of p_mode.
    DYNAMIC_OFFSET() As Double 'DYNAMIC_OFFSET(MaxJobCountInVbt, MaxTestType) As Double 20230526
    SRAM_VTH_SPEC(1) As Double '''SRAM_VTH_SPEC(0): for CP1 BV binSearch and postBinCut/OutsideBinCut, SRAM_VTH_SPEC(1): for CP1 HBV and non-CP1 BV/HBV.
    OutsideBinCut_OTHER_VOLTAGE() As String 'OutsideBinCut_OTHER_VOLTAGE(MaxBincutPowerdomainCount) As String
    OutsideBinCut_HVCC_OTHER_VOLTAGE() As String 'OutsideBinCut_HVCC_OTHER_VOLTAGE(MaxBincutPowerdomainCount) As String
    OutsideBinCut_Addtional_OTHER_VOLTAGE() As String 'OutsideBinCut_Addtional_OTHER_VOLTAGE(MaxBincutPowerdomainCount, MaxAdditionalModeCount) As String
    OutsideBinCut_HVCC_Addtional_OTHER_VOLTAGE() As String 'OutsideBinCut_HVCC_Addtional_OTHER_VOLTAGE(MaxBincutPowerdomainCount, MaxAdditionalModeCount) As String
    OutsideBinCut_EquationVoltages() As String 'OutsideBinCut_EquationVoltages(MaxBincutPowerdomainCount) As String
    DC_Bincut_X_X_X() As String 'WWWWW20220631 DC_Bincut_X_X_X() As String
End Type

'===========================================================================================================
' The data structure is for combining all BinCut Tables
'===========================================================================================================
'''20191127: Modified to add "powerpin" for the revised initVddBinTable.
'''20191227: Modified to add "Allow_Equal" for allow equal of the performance mode.
'''20200501: Modified to add "INTP_SKIPTEST" to skip interpolation tests of p_mode.
'''20210414: Modified to add "is_for_BinSearch as Boolean" for AllBinCut(p_mode).
'''20210701: Modified to add "listed_in_Efuse_BDF As Boolean" for AllBinCut(p_mode).
Public Type ALL_BINCUT_TYPE '''For store data related to all BinCut Table.
    Used As Boolean
    Mode_Step As Long
    powerPin As String
    PREVIOUS_Performance_Mode As Integer
    Allow_Equal As Integer
    IDS_MAPPING As String       ''' 20211202: For IDS mapping for CPU_SRAM
    IDS_MAPPING_LIMIT As Double ''' 20230314: Added for printing IDS limit in Adjust_Vddbinning
    IDS_CP_LIMIT As Double
    IDS_FT_LIMIT As Double
    IDS_QA_LIMIT As Double
    IDS_FT2_LIMIT As Double
    IDS_FT2_QA_LIMIT As Double
    TRACKINGPOWER As String
    INTP_SKIPTEST As Boolean
    is_for_BinSearch As Boolean
    listed_in_Efuse_BDF As Boolean '''Check Efuse category of p_mode and update this...
    ''20230901: Added saving bin1 EQN, for pre-flow in bincut
    Last_Bin1_EQN As Long
End Type

'===========================================================================================================
' The data structure is for storing the VDD Binning search result
'===========================================================================================================
'''20200422: Added "tested" for checking if p_mode is tested or not.
'''20210120: Added "step_1stPass_in_IDS_Zone" to store the first pass step in Dynamic IDS Zone for each p_mode. We can use it to find the correspondent PassBinCut number.
'''20210526: Added "is_Monotonicity_Offset_triggered" for Monotonicity_Offset check because C651 Si revised the check rules.
'''20210809: Modified to remove the redundant property "ALL_SITE_MIN As New SiteDouble" from Public Type VBIN_RESULT_TYPE.
'''20210906: Modified to remove the redundant property "IDS As New SiteDouble" from Public Type VBIN_RESULT_TYPE.
Public Type VBIN_RESULT_TYPE '''//liki 1022
    tested As New SiteBoolean
    passBinCut As New SiteLong
    GRADE As New SiteDouble
    GRADEVDD As New SiteDouble
    step_in_BinCut As New SiteLong      '''the step is the address number for storing the EQ (step from 0 and EQ from 1).
    step_in_IDS_Zone As New SiteLong    '''the step_ids_zone is the step in the IDS Zone.
    step_1stPass_in_IDS_Zone As New SiteLong
    FLAGFAIL As New SiteBoolean
    DSSC_Dec As New SiteLong
    is_Monotonicity_Offset_triggered As New SiteBoolean
    step_BINX_MFSTP As New SiteLong
End Type

'===========================================================================================================
' The data structure is for storing data of the IDS Zone
'===========================================================================================================
'''20160512: Modified to add "voltage" for Allen's request.
'''20210405: Modified to remove "PassBinCutList_per_Zone(Max_IDS_Zone) As Long"
'''20210722: Modified to use VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Zone, Max_IDS_Step) and DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Step) for GradeVDD.
Public Type VBIN_IDS_ZONE
    Used As Boolean
    Ids_range() As Double 'Ids_range(Max_IDS_Zone, MaxTestType) As Double20230526
    IDS_Start_EQ_Num() As Long 'IDS_Start_EQ_Num(Max_IDS_Zone, MaxTestType) As Long 20230526
    IDS_START_STEP() As Long 'IDS_START_STEP(Max_IDS_Zone, MaxTestType) As Long 20230526
    IDS_RANGE_COUNT() As Long ' IDS_RANGE_COUNT() As Long 20230526
    c() As Double 'c(Max_IDS_Zone, Max_IDS_Step) As Double 20230526
    m() As Double 'm(Max_IDS_Zone, Max_IDS_Step) As Double 20230526
    passBinCut() As Long ' passBinCut(Max_IDS_Zone, Max_IDS_Step) As Long 20230526
    EQ_Num() As Long 'EQ_Num(Max_IDS_Zone, Max_IDS_Step) As Long 20230526
    Voltage() As New SiteDouble 'Voltage(Max_IDS_Zone, Max_IDS_Step) As New SiteDouble 20230526
    Product_Voltage() As New SiteDouble '''for GradeVDD Product_Voltage(Max_IDS_Zone, Max_IDS_Step) As New SiteDouble 20230526
    Max_Step() As Long 'Max_Step(Max_IDS_Zone) As Long 20230526
    IDS_ZONE_NUMBER As New SiteLong
End Type

'===========================================================================================================
' The data structure is for storing data of the IDS Zone
'===========================================================================================================
'''20160512: Modified to add "voltage" for Allen's request.
'''20210223: Modified to add "step_Mapping" for mapping DYNAMIC_VBIN_IDS_ZONE(p_mode).step_Mapping(BinNum, EQN) to step in DYNAMIC_IDS_Zone.
'''20210407: Modified to add "interpolated as new SiteBoolean" and "step_Interpolated_Start as new SiteLong" for "Public Type DYNAMIC_VBIN_IDS_ZONE".
'''20210722: Modified to use VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Zone, Max_IDS_Step) and DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Step) for GradeVDD.
'''20210810: Modified to add the property "step_Lowest As New SiteLong" to Public Type DYNAMIC_VBIN_IDS_ZONE.
'''20210812: Modified to rename the property "step_lowest As New SiteLong" as "step_inherit As New SiteLong".
Public Type DYNAMIC_VBIN_IDS_ZONE
    Used As New SiteBoolean
    IDS_Start_EQ_Num() As New SiteLong 'IDS_Start_EQ_Num(MaxTestType) As New SiteLong 20230526
    IDS_START_STEP() As New SiteLong ' IDS_START_STEP() As New SiteLong 20230526
    c() As New SiteDouble 'c(Max_IDS_Step) As New SiteDouble
    m() As New SiteDouble 'm(Max_IDS_Step) As New SiteDouble
    passBinCut() As New SiteLong 'passBinCut(Max_IDS_Step) As New SiteLong
    EQ_Num() As New SiteLong 'EQ_Num(Max_IDS_Step) As New SiteLong
    step_Mapping() As New SiteLong '''step_Mapping(PassBin,EQN) 'step_Mapping(MaxPassBinCut, Max_IDS_Step + 1) As New SiteLong 20230526
    Voltage() As New SiteDouble 'Voltage(Max_IDS_Step) As New SiteDouble
    Product_Voltage() As New SiteDouble '''for GradeVDD 'Product_Voltage(Max_IDS_Step) As New SiteDouble
    Max_Step As New SiteLong
    IDS_ZONE_NUMBER As New SiteLong
    interpolated As New SiteBoolean
    step_Interpolated_Start As New SiteLong
    step_inherit As New SiteLong
End Type

'===========================================================================================================
' The data structure is for storing data of the IDS Distribution Table (generated by TSMC Jack on 20160614)
'===========================================================================================================
'''20210405: Modified to remove "PassBinCutList_per_Zone(Max_IDS_Zone) As Long"
Public Type IDS_Distribution_TYPE
    Used As Boolean
    range() As Double 'range(Max_IDS_Zone, MaxTestType) As Double 20230526
    Start_Bin() As Long 'Start_Bin(Max_IDS_Zone, MaxTestType) As Long 20230526
    START_STEP() As Long 'START_STEP(Max_IDS_Zone, MaxTestType) As Long 20230526
    RANGE_COUNT As Long
End Type

Public BinCut() As BINCUT_TYPE 'BinCut(MaxPerformanceModeCount, MaxPassBinCut) As BINCUT_TYPE 20230526
Public AllBinCut() As ALL_BINCUT_TYPE 'MaxPerformanceModeCount down
Public VBIN_RESULT() As VBIN_RESULT_TYPE
Public VBIN_IDS_ZONE() As VBIN_IDS_ZONE
Public VBIN_IDS_ZONE_Temp() As VBIN_IDS_ZONE                     '(generated by TSMC Jack on 20160614)
Public IDS_Distribution_Table() As IDS_Distribution_TYPE         '(generated by TSMC Jack on 20160614)
Public DYNAMIC_VBIN_IDS_ZONE() As DYNAMIC_VBIN_IDS_ZONE 'MaxPerformanceModeCount up

Public Max_V_Step_per_IDS_Zone As Long '''for calculating the Max steps count for IDS Zone
Public RestoredSites As New SiteBoolean

'''20190606: Modified for CPIDS_Spec_OtherRail and FTIDS_Spec_OtherRail.
Public CPIDS_Spec() As Double 'CPIDS_Spec(1 To MaxBincutPowerdomainCount, 1 To MaxPassBinCut) As Double 20230526
Public FTIDS_Spec() As Double 'Public FTIDS_Spec(1 To MaxBincutPowerdomainCount, 1 To MaxPassBinCut) As Double 20230526
Public gb_IDS_hi_limit() As Double 'gb_IDS_hi_limit(1 To MaxBincutPowerdomainCount, 1 To MaxPassBinCut) As Double 20230526

Public Binx_fail_flag As New SiteBoolean    'for record Binx failed item for binning
Public Binx_fail_power As New SiteVariant   'for record Binx failed item for binning
Public Biny_fail_flag As New SiteBoolean    'for record Biny failed item for binning
Public Biny_fail_power As New SiteVariant   'for record Biny failed item for binning

'=======================================================================================================
' Define the different job in vbt lib
'=======================================================================================================
'''//For current TestJob Mapping in IGXL to BinCut testJob(defined in header of sheet "Non_Binning_Rail").
Public bincutJobName As String

'''BinCut testJob names are defined in the sheet "Non_Binning_Rail".
'''Remember to check the constant "MaxJobCountInVbt".
Enum BinCutJobDefinition
    CP1 = 0
    CP2 = 1
    FT1 = 2
    FT2 = 3
    QA = 4
    Total_Count = 5
    COND_ERROR = 6
End Enum

'''//For Selsrm_Mapping_Table
'''20190906: Modified the parsing method for the different SELSRAM DSSC bit length.
'''20191210: Added selsramPin and selsramSramPin.
Public Const SELSRAM_EXPAND_CYCLE = 1
Public selsramPingroup() As String
Public selsramSramPingroup() As String
Public selsramLogicPingroup() As String
Public selsramLogicPinalphagroup() As String
Public selsramPin As String
Public selsramLogicPin As String
Public selsramLogicPinalpha As String
Public selsramSramPin As String

Public Type SELSRAM_Bit_Table
    blockName As String
    bitCount() As String
    Pattern As String
    logic_Pin() As String
    SRAM_PIN() As String
    SelSrm1() As String
    SelSrm0() As String
    alpha() As String
    COMMENT() As String
    DigSrc_Assignment() As String
    DuplicateBlock As Boolean
End Type

Public SelsramMapping() As SELSRAM_Bit_Table
Public StrAry_DuplicateBlock() As String

'''//For parsing header of Power_Binning Table
'''Binned Mode
Public dict_Binned_Mode_Ratio2Idx As New Dictionary     '''Ratio(Name) -> Index to position of array.
Public dict_Binned_Mode_Ratio2Column As New Dictionary  '''Ratio(Name) -> column.
Public dict_Binned_Mode_Column2Ratio As New Dictionary  '''column -> Ratio(Name).
'''Other Mode
Public dict_Other_Mode_Ratio2Idx As New Dictionary      '''Ratio(Name) -> Index to position of array.
Public dict_Other_Mode_Ratio2Column As New Dictionary   '''Ratio(Name) -> column.
Public dict_Other_Mode_Column2Ratio As New Dictionary   '''column -> Ratio(Name).

'''//For Ratio2Idx mapping of Power Binning Table
Public Binned_Ratio_Name() As String
Public Other_Ratio_Name() As String

'''//PowerBinning: Condition -> Spec -> Sheet -> Ratio.
'''//PowerBinning table with ratios: A, B, C, D, E.
'''20200831: Modified to define the array size of PwrBin_Sheet.
'''20201111: Modified for the new format of PowerBinning tables.
Public Type PWRBIN_RATIO_Type
    ratio() As Variant
    pmode As String
End Type

Public Type PWRBIN_SHEET_Type
    Binned_Mode() As PWRBIN_RATIO_Type
    Other_Mode() As PWRBIN_RATIO_Type
    cnt_Binned_Mode As Integer
    cnt_Other_Mode As Integer
    offset As Double
    Spec As Double
    sheetName As String
End Type

Public PwrBin_Sheet() As PWRBIN_SHEET_Type              '''array to store PowerBinning sheetName.
Public PwrBin_SheetCnt As Integer                       '''sheetCnt of powerBinning sheets.
Public PwrBin_SheetnameDict As New Dictionary           '''dictionary of PowerBinning sheetName2enum.
Public PwrBin_SpecIdx2SpecNameDict As New Dictionary    '''dictionary of PowerBinning sheetName2enum.

Public Type PWRBIN_SPEC_Type
    TestName As String                                  '''ex: bin1_low_power, bin1_high_power.
    fusePwrbin As String                                '''PASS: power_binning.
    fuseValue As String                                 '''PASS: fuse_name2.
    specUsed() As Boolean
    specCustomized() As Double                          '''overwrite the spec value.
    haveNextSpec As Boolean                             '''have next spec when same passbin and same Harvest_bin.
    idxAllSpec As Integer
End Type

Public Type PWRBIN_CONDITION_Type
    passBinCut As Integer                               '''Efuse product_identifier (BinCut current passbin number).
    harvestUsed As Boolean
    harvestBin() As Long                                  '''Harvest_bin.
    TestSpec() As PWRBIN_SPEC_Type
End Type

Public AllPwrBin() As PWRBIN_CONDITION_Type
Public gb_str_EfuseCategory_for_powerbinning As String

'''For Capture Memory(CMEM)
'''20201126: Modified to set IfStoreData as siteBoolean
'''20210225: Modified to collect CMEM_VectorData/CMEM_CycleData/CMEM_PatRange/CMEM_PatName for each site. Revised by Oscar and PCLINZG.
Type CMEM_StoreData
    CMEM_VectorData As New SiteVariant
    CMEM_CycleData As New SiteVariant
    CMEM_IndexData As New SiteVariant
    CMEM_PinData As New SiteVariant
    CMEM_PatRange As New SiteVariant
    CMEM_PatName As New SiteVariant
    IfStoreData As New SiteBoolean
End Type

'''//Efuse IDS resolution is defined in "Efuse_BitDef_Table" in Test Plan and Test Program.
'''20190514: Removed the hard-code "I_VDD_***" by the data type "IDS_for_BinCut(VddBinStr2Enum(powerPin)).Real".
Public Type IDS_value
    Real As New SiteDouble
    '''DcTest As New SiteDouble
    ids_name() As String 'ids_name(MaxSiteCount - 1) As String 20230526
End Type

Public IDS_for_BinCut() As IDS_value 'Public IDS_for_BinCut(1 To MaxBincutPowerdomainCount) As IDS_value

'''20190215: Added for SyncUp_DCVS_Output.
Public SyncUp_PowerPin_Group As String

'''20191002: For printing BinCut voltage type in the datalog.
'''Use MaxBincutVoltageType to decide the max number of BinCut voltage type.
'''Remember to maintain the vbt function "initBincutVoltageType".
Enum BincutVoltageType
    None = 0
    InitialVoltage = 1
    SafeVoltage = 2
    PayloadVoltage = 3
    PostbincutBinningpower = 4
    PostbincutAllpower = 5
    Total_Count = 6
End Enum

Public BincutVoltageTypeName(BincutVoltageType.Total_Count) As String 'Public BincutVoltageTypeName(MaxBincutVoltageType) As String

'''//For the arguments "Adjust_Max_Enable","Adjust_Min_Enable","Adjust_Power_Max_list","Adjust_Power_Min_list" of the test instance "Adjust_VddBinning".
'''//Check if "MaxPV(pmode0/pmode1)" is in the column "Comment" of sheet "Vdd_Binning_Def" or not.
Public Flag_Adjust_Max_Enable As Boolean
Public Flag_Adjust_Min_Enable As Boolean
Public Adjust_Power_Max_pmode As String
Public Adjust_Power_Min_pmode As String

'''20200130: Created to store the init and payload voltages.
Public BinCut_Init_Voltage() As New SiteDouble 'Public BinCut_Init_Voltage(MaxBincutPowerdomainCount) As New SiteDouble
Public BinCut_Payload_Voltage() As New SiteDouble 'Public BinCut_Payload_Voltage(MaxBincutPowerdomainCount) As New SiteDouble '
Public Previous_Payload_Voltage() As New SiteDouble 'Public Previous_Payload_Voltage(MaxBincutPowerdomainCount) As New SiteDouble '

Public Type DC_Mapping_Table
    LV_DC_Value As String
    HVCC_DC_Value As String
    Addtional_LV_DC_Value() As String 'Addtional_LV_DC_Value(MaxAdditionalModeCount) As String
    Addtional_HVCC_DC_Value() As String 'Addtional_HVCC_DC_Value(MaxAdditionalModeCount) As String
End Type

Public DC_Mapping() As DC_Mapping_Table 'DC_Mapping(MaxPerformanceModeCount, MaxPassBinCut) As DC_Mapping_Table 20230526
'''20220127,HarvPinGrp relation global variable
Public Type HarvPinGrpTyp
    keyword As String
    MappingIndex As Long
    HarvFailResult() As New SiteBoolean
End Type
Public Type HarvPinFlag_Table
    Pattern As String
    HarvPinGrpConditionArr() As String
    HarvPinGrpCondFailFlagArr() As String
    HarvPinGrpComFailArr() As String
    Enable_otherPinGrp As Boolean
End Type

Public HarvPinFlagMapping() As HarvPinFlag_Table


'''20220418, Static_MappingTable Parsing
Public Type Static_Mapping_Table
    Selsrm_Value As String
    HVCC_Selsrm_Value As String
    Addtional_Selsrm_Value() As String
    HVCC_Addtional_Selsrm_Value() As String
End Type
Public Static_Mapping() As Static_Mapping_Table

'''//Data structure for Instance Info and Instance Step Control.
'''20201027: Modfied to create data structure to store instance info.
'''20201103: Modified to move "Dim stepcount As Long" and "Dim stepcountMax As Long" into "Public Type Instance_Info".
'''20210513: Modified to set inst_info.Harvest_Core_DigSrc_Pin and inst_info.Harvest_Core_DigSrc_SignalName.
'''20210528: Modified to add Pmode_keyword and By_Mode into Public Type Instance_Info to store Harvest Pattern_Pmode and By_mode.
'''20210728: Modified to move "Dim Sram_Vth(MaxBincutPowerdomainCount) As New SiteDouble" into "Public Type Instance_Info".
'''20210830: Modified to add "HarvestBinningFlag as String" for Harvest in BinCut, as requested by C651 Toby.
'''20210901: Modified to move "Step_GradeFound As New SiteLong" from Public Type Instance_Step_Control to the vbt function Update_VBinResult_by_Step.
'''20210903: Modified to merge properties of "Public Type Instance_Step_Control" into Public Type Instance_Info "Public Type Instance_Info".
Public Type Instance_Info
    inst_name As String
    Test_Type As testType
    performance_mode As String                      '''The full string from the argument "performance_mode" of the test instance. It might contain the main "performance mode" and "additional mode", ex: "VDD_SOC_MS001_GPU".
    p_mode As Integer
    powerDomain As String                           '''powerDomain of the binning performance mode.
    Special_Voltage_setup As Boolean                '''If the performance mode has the additional mode in "Non Binning Rail", set it to true.
    addi_mode As Integer
    jobIdx As Integer
    offsetTestTypeIdx As Integer
    is_BinSearch As Boolean                         '''True: BinSearch; False: Functional Test (Pass/Fail only).
    '''IDS value
    ids_current As New SiteDouble                   '''IDS values of the binning powerDomain.
    '''Step-loop control for BinCut search
    count_Step As Long
    maxStep As Long
    Active_site_count As Long
    Step_Current As New SiteLong
    '''20210903: Modified to merge properties of "Public Type Instance_Step_Control" into Public Type Instance_Info "Public Type Instance_Info".
    step_Start As New SiteLong
    step_Stop As New SiteLong
    gradeAlg As New SiteLong
    All_Site_Mask As Long
    AllSiteFailPatt As Long
    On_StopVoltage_Mask As Long
    grade_found As New SiteBoolean
    AnySiteGradeFound As Boolean
    Grade_Not_Found_Mask As Long
    Grade_Found_Mask As Long
    IDS_ZONE_NUMBER As New SiteLong
    All_Patt_Pass As New SiteBoolean
    '''Capture Memory (CMEM)
    enable_CMEM_Collection As Boolean               '''Added the flag to Enable CMEM collection for FFC, requested by Si. 20190611.
    PrintSize As Long
    Step_CMEM_Data() As CMEM_StoreData
    BC_CMEM_StoreData() As CMEM_StoreData
    str_patternPins() As String
    '''COFInstance
    enable_COFInstance As Boolean                   '''Added the flag to COFInstance, requested by Si. 20201015.
    enable_PerEqnLog As Boolean                     '''Added the flag to print EQN log for COFInstance, requested by Si. 20201020.
    '''result_mode
    result_mode As tlResultMode
    '''decompose patset
    enable_DecomposePatt As Boolean                 '''Must set it as "True" if pattern set is decomposed.
    '''PrePatt and FuncPatt
    PrePatt As String
    FuncPat As String
    ary_PrePatt_decomposed() As String
    ary_FuncPat_decomposed() As String
    count_PrePatt_decomposed As Long
    count_FuncPat_decomposed As Long
    '''pattern pass/fail
    PrePattPass As New SiteBoolean
    funcPatPass As New SiteBoolean
    sitePatPass As New SiteBoolean
    '''DCVS output
    previousDcvsOutput As Integer
    currentDcvsOutput As Integer
    '''print BV
    is_BV_Safe_Voltage_printed As Boolean
    is_BV_Payload_Voltage_printed As Boolean
    '''Dynamic_Offset
    str_dynamic_offset() As String 'str_dynamic_offset(MaxSiteCount - 1) As String  20230526
    '''SELSRM
    selsrm_DigSrc_Pin As New PinList
    selsrm_DigSrc_SignalName As String
    patt_SelsrmDigSrc_decomposed_from_PrePatt As String
    patt_SelsrmDigSrc_decomposed_from_FuncPat As String
    patt_SelsrmDigSrc_single As String
    idxBlock_Selsrm_PrePatt As Integer
    idxBlock_Selsrm_FuncPat As Integer
    idxBlock_Selsrm_singlePatt As Integer
    str_Selsrm_DSSC_Info() As String 'str_Selsrm_DSSC_Info(MaxSiteCount - 1) As String 20230526
    str_Selsrm_DSSC_Bit() As String 'str_Selsrm_DSSC_Bit(MaxSiteCount - 1) As String 20230526
    voltage_SelsrmBitCalc() As New SiteDouble '''added for Selsrm Bit calculation, 20201111. 'MaxBincutPowerdomainCount
    sram_Vth() As New SiteDouble '''added for Selsrm Bit calculation, 20210728. 'sram_Vth(MaxBincutPowerdomainCount) As New SiteDouble
    '''Harvest Core DSSC (FSTP and MultiFSTP)
    Pattern_Pmode As String '''ex: MGX001, MGX003, MGX008.
    By_Mode As String '''ex: X4, X6, X10
    Harvest_Core_DigSrc_Pin As New PinList
    Harvest_Core_DigSrc_SignalName As String
    '''HarvestBinning
    HarvestBinningFlag As String
    '''DevChar (CZ Shmoo)
    is_DevChar_Running As Boolean
    DevChar_Setup As String
    get_DevChar_Precondition As Boolean
    '''20220125,HarvPinGrp relation global variable
    enable_HarvPinGrp As Boolean
    HarvPinGrpCheckData() As HarvPinGrpTyp
    HarvPinGrpCommnFailResult As New SiteBoolean
    '''20230926: Added to save non-harvest pattern result in harvest instance
    HarvPinGrp_NonHarvFail As New SiteBoolean
    HarvStartWithBin1 As New SiteBoolean
    '''20220125,Special Search from Bin1EQ1
    StartBin1E1_Flag As New SiteBoolean
    '''20220407, disable compare harvPinGroup when doing regular BV search
    patCnt_HarvPinGrp As Long
    MultiFSTP_Enable As Boolean
    digSrcLabel() As String
    digSrcPatterns() As String
    idxDigSrcPattern As Long
    SingleFSTP_Offset As Long '20240226 Temp SingleFSTP solution by add a argument of GradeSearch_VT
    Harvest_With_BIN1E1 As Boolean '20240313 Harvest on BIN1 E1 will force VBINRESULT to BIN1 E1
    Instance_Offset As Long '20240417 Offset Voltage add by instance
    InitVmain_PayloadValt As Boolean '20240611 judge init for Vmain Payload for Valt in FuncPat
    Pat2DCVSOutput As New Dictionary '20240611 store pattern need to Valt or Vmain
    'F_PayloadVoltage_set As Boolean  '20240611 record whether Bincut voltage set or not
    InitValt_CorePowerStored() As New SiteDouble
End Type

'''//Data structure for COFInstance
'''20201016: Modfied to save EQN-based BinCut Payload voltage of binning P_mode. Requested by C651 Si Li.
Public Type Patt_COFInstance
    Pattern As String
    is_payload_pattern As Boolean
    grade_found As New SiteBoolean
    PASSBIN As New SiteLong
    EQN As New SiteLong
    Voltage As New SiteDouble
End Type
Public Info_COFInstance() As Patt_COFInstance

'''//
'''20200107:Add for recording first change bin mode by Si.
'''20210910: Modified to revise the format of FirstChangeBinInfo, as requested by C651 Si and TSMC ZYLINI.
Public Type FstChgBinInfo
    FirstChangeBinMode As New SiteLong
    FirstChangeBinType As New SiteLong
    str_Pmode_Test() As String 'str_Pmode_Test(MaxSiteCount - 1) As String
End Type
Public FirstChangeBinInfo As FstChgBinInfo

'''20210510: Modified to parse HARV_Pmode_Table and HARVMappingTable for Harvest Core DSSC.
Public dict_Pmode2ByMode As New Dictionary
Public dict_ByMode2Index As New Dictionary
Public dict_DisableCore2FailFlag As New Dictionary
Public dict_FailFlag2DisableCore As New Dictionary
Public dict_FailFlagOfDisableCore2DevCondition As New Dictionary
Public dict_HarvestCoreGroup2Index As New Dictionary
Public dict_EnableCore2Fstp As New Dictionary
Public dict_Fstp2EnableCore As New Dictionary
Public strAry_HarvestCoreGroupName() As String
Public strAry_HarvestCoreFstpName() As String
Public dict_offset As New Dictionary

'''20210723: Dictionary of Harv_Core_Map_Table
Public dict_HarvCoreMapTable As New Dictionary
Public dict_PwrBinning_Col_HarvBin As New Dictionary

Public Type Harv_Core_FailFlag_CoreGroup
    MainCore As Long
    Failflag As String
    DevCondition As Boolean
    overWriteSeq() As Long
    bitStart_overWriteSeq As Long
    bitStop_overWriteSeq As Long
    dict_GroupName2CoreGroup As New Dictionary
End Type

Public Type Harv_Core_ByMode_Condition
    Condition() As Harv_Core_FailFlag_CoreGroup
End Type

Public HarvCoreByMode() As Harv_Core_ByMode_Condition

Public Type Patt_Harv_Core_DSSC_Source
    Pattern() As String
    bitSeq_Core As New Dictionary
    cnt_BitSequence As Long
End Type

Public HarvCoreDSSC_BitSequence() As Patt_Harv_Core_DSSC_Source

'''SiteMask for MultiFSTP in BinCut search.
'''20210525: Modified to add siteMasks for MultiFSTP in CP1.
'''20210530: Modified to use gb_sitePassBin_original to save CurrentPassBinCutNum before MultiFSTP instances.
Public gb_siteMask_original As New SiteBoolean
Public gb_siteMask_current As New SiteBoolean
Public gb_sitePassBin_original As New SiteLong

'''********************'''
'''For the special case
'''********************'''
'''For HardIP ELB/ILB/TMPS HardIP call instance.
'''Warning!!! Remember to check if BV_Pass is used in LIB_HardIP\HardIP_WriteFuncResult.
'Public BV_Pass As New SiteBoolean

'' 20211224: Added to Print Bincut Sheet Name in Print_Bincut_Config (Originally from JadeC/D, 20210202 Jack)
Public gS_EquationVoltages_sheet As String ''20220321
Public gS_Non_Binning_Rail_sheet As String
Public gS_Vdd_Binning_Def_appA_1 As String
Public gS_Vdd_Binning_Def_appA_2 As String
Public gS_Vdd_Binning_Def_appA_3 As String
Public gS_PwrSeq_ShtName As String

'''20220622: Added for parsing bin table for BinCut
Public Sortbin2hardbinDict As New Dictionary
Public Sortbin2PmodeDict As New Dictionary
Public Sortbin2ResultDict As New Dictionary

'''20230201: Added for ffc on UFP
Public ary_FuncPat_decomposed_FFC() As String

'20230504: Added for parsing sheet defining BinX EQNs and IDS value
Public Flag_BinX_Info_Parsed As Boolean
Public Pmode2BinX_EQN_Dict As New Dictionary
Public Pmode2BinX_IDS_Dict As New Dictionary
Public Pmode2BinX_IDSHot_Dict As New Dictionary
'20240118: Added for save bincut pre-flow bin number
Public Flag_Fuse_BinX As New SiteBoolean

'''20230721: Added for T-Tah harvest flag
Public Enum harvestFlagType
    PinHarvest = 0
    OtherPin = 1
    NonHarvest = 2
    GFXHarvest = 3
End Enum

