Attribute VB_Name = "LIB_Common_GlobalConstant"
Option Explicit
'variable declaration
'Public Const Version_Lib_GlobalConstant = "0.1"  'lib version
Public Const Version_Lib_GlobalConstant = "0.2"  'add, remove efuse variant for Cayman, remove unused items

' Setup for DebugPrint
Public Const AllDCVIPinlist = "All_DCVI"
Public Const AllPowerPinlist = "All_Power"
Public Const CorePowerPinlist = "CorePower"
Public Const All_DigitalPinlist = "All_Digital" 'need to remove refclk pins
Public Const All_DigitalPinlist_Disc = "All_Digital_Disc"   'need to remove refclk, PA pins pins
Public Const All_Utility_list = "All_Utilities"
'CHWu modify 10/14 remove useless pin group
'Public Const PinGrouplist = "Pins_0pv,Pins_0p4v,Pins_0p8v,Pins_0p9v,Pins_1p1v,Pins_1p2v,Pins_1p8v,Pins_3p3v"    ',DDR_IO_GP,DDR_Vref,Efuse_Data_Out,PLL_Pins_1p8v,LPDP_IO_GP,LPDP_TX3_Diff,PCIE_IO_GP,MIPI_IO_GP,PCIE_REF,Pcie_txrx_io,Group_A,gpio20,SEP_SPI_MISO,GPIO_leak,ULPI_DIR"
Public Const PinGrouplist = "Pins_1p1v,Pins_1p2v,Pins_1p8v"    ',DDR_IO_GP,DDR_Vref,Efuse_Data_Out,PLL_Pins_1p8v,LPDP_IO_GP,LPDP_TX3_Diff,PCIE_IO_GP,MIPI_IO_GP,PCIE_REF,Pcie_txrx_io,Group_A,gpio20,SEP_SPI_MISO,GPIO_leak,ULPI_DIR"
Public Const XI0_GP = vbNullString
Public Const XI0_Diff_GP = "XI0_Diff_PA"
Public Const RTCLK_GP = "RT_CLK32768_PA"
Public Const RTCLK_Diff_GP = vbNullString
Public FreeRunFreq_debug As Double
Public clock_Vih_debug As Double
Public clock_Vil_debug As Double
Public CurrentXi0Freq As Double
Public DebugPrintFlag_Chk As Boolean

'Avoid hard rule in ceentral
Public Const XI0_Diff_Port = "XI0_Diff_Port"
Public Const RT_CLK32768_Port = "RT_CLK32768_Port"
Public Const REFCLK_XI0 = "REFCLK_XI0"
Public Const REFCLK_RT_CLK32768 = "REFCLK_RT_CLK32768"
Public Const XI0_Shmoo = "XI0_Shmoo" 'XI0_Shmoo_Freq_VAR
Public Const XI0_Diff = "XI0_Diff" 'XI0_Diff_Freq_VAR
Public Const RT_CLK32768_Diff_Port = "RT_CLK32768_Diff_Port" 'RT_CLK32768_Diff_Port_PowerSequence_GLB
Public Const XO0_Diff_Port = "XO0_Diff_Port" 'XO0_Diff_Port_PowerSequence_GLB
Public Const XO0 = "XO0"
Public Const XI0 = "XI0"

'universal
Public glb_TesterType As String
Public Const MaxNumSite = 8
Public currentJobName As String
Public CurrentChannelMap As String
Public gL_ProductionTemp As String
Public gS_SPI_Version As String    'SPIROM version printing
Public TestProgram_Day_Code As String
Public Profile_Header As String

'Control flags
Public Flag_RAK_INIT As Boolean
Public Flag_RSCR_INIT As Boolean
Public Flag_Shmoo_INIT As Boolean
Public Flag_MBISTFailBlock_INIT As Boolean
Public Flag_GetChannelType As Boolean
Public Flag_GetPowerSeq As Boolean
Public Flag_PowerSeq As Boolean
Public Flag_PowerDownSeq As Boolean
Public Flag_CZMappingTable_Check As Boolean
Public Flag_DTS_function As Boolean
Public Flag_GetCurrentProfile As Boolean
'Public Flag_Uart_INIT As Boolean

'nWire Setup, dual port
Public Const XI0_PA_Pin = "XI0_PA, XO0_PA"
Public Const XI0_1_PA_Pin = "RT_CLK32768_PA"
Public Const XI0_PA_Refclk_Pin = "REFCLK1"
Public Const XI0_1_PA_Refclk_Pin = "REFCLK2"
Public Const Clock_Port = "Clock_Port"
Public Const Clock_Port1 = "RTCLK_Port"
Public Const XI0_ref_VOH = 1.8  'use 1.8v buffer
Public Const XI0_ref_VOL = 0  'use 1.8v buffer
Public Const Relay_Off_nWire = "K0"
Public Const Relay_On_nWire = vbNullString
Public Const Relay_Off_SupportBoard = vbNullString
Public Const Relay_On_SupportBoard = vbNullString
Public Const Level_nWire = "Levels_nWire_XI0"
Public Const Level_nWire_Diff = "Levels_nWire_XI0_Diff"
Public Const TSB_nWire = "TSB_nWire_XI0"

'====================================================
'=   Define the variables for ECID Fuse data        =
'====================================================
Public XCoord As New SiteLong
Public YCoord As New SiteLong
Public WaferID As Long
Public lotid As String
Public HramWaferId As New SiteLong
Public HramLotId As New SiteVariant
Public HramXCoord As New SiteLong
Public HramYCoord As New SiteLong

Public TMPS_TD1_1 As New SiteLong
Public TMPS_TD1_2 As New SiteLong
Public TMPS_TD1_3 As New SiteLong
Public TMPS_TD1_4 As New SiteLong
Public TMPS_TD1_5 As New SiteLong
Public TMPS_TD1_6 As New SiteLong
Public TMPS_TD1_7 As New SiteLong

Public ADC_trim_V3 As New SiteLong
Public ADC_trim_V3_ECID As New SiteDouble
Public REFERENCE_CTRL_25C As New SiteLong
Public VOLTAGE_TRIM_BITS As New SiteLong
Public TEMP_TRIM_BITS1 As New SiteLong

'====================================================
'=  ECID fuse test flags                            =
'====================================================
Public FailFlag_untrim25c As New SiteBoolean 'TMPS TD1 25C
Public FailFlag_ADC25C As New SiteBoolean
Public FailFlag_Freq_Detect As New SiteBoolean

'====================================================
'=   Define the variables for Config Fuse data      =
'====================================================
Public Synth_Trim As New SiteLong
Public TRIMG_SOC_0 As New SiteLong
Public TRIMO_SOC_0 As New SiteLong
Public TRIMG_SOC_1 As New SiteLong
Public TRIMO_SOC_1 As New SiteLong
Public TRIMG_SOC_2 As New SiteLong
Public TRIMO_SOC_2 As New SiteLong
Public TRIMG_SOC_3 As New SiteLong
Public TRIMO_SOC_3 As New SiteLong


'[  for SPI - Define the IDS code resolution according to Table 32 and 33 in Test Plan ]
Public I_VDD_CPU_SPI As New SiteDouble
Public I_VDD_GPU_SPI As New SiteDouble
Public I_VDD_SOC_SPI As New SiteDouble
Public I_VDD_FIXED_SPI As New SiteDouble
Public I_VDD_CPU_SRAM_SPI As New SiteDouble
Public I_VDD_GPU_SRAM_SPI As New SiteDouble
Public I_VDD_LOW_SPI As New SiteDouble

'use in vdd binning VBT codes
Public I_VDD_CPU_IDS_Check As New SiteDouble
Public I_VDD_GPU_IDS_Check As New SiteDouble
Public I_VDD_SOC_IDS_Check As New SiteDouble
Public I_VDD_FIXED_IDS_Check As New SiteDouble
Public I_VDD_CPU_SRAM_IDS_Check As New SiteDouble
Public I_VDD_GPU_SRAM_IDS_Check As New SiteDouble
Public I_VDD_LOW_IDS_Check As New SiteDouble

Public IDS_CPU_Decimal As New SiteLong
Public IDS_GPU_Decimal As New SiteLong
Public IDS_SOC_Decimal As New SiteLong
Public IDS_FIXED_Decimal As New SiteLong
Public IDS_CPU_SRAM_Decimal As New SiteLong
Public IDS_GPU_SRAM_Decimal As New SiteLong
Public IDS_LOW_Decimal As New SiteLong

'20150610 update
Public IDS_CPU_Resolution As Double      ' 0.0002    ''0.2mA
Public IDS_GPU_Resolution As Double      ' 0.0002    ''0.2mA
Public IDS_SOC_Resolution As Double      ' 0.0001    ''0.1mA
Public IDS_FIXED_Resolution As Double    ' 0.0001    ''0.1mA
Public IDS_CPU_SRAM_Resolution As Double ' 0.0001    ''0.1mA
Public IDS_GPU_SRAM_Resolution As Double ' 0.0001    ''0.1mA
Public IDS_LOW_Resolution As Double      ' 0.0001    ''0.1mA

'20150121 define the MaxDecimal according to test plan
Public IDS_CPU_MaxDecimal As Double
Public IDS_GPU_MaxDecimal As Double
Public IDS_SOC_MaxDecimal As Double
Public IDS_FIXED_MaxDecimal As Double
Public IDS_CPU_SRAM_MaxDecimal As Double
Public IDS_GPU_SRAM_MaxDecimal As Double
Public IDS_LOW_MaxDecimal As Double


Public DPTX_LPDP0_PLL_FCAL As New SiteLong
Public PCIE_REFPLL_FCAL_VCO_DIGCTRL As New SiteLong
Public LPDP_C_RX As New SiteLong
Public LS3B As New SiteLong

'''20150207 add FCAL_VCO_DIGCTRL    'old Elba
Public FCAL_VCO_DIGCTRL_Decimal As New SiteLong
Public PCIE_FCAL_VCO_DIGCTRL_1st_Value As New SiteLong
Public PCIE_FCAL_VCO_DIGCTRL_2nd_Value As New SiteLong

'''20150819 add PLL_CPU_KVCO    'Cayman
Public PLL_CPU_KVCO_Decimal As New SiteLong


'''20150819 add PLL_LPDP_FCAL    'Cayman
Public PLL_LPDP_FCAL_Decimal As New SiteLong

'''20160725 add ADCLK trim    'Skye
Public PLL_GPU_FCAL_Decimal As New SiteLong
Public pblk_PLL_CFG1_kvco_trim_Decimal As New SiteLong
Public eblk_PLL_CFG1_kvco_trim_Decimal As New SiteLong

'====================================================
'=  Config fuse test flags                          =
'====================================================
Public FailFlag_Fcal_LPDP As New SiteBoolean
Public FailFlag_TrimVerify85c As New SiteBoolean    'trimG, trimO
Public FailFlag_Freq_Synth As New SiteBoolean
Public FailFlag_FCAL_VCO As New SiteBoolean

'Real VddBinning Check revision
Public Const Real_VddBinning_version = 99

'====================================================
'=    Define the variables for UDR Fuse data        =
'====================================================
Public TRIMG_CPU_0 As New SiteLong
Public TRIMO_CPU_0 As New SiteLong
Public TRIMG_CPU_1 As New SiteLong
Public TRIMO_CPU_1 As New SiteLong
Public TRIMG_CPU_2 As New SiteLong
Public TRIMO_CPU_2 As New SiteLong
Public ADC_vTrim As New SiteLong
Public ADC_tTrim As New SiteLong
Public pllTrimFusedBit As New SiteLong
Public PLL_CFG1_kvco_trim As New SiteLong
Public ADCLK_SCR2_vsns_cal_fuse As New SiteLong


'====================================================
'=  UDR fuse test flags                             =
'====================================================
Public FailFlag_ADC85C As New SiteBoolean
Public FailFlag_pllTrimFusedBit As New SiteBoolean
Public FailFlag_PLL_CFG1_kvco_trim As New SiteBoolean
Public FailFlag_ADCLK_SCR2_vsns_cal_fuse As New SiteBoolean

'20150319 add CPU PLL Fcal  'old Elba
''Public CPU_PLL_Fcal_Decimal As New SiteLong
''Public CPU_PLL_Fcal_V1_Decimal As New SiteLong

'====================================================
'=   Define the variables for Sensor Fuse data      =
'====================================================
Public AFREQ_CTRL_EN As New SiteLong
Public LATENCY As New SiteLong
Public Freq_Det_Precision As New SiteLong
Public Freq_Det_Decimal As New SiteLong
Public DFREQ_CTRL_EN As New SiteLong
Public DFREQ_CTRL_OFFSET As New SiteLong
Public SEN_SOC_TRIMG_0 As New SiteLong ''''should be equal to TRIMG_SOC_0
Public SEN_SOC_TRIMO_0 As New SiteLong ''''should be equal to TRIMO_SOC_0

'Public gS_SEN_CRC_HexStr As New SiteVariant



'====================================================
'=   Define the variables for IEDA registry     =
'====================================================
Public gS_TMPS1_Untrim As New SiteVariant
Public gS_TMPS2_Untrim As New SiteVariant
Public gS_TMPS3_Untrim As New SiteVariant
Public gS_TMPS4_Untrim As New SiteVariant
Public gS_TMPS5_Untrim As New SiteVariant
Public gS_TMPS6_Untrim As New SiteVariant
Public gS_TMPS7_Untrim As New SiteVariant
Public gS_TMPS8_Untrim As New SiteVariant
Public gS_TMPS9_Untrim As New SiteVariant
Public gS_TMPS10_Untrim As New SiteVariant
Public gS_TMPS11_Untrim As New SiteVariant
Public gS_TMPS12_Untrim As New SiteVariant
Public gS_TMPS13_Untrim As New SiteVariant
Public gS_TMPS14_Untrim As New SiteVariant

Public gS_TMPS1_Trim As New SiteVariant
Public gS_TMPS2_Trim As New SiteVariant
Public gS_TMPS3_Trim As New SiteVariant
Public gS_TMPS4_Trim As New SiteVariant
Public gS_TMPS5_Trim As New SiteVariant
Public gS_TMPS6_Trim As New SiteVariant
Public gS_TMPS7_Trim As New SiteVariant
Public gS_TMPS8_Trim As New SiteVariant
Public gS_TMPS9_Trim As New SiteVariant
Public gS_TMPS10_Trim As New SiteVariant
Public gS_TMPS11_Trim As New SiteVariant
Public gS_TMPS12_Trim As New SiteVariant
Public gS_TMPS13_Trim As New SiteVariant
Public gS_TMPS14_Trim As New SiteVariant

Public gS_TMPS1 As New SiteVariant
Public gS_TMPS2 As New SiteVariant
Public gS_TMPS3 As New SiteVariant
Public gS_TMPS4 As New SiteVariant
Public gS_TMPS5 As New SiteVariant
Public gS_TMPS6 As New SiteVariant
Public gS_TMPS7 As New SiteVariant
Public gS_TMPS8 As New SiteVariant
Public gS_TMPS9 As New SiteVariant
Public gS_TMPS10 As New SiteVariant
Public gS_TMPS11 As New SiteVariant
Public gS_TMPS12 As New SiteVariant
Public gS_TMPS13 As New SiteVariant
Public gS_TMPS14 As New SiteVariant

'====================================================
'=   Define the variables for alarm happened     =
'====================================================
Public alarmFail As New SiteBoolean

'====================================================
'=   Define the variables for mbist loop     =
'====================================================
Public mbist_sheet_init As Boolean
Public currentBlock_loopCnt As Integer
Public currentAPK_loopCnt As Integer
'====================================================
'=   Define DAC Trim Flag                           =
'====================================================
'Public DACInitialFlag As Boolean

'========================
'HardIP Test Name
'========================
Public gl_Tname_Alg_Index As Long
Public gl_Tname_Meas As String
'----------------20180523----------------
Public gl_Tname_Meas_FromFlow() As String
Public gl_Tname_Alg As String
Public gl_Sweep_Name As String
Public gl_SweepY_Name As String
'=====================================================
'20171207 - HardIP use Standard Test Name Format Flag,Roger add
Public gl_UseStandardTestName_Flag As Boolean
'=====================================================
Public gl_Disable_HIP_debug_log As Boolean
Public gl_Disable_IDS_AutoRange_log As Boolean

Public XVal As Double
Public YVal As Double
Public gl_flag_end_shmoo As Boolean
Public gl_flag_CZ_Nominal_Measured_1st_Point As Boolean

Public gl_FlowForLoop_DigSrc_SweepCode As String

'========================
'Powerup/down sequence
'========================
Public Type PowerSeqeuce
    nWireSeq() As Integer
    nWirePort() As String
    IOHSeq() As Integer
    IOHPin() As String
    IOLSeq() As Integer
    IOLPin() As String
    IOHZSeq() As Integer
    IOHZPin() As String
    PowerSeqPin() As String
    PowerSeq99Log As String
    PowerSeqNCLog As String
End Type

Public Const PowerUpSeqName As String = "_PowerSequence"
Public Const PowerDownSeqName As String = "_PowerDownSequence"
Public Const IOInitHi As String = "_GLB_InitHi"
Public Const IOInitLo As String = "_GLB_InitLO"
Public Const IOInitHZ As String = "_GLB_initHZ"
Public pwrUpSeq As PowerSeqeuce
Public pwrDownSeq As PowerSeqeuce
Public nonPowerSeqNum As New Dictionary
Public maxSeqNum As Integer
Public powerUpEnable As Boolean
Public powerDownEnable As Boolean
Public powerUpDone As New SiteBoolean
Public powerUpPins As String
'================================================
''Public site As Variant 'Carter, 20240304

Public gL_License_check As Long
Public glb_TestInstance As String

'''''Add by TY
Public glb_TestTimeCollect_Dict As New Dictionary
'''''Add by TY
Public glb_CurrentProfile_Dict As New Dictionary
Public glb_ExecutionProfile_Dict As New Dictionary

'========================
'Power Pin Information
'========================
Public Pin_range_ary() As AutoRange_Info
Public PowerPin_range_ary() As AutoRange_Info
Type AutoRange_Info
    MergedN As Long
    Init_step As Long
    PinName As String
    MergeType As String
    PinMapType As String
    ChanMapType As String
    HiLimit As Double
    LoLimit As Double
    PowerSeq As Double
    PowerDownSeq As Double
    Range_List() As Double
    Accuracy_List() As Double
    WaitTime_List() As Double
    Init_CurrentRange As Double
    Init_Source_FoldLimit As Double
    MinIFoldLimit As Double
End Type

Public Flag_IDSMappingTable As Boolean
Public IDS_MAPPING() As IDS_Mapping_Info
Type IDS_Mapping_Info
    Stage As String
    MappingDict As New Dictionary
    StageCell As Long
    cnt As Long
    InstanceName As String 'Added for consider InstanceNamme 20240612
End Type

'========================
'Align AP,LCD,RF VBT
'========================
Public Const gB_Enable_AP = True
Public Const gB_Enable_LCD = False
Public Const gB_Enable_RF = False

Public Enum ProjType
    Type_AP = 0
    Type_LCD = 1
    Type_RF = 2
End Enum

Public Enum BlockType
    Type_Conti = 0
    Type_PowerShort
    Type_IDS
    Type_HIP
End Enum

Public GlbCustomizeSet As CustomizeSetting
Public Name_Flag As Integer

Public gl_Flag_1st_contact_R As Boolean
Public First_Contact_R As New PinListData

Public PinLevelName() As String
Public ParsePreConditionDone As Boolean
Public PreConditionInfo() As preCondition
Public dic_forceV As New Dictionary
Public dic_VIL As New Dictionary
Public dic_mapCritcalPin As New Dictionary
Public CurrentPatIndex As Long
Type preCondition
    ''String Type - "pin1,pin2,pin3,pin4,..."
    PatternName As String
    PatternVersion As String
    PatCell As Long
    XPins() As String
    UnusedPins() As String
    InputPins As String
    CriticalPins As String
    ''Array
    ''Arr(0) - "pin1"
    ''Arr(1) - "pin2"
    ''Arr(2) - "pin3" ....
    AllPins() As String
    Arr_XPins() As String
    Arr_UnusedPins() As String
    Arr_AllPins() As String
    Arr_InputPins() As String
    Arr_CriticalPins() As String
End Type

Public ENG_SweepPin As Boolean
Public Const Enable_SweepPinPrintOut = False
Public dic_IgnorePin As New Dictionary  ''20211109
Public Const Threshold_CriticalPin = 5
Public Const Threshold_EnableSweepPin = 5
Public Const Threshold_VT_Half = 0.000001
Public Const Threshold_VT_Vil = 0.0000002
Public FindIDSPattern As Boolean
Public ENG_Limit As Boolean
Public Const SkipAllPinToHalf = True
Public isDebugMode As Boolean
Public gl_isCheckPreConditionVerDone As Boolean
Public gl_isParsePatSetAll As Boolean
Public PatSetAllDic As New Dictionary
Public gl_isParsePreConditionDone As Boolean

Public glb_SVN As String
Public glb_ModuleName As String

Public Enum Error_Warning_Info
    Error_Info = 0
    Warning_Info = 1
End Enum

Public Const gl_HRAMmaxDepth = 512        'Flex UP1600 max depth 512    'Plus org max depth 16k, but use 512 to let speed faster

Public Enum DisconnectPinType
    UnusedIoPin = 0
    Xpin = 1
End Enum

Public gl_PreCondition_Dic As New Dictionary

Public gl_ePreConditionError As PreConditionError

Public Enum PreConditionError
    PreConditionPass = 0
    SheetNotExist = 1
    PatVerDifferent = 2
    PatRepetition = 3
    ParseSheetError = 4
End Enum

Public gl_isPatAndCZmappingTable As Boolean
Public gl_isFind_nWire_Pin As Boolean

'Move from LIB_HardIP.bas to here
Public gl_GetInstrument_Dic As New Dictionary
Public gl_GetInstrumentType_Dic As New Dictionary
Public gl_dicPowerPinIndex As New Dictionary

Public gl_isParExecutionProfileSheet As Boolean
Public gl_EnableCurrentProfile As Boolean
Public gl_EnableVoltageProfile As Boolean
Public glb_ProfileFilter As Boolean
Public glb_DisablePlot_IProfile As Boolean
Public glb_Print_IProfileValue As Boolean

Public gl_nWireFreq As Double
Public glb_EVS_Disable_Printout As Boolean

Public gl_isCheckClampLimit As ContiClampCheckType   '20230307 check conti clamp use
Public Enum ContiClampCheckType
    CheckInit = 0
    CheckFail = 1
    CheckPass = 2
End Enum
Public gl_saPatSetAllPrintInfo() As String

Public Const glbConstIns_HEXVS = "HEXVS"
Public Const glbConstIns_VSM = "VSM"
Public Const glbConstIns_VHDVS = "VHDVS"
Public Const glbConstIns_VS800MA = "VS-800MA"
Public Const glbConstIns_VS5A = "VS-5A"
Public Const glbConstIns_DC07 = "DC-07"
Public Const glbConstIns_DC30 = "DC-30"
Public Const glbConstIns_DC75 = "DC-75"
Public Const glbConstIns_UP2200 = "HSDP"
Public Const glbConstIns_UP1600 = "HSD-U"

Public gl_nWireFreq_Value_Dict As New Dictionary ''Add for CZ PrintShmooInfo, 20230316
Public gl_nWireFreq_AC_Dict As New Dictionary ''Add for CZ PrintShmooInfo, 20230316
Public gl_bTTRDisableAlarm As Boolean

'==== print information ====
'add for print version
Public glb_isVersion_Info As Boolean
Public gls_PatVersion As String
Public gls_TestPlanVersion As String
Public gls_DC_Table_Ver As String
Public gls_SCGHVersion As String

Public glb_isOS_Info As Boolean
Public gls_OperatingSystem As String
Public gls_NumberOfProcessors As String
Public gls_ProcessingBits As String
Public gls_PhysicalMemory As String

Public glb_isIGXL_Info As Boolean
Public gls_SoftwareVersion As String
Public gls_SoftwareBuild As String

Public glb_isRunOptions_Info As Boolean
Public gls_DoAll As String
Public gls_OverrideFailStop As String
Public gls_Assume As String
Public gls_AssumeSiteDisable As String
Public glb_isParsing_EW As Boolean
Public gls_Active_EW As String
Public gls_Disable_EW As String
'==== print information ====

'Harvest global variable
Public Harvest_Pin_From_Table_Flag As Boolean 'Add this boolean for ATPG HarvestPinGrpFlagTable
Public EnableCoreHarvest As Boolean 'Add this boolean to decide whether ATPG Harvest Fail Flag Set TRUE
Public EnableCoreMask As Boolean 'Add this boolean to decide whether DisableCompare ATPG Harvest Pin/pin group
Public glb_Harvest_Fail_Flag_From_eFuseRead_Dict As New Dictionary

'20240124: New MFSTP and user defined function
Public Type UserDefinedFunctionType
    digSrcLabel() As String
    digSrcPatterns() As String
End Type
Public DigSrcPatternDict As New Dictionary

Public Find_shmoo_hole_low_power_twice_hvcc_step() As New SiteLong '20231018
Public Find_shmoo_hole_low_power_twice_hvcc_voltage() As New SiteDouble '20231018
'Public find_shmoo_hole As SiteBoolean '20231018
Public gls_PowerDownDisconectIOPin As String

Public ALL_Power_DCVS_pins As String
Public ALL_Power_DCVI_pins As String
Public Core_Power_DCVS_pins As String
Public Core_Power_DCVI_pins As String

Type Excution_Profile_Info
    sFlowName As String
    ProcessItemDict As New Dictionary   'Inatance -> Process Time
    dTotalProcessTime As Double
End Type

Public glb_ApplyLevelTiming_FRC_Flag As Boolean 'ApplyLevelTiming one time each touch down for SBC Free Running Clock Check
Public glb_FlagCheckingFRCClock As Boolean 'For SBC Free Running Clock check

'''' PPMU test walkingZ fail pin
Public gldic_ComposeFailPins As New Dictionary
Public glb_CheckIDSMappingTable_With_Fuse As Boolean

'==== SFC ====
Public glb_isParsingSFC As Boolean
Public glb_isSFC_Enabled As Boolean
Public glb_SFC_Scan_Check As Boolean

Type TableScanPatterns
    instancename As String
    FailCycleCount As Integer
    ScanPatternName As String
End Type

Public ScanPatternsList() As TableScanPatterns
'==== SFC ====

Public F_p2pWT As Boolean ' Add for CRWT

' Binout default value
Public Const glb_SortBin_Bin0 = 9987
Public Const glb_HardBin_Bin0 = 15

Public glb_SpecialPin_dic As New Dictionary 'Added for store special pins pair

'==== Harvest Instance Check ====
Public glb_CheckHarvestInst As Boolean
'Current Profile new feature 20241009
Public gls_DoingCurrentProfile As String
Type CurrentCaptureFlowInfo
    s_FlowName As String
    dic_InstanceInfo As New Dictionary
    s_CapturePins As String
    d_SampleRate As Double
    dic_DupicateInst As New Dictionary
    b_IsRunning As Boolean
End Type

Public glb_CurrentCaptureFlowInfo As CurrentCaptureFlowInfo
Public gl_CaptureFlowInfoAry() As CurrentCaptureFlowInfo
Public gldic_Lost_Header_Footer_Flow As New Dictionary


Public glb_CheckFlowHeaderFooter As Boolean
Type CurrentProfilePlot_Info
    s_StartInstance As String
    s_EndInstance As String
    s_PlotPins As String
    d_SampleRate As Double
    d_PlotTime As Double
    b_isPlotFlow As Boolean
    n_StartIndexDuplicateIndex As Long
    n_EndIndexDuplicateIndex As Long
End Type
Public gldic_CaptureFlowInfo As New Dictionary
Public glb_Boolean_export As Boolean
Public glb_WhatToCapture As String

Public T_ProfileStart As Double
Public profileAction As New Dictionary
Public check_profile As Boolean
Public profile_count As Integer
Public Profile_byflow As Boolean
Public glb_HexVs_0p1range_waittime As Double
