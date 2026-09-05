Attribute VB_Name = "VBT_LIB_HardIP"
#Const isUFP = True
' dummy comment
Option Explicit

Public gl_Sweep_V_pin As String 'VTHSENSOR_Calibration Tname using
Public gl_ForceVlotagesweep_VTHSENSOR As New SiteDouble  'VTHSENSOR_Calibration Tname using
Public gl_TName_Pat As String           'Roger add, For TName
Public gl_SweepNum As String
Private m_InstanceName As String
'Public gl_Flag_HardIP_Characterization_1stRun As Boolean
'Public DACInitialFlag As Boolean
'==========================================================Roger Add,for power sweep
Type Power_Sweep
    Loop_count As Long
    Loop_Index_Name As String
    PinName As String
    from As String
    stop As String
    step As String
    Count As Long
    key As String
End Type
'==========================================================
'' 20151229 - hard ip dssc code, add rd-rd9
Public Src_DSPWave As New DSPWave
Public Src1_DSPWave As New DSPWave
Public Src2_DSPWave As New DSPWave
Public Src3_DSPWave As New DSPWave
Public Src4_DSPWave As New DSPWave
Public Src5_DSPWave As New DSPWave
Public Src6_DSPWave As New DSPWave
Public Src7_DSPWave As New DSPWave
Public Src8_DSPWave As New DSPWave
Public Src9_DSPWave As New DSPWave

'' 20160121 - hard ip dssc code, add rd10-rd12 for Starling
Public Src10_DSPWave As New DSPWave
Public Src11_DSPWave As New DSPWave
Public Src12_DSPWave As New DSPWave

Public CP_Card_RAK As New PinListData
Public FT_Card_RAK As New PinListData
Public WLFT1_Card_RAK As New PinListData
Public CurrentJob_Card_RAK As New PinListData
Public FourceV As Double

Public ADDRIO_Norm_Y_T1() As Double
Public ADDRIO_Norm_Y_T2() As Double
Public ADDRIO_Norm_Y_T1_T2_ReadOnce_Flag As Boolean

Enum CalculateMethodSetup_PPMU
    PPMU_DEFAULT = 0
    'PPMU_STORE_I = 1
    VIR_DIFF_PN_ABS = 4
    VIR_DIFF_PN = 5
    VIR_VOD_VOCM_XI0_Off = 8
    VIR_VOD_VOCM_PN = 9
    VIR_DDIO = 10
End Enum

Enum CalculateMethodSetup_DSPWave
     DigCap_DEFAULT_SETUP = 0
     DigCap_MultiPinsOperation = 1
End Enum

'Enum InstrumentSpecialSetup_PPMU
'     DEFAULT_SETUP = 0
'     PPMU_SerialMeasurement = 1
'     DigitalConnectPPMU2 = 2 ' 20160204
'End Enum

Enum CalculateMethodSetup
    DEFAULT_SINGLE = 0
    DIFF_1ST = 1
    DIFF_2ND = 2
    DIFF_PT12 = 3
    DIFF_PN = 4
    DIFF_DCO = 5
    DIFF_DAC = 6
    Force_VDD12_RX = 7
    RATIO_FREQ = 8
    'Different with Sicily ,20200423, Oscar
    PPMU_TestLimit_TTR = 15
    Average_voltage = 10
    PPMU_STORE_I = 11
    VIR_DIFF_PN_ABS = 12
    VIR_DIFF_PN = 13
    VIR_VOD_VOCM_XI0_Off = 14
    'Different with Sicily ,20200423, Oscar
    VIR_VOD_VOCM_PN = 9
    VIR_DDIO = 16
''    DigCap_MultiPinsOperation = 4
    DigCap_DEFAULT_SETUP = 17
    DigCap_MultiPinsOperation = 18
    'TTR,20200423, Oscar
    BYPASS_LIMIT = 19
    'Add in C-Chop
    ABS_Current = 20
    IO_DCDS8_differenceLimit_by1P2or1P8 = 21
End Enum


Type DSSC_CodeSearchCond
    MeasureValue As New SiteDouble
    SearchCode As New SiteLong
    TargetCodeFind As New SiteBoolean
    TransitionPoint As New SiteVariant
    PatternPass As New SiteBoolean
End Type

Enum InstrumentSpecialSetup
     DEFAULT_SETUP = 0
     DigitalConnectPPMU = 1
     'Different with Sicily ,20200423, Oscar
     PPMU_SerialMeasurement = 3 'For turks HIP_USB
     'Different with Sicily ,20200423, Oscar
     DigitalConnectPPMU2 = 2 'For turks HIP_USB
     PPMU_AccurateMeasurement = 4
     PPMU_2mA_Force_I_Range = 5
     PPMU_200uA_Force_I_Range = 6
     PPMU_20uA_Force_I_Range = 7
     EUSB_T10T11_Split_Force_I_Range = 8
     HighAccuracyVoltage = 9
     HighAccuracyVoltageHiz = 10
     Bypass_Setup = 11
         IO_DS_VOHVOL_with_RAK = 16  'New Method for IO_DS_VOHVOL_with_RAK --20220421
End Enum

Public Stored_MeasI_PPMU As New PinListData

'' 20151117 - Event source combine HiZ/VT mode
Enum EventSourceWithTerminationMode 'not use, for reference.
     BOTH_VT = 1
     VOH_VT = 2
     VOL_VT = 3
     BOTH_HIZ = 4
     VOH_HIZ = 5
     VOL_HIZ = 6
End Enum

Enum Enum_RAK
     Default = 0
     R_TraceOnly = 1
     R_PathWithContact = 2
     R_TracePN = 3
End Enum

'' 20151224 - Print measured frequency during shmoo if need
Public G_MeasFreqForCZ As New PinListData

Public SweepWhileFirstTimeFlag As New SiteBoolean ''VBT_LIB_HardIP
Public SourceCode_First6Bit As New SiteVariant ''VBT_LIB_HardIP
Public SourceCode_Last6Bit_MSB As New SiteVariant ''VBT_LIB_HardIP
Public SourceCode_Last6Bit As New SiteVariant ''VBT_LIB_HardIP
Public SourceCode_First6Bit_MSB As New SiteVariant ''VBT_LIB_HardIP
Public SourceCode_First6Bit_Code0To4 As New SiteVariant ''VBT_LIB_HardIP
Public SourceCode_Last6Bit_Code0To4 As New SiteVariant ''VBT_LIB_HardIP
Public Dec_First6Bit_Code0To4 As New SiteDouble ''VBT_LIB_HardIP
Public Dec_Last6Bit_Code0To4 As New SiteDouble ''VBT_LIB_HardIP
Public Final_Dec_First6Bit_Code0To4 As New SiteDouble ''VBT_LIB_HardIP
Public Final_Dec_Last6Bit_Code0To4 As New SiteDouble ''VBT_LIB_HardIP
Public PostiveIndex As New SiteLong ''VBT_LIB_HardIP
Public Source12Bits As New SiteVariant ''VBT_LIB_HardIP
Public Final_point_flag As New SiteBoolean ''VBT_LIB_HardIP
Public Final_point As New SiteLong ''VBT_LIB_HardIP
Public NegativeIndex As New SiteLong ''VBT_LIB_HardIP
Public Imped_LowLimit As Double ''VBT_LIB_HardIP, ImpedanceMeasurement_2Point
Public Imped_HighLimit As Double ''VBT_LIB_HardIP, ImpedanceMeasurement_2Point

Public G_pld_DigCapInfo() As New PinListData ''VBT_LIB_HardIP

Enum DDR_Eye_setup
    DDR_EYE_False = 0
    DDR_EYE_1ST = 1
    DDR_EYE_2ND = 2
End Enum

''20160729 - Use global value to denfine default setting
Public Const pc_Def_VFI_FreqInterval = 0.001
Public Const pc_Def_VFI_FreqThresholdPercentage = 0.5
Public Const pc_Def_VFI_MeasCurrRange = 0.02

Public Const pc_Def_VIR_MeasCurrRange = 0.05

Public Const pc_Def_VFI_UVI80_VoltCalmp = 6
Public Const pc_Def_VFI_UVI80_InitialVal_FI = 0
Public Const pc_Def_VFI_UVI80_ReadPoint = 2
Public Const pc_Def_VFI_UVI80_VoltageRange = 7
Public Const pc_Def_VFI_UVI80__InitialVal_FI = 0

Public Const pc_Def_Default_Range_By_Instrument = 0
Public Const pc_Def_UVI80_Init_MeasCurrRange = 0.2

Public Const pc_Def_PPMU_InitialValue_FI = 0
Public Const pc_Def_PPMU_InitialValue_FI_Range = 0.05
Public Const pc_Def_PPMU_Max_InitialValue_FI_Range = 0.05
Public Const pc_Def_PPMU_InitialValue_FV = 0
Public Const pc_Def_PPMU_InitialValue_FV_Range = 0
Public Const pc_Def_PPMU_FI_Range_200uA = 0.0002
Public Const pc_Def_PPMU_ReadPoint_FreqDC = 20
Public Const pc_Def_PPMU_ReadPoint = 10
Public Const pc_Def_PPMU_ClampVHi = 2
'Different with Sicily ,20200423, Oscar
Public Const pc_Def_PPMU_Digital_MaxCurrRange = 0.05

Public Const pc_Def_DSSC_Amplitude = 1

Public Const pc_Def_HexVS_ReadPoint = 10000

Public Const pc_Def_UVS256_ReadPoint = 1
Public Const pc_Def_UVS256_CurrentRangeRatio = 0.09

'---------------------------------------UFP_Corr fix 20200413
Public Const pc_Def_UVS256HP_Init_MeasCurrRange = 0.02
Public Const pc_Def_UVS64_Init_MeasCurrRange = 0.02
Public Const pc_Def_UVS256HP_ReadPoint = 1
Public Const pc_Def_UVS64_ReadPoint = 1
Public Const pc_Def_UFP_DCVS_VoltageRange = 3           'For UVS256HP/UVS64 setup on UFP --20230901
Public Const pc_Def_UFP_DCVS_Relax_VoltageRange = 5.5   'For UVS256HP/UVS64 setup on UFP --20230901
'---------------------------------------UFP_Corr fix 20200413

Public Const pc_Def_Power1p2 = 1.2

Public Const pc_Def_DCTIME_InitialCurrent = 0
Public Const pc_Def_DCTIME_SampleSize = 10
  
Public Const pc_Def_DiffMeter_HWAverageSize = 64
Public Const pc_Def_DiffMeter_VoltRange = 1.4
Public Const pc_Def_DiffMeter_ReadPoint = 100000
Public Const pc_Def_HardIP_PatGenTimeout = 10

Public Const pc_Str_InitOff_Pins = "All_Digital"

Public Const pc_Def_VFI_MI_WaitTime = 1 * ms
Public Const pc_Def_VFI_MI_WaitTime_PPMU = 100 * us
Public Const pc_Def_VFI_MI_WaitTime_UVI80 = 1 * ms

'Public FlowShmooString_GLB As String               '20170717 already defined in Lib_Digital_Shmoo.bas

Public Const PPMU_SettlingTime = 0.00001
Type MeasTrimImpedInfo
    Pat As String
    MeasPinsAry() As String
    IsDifferential As Boolean
End Type

''20170405-Global string to record all functional result if flow for loop specified (gray code)
Public gs_RecordGrayCodeTestResult As String

'' 20170523
Public Const pc_Def_VFI_ForceI_Val = 0

''20170518-Global string to Customize DigCap
Public MTR_CusDigCap As String
Public MTR_VIN As String
''20170524-Global string to T4P3
Public MTR_T4P3_MeasVolt_P(1023) As New PinListData
Public MTR_T4P3_MeasVolt_N(1023) As New PinListData
Public MTR_T4P3_DigSrc(255) As String
Public StoreIndex_MTR_P As Integer
Public StoreIndex_MTR_N As Integer
''20170605-Globastring forMetrology T2P6 CMRR
Public CMRR_VIN As Double
''20170606-Globastring forMetrology T2P6 CMRR average
Public CMRR_Average(71) As New DSPWave
Public StoreIndex_CMRR_Average As Integer
Public PSRR_Average(35) As New DSPWave
Public StoreIndex_PSRR_Average As Integer
'Store for Avg Dic for metrology 20170711
Public MetroAvgDic As New Dictionary

Public gl_CZ_FlowTestNameIndex As Long
Public gl_DSSC_OUT_STR As String
Public gl_DSSC_CALC_STR As String
Public gl_Sweep_Glb_TName As String '' 20190529 - Add for sweep force V

Public gl_Flag_HardIP_Characterization_1stRun As Boolean
Public gl_Flag_HardIP_Trim_Set_PrePoint As Boolean
Public gl_Flag_HardIP_Trim_Set_PostPoint As Boolean
Public gl_Flag_HardIP_Disable_Functional_Result As Boolean
'Sweep name for Sweepsrc
Public Sweepnameforsweep() As String
Public temp_CUS_String As String
Public srcnameindex As Integer
Public HIP_TTR_Enable_Control_Flag as boolean	'[20240207][All][Brian] Check the enable for one time
Type EyeDiagram
    value(63) As New SiteVariant
End Type

'Add from Turks, for Calc.
Public sweep_power_val_per_loop_count As String

'Add from Sicily, for Sicily/Tonga/C-Chop, MTR processing on PTE function.
Type MTRSNS_Matrix
    ROT_Matrix() As Double
    ROV_Matrix() As Double
    ROT_a_max_min_Matrix() As Double
    ROV_a_max_min_Matrix() As Double
End Type
Public MetrologySense_Matrix() As MTRSNS_Matrix
Public glb_DSSCSetup As String ''RF DSSC, 221212
'Use for SweepVoltage function --20221018
Public gl_sweepVoltage_Loop_Indx As Long                                  'Record Sweep Voltage Loop Index
Public gl_sweepCode_Loop_Indx As Long
Public gl_sweepVoltage_Value As Double


Public Function Meas_Vdiff_func(Optional patset As Pattern, Optional DisableComparePins As PinList, Optional DisableConnectPins As PinList, Optional TestSequence As String, Optional CPUA_Flag_In_Pat As Boolean, _
    Optional PpmuMeasureP_Pin As String, Optional PpmuMeasureN_Pin As String, _
    Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, _
    Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As String, Optional DigSrc_Sample_Size As String, Optional DigSrc_Equation As String, Optional digsrc_assignment As String, Optional DigSrc_FlowForLoopIntegerName As String = vbNullString, _
    Optional ForceV1p As String, Optional ForceV2p As String, Optional ForceV1n As String, Optional ForceV2n As String, _
    Optional CUS_Str_MainProgram As String = vbNullString, Optional CUS_Str_DigCapData As String = vbNullString, Optional CUS_Str_DigSrcData As String = vbNullString, _
    Optional Interpose_PrePat As String, Optional Interpose_PreMeas As String, Optional Interpose_PostTest As String, _
    Optional Meas_StoreName As String, Optional Calc_Eqn As String, Optional TestLimitPerPin_VFI As String = "FFF", _
    Optional DSSCSetup As String, Optional Validating_ As Boolean) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    
    Dim PatCount As Long
    Dim PattArray() As String
    
    If Validating_ Then
        Call PrLoadPattern(patset.value)
        Exit Function    ' Exit after validation
    End If

    Call HardIP_InitialSetupForPatgen


    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)
        
    Dim pat_count As Long
    Dim i As Long, k As Long, j As Long
    Dim TestOptLen As Integer
    Dim TestSequenceArray() As String
    Dim TestOption As Variant
    Dim Ts As Variant
    Dim TestSeqNum As Integer
    Dim testnum As Long
    Dim TestNumber As Long
    Dim patt_ary() As String
    Dim InDSPWave As New DSPWave
    Dim OutDspWave As New DSPWave
    Dim show_Dec As String
    Dim show_out As String
    Dim site As Variant

    Dim patt As Variant
    Dim Pat As String
    Dim HighLimitVal() As Double
    Dim LowLimitVal() As Double
    Dim Idiff As New SiteDouble, Vdiff As New SiteDouble, Vocm As New SiteDouble
    Dim TxVDD As New SiteDouble, TxVss As New SiteDouble
    Dim MeasIp1 As New PinListData
    Dim MeasIp2 As New PinListData
    Dim MeasIn1 As New PinListData
    Dim MeasIn2 As New PinListData
    Dim MeasVdiff As New PinListData
    Dim MeasVocm As New PinListData
    Dim ForceV1pAry() As String
    Dim ForceV2pAry() As String
    Dim ForceV1nAry() As String
    Dim ForceV2nAry() As String
                                                                                                                                                                                                                                                               
    Dim Zp As New SiteDouble
    Dim Zn As New SiteDouble
    
    Dim TestNameInput As String
    Dim OutputTname_format() As String
 
    Dim TestPinArrayP() As String, TestPinArrayN() As String
    Dim pin As Variant

    ''20160923 - Analyze Interpose_PreMeas to force setting with different sequence.
    Dim Interpose_PreMeas_Ary() As String
    Interpose_PreMeas_Ary = Split(Interpose_PreMeas, "|")
    
    ''Defined for TTR
    Dim DiffPins As New PinList
    
    Dim TName_Ary() As String
    Dim Tname As String
    'Dim Inst_Name_Str As String: Inst_Name_Str = glb_TestInstance
    '----180524----------------------------------------------------------------
    Call GetFlowTName
    
    DiffPins = PpmuMeasureP_Pin & "," & PpmuMeasureN_Pin
                                                                                                                                                                                                                                                           
    TestSequenceArray = Split(TestSequence, ",")

    TestPinArrayP = Split(PpmuMeasureP_Pin, "+")
    TestPinArrayN = Split(PpmuMeasureN_Pin, "+")
    
    ForceV1pAry = Split(ForceV1p, "+")
    ForceV2pAry = Split(ForceV2p, "+")
    ForceV1nAry = Split(ForceV1n, "+")
    ForceV2nAry = Split(ForceV2n, "+")
    
    Call GetFlowSingleUseLimit(HighLimitVal, LowLimitVal)  ''20141219
    Instance_Data.Is_PreCheck_Func = False
    Call ProcessInputToGLB(patset, , , , , , , , , , , , , , , , , , , , DigSrc_pin, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, DigSrc_FlowForLoopIntegerName)
    ''20161130-Get test name from flow table
    Dim FlowTestNme() As String
    If TPModeAsCharz_GLB Then
        gl_CZ_FlowTestName_Counter = 0
        Call GetFlowTestName(FlowTestNme)
    End If
                                                                                                                                                                                                                                                               
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    '' 20160923 - Add Interpose_PrePat entry point
    If Interpose_PrePat <> "" Then
        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    End If
    
    'TName_Ary = Split(gl_Tname_Meas, "+")
    
'    If (UBound(TestSequenceArray) > UBound(TName_Ary)) Then
'        ReDim Preserve TName_Ary(UBound(TestSequenceArray)) As String
'
'    End If
    
    TheHdw.Patterns(patset).Load
    
    Call PatternBurstCheckAndSplit(patset.value, patt_ary, pat_count)
    
    If (DisableConnectPins <> "") Then TheHdw.Digital.Pins(DisableConnectPins).Disconnect
    If (DisableComparePins <> "") Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = True
    
    ''20161107-Return sweep test name
    Dim Rtn_SweepTestName As String
    Rtn_SweepTestName = vbNullString
    
    ''========================================================================================
    ''20170203 - Analyze Meas_StoreName and store the measurement for futher use.
    Dim Rtn_MeasVolt As New PinListData
    Dim MeasStoreName_Ary() As String
    MeasStoreName_Ary = Split(Meas_StoreName, "+")
    Dim Store_Rtn_Meas() As New PinListData
    Dim SoreMaxNum As Long
    Dim StoreIndex As Long
    ''20170123-Get how many store name in MeasStoreName_Ary
    If Meas_StoreName <> "" Then
        SoreMaxNum = 0
        For i = 0 To UBound(MeasStoreName_Ary)
            If MeasStoreName_Ary(i) <> "" Then
                SoreMaxNum = SoreMaxNum + 1
            End If
        Next i
         ReDim Store_Rtn_Meas(SoreMaxNum - 1) As New PinListData
         StoreIndex = 0
     End If
    ''========================================================================================
    '20230714 sync RF feature
    #If RF = True Then
        glb_DSSCSetup = DSSCSetup
    #End If
                                                                                                                                                                                                                                                               
    For Each patt In patt_ary
        
        Pat = CStr(patt)
        
        TheHdw.Patterns(Pat).Load
        
        'If theexec.DataManager.InstanceName = "LPDPTX_1D_PP_CEBA0_S_FULP_AN_TX00_DCT_JTG_VMX_ALLFV_SI_DPTX_1D_NV" Then Stop
        
        Call GeneralDigSrcSettingWithBurst(Pat, DigSrc_pin, InDSPWave)

        If TPModeAsCharz_GLB = True Then
            If Rtn_SweepTestName <> "" Then
''                Rtn_SweepTestName = "_" & Rtn_SweepTestName
                For i = 0 To UBound(FlowTestNme)
                    FlowTestNme(i) = Replace(LCase(FlowTestNme(i)), "sweepcode", Rtn_SweepTestName)
                Next i
            End If
        End If
        
        Call GeneralDigCapSetting(Pat, DigCap_Pin, DigCap_Sample_Size, OutDspWave)
        
        Call PrintDigCapSetting(DigCap_Pin, DigCap_Sample_Size, CUS_Str_DigCapData)
        
        '' 20160713 - If no cpuflags in the test item modify the code to run pattern by using .test
        If (CPUA_Flag_In_Pat) Then
            Call TheHdw.Patterns(Pat).start
        Else
            Call TheHdw.Patterns(Pat).test(pfAlways, 0)
        End If
        
        TestSeqNum = 0
        Dim Force_idx As Integer
        Force_idx = 0
        For Each Ts In TestSequenceArray
'            If (UBound(TName_Ary) < TestSeqNum) Then
'                TName = ""
'            Else
'                TName = TName_Ary(TestSeqNum)
'            End If


            ''20150907 - Only need CPUA_Flag_In_Pat to do control
            If (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0) 'Meas during CPUA loop
            Else
                Call TheHdw.Digital.Patgen.HaltWait 'Pattern without CPUA loop
            End If
            
            TestOptLen = Len(Ts)
            
            ''20160923 - Add Interpose_PreMeas entry point by each sequence
            If Interpose_PreMeas <> "" Then
                If UBound(Interpose_PreMeas_Ary) = 0 Then
                    Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
                ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                    Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
                Else
                'Do nothing
                End If
            End If
                                                                                                                                                                                                                                                               
            For k = 1 To TestOptLen
                TestOption = mid(Ts, k, 1)
                For Each site In TheExec.sites.Active: testnum = TheExec.sites.item(site).TestNumber: Exit For: Next site
                TheHdw.Digital.Pins(Replace(DiffPins, "+", ",")).Disconnect
                
                With TheHdw.PPMU.Pins(TestPinArrayP(TestSeqNum))
                    .ForceV ForceV1pAry(Force_idx)
''                   .ClampVHi = 2
                    .ClampVHi = pc_Def_PPMU_ClampVHi
                    .Connect
                    .Gate = tlOn
                End With
                'Update PinName datalogging, From C-Chop, 20200423, Oscar
                TheExec.Datalog.WriteComment "PinName : " & CStr(TestPinArrayP(TestSeqNum)) & " , ForceV1p : " & CStr(Format(TheHdw.PPMU.Pins(TestPinArrayP(TestSeqNum)).Voltage.value, "0.000"))
               
                With TheHdw.PPMU.Pins(TestPinArrayN(TestSeqNum))
                    .ForceV ForceV1nAry(Force_idx)
''                   .ClampVHi = 2
                    .ClampVHi = pc_Def_PPMU_ClampVHi
                    .Connect
                    .Gate = tlOn
                End With
                'Update PinName datalogging, From C-Chop, 20200423, Oscar
                TheExec.Datalog.WriteComment "PinName : " & CStr(TestPinArrayN(TestSeqNum)) & " , ForceV1n : " & CStr(Format(TheHdw.PPMU.Pins(TestPinArrayN(TestSeqNum)).Voltage.value, "0.000"))
                 
                TheHdw.Wait 0.001
                DebugPrintFunc_PPMU CStr(TestPinArrayP(TestSeqNum))
''                MeasIp1 = TheHdw.PPMU.Pins(TestPinArrayP(k - 1)).Read(tlPPMUReadMeasurements, 10)
                MeasIp1 = TheHdw.PPMU.Pins(TestPinArrayP(TestSeqNum)).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
                DebugPrintFunc_PPMU CStr(TestPinArrayN(TestSeqNum))
''                MeasIn1 = TheHdw.PPMU.Pins(TestPinArrayN(k - 1)).Read(tlPPMUReadMeasurements, 10)
                MeasIn1 = TheHdw.PPMU.Pins(TestPinArrayN(TestSeqNum)).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
                 
                TheHdw.PPMU.Pins(TestPinArrayP(TestSeqNum)).ForceV ForceV2pAry(Force_idx)
                TheHdw.PPMU.Pins(TestPinArrayN(TestSeqNum)).ForceV ForceV2nAry(Force_idx)
                TheHdw.Wait 0.001
                'Update PinName datalogging, From C-Chop, 20200423, Oscar
                TheExec.Datalog.WriteComment "PinName : " & CStr(TestPinArrayP(TestSeqNum)) & " , ForceV2p : " & CStr(Format(TheHdw.PPMU.Pins(TestPinArrayP(TestSeqNum)).Voltage.value, "0.000"))
                TheExec.Datalog.WriteComment "PinName : " & CStr(TestPinArrayN(TestSeqNum)) & " , ForceV2n : " & CStr(Format(TheHdw.PPMU.Pins(TestPinArrayN(TestSeqNum)).Voltage.value, "0.000"))
                 
                DebugPrintFunc_PPMU CStr(TestPinArrayP(TestSeqNum))
''                MeasIp2 = TheHdw.PPMU.Pins(TestPinArrayP(k - 1)).Read(tlPPMUReadMeasurements, 10)
                MeasIp2 = TheHdw.PPMU.Pins(TestPinArrayP(TestSeqNum)).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
                DebugPrintFunc_PPMU CStr(TestPinArrayN(TestSeqNum))
''                MeasIn2 = TheHdw.PPMU.Pins(TestPinArrayN(k - 1)).Read(tlPPMUReadMeasurements, 10)
                MeasIn2 = TheHdw.PPMU.Pins(TestPinArrayN(TestSeqNum)).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)

                MeasVdiff = MeasIp1
                MeasVocm = MeasIp1
                
                'Dim RakCh_p() As Double
                'Dim RakCh_n() As Double
                Dim RAKVal_p As Double
                Dim RAKVal_n As Double
                 
                For Each site In TheExec.sites
                    For i = 0 To MeasIp1.Pins.Count - 1
                        'RakCh_p = TheHdw.PPMU.ReadRakValuesByPinnames(MeasIp1.Pins(i), site)
                        'RakCh_n = TheHdw.PPMU.ReadRakValuesByPinnames(MeasIn1.Pins(i), site)
                        
                        If (MeasIp1.Pins(i).value(site) - MeasIp2.Pins(i).value(site)) = 0 Then MeasIp1.Pins(i).value(site) = MeasIp1.Pins(i).value(site) + 0.000000001
                        If (MeasIn1.Pins(i).value(site) - MeasIn2.Pins(i).value(site)) = 0 Then MeasIn1.Pins(i).value(site) = MeasIn1.Pins(i).value(site) + 0.000000001
                       
                        Zp(site) = (CDbl(ForceV1pAry(Force_idx)) - CDbl(ForceV2pAry(Force_idx))) / (MeasIp1.Pins(i).value(site) - MeasIp2.Pins(i).value(site))
                        Zn(site) = (CDbl(ForceV1nAry(Force_idx)) - CDbl(ForceV2nAry(Force_idx))) / (MeasIn1.Pins(i).value(site) - MeasIn2.Pins(i).value(site))
                        TxVDD(site) = CDbl(ForceV2pAry(Force_idx)) + Zp(site) * MeasIp2.Pins(i).value(site) * (-1)
                        TxVss(site) = ForceV2nAry(Force_idx) - Zn(site) * MeasIn2.Pins(i).value(site)
                        
                        RAKVal_p = (CurrentJob_Card_RAK.Pins(MeasIp1.Pins(i)).value(site))
                        RAKVal_n = (CurrentJob_Card_RAK.Pins(MeasIn1.Pins(i)).value(site))
                        
                        Idiff(site) = (TxVDD(site) - TxVss(site)) / (Zp(site) - (RAKVal_p) + 100 + Zn(site) - (RAKVal_n))
                        Vdiff(site) = Abs((TxVDD(site) - (Zp(site) - RAKVal_p) * Idiff(site)) - (TxVss(site) + (Zn(site) - RAKVal_n) * Idiff(site)))
                        Vocm(site) = 0.5 * ((TxVDD(site) - (Zp(site) - RAKVal_p) * Idiff(site)) + (TxVss(site) + (Zn(site) - RAKVal_n) * Idiff(site)))
                        
                        MeasVdiff.Pins(i).value = Vdiff(site)
                        MeasVocm.Pins(i).value = Vocm(site)
 
                    Next i
                Next site
                   
                For Each site In TheExec.sites.Active: testnum = TheExec.sites.item(site).TestNumber: Exit For: Next site
                
                ''20160906 - Check store measurement or not
                If Meas_StoreName <> "" Then
                    If MeasStoreName_Ary(TestSeqNum) <> "" Then
                        Store_Rtn_Meas(StoreIndex) = MeasVdiff
                        Call StoreDataAllType(MeasStoreName_Ary(TestSeqNum), Store_Rtn_Meas(StoreIndex))
                        StoreIndex = StoreIndex + 1
                    End If
                End If
                

                
                '' Added TestLimitPerPin 20170814
                Dim p As Long

                For p = 0 To MeasVdiff.Pins.Count - 1
                    TestNameInput = Report_TName_From_Instance("Vdiff", MeasVdiff.Pins(p), , TestSeqNum, p)
                    TheExec.Flow.TestLimit MeasVdiff.Pins(p), , , scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=TestNameInput, ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                Next p
                If UCase(CUS_Str_MainProgram) = "VOCM" Then
                    For p = 0 To MeasVocm.Pins.Count - 1
                        TestNameInput = Report_TName_From_Instance("Vocm", MeasVocm.Pins(p), , TestSeqNum, p)
                        TheExec.Flow.TestLimit MeasVocm.Pins(p), , , scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=TestNameInput, ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                    Next p
                End If
                If TheExec.sites.Active.Count = 0 Then Exit Function

            Next k
            ''20161206-Restore force condiction after measurement

            If Interpose_PreMeas <> "" Then
                If UBound(Interpose_PreMeas_Ary) = 0 Then
                    Call SetForceCondition("RESTOREPREMEAS")
                ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                    Call SetForceCondition("RESTOREPREMEAS")
                Else
                'Do nothing
                End If
            End If
            
            TestSeqNum = TestSeqNum + 1
            
            If (CPUA_Flag_In_Pat) Then Call TheHdw.Digital.Patgen.Continue(0, cpuA)                'Jump out CPUA loop
            
            Force_idx = Force_idx + 1
        
        Next Ts
                                                                                                                                                                                                                                                               
        If DebugPrintEnable = True Then TheExec.Datalog.WriteComment "  Pattern(" & pat_count & "): " & Pat & vbNullString
        TheHdw.Digital.Patgen.HaltWait ' haltwait at patten end
        pat_count = pat_count + 1
    Next patt
    
    '' 20160923 - Add Interpose_PostTest entry point
    Call SetForceCondition(Interpose_PostTest)
    
    ''Comment by Martin for TTR
    TheHdw.PPMU.Pins(Replace(DiffPins, "+", ",")).ForceV pc_Def_PPMU_InitialValue_FI
    TheHdw.PPMU.Pins(Replace(DiffPins, "+", ",")).Gate = tlOff
    TheHdw.PPMU.Pins(Replace(DiffPins, "+", ",")).Disconnect
    TheHdw.Digital.Pins(Replace(DiffPins, "+", ",")).Connect
     
    '' 20160211 - Process DigCapData by using DSP
    If DigCap_Sample_Size <> 0 Then
        Dim DigCapPinAry() As String, NumberPins As Long
        Call TheExec.DataManager.DecomposePinList(DigCap_Pin, DigCapPinAry(), NumberPins)
        
        If NumberPins > 1 Then
            Call CreateSimulateDataDSPWave_Parallel(OutDspWave, DigCap_Sample_Size)
            Call DigCapDataProcessByDSP_Parallel(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, NumberPins, , DigCap_Pin.value)
        ElseIf NumberPins = 1 Then
            Call CreateSimulateDataDSPWave(OutDspWave, DigCap_Sample_Size, DigCap_DataWidth)
            Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDspWave)
            Call DigCapDataProcessByDSP(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, DigCap_DataWidth, , , DigCap_Pin.value)
        Else
        'Do nothing
        End If
    End If
    
    If (DisableConnectPins <> "") Then TheHdw.Digital.Pins(DisableConnectPins).Connect
    If DisableComparePins <> "" Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = False
                                                                                                                                                                                                                                                               
    '' 20160907 - Process calculate equation by dictionary.
    If Calc_Eqn <> "" Then
        Call ProcessCalcEquation(Calc_Eqn)
    End If
                                                                                                                                                                                                                                                               
    '' 20160713 - Call write functional result if cpu flag in pattern
    If (CPUA_Flag_In_Pat) Then
        Call HardIP_WriteFuncResult(, , glb_TestInstance)
    End If
                                                                                                                                                                                                                                                               
    DebugPrintFunc patset.value  ' print all debug information
    
    If Interpose_PrePat <> "" Then
        Call SetForceCondition("RESTOREPREPAT")
    End If
    'Alarm check, from Sicily ,20200423, Oscar
    ' Check implicit alarms
    TheHdw.Alarms.Check

    Exit Function
                                                                                                                                                                                                                                                               

errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP", "Meas_Vdiff_func") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
                                                        
Public Function Meas_VIR_IO_Universal_func(Optional patset As Pattern, Optional TestSequence As String, Optional CPUA_Flag_In_Pat As Boolean, _
Optional DisableComparePins As PinList, Optional DisableConnectPins As PinList, Optional DisableFRC As Boolean = False, Optional FRCPortName As String, _
Optional Measure_Pin_PPMU As String, Optional ForceV As String, Optional ForceI As String, Optional MeasureI_Range As String = "0.05", _
Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, Optional DigCap_DSPWaveSetting As CalculateMethodSetup = 0, _
Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As Long, Optional DigSrc_Sample_Size As Long, Optional DigSrc_Equation As String, Optional digsrc_assignment As String, Optional DigSrc_FlowForLoopIntegerName As String = vbNullString, _
Optional InstSpecialSetting As InstrumentSpecialSetup = 0, Optional SpecialCalcValSetting As CalculateMethodSetup = 0, Optional RAK_Flag As Enum_RAK = 0, _
Optional CUS_Str_MainProgram As String = vbNullString, Optional CUS_Str_DigCapData As String = vbNullString, Optional CUS_Str_DigSrcData As String, _
Optional Flag_SingleLimit As Boolean = False, Optional TestLimitPerPin_VIR As String = "FFF", _
Optional ForceFunctional_Flag As Boolean = False, _
Optional Meas_StoreName As String, Optional Calc_Eqn As String, _
Optional Interpose_PrePat As String, Optional Interpose_PreMeas As String, Optional Interpose_PostTest As String, Optional WaitTime_VIRZ As String, Optional Validating_ As Boolean) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

''Optional b_ProcessDigCapByDSP As Boolean = False, _
''==================================================================================
'' 20150621 - Check with CCWu: FRCPortName As String, Optional DisableFRC As Boolean = False not use in this function
'' 20150717 - Impedance measurement by using 2 point measure method, Define "Z" for TestSequence - On going
''                - EX: Pin1, Pin2 + Pin3, Pin4     V1, V2 + V3, V4
''                - V1 and V2 use for Pin1 of impedence measurement
''                - V1 and V2 use for Pin2 of impedence measurement
'' 20150717 - Get I from previous item and apply the current value to next item, use enum for the feature
''                - EX: TestSequence: "V,V,V"
''                  If second V want to apply calcuated I value that Force I value argument should be "0,keyword,0"
'' 20150727 - MeasureI_Range is use for test sequence "I", "R" and "Z"
''==================================================================================
    Dim site As Variant
    Dim PattArray() As String, PatCount As Long
    
    If Validating_ Then
        Call PrLoadPattern(patset.value)
        Exit Function    ' Exit after validation
    End If

    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)
        
    Call HardIP_InitialSetupForPatgen
    
    Dim i As Long, j As Long, k As Long
    Dim TestOptLen As Integer
    Dim TestSequenceArray() As String, ForceISequenceArray() As String, ForceVSequenceArray() As String
    Dim TestOption As Variant, Ts As Variant, TestSeqNum As Integer
    Dim TestPinArrayIV() As String, TestIrange() As String
    Dim TestSeqNumIdx As Long
    Dim InDSPWave As New DSPWave, OutDspWave As New DSPWave
    Dim ShowDec As String, ShowOut As String
    Dim Pat As String, patt As Variant
    Dim HighLimitVal() As Double, LowLimitVal() As Double
    Dim Rtn_MeasVolt As New PinListData, Rtn_MeasVolt_CUS_R As New PinListData, Rtn_MeasCurr As New PinListData
    Dim FlowForLoopName() As String   ' Sequences : Code , Voltage , Loop Index
    Dim MeasStoreName_Ary() As String
    Dim Interpose_PreMeas_Ary() As String
    'Dim Inst_Name_Str As String: Inst_Name_Str = glb_TestInstance  '20170728 Added for HardIP_WriteFuncResult Output
    Dim WaitTime_VIRZ_Ary() As String
    
''    Dim RTN_InterposeString As String
    Dim OutputTname() As String
    
    
    '----180524----------------------------------------------------------------
    Call GetFlowTName
    
    If ForceI Like "*@*" Then
    ForceI = Replace(ForceI, "@", vbNullString)
    End If
    '''''========================================================================================
    If WaitTime_VIRZ <> "" Then   ' update wait parameter for V,I,R,Z
        WaitTime_VIRZ_Ary = Split(WaitTime_VIRZ, ",")
        If UBound(WaitTime_VIRZ_Ary) = 0 Then
            ReDim Preserve WaitTime_VIRZ_Ary(3) As String
            WaitTime_VIRZ_Ary(1) = WaitTime_VIRZ_Ary(0)
            WaitTime_VIRZ_Ary(2) = WaitTime_VIRZ_Ary(0)
            WaitTime_VIRZ_Ary(3) = WaitTime_VIRZ_Ary(0)
        ElseIf UBound(WaitTime_VIRZ_Ary) < 3 Then
            ReDim Preserve WaitTime_VIRZ_Ary(3) As String
        Else
        'Do nothing
        End If
    Else
        ReDim WaitTime_VIRZ_Ary(3) As String
    End If
    '''''========================================================================================
    If Measure_Pin_PPMU Like "*@*" Then Measure_Pin_PPMU = Replace(Measure_Pin_PPMU, "@", vbNullString)
    Shmoo_Pattern = patset.value

    Call tl_PinListDataSort(True)
    
    If (InStr(MeasureI_Range, ":") <> 0) Then MeasureI_Range = Select_MeasIRange(MeasureI_Range, CurrentJobName_U)  ' support different Meter_Range in different Job, add by Roger 20170628
    
    Call VIR_AnalyzedInputStringByAt(Measure_Pin_PPMU, ForceV, ForceI, MeasureI_Range)
   
    If ForceI = "" Then ForceI = 0
    If ForceV = "" Then ForceV = 0
    If MeasureI_Range = "" Then MeasureI_Range = pc_Def_VIR_MeasCurrRange
    
    Call VIR_CheckForceVal(ForceI, ForceV)

    Call VIR_ProcessInputString(TestSequence, ForceI, ForceV, Measure_Pin_PPMU, MeasureI_Range, Meas_StoreName, Interpose_PreMeas, _
                                              TestSequenceArray(), ForceISequenceArray(), ForceVSequenceArray(), TestPinArrayIV(), _
                                              TestIrange(), MeasStoreName_Ary(), Interpose_PreMeas_Ary())
 
'    Call HIP_Evaluate_ForceVal_New(ForceVSequenceArray())
'    Call HIP_Evaluate_ForceVal_New(ForceISequenceArray())
    ''========================================================================================
    Dim Store_Rtn_Meas() As New PinListData
    Dim SoreMaxNum As Long
    Dim StoreIndex As Long
    ''20170123-Get how many store name in MeasStoreName_Ary
    If Meas_StoreName <> "" Then
        SoreMaxNum = 0
        For i = 0 To UBound(MeasStoreName_Ary)
            If MeasStoreName_Ary(i) <> "" Then
                SoreMaxNum = SoreMaxNum + 1
            End If
        Next i
         ReDim Store_Rtn_Meas(SoreMaxNum - 1) As New PinListData
         StoreIndex = 0
     End If
    ''========================================================================================
     ''20170807 - CZ test name index
    gl_CZ_FlowTestNameIndex = 0
    ''========================================================================================
    
    Call GetFlowSingleUseLimit(HighLimitVal, LowLimitVal)
    
    ''20161130-Get test name from flow table
    Dim FlowTestNme() As String
    If TPModeAsCharz_GLB Then
        gl_CZ_FlowTestName_Counter = 0
        Call GetFlowTestName(FlowTestNme)
    End If
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    '' 20160923 - Add Interpose_PrePat entry point
    If Interpose_PrePat <> "" Then
        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    End If
    

    ''20161205 - Force_Flow_Shmoo_Condition
    If TheExec.sites.item(site).SiteVariableValue("Flow_Shmoo_DevCharSetup") <> "" Then Force_Flow_Shmoo_Condition
    'Do Flow Shmoo
    
    If patset.value <> "" Then
         gl_TName_Pat = patset.value
        Shmoo_Pattern = patset.value '' 20170808 add for shmoo pattern name print
        TheHdw.Patterns(patset).Load
        Call PATT_GetPatListFromPatternSet(patset.value, PattArray, PatCount)
    Else
        ReDim PattArray(0)
        PattArray(0) = vbNullString
    End If
    
    If (DisableConnectPins <> "") Then TheHdw.Digital.Pins(DisableConnectPins).Disconnect
    If (DisableComparePins <> "") Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = True

    Dim Rtn_SweepTestName As String
    Rtn_SweepTestName = vbNullString
    
    For Each patt In PattArray
        If patt <> "" Then

        Pat = CStr(patt)
        
        Call GeneralDigSrcSetting(Pat, DigSrc_pin, DigSrc_Sample_Size, DigSrc_DataWidth, DigSrc_Equation, digsrc_assignment, _
                                               DigSrc_FlowForLoopIntegerName, CUS_Str_DigSrcData, InDSPWave, Rtn_SweepTestName)

        If TPModeAsCharz_GLB = True Then
            If Rtn_SweepTestName <> "" Then
''                Rtn_SweepTestName = "_" & Rtn_SweepTestName
                For i = 0 To UBound(FlowTestNme)
                    FlowTestNme(i) = Replace(LCase(FlowTestNme(i)), "sweepcode", Rtn_SweepTestName)
                Next i
            Else
                Call SimulateFlowForSweep(FlowShmooString_GLB)
                If FlowShmooString_GLB <> "" Then
                    For i = 0 To UBound(FlowTestNme)
                        FlowTestNme(i) = Replace(LCase(FlowTestNme(i)), "sweepvoltage", FlowShmooString_GLB)
                    Next i
                End If
            End If
        End If

        Set OutDspWave = Nothing
        Call GeneralDigCapSetting(Pat, DigCap_Pin, DigCap_Sample_Size, OutDspWave)
        
        Call PrintDigCapSetting(DigCap_Pin, DigCap_Sample_Size, CUS_Str_DigCapData)
        
        ''20160306-Sweep Volt
        If UCase(DigSrc_FlowForLoopIntegerName) = "SWEEP_V" Then
            Call Cust_Sweep_V
        End If
        
        '' 20160713 - If no cpuflags in the test item modify the code to run pattern by using .test
        If (CPUA_Flag_In_Pat) Then
            Call TheHdw.Patterns(Pat).start
        Else
            Call TheHdw.Patterns(Pat).test(pfAlways, 0)
        End If
        End If
        
        TestSeqNum = 0
        
        Call ProcessTestNameInputString(OutputTname, UBound(TestSequenceArray))
        For Each Ts In TestSequenceArray
            
            ''20150907 - Only need CPUA_Flag_In_Pat to do control
            If (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0) 'Meas during CPUA loop
            Else
                Call TheHdw.Digital.Patgen.HaltWait 'Pattern without CPUA loop
            End If
            
            ''20160923 - Add Interpose_PreMeas entry point by each sequence
            If Interpose_PreMeas <> "" Then
                If UBound(Interpose_PreMeas_Ary) = 0 Then
                    Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
                ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                    Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
                Else
                'Do nothing
                End If
            End If
            
            TestOptLen = Len(Ts)
            
            TestSeqNumIdx = TestSeqNum
            
            For k = 1 To TestOptLen
                
                TestOption = mid(Ts, k, 1)
                
                '' 20160106 - If "ForceFunctional_Flag" = True to let TestOption = "N" to make the test instance only run functional test
                If ForceFunctional_Flag = True Then
                    TestOption = "N"
                End If
                
                '' 20160705 - If second case is N that will cause error
                If (Measure_Pin_PPMU <> "") Then
                    Call Meas_VIR_IO_PreSetupBeforeMeasurement(TestPinArrayIV, TestSeqNumIdx)
                    
                    Select Case UCase(TestOption)
                    
                        Case "V"
                        
                            Call IO_HardIP_PPMU_Measure_V(TestPinArrayIV, TestSeqNum, TestSeqNumIdx, ForceISequenceArray, _
                                    k, Pat, Flag_SingleLimit, HighLimitVal(0), LowLimitVal(0), TestLimitPerPin_VIR, Rtn_MeasVolt, FlowTestNme, _
                                    SpecialCalcValSetting, InstSpecialSetting, RAK_Flag, CUS_Str_MainProgram, Rtn_SweepTestName, OutputTname(TestSeqNum), WaitTime_VIRZ_Ary(0))
 
                             ''20160906 - Check store measurement or not
                            If Meas_StoreName <> "" Then
                                If MeasStoreName_Ary(TestSeqNum) <> "" Then
                                    Store_Rtn_Meas(StoreIndex) = Rtn_MeasVolt
                                    Call StoreDataAllType(MeasStoreName_Ary(TestSeqNum), Store_Rtn_Meas(StoreIndex))
                                    StoreIndex = StoreIndex + 1
                                End If
                            End If
 
''                             ''20151028  CUS_MeasV_And_CalR -- TYCHENGG
''                            ''========================================================================================
''                            If (UCase(CUS_Str_MainProgram) Like "*CALR*") Then
''                                Call CUS_VIR_MainProgram_MeasV_CalR(TestPinArrayIV, TestSeqNum, CUS_CalR_Seq_Ary, ForceISequenceArray, Rtn_MeasVolt_CUS_R, CUS_CalR_VDD)
''                            End If
''                            ''========================================================================================
                            
                        Case "I"
                            
                            If DisableFRC = True Then FreeRunClk_Disable (FRCPortName)
                            
                            Call IO_HardIP_PPMU_Measure_I(TestPinArrayIV, TestSeqNum, TestSeqNumIdx, ForceVSequenceArray, _
                                    k, Pat, Flag_SingleLimit, HighLimitVal(0), LowLimitVal(0), TestLimitPerPin_VIR, TestIrange, FlowTestNme, CUS_Str_MainProgram, SpecialCalcValSetting, Rtn_MeasCurr, Rtn_SweepTestName, InstSpecialSetting, OutputTname(TestSeqNum), WaitTime_VIRZ_Ary(1))
                            
                            ''20160906 - Check store measurement or not
                            If Meas_StoreName <> "" Then
                                If MeasStoreName_Ary(TestSeqNum) <> "" Then
                                    Store_Rtn_Meas(StoreIndex) = Rtn_MeasCurr
                                    Call StoreDataAllType(MeasStoreName_Ary(TestSeqNum), Store_Rtn_Meas(StoreIndex))
                                    StoreIndex = StoreIndex + 1
                                End If
                            End If
                            
                        Case "R"
                            
                            Call IO_HardIP_PPMU_Measure_R(TestPinArrayIV, TestSeqNum, TestSeqNumIdx, ForceVSequenceArray, _
                                    k, Pat, Flag_SingleLimit, HighLimitVal(0), LowLimitVal(0), TestLimitPerPin_VIR, TestIrange, FlowTestNme, RAK_Flag, Rtn_SweepTestName, CUS_Str_MainProgram, OutputTname(TestSeqNum), WaitTime_VIRZ_Ary(2), SpecialCalcValSetting)
                        
                        Case "Z"
                            
                            If (Len(Ts) <> 1) Then ForceVSequenceArray(TestSeqNum) = ForceVSequenceArray(TestSeqNum) & ";sweep"

                            Call IO_HardIP_PPMU_Measure_Z(TestPinArrayIV, TestSeqNum, TestSeqNumIdx, ForceVSequenceArray, _
                                    k, Pat, Flag_SingleLimit, HighLimitVal(0), LowLimitVal(0), TestLimitPerPin_VIR, TestIrange, FlowTestNme, RAK_Flag, Rtn_SweepTestName, OutputTname(TestSeqNum), WaitTime_VIRZ_Ary(3))
                                    
                        Case "N"
                        
                        Case Else
                            TheExec.Datalog.WriteComment "Error Test Option, please select V, I or R"
                    
                    End Select
                    
                    Call Meas_VIR_IO_PostSetupAfterMeasurement(TestPinArrayIV, TestSeqNumIdx)
                End If
            Next k
            
            ''20161206-Restore force condiction after measurement

            If Interpose_PreMeas <> "" Then
                If UBound(Interpose_PreMeas_Ary) = 0 Then
                    Call SetForceCondition("RESTOREPREMEAS")
                ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                    Call SetForceCondition("RESTOREPREMEAS")
                Else
                'Do nothing
                End If
            End If
            
            TestSeqNum = TestSeqNum + 1
            
            If (CPUA_Flag_In_Pat) Then Call TheHdw.Digital.Patgen.Continue(0, cpuA)
            
        Next Ts
        
        If DebugPrintEnable = True Then TheExec.Datalog.WriteComment "  Pattern(" & PatCount & "): " & Pat & vbNullString
                
        TheHdw.Digital.Patgen.HaltWait ' haltwait at patten end
        
        PatCount = PatCount + 1
        
        '' 20160923 - Add Interpose_PostTest entry point
        Call SetForceCondition(Interpose_PostTest)
    
'        If gl_FlowForLoop_DigSrc_SweepCode <> "" Then   '20180509 TER add
'            gl_FlowForLoop_DigSrc_SweepCode = ""
'        End If
    
        '' 20160211 - Process DigCapData by using DSP
''        If b_ProcessDigCapByDSP = True Then
            If DigCap_Sample_Size <> 0 Then
                Dim DigCapPinAry() As String, NumberPins As Long
                Call TheExec.DataManager.DecomposePinList(DigCap_Pin, DigCapPinAry(), NumberPins)
                
                If NumberPins > 1 Then
                    Call CreateSimulateDataDSPWave_Parallel(OutDspWave, DigCap_Sample_Size)
                    Call DigCapDataProcessByDSP_Parallel(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, NumberPins, , DigCap_Pin.value)
                ElseIf NumberPins = 1 Then
                    Call CreateSimulateDataDSPWave(OutDspWave, DigCap_Sample_Size, DigCap_DataWidth)
                    Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDspWave)
                    Call DigCapDataProcessByDSP(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, DigCap_DataWidth, , , DigCap_Pin.value)
                Else
                'Do nothing
                End If
            End If
            
        If gl_FlowForLoop_DigSrc_SweepCode <> "" Then   '20180814 TER add
            gl_FlowForLoop_DigSrc_SweepCode = vbNullString
            gl_FlowForLoop_DigSrc_SweepCode_Dec = vbNullString '20190613 CT add for Decimal value printing
        End If

''        End If
    Next patt
    
    
    If (DisableConnectPins <> "") Then TheHdw.Digital.Pins(DisableConnectPins).Connect
    If DisableComparePins <> "" Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = False
    
    If DisableFRC = True Then
        Call ReStart_FRC(FRCPortName)
    End If
    '' 20160907 - Process calculate equation by dictionary.
    If Calc_Eqn <> "" Then
        Call ProcessCalcEquation(Calc_Eqn)
    End If

    '' 20160713 - Call write functional result if cpu flag in pattern
    If (CPUA_Flag_In_Pat) Then
        Call HardIP_WriteFuncResult(, , glb_TestInstance)
    End If
    
    DebugPrintFunc patset.value  ' print all debug information
    
    If Interpose_PrePat <> "" Then
        Call SetForceCondition("RESTOREPREPAT")
    End If
    
Exit Function


errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP", "Meas_VIR_IO_Universal_func") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function



'[20230721][All][Neil] Move Interpose_PostTest to out of Pattern loop level
'[20231107][All][Neil] Duty Cycle Test Methodology on Digital and DCVI instrument
Public Function Meas_FreqVoltCurr_Universal_func(Optional patset As Pattern, Optional TestSequence As String, Optional CPUA_Flag_In_Pat As Boolean, _
Optional DisableComparePins As PinList, Optional DisableConnectPins As PinList, Optional DisableFRC As Boolean = False, Optional FRCPortName As String, _
Optional MeasV_PinS As String, _
Optional MeasF_PinS_SingleEnd As String, Optional MeasF_Interval As String, Optional MeasF_EventSourceWithTerminationMode As EventSourceWithTerminationMode, Optional MeasF_Flag_MeasureThreshold As Boolean = False, Optional MeasF_ThresholdPercentage As Double = 0.5, Optional MeasF_WaitTime As String, _
Optional MeasI_pinS As String, Optional MeasI_Range As String, Optional MeasI_WaitTime As String, _
Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, _
Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As String, Optional DigSrc_Sample_Size As String, Optional DigSrc_Equation As String, Optional digsrc_assignment As String, Optional DigSrc_FlowForLoopIntegerName As String = vbNullString, _
Optional SpecialCalcValSetting As CalculateMethodSetup = 0, _
Optional InstSpecialSetting As InstrumentSpecialSetup = 0, _
Optional CUS_Str_MainProgram As String = vbNullString, Optional CUS_Str_DigCapData As String = vbNullString, Optional CUS_Str_DigSrcData As String = vbNullString, _
Optional Flag_SingleLimit As Boolean = False, Optional Spare_Argument As String, _
Optional MeasF_PinS_Differential As String, Optional ForceFunctional_Flag As Boolean = False, _
Optional MeasF_WalkingStrobe_Flag As Boolean, Optional MeasF_WalkingStrobe_StartV As Double, Optional MeasF_WalkingStrobe_EndV As Double, Optional MeasF_WalkingStrobe_StepVoltage As Double, Optional MeasF_WalkingStrobe_BothVohVolDiffV As Double, Optional MeasF_WalkingStrobe_interval As Double, Optional MeasF_WalkingStrobe_miniFreq As Double, _
Optional Meas_StoreName As String, Optional Calc_Eqn As String, _
Optional Interpose_PrePat As String, Optional Interpose_PreMeas As String, Optional Interpose_PostMeas As String, Optional Interpose_PostTest As String, Optional CharSetName As String, _
Optional ForceV_Val As String, Optional ForceI_Val As String, Optional UVI80_MeasV_WaitTime As String = vbNullString, _
Optional RAK_Flag As Enum_RAK, Optional WaitTime_VIRZ As String, Optional MSB_First_Flag As Boolean = False, Optional BV_Enable As Boolean, Optional DSSCSetup As String, Optional Validating_ As Boolean) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
 
    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)
    
        'TTR,20200423, Oscar
    'Dim Inst_Name_Str As String: Inst_Name_Str = glb_TestInstance  '20170728 Added for HardIP_WriteFuncResult Output

        
''        Dim TestTime_D As Double
    m_InstanceName = LCase(glb_TestInstance)
''    TestTimeCollect_Start TestTime_D, m_InstanceName, Validating_
''    ProfileCollect_Start m_InstanceName, Validating_
    
    '=======Turks add to bypass HIP_Universal shmoo post body=======
    Dim DevChar_Setup As String
    
    If TheExec.DevChar.Setups.IsRunning = True Then
        DevChar_Setup = TheExec.DevChar.Setups.ActiveSetupName
        If TheExec.DevChar.Results(DevChar_Setup).StartTime Like "1/1/0001*" Or TheExec.DevChar.Results(DevChar_Setup).StartTime Like "0001/1/1*" Then ' initial run of shmoo, not the first point
            Shmoo_End = False
        End If
    End If
    If gl_EnableCurrentProfile Or gl_EnableVoltageProfile Then
        ProfileCollect_Start glb_TestInstance
    End If
    '====================================ADC===========================
'        Analog Input, From Sicily, 20200423, Oscar
'    If TheExec.DataManager.instanceName = "ADC_PEAKMAXT7_PP_SCYA0_C_FULP_AN_AQ00_DLL_JTG_IMX_ALLFV_SI_PEAKMAX_T7_NV" Or TheExec.DataManager.instanceName = "ADC_PEAKMINT8_PP_SCYA0_C_FULP_AN_AQ00_DLL_JTG_IMX_ALLFV_SI_PEAKMIN_T8_NV" Then
'                thehdw.DCVI.Pins("ANALOGMUX_OUT").Source.Signals.Add "MySquare"
'                thehdw.DCVI.Pins("ANALOGMUX_OUT").Source.Signals("MySquare").WaveDefinitionName = "square"
'                With thehdw.DCVI.Pins("ANALOGMUX_OUT").Source.Signals.Item("MySquare")
'                    .mode = tlDCVIModeVoltage
'                    .mode.mode = tlSignalModeUseValue
'                    'sine
'                    .Amplitude = 0.075
'                    .Offset = 0.5
'                    .ApplyMixedSignalTiming ("Dcvi1")
'                    .LoadSettings
'                End With
'                thehdw.DCVI.Pins("ANALOGMUX_OUT").Connect tlDCVIConnectDefault
'                thehdw.DCVI.Pins("ANALOGMUX_OUT").Gate = True
'    End If
    
    '====================================ADC===========================
    
    If TheExec.DevChar.Setups.IsRunning = True And Shmoo_End = True Then
        Exit Function
    End If
    Dim PatCount As Long
    Dim PattArray() As String
    Dim Loopendnumber() As String
    
    If Validating_ Then
        Call PrLoadPattern(patset.value)
        Exit Function    ' Exit after validation
    End If
    
    If InStr(CUS_Str_MainProgram, "vt_sweep") > 0 And AMP_EYE_VT_CZ_Flag = True Then
        Call SWEEP_VT(CUS_Str_MainProgram, Interpose_PrePat)
    Else
        gl_Sweep_vt = vbNullString
    End If
    
    If InStr(UCase(CUS_Str_MainProgram), UCase("V_Sweep")) > 0 Then
        Call SWEEP_V(CUS_Str_MainProgram, Interpose_PrePat)
    Else
        gl_Sweep_vt = vbNullString
    End If

    
    Call HardIP_InitialSetupForPatgen
    If MSB_First_Flag = True Then
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Note: DSSC MSB First. Please check pattern content"
    End If
    
    
    Call ShmooEndFunction  ' For Shmoo DIgsrc
    
''    m_InstanceName = LCase(theexec.DataManager.instancename)
    
    
    Dim i As Long, j As Long, k As Long
    Dim TestOptLen As Integer
    Dim TestSequenceArray() As String, MeasPinAry_V() As String, MeasPinAry_F() As String, MeasPinAry_I() As String, MeasPinAry_IRange() As String
    Dim MeasPinAry_F_Differential() As String
    Dim MeasureF_Pin_Differential As New PinList
    Dim Ts As Variant, TestOption As Variant, site As Variant
    Dim TestSeqNum As Integer
    Dim MeasureV_pin As New PinList, MeasureF_Pin_SingleEnd As New PinList, MeasureI_pin As New PinList
    Dim MeasureI_Pin_CurrentRange As String
    Dim testnum As Long
    Dim InDSPWave As New DSPWave, OutDspWave As New DSPWave
    Dim ShowDec As String, ShowOut As String
    Dim patt As Variant
    Dim Pat As String
    Dim HighLimitVal() As Double, LowLimitVal() As Double
    Dim MeasureV_Pin_PPMU As String, MeasureV_Pin_UVI80 As String
    Dim d_MeasF_Interval As Double
    Dim FreqPinsCheckType() As String
    Dim ThisPinType As String
    Dim MeasF_EventSource As FreqCtrEventSrcSel
    Dim MeasF_EnableVtMode As Boolean
    Dim Split_F_Str() As String
    Dim restore_Flag As Boolean
    
    
    ''20160906 - Return measurement to directionary if needed
    Dim Rtn_MeasVolt As New PinListData, Rtn_MeasCurr As New PinListData, Rtn_MeasFreq As New PinListData
    Dim MeasStoreName_Ary() As String
    Dim Interpose_PreMeas_Ary() As String
    Dim Interpose_PostMeas_Ary() As String
''    Dim RTN_InterposeString As String
    
    Dim CheckDSPWave As New DSPWave
    Dim Sweep_Enable As Boolean: Sweep_Enable = False
    Dim Sweep_Loop_Calc_Eqn As String: Sweep_Loop_Calc_Eqn = vbNullString
'    Dim Sweep_Calc_Eqn As Boolean: Sweep_Calc_Eqn = False
    Dim Sweep_Calc_Eqn_index As String: Sweep_Calc_Eqn_index = vbNullString
    Dim Sweep_Dictionary As String: Sweep_Dictionary = vbNullString
    Dim Sweep_Calc_Eqn As String: Sweep_Calc_Eqn = vbNullString
    
    Dim OutputTname() As String

    Call tl_PinListDataSort(True)

    Processing_MeasPins MeasV_PinS, MeasI_pinS, MeasF_PinS_SingleEnd            'Jade Edited by 20200116

'**************************************************
'SeaHawk Edited by 20190606
    Dim SpecialUsePatName As String
    SpecialUsePatName = CStr(patset)
    If CUS_Str_DigCapData Like "*Special_DigCapData_Setting*" Then
        gl_SpecialString = vbNullString
        SpecialUsePatName = SpecialUsePatName & "_SpecialDigCap"
        Public_GetStoredString (SpecialUsePatName)
        CUS_Str_DigCapData = gl_SpecialString
    End If
    
    If CUS_Str_DigSrcData Like "*Special_DigSrcData_Setting*" Then
        gl_SpecialString = vbNullString
        SpecialUsePatName = SpecialUsePatName & "_SpecialDigSrc"
        Public_GetStoredString (SpecialUsePatName)
        CUS_Str_DigSrcData = gl_SpecialString
    End If
'**************************************************
    
       If gl_Flag_HardIP_Characterization_1stRun = False Then 'Then: Exit Function
        If TheExec.DevChar.Setups.IsRunning = True And CStr(TheExec.DevChar.Setups.ActiveSetupName) Like "*SweepDigSrc*" Then
            Call ReDefineDigSrcForCharacterization(digsrc_assignment)
        End If
      End If
    
    '================================================================ Roger
    If InStr(1, LCase(Interpose_PrePat), "sweep:") <> 0 Then
        Dim Sweep_Info() As Power_Sweep
        Dim Sweep_CUS_Str_DigCapData As String
        Call SortSweepInfo(Sweep_Info, Interpose_PrePat)
        Sweep_Enable = True
        Sweep_CUS_Str_DigCapData = CUS_Str_DigCapData
        Sweep_Calc_Eqn = Calc_Eqn
    End If
    '================================================================
    ''20170322-Store MeasF mid value for VT
    Dim SplitFreqVtValue() As String
    Dim DictKey_StoreVT As String
    Dim Dict_VT_Value As New SiteDouble
    
    
    'For Register Assign sheet, adding Meas_StoreName with Offset, From JadeCdie, 20201023, Oscar
    Call Reg_Assign_Processing(DigSrc_Equation, digsrc_assignment, CUS_Str_DigCapData, Calc_Eqn, CUS_Str_DigSrcData, MeasV_PinS, MeasI_pinS, MeasF_PinS_SingleEnd, Interpose_PreMeas, Meas_StoreName, CUS_Str_MainProgram, DigSrc_FlowForLoopIntegerName)
    
    'If (UCase(MeasI_Range) Like "*CP*:*" Or UCase(MeasI_Range) Like "*FT*:*") Then MeasI_Range = Select_MeasIRange(MeasI_Range, CurrentJobName_U)   ' support different Meter_Range in different Job, add by Roger 20170628
    
    '' 20160201 - Check input argumenets whether have "@" in the first character. Add it If no "@" in the beginning. Then remove it to process fomat.
    'TTR,20200423, Oscar
    'Call VFI_AnalyzedInputStringByAt(MeasV_PinS, MeasF_PinS_SingleEnd, MeasI_pinS, MeasI_Range, MeasF_PinS_Differential, ForceV_Val, ForceI_Val)
    
    Dim ForceV_Val_Ary() As String
    Dim ForceI_Val_Ary() As String
    Dim MeasurePin_ForceV_Val As String
    Dim MeasurePin_ForceI_Val As String
    Dim MeasI_WaitTime_Ary() As String
    Dim MeasF_WaitTime_Ary() As String
    Dim UVI80_MeasV_WaitTime_Ary() As String
    
    
    

    '----------------------------20180523
    
    'Roger New,20180510 TName
    '--------------------------------------------------------------------
    'Call GetFlowTName
    
    '----------------------------20180523
        
    'Call VFI_ProcessInputString(TestSequence, MeasV_PinS, MeasI_pinS, MeasF_PinS_SingleEnd, MeasF_PinS_Differential, MeasI_Range, Meas_StoreName, Interpose_PreMeas, _
                                            ForceV_Val, ForceI_Val, _
                                            TestSequenceArray(), MeasPinAry_V(), MeasPinAry_I(), MeasPinAry_F(), _
                                            MeasPinAry_F_Differential(), MeasPinAry_IRange(), MeasStoreName_Ary(), Interpose_PreMeas_Ary(), ForceV_Val_Ary(), ForceI_Val_Ary())

    'Call VFI_ProcessWaitTimeString(MeasI_WaitTime, MeasF_WaitTime, UVI80_MeasV_WaitTime, MeasI_WaitTime_Ary(), MeasF_WaitTime_Ary(), UVI80_MeasV_WaitTime_Ary(), TestSequenceArray())
                                            
    
'    Call HIP_Evaluate_ForceVal(ForceV_Val_Ary())
'
'    Call HIP_Evaluate_ForceVal(ForceI_Val_Ary())
    
    ''20170807 - CZ test name index
'    gl_CZ_FlowTestNameIndex = 0

'    Call Freq_ProcessEventSourceTerminationMode(MeasF_EventSourceWithTerminationMode, MeasF_EventSource, MeasF_EnableVtMode)
    
    ''20141219 Get use-limit from flow table
'    Call GetFlowSingleUseLimit(HighLimitVal, LowLimitVal)
    
    ''20161130-Get test name from flow table
    Dim FlowTestNme() As String
    ''========================================================================================
''    Dim Store_Rtn_Meas() As New PinListData
''    Dim SoreMaxNum As Long
''    Dim StoreIndex As Long
''    ''20170123-Get how many store name in MeasStoreName_Ary
''    If Meas_StoreName <> "" Then
''        SoreMaxNum = 0
''        For i = 0 To UBound(MeasStoreName_Ary)
''            If MeasStoreName_Ary(i) <> "" Then
''               SoreMaxNum = SoreMaxNum + 1
''            End If
''        Next i
''         ReDim Store_Rtn_Meas(SoreMaxNum - 1) As New PinListData
''         StoreIndex = 0
''     End If
    ''========================================================================================
    If TheExec.DevChar.Setups.IsRunning Then
        If CharSetName <> "" And InStr(UCase(Interpose_PrePat), ":TERM:") <> 0 Then
        'HIO:can not applylevelsTiming for the first point  of run_shmoo
        ElseIf TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.axes.Contains(tlDevCharShmooAxis_Y) Then
            If gl_Flag_HardIP_Trim_Set_PrePoint And Not (gl_Flag_HardIP_Characterization_1stRun) Then
                TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
                Call TheExec.Overlays.ApplyUniformSpecToHW(XI0_Shmoo & "_Freq_VAR", TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.axes(tlDevCharShmooAxis_Y).value)
            ElseIf gl_Flag_HardIP_Trim_Set_PostPoint Then
                TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
                Call TheExec.Overlays.ApplyUniformSpecToHW(XI0_Shmoo & "_Freq_VAR", 24000000#)
            Else
                TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
            End If
        Else
            TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
        End If
    ElseIf BV_Enable Then
    Else
        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    End If
	
    If gl_Disable_HIP_debug_log = False Then
    	DatalogContextInformation
    End If	

    Dim Loop_Idx As Long
    Dim Loop_count As Long
    Dim Loop_Init As Long
    Dim Loop_Max As Long
    Dim Loop_Step As Long
    Dim Loop_BitNum As Long
    Dim Loop_RegName As String
    Dim SplitLoop_RegName() As String
    Dim Split_Loop_DigSrc_Str() As String
    Dim binstr As String
    Dim Loop_SplitByComma() As String
    Dim Loop_SplitByEqual() As String
    Dim Loop_Digsrc_name As String
    Dim Loop_MeasV_PinS_name As String
    Dim Split_DigSrc_Equation() As String
    Dim Split_MeasV_PinS_By_Comma() As String
    Dim Split_MeasV_PinS_By_Colon() As String
    Dim MeasV_PinS_Temp As String
    Dim MeasV_PinS_Temp1 As String
    
    Loop_Idx = 0
    Loop_Init = 0
    Loop_Max = 0
    Loop_Step = 1
    
    If (Sweep_Enable = True) Then
        Loop_Max = Sweep_Info(0).Count - 1
    
    End If
    Dim timer_ As Double
    Dim Loop_count2 As Long: Loop_count2 = 0
    Dim Loop_Max2 As Long: Loop_Max2 = 0
    Dim Loop_Init2 As Long: Loop_Init2 = 0
    Dim sweepVoltage_Loop_Max As Long: sweepVoltage_Loop_Max = 0 'For Store sweep voltage max count
    Dim sweepCode_Loop_Max As Long: sweepCode_Loop_Max = 0       'For Store sweep code max count
    Dim sweepCode_temp_StrAry() As String                        'For Parsing SweepCode set information
    Dim sweepVoltage_ary() As String
    Dim Loop_Priority As Long                                    '0: Sweep_Code_First / 1:Sweep_voltage_First
    Dim sweepCode_list As Long
    gl_sweepVoltage_Loop_Indx = 0                                'Use for record SweepVoltage loop index[Inner_Loop]
    gl_sweepCode_Loop_Indx = 0                                   'Use for record SweepCode loop index[Inner_Loop]    
    'timer_ = theexec.Timer()
'                                                                                                                                                                Dim VBT_LIB_HardIP_ProfileMark_1164 As Long: VBT_LIB_HardIP_ProfileMark_1164 = ProfileMarkEnter(2, instance_name & "_" & "ProcessInputToGLB&Module=VBT_LIB_HardIP&ProcName=Meas_FreqVoltCurr_Universal_func&LineNumber=1158")   ' Profile Mark
    
''If TheExec.DataManager.instanceName Like "*D2DEXLBK*" Then   'CSHOX for D2D ELB too long
If DigSrc_Equation Like "*Duplicate*" Then    '' Duplicated long string case, From C-Chop
'' 190814 re-edit set a format to loop the DigSrc assignment : Duplicate:[512](Loop count):[D2D_PHY__ZCPU_ZCPD__zcpu+D2D_PHY__ZCPU_ZCPD__zcpd](LoopName)
    Split_DigSrc_Equation = Split(DigSrc_Equation, ":")
    Loop_Digsrc_name = Split_DigSrc_Equation(2) ''LoopName
    For i = 0 To CInt(Split_DigSrc_Equation(1)) - 1   ''LoopCount
        If i = 0 Then
            DigSrc_Equation = Loop_Digsrc_name
        Else
            DigSrc_Equation = DigSrc_Equation & "+" & Loop_Digsrc_name
        End If
    Next i
End If
   
Dim loop_i_MeasV As Long
'' Barry edit for ADCLK_CZ T10 MeasV_PinS string too long
'' 200310 re-edit set a format to loop the MeasV_PinS
'' EX:Duplicate:[64](Loop count):[+ANALOGMUX_OUT0+ANALOGMUX_OUT0](LoopName);Duplicate:[128](Loop count):[+ANALOGMUX_OUT1+ANALOGMUX_OUT1](LoopName)
If MeasV_PinS <> "" And InStr(UCase(MeasV_PinS), UCase("Duplicate")) <> 0 Then

    Split_MeasV_PinS_By_Comma = Split(MeasV_PinS, ";")
    For loop_i_MeasV = 0 To UBound(Split_MeasV_PinS_By_Comma)
        Split_MeasV_PinS_By_Colon = Split(Split_MeasV_PinS_By_Comma(loop_i_MeasV), ":")
        Loop_MeasV_PinS_name = Split_MeasV_PinS_By_Colon(2) ''LoopName
        For i = 0 To CInt(Split_MeasV_PinS_By_Colon(1)) - 1   ''LoopCount
            If i = 0 Then
                MeasV_PinS_Temp = Loop_MeasV_PinS_name
            Else
                MeasV_PinS_Temp = MeasV_PinS_Temp & "+" & Loop_MeasV_PinS_name
            End If
        Next i
        
        If loop_i_MeasV = 0 Then
            MeasV_PinS = MeasV_PinS_Temp
            MeasV_PinS_Temp1 = MeasV_PinS
        Else
            MeasV_PinS = MeasV_PinS_Temp1 & "+" & MeasV_PinS_Temp
            MeasV_PinS_Temp1 = MeasV_PinS
        End If
        
    Next loop_i_MeasV
    
End If

    'TTR,20200423, Oscar
    '----------Oscar,  20200214 skip parsing parameters
    Dim InstanceNumber As Long
    If GLB_InstanceParaDict.Exists(glb_TestInstance) And EnableFieldProcesingTTR = True Then
        InstanceNumber = GLB_InstanceParaDict(glb_TestInstance)
        Instance_Data = GLB_InstanceProcessedInfo(InstanceNumber).SavedInstance_Data
        TestConditionSeqData = GLB_InstanceProcessedInfo(InstanceNumber).SavedTestConditionSeqData
    Else
        Call ProcessInputToGLB(patset, TestSequence, CPUA_Flag_In_Pat, DisableComparePins, DisableConnectPins, DisableFRC, FRCPortName, MeasV_PinS, MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, MeasF_Flag_MeasureThreshold, _
                            MeasF_ThresholdPercentage, MeasF_WaitTime, MeasI_pinS, MeasI_Range, MeasI_WaitTime, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, DigSrc_pin, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, _
                            DigSrc_FlowForLoopIntegerName, SpecialCalcValSetting, InstSpecialSetting, CUS_Str_MainProgram, CUS_Str_DigCapData, CUS_Str_DigSrcData, Flag_SingleLimit, Spare_Argument, MeasF_PinS_Differential, ForceFunctional_Flag, _
                            MeasF_WalkingStrobe_Flag, MeasF_WalkingStrobe_StartV, MeasF_WalkingStrobe_EndV, MeasF_WalkingStrobe_StepVoltage, MeasF_WalkingStrobe_BothVohVolDiffV, MeasF_WalkingStrobe_interval, MeasF_WalkingStrobe_miniFreq, Meas_StoreName, _
                            Calc_Eqn, Interpose_PrePat, Interpose_PreMeas, Interpose_PostTest, CharSetName, ForceV_Val, ForceI_Val, UVI80_MeasV_WaitTime, RAK_Flag, WaitTime_VIRZ, Interpose_PostMeas)
        If EnableFieldProcesingTTR = True Then
            InstanceNumber = GLB_InstanceParaDict.Count
            ReDim Preserve GLB_InstanceProcessedInfo(InstanceNumber) As InstanceProcessedInfo
            GLB_InstanceProcessedInfo(InstanceNumber).SavedInstance_Data = Instance_Data
            GLB_InstanceProcessedInfo(InstanceNumber).SavedTestConditionSeqData = TestConditionSeqData
            GLB_InstanceParaDict.Add glb_TestInstance, InstanceNumber
        End If
    End If
    '----------Oscar,  20200214 skip parsing parameters
'                                                                                                                                                                ProfileMarkLeave VBT_LIB_HardIP_ProfileMark_1164    ' Profile Mark
    '[20231122][All][Neil]Record VDD pin name for measure IOL result on GPIO item
    gl_GPIO_VDD_Pin = vbNullString
    If UCase(ForceV_Val) Like "*_VAR*" Then
        'Ex. ForceV_Val=_VDDIO12_GRP0_VAR*0.2:_VDDIO12_GRP0_VAR*0.8|_VDDIO12_GRP0_VAR*0.8:_VDDIO12_GRP0_VAR*0.2
        gl_GPIO_VDD_Pin = Split(ForceV_Val, "_VAR")(0) '[_VDDIO12_GRP0]
        If InStr(gl_GPIO_VDD_Pin, "*") <> 0 Then
            '[Format]=_VDDIO12_GRP0_VAR*0.2 -> VDDIO12_GRP0
            gl_GPIO_VDD_Pin = right(gl_GPIO_VDD_Pin, Len(gl_GPIO_VDD_Pin) - 1 - InStr(gl_GPIO_VDD_Pin, "*"))
        Else
            '[Format]=0.2*_VDDIO12_GRP0_VAR -> VDDIO12_GRP0
            gl_GPIO_VDD_Pin = right(gl_GPIO_VDD_Pin, Len(gl_GPIO_VDD_Pin) - 1)
        End If
    End If	
	
    If TestSequence = "" Then                       '20170714
        ReDim TestSequenceArray(0) As String
        TestSequenceArray(0) = TestSequence
    Else
        TestSequenceArray = Split(TestSequence, ",")
    End If
    
    Interpose_PreMeas_Ary = SplitInputCondition(Interpose_PreMeas, "|") ''Carter, 20190616
    Dim PreMeas_Ary() As String
    If Interpose_PreMeas <> "" Then
        PreMeas_Ary = ParseData_InterPose(Interpose_PreMeas_Ary, TestSequenceArray)
    End If
    Dim PostMeas_Ary() As String
    
    
    Interpose_PostMeas_Ary = SplitInputCondition(Interpose_PostMeas, "|")
    If Interpose_PostMeas <> "" Then
        PostMeas_Ary = ParseData_InterPose(Interpose_PostMeas_Ary, TestSequenceArray)
    End If
    
    'theexec.Datalog.WriteComment "ProcessInputToGLB Time : " & FormatNumber(theexec.Timer(timer_), 6) & ":" & theexec.DataManager.instanceName & ":" & TestSequence & ":" & CStr(DigSrc_Sample_Size) & ":" & DigSrc_Equation & ":" & DigSrc_Assignment
    
    If InStr(UCase(CUS_Str_MainProgram), UCase("Loop_DigSrc")) <> 0 Then
        meas_val_delay_instance_name = vbNullString
        'Ex: CUS_Str_MainProgram ==> Loop_DigSrc;1;119;10;32;ddr0_mdll0_lsw:ddr0_mdll1_lsw:ddr1_mdll0_lsw:ddr1_mdll1_lsw
        Split_Loop_DigSrc_Str = Split(CUS_Str_MainProgram, ";")
        Loop_Init = Split_Loop_DigSrc_Str(1)
        Loop_Max = Split_Loop_DigSrc_Str(2)
        Loop_Step = Split_Loop_DigSrc_Str(3)
        Loop_BitNum = Split_Loop_DigSrc_Str(4)
        Loop_RegName = Split_Loop_DigSrc_Str(5)
        SplitLoop_RegName = Split(Loop_RegName, ":")
        'Check if exceed split(DigSrc,"@")
        If UBound(Split_Loop_DigSrc_Str) > 5 Then
            Loopendnumber = Split(Split_Loop_DigSrc_Str(6), "$")
        End If
    End If
    
    Dim loop_i As Long, Loop_j As Long
    Dim Temp_Equal_Str As String
    Dim Final_Comma_Str As String
    Temp_Equal_Str = vbNullString
    Final_Comma_Str = vbNullString
    Dim Split_CUS_Str_MainProgram_Str() As String
    Dim srcsweeparray() As String
    
    temp_CUS_String = vbNullString
    glb_s_sweepsrc_DigSrcAssignment = vbNullString
'    If InStr(UCase(CUS_Str_MainProgram), UCase("sweepsrc:")) <> 0 Then
'        'Dim srcsweeparray() As String
'        ''temp_CUS_String = CUS_Str_MainProgram
'        Split_CUS_Str_MainProgram_Str = Split(CUS_Str_MainProgram, ";")
'        For i = 0 To UBound(Split_CUS_Str_MainProgram_Str)
'            If InStr(Split_CUS_Str_MainProgram_Str(i), "=") <> 0 Then
'                If temp_CUS_String = "" Then
'                    temp_CUS_String = Split_CUS_Str_MainProgram_Str(i)
'                Else
'                    temp_CUS_String = temp_CUS_String & ";" & Split_CUS_Str_MainProgram_Str(i)
'                End If
'            End If
'        Next i
'        glb_s_sweepsrc_DigSrcAssignment = LCase(digsrc_assignment)
'        Call ProcessSweepString(digsrc_assignment, temp_CUS_String, srcsweeparray, Loop_Max)
'        srcnameindex = 0
'        Sweepnameforsweep = srcsweeparray
'        Loop_Max = Loop_Max - 1
'    End If

    If InStr(UCase(CUS_Str_MainProgram), UCase("sweepsrc:")) <> 0 Or InStr(UCase(CUS_Str_MainProgram), UCase("sweepvoltage")) <> 0 Then
        Split_CUS_Str_MainProgram_Str = Split(CUS_Str_MainProgram, ";")
        For i = 0 To UBound(Split_CUS_Str_MainProgram_Str)
            'Decide to Sweep_Code_First or Sweep_Voltage_First flag
            '0:Sweep_Code_First
            '1:Sweep_voltage_First
            If i = 0 Then
                If InStr(UCase(Split_CUS_Str_MainProgram_Str(i)), UCase("sweepvoltage")) <> 0 Then
                    Loop_Priority = 1  'Sweep_voltage_First
                Else
                    Loop_Priority = 0  'Sweep_Code_First
                End If
            End If
            If Not InStr(LCase(Split_CUS_Str_MainProgram_Str(i)), "sweepvoltage") <> 0 Then
                If InStr(Split_CUS_Str_MainProgram_Str(i), "=") <> 0 Then
                    If temp_CUS_String = "" Then
                        temp_CUS_String = Split_CUS_Str_MainProgram_Str(i)
                    Else
                        temp_CUS_String = temp_CUS_String & ";" & Split_CUS_Str_MainProgram_Str(i)
                    End If
                End If
            Else
                '==== New SweepVoltage function ====  '20221018
                Call ProcessSweepVoltageString(Split_CUS_Str_MainProgram_Str(i), sweepVoltage_ary, sweepVoltage_Loop_Max)   'Parsing Sweepvoltage information -- 20221018
                gl_sweepVoltage_Value = 0
            End If
        Next i
        glb_s_sweepsrc_DigSrcAssignment = LCase(digsrc_assignment)  '''For SweepSrc scenario -- 20221018
        
        'Call ProcessSweepString(DigSrc_Assignment, temp_CUS_String, srcsweeparray, Loop_Max)
        If digsrc_assignment <> "" Then
        Call ProcessSweepString(digsrc_assignment, temp_CUS_String, srcsweeparray, sweepCode_Loop_Max)    'Parsing SweepCode information -- 20221018
        srcnameindex = 0
        Sweepnameforsweep = srcsweeparray
        'Loop_Max = Loop_Max - 1
        sweepCode_Loop_Max = sweepCode_Loop_Max - 1         'Record sweep code max loop count -- 20221018
        End If
    End If	
	
    '[20231107][All][Neil] New sweep DigSrc method : sweep multi variable register code in single loop
    If InStr(UCase(CUS_Str_MainProgram), UCase("sweepsrc_singleloop:")) <> 0 Then
        Split_CUS_Str_MainProgram_Str = Split(CUS_Str_MainProgram, ";")
        For i = 0 To UBound(Split_CUS_Str_MainProgram_Str)
            If InStr(Split_CUS_Str_MainProgram_Str(i), "=") <> 0 Then
                If temp_CUS_String = "" Then
                    temp_CUS_String = Split_CUS_Str_MainProgram_Str(i)
                Else
                    temp_CUS_String = temp_CUS_String & ";" & Split_CUS_Str_MainProgram_Str(i)
                End If
            End If
        Next i
        glb_s_sweepsrc_DigSrcAssignment = LCase(digsrc_assignment)  
            
        Call ProcessSweepString_SingleLoop(digsrc_assignment, temp_CUS_String, srcsweeparray, Loop_Max)
        srcnameindex = 0
        Sweepnameforsweep = srcsweeparray
        Loop_Max = Loop_Max - 1
    End If
    
    'Decide to Sweep_Code_First or Sweep_Voltage_First flag
    '0:Sweep_Code_First/ 1:Sweep_Voltage_First -- 20221018
    If Loop_Priority = 0 Then
        Loop_Init = 0
        Loop_Init2 = 0
        Loop_Max = sweepCode_Loop_Max
        Loop_Max2 = sweepVoltage_Loop_Max
    Else
        Loop_Init = 0
        Loop_Init2 = 0
        Loop_Max = sweepVoltage_Loop_Max
        Loop_Max2 = sweepCode_Loop_Max
    End If
    '=== New Add a Loop for sweep voltage/code complex usage ===	
	
    For Loop_count = Loop_Init To Loop_Max
	For Loop_count2 = Loop_Init2 To Loop_Max2        '2nd Inner Loop level
        If InStr(UCase(CUS_Str_MainProgram), UCase("loopendnum")) <> 0 And Loop_count = Loop_Max Then
              
            Loop_count = Loopendnumber(1)
        
        End If
        
        If InStr(UCase(CUS_Str_MainProgram), UCase("Loop_DigSrc")) <> 0 Then
              Call Public_AddStoredString(Split_Loop_DigSrc_Str(0), CStr(Loop_count))
        End If
      
        
        If InStr(UCase(CUS_Str_MainProgram), UCase("Calc_Freq_SDLL_SWP")) <> 0 Then gl_Tname_Alg_Index = Loop_count
        
        'TypeName (Loop_count)
        
        If (Sweep_Enable = True) Then
            CUS_Str_DigCapData = Sweep_CUS_Str_DigCapData
            
            Call SetForceSweepVoltAndTName(Sweep_Info, CUS_Str_DigCapData, Loop_count)
            
            If InStr(UCase(glb_TestInstance), "MTRGR_T2P6") <> 0 Or InStr(UCase(glb_TestInstance), "MTRGR_T2P7") <> 0 Then
                Calc_Eqn = Replace(Calc_Eqn, Replace(Split(Split(Calc_Eqn, ":")(2), "(")(1), ")", vbNullString), Split(CUS_Str_DigCapData, ":")(2))
                CUS_Str_DigCapData = Replace(CUS_Str_DigCapData, Split(CUS_Str_DigCapData, ":")(1), Split(CUS_Str_DigCapData, ":")(1) & CStr(Loop_count))
            Else
                Calc_Eqn = Sweep_Calc_Eqn & "," & CStr(Loop_count)
            End If
        End If
        
'        If InStr(UCase(CUS_Str_MainProgram), UCase("sweepsrc:")) Or InStr(UCase(CUS_Str_MainProgram), UCase("specialsorce")) <> 0 Then
'            TheExec.Flow.TestLimitIndex = 0
'            digsrc_assignment = srcsweeparray(Loop_count)
'            Instance_Data.digsrc_assignment = digsrc_assignment
'        End If
 
        If Loop_Priority = 0 Then
                If InStr(UCase(CUS_Str_MainProgram), UCase("sweepsrc:")) Or InStr(UCase(CUS_Str_MainProgram), UCase("sweepsrc_singleloop:")) Or InStr(UCase(CUS_Str_MainProgram), UCase("specialsorce")) <> 0 Then
                    TheExec.Flow.TestLimitIndex = 0
                    digsrc_assignment = srcsweeparray(Loop_count)
                    Instance_Data.digsrc_assignment = srcsweeparray(Loop_count)
                    gl_sweepCode_Loop_Indx = Loop_count         'Record Sweep Code loop index to global -- 20231013
                End If
                ''=============== For Sweep Voltage in Interpose function ===============
                If InStr(LCase(CUS_Str_MainProgram), "sweepvoltage") <> 0 Then
                    gl_sweepVoltage_Loop_Indx = Loop_count2     'Record Sweep Voltage loop index to global -- 20231013
                    gl_sweepVoltage_Value = CDbl(sweepVoltage_ary(Loop_count2))
                End If
                ''=============== For Sweep DigSrc new method ===============
                'New sweep DigSrc method : sweep multi variable register code in single loop  -- 20230224
                If InStr(UCase(CUS_Str_MainProgram), UCase("sweepsrc_singleloop:")) <> 0 Then
                    TheExec.Flow.TestLimitIndex = 0
                    digsrc_assignment = srcsweeparray(Loop_count)
                    Instance_Data.digsrc_assignment = srcsweeparray(Loop_count)
                    gl_sweepCode_Loop_Indx = Loop_count          'Record Sweep Code loop index to global -- 20231013
                End If
        Else
                If InStr(UCase(CUS_Str_MainProgram), UCase("sweepsrc:")) Or InStr(UCase(CUS_Str_MainProgram), UCase("sweepsrc_singleloop:")) Or InStr(UCase(CUS_Str_MainProgram), UCase("specialsorce")) <> 0 Then
                    TheExec.Flow.TestLimitIndex = 0
                    digsrc_assignment = srcsweeparray(Loop_count2)
                    Instance_Data.digsrc_assignment = srcsweeparray(Loop_count2)
                    gl_sweepCode_Loop_Indx = Loop_count2          'Record Sweep Code loop index to global -- 20231013
                End If
                ''=============== For Sweep Voltage in Interpose function ===============
                If InStr(LCase(CUS_Str_MainProgram), "sweepvoltage") <> 0 Then
                    gl_sweepVoltage_Loop_Indx = Loop_count        'Record Sweep Voltage loop index to global -- 20231013
                    gl_sweepVoltage_Value = CDbl(sweepVoltage_ary(Loop_count))
                End If
            
                ''=============== For Sweep DigSrc new method ===============
                'New sweep DigSrc method : sweep multi variable register code in single loop  -- 20230224
                If InStr(UCase(CUS_Str_MainProgram), UCase("sweepsrc_singleloop:")) <> 0 Then
                    TheExec.Flow.TestLimitIndex = 0
                    digsrc_assignment = srcsweeparray(Loop_count)
                    Instance_Data.digsrc_assignment = srcsweeparray(Loop_count)
                    gl_sweepCode_Loop_Indx = Loop_count             'Record Sweep Code loop index to global -- 20231013
                End If
        End If
        '======= New structure for Sweep Code/Voltage Priority ======
		
        If CUS_Str_MainProgram <> "" And InStr(UCase(CUS_Str_MainProgram), UCase("Loop_DigSrc")) <> 0 Then
            binstr = Dec2BinStr32Bit_Rev(Loop_BitNum, Loop_count)
            Loop_SplitByComma = Split(digsrc_assignment, ";")
            
            For loop_i = 0 To UBound(Loop_SplitByComma)
                Loop_SplitByEqual = Split(Loop_SplitByComma(loop_i), "=")
                For Loop_j = 0 To UBound(SplitLoop_RegName)
                    If UCase(Loop_SplitByEqual(0)) = UCase(SplitLoop_RegName(Loop_j)) Then
                        Loop_SplitByEqual(1) = binstr
                        Temp_Equal_Str = Loop_SplitByEqual(0) & "=" & Loop_SplitByEqual(1)
                        Exit For
                    Else
                        Temp_Equal_Str = Loop_SplitByEqual(0) & "=" & Loop_SplitByEqual(1)
                    End If
                Next Loop_j
                If loop_i = 0 Then
                    Final_Comma_Str = Temp_Equal_Str
                Else
                    Final_Comma_Str = Final_Comma_Str & ";" & Temp_Equal_Str
                End If
            Next loop_i
            digsrc_assignment = Final_Comma_Str
        End If
        
        
        '' 20190529 - Add for sweep force V
        If InStr(Interpose_PrePat, "x_sweep") <> 0 Then
            gl_Sweep_Glb_TName = CDbl(val(TheExec.Flow.var("x_sweep").value)) / 1000
            Interpose_PrePat = Replace(Interpose_PrePat, "x_sweep", gl_Sweep_Glb_TName)
            
            'USB_DP:V:x_sweep;Sweep_Name:
            If InStr(Interpose_PrePat, "Sweep_Name") <> 0 Then
                Interpose_PrePat = Replace(Interpose_PrePat, "Sweep_Name:", vbNullString)
            End If
        End If
        
        If InStr(Interpose_PrePat, "x_power_sweep") <> 0 Then
            gl_Sweep_Glb_TName = CDbl(val(TheExec.Flow.var("x_power_sweep").value)) / 1000
            Interpose_PrePat = Replace(Interpose_PrePat, "x_power_sweep", gl_Sweep_Glb_TName)
            
            If InStr(Interpose_PrePat, "Sweep_Name") <> 0 Then
                Interpose_PrePat = Replace(Interpose_PrePat, "Sweep_Name:", vbNullString)
            End If
        End If
        
          
        '' 20190531 - Add for sweep Volt by shmoo
        If InStr(Interpose_PrePat, "Volt_sweep_GLB") <> 0 Then
        For Each site In TheExec.sites
            gl_Sweep_Glb_TName = CDbl(TheExec.Specs.Globals("Volt_sweep_GLB").CurrentValue)
            Exit For
        Next site
            Interpose_PrePat = Replace(Interpose_PrePat, "Volt_sweep_GLB", gl_Sweep_Glb_TName)
            

            If InStr(Interpose_PrePat, "Sweep_Name") <> 0 Then
                Interpose_PrePat = Replace(Interpose_PrePat, "Sweep_Name:", vbNullString)
            End If
        End If
         
        If gl_Flag_HardIP_Characterization_1stRun Then: Exit Function
        
        '20230714 sync RF feature
        #If RF = True Then
             glb_DSSCSetup = DSSCSetup
        #End If

        '' 20160923 - Add Interpose_PrePat entry point
        If Interpose_PrePat <> "" Then
            Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
        End If

        'New Method for IO_DS_VOHVOL_with_RAK '-- 20220421
        If Instance_Data.InstSpecialSetting = IO_DS_VOHVOL_with_RAK Then
            Call IO_HardIP_DS_VOH_VOL_SETUP(UCase(glb_TestInstance))
        End If
        
''        If gl_Flag_HardIP_Characterization_1stRun Then: Exit Function ''' Update in Crete @210805
            
        ''20161205 - Force_Flow_Shmoo_Condition
        If TheExec.sites.item(site).SiteVariableValue("Flow_Shmoo_DevCharSetup") <> "" Then Force_Flow_Shmoo_Condition
        'Do Flow Shmoo
        
        If patset.value <> "" Then
            Shmoo_Pattern = patset.value '' 20170808 add for shmoo pattern name print
            TheHdw.Patterns(patset).Load
            Call PatternBurstCheckAndSplit(patset.value, PattArray, PatCount)
'''            Call PATT_GetPatListFromPatternSet(PatSet.value, PattArray, PatCount)
        Else
            ReDim PattArray(0)
            PattArray(0) = vbNullString
        End If
            
        If (DisableConnectPins <> "") Then TheHdw.Digital.Pins(DisableConnectPins).Disconnect
        If (DisableComparePins <> "") Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = True
        
        ''20161107-Return sweep test name
        Dim Rtn_SweepTestName As String
        Rtn_SweepTestName = vbNullString
        gl_TName_Pat = patset.value
        
        Dim current_pat_index As Integer
        current_pat_index = 0
        
        
        '20191003 add for CPM with Multi_Init(DigSrc)_PL
        Dim DigSrc_Equation_temp_array() As String
        Dim DigSrc_Assignment_temp_array() As String
        Dim DigSec_Multi_Init_PL__Seq_index As Long: DigSec_Multi_Init_PL__Seq_index = 0
       
        If (InStr(UCase(CUS_Str_MainProgram), "CPM_MULTI_INIT_DIGSRC") > 0) Then
            DigSrc_Equation_temp_array = Split(DigSrc_Equation, "|")
            DigSrc_Assignment_temp_array = Split(digsrc_assignment, "|")
        End If
       
        
        
        For Each patt In PattArray
            If patt <> "" Then
                TheExec.Flow.TestLimitIndex = 0
                Pat = CStr(patt)
                TheHdw.Patterns(Pat).Load
'                                                                                                                                                            Dim VBT_LIB_HardIP_ProfileMark_1267 As Long: VBT_LIB_HardIP_ProfileMark_1267 = ProfileMarkEnter(2, instance_name & "_" & "GenDigSrc&Module=VBT_LIB_HardIP&ProcName=Meas_FreqVoltCurr_Universal_func&LineNumber=1264")    ' Profile Mark
                If (InStr(UCase(CUS_Str_MainProgram), "CPM_INIT_DIGSRC") > 0) And (UCase(patt) Like "*_PL*") Then

                Else
                    '20191003 add for CPM with Multi_Init(DigSrc)_PL
                    If (InStr(UCase(CUS_Str_MainProgram), "CPM_MULTI_INIT_DIGSRC") > 0) Then
                        DigSrc_Equation = DigSrc_Equation_temp_array(DigSec_Multi_Init_PL__Seq_index)
                        digsrc_assignment = DigSrc_Assignment_temp_array(DigSec_Multi_Init_PL__Seq_index)
                        DigSec_Multi_Init_PL__Seq_index = DigSec_Multi_Init_PL__Seq_index + 1
                    End If
                    
                    
                
                Set InDSPWave = Nothing
                
'*******************************New Feature for trimcode table*******************************
'Added by  20190509
                If LCase(digsrc_assignment) Like "*table*" Then digsrc_assignment = digsrc_assignment & "_" & CStr(TheExec.Flow.var("SrcCodeIndx").value)
'********************************************************************************************
                Call GeneralDigSrcSettingWithBurst(LCase(patt), DigSrc_pin, InDSPWave, Rtn_SweepTestName, MSB_First_Flag)       'Fix MSB_First_Flag import -- 20230203
                End If
                Set OutDspWave = Nothing
                Call GeneralDigCapSetting(Pat, DigCap_Pin, DigCap_Sample_Size, OutDspWave)
                 
                Call PrintDigCapSetting(DigCap_Pin, DigCap_Sample_Size, CUS_Str_DigCapData)
                

                If InStr(UCase(CUS_Str_MainProgram), "MTR_UVI80_SETUP") <> 0 Then
                    Call MTR_UVI80_Setup
                End If
            
            
            
                Dim SplitByCommaStr() As String
                Dim ForcePin_X As String
                Dim ForcePin_Y As String
                Dim SweepIndexStr_X As String
                Dim ForceVal_X As Double
                If LCase(CUS_Str_MainProgram) Like "*x_sweep*" Then
                    SplitByCommaStr = Split(CUS_Str_MainProgram, ",")
                    SweepIndexStr_X = SplitByCommaStr(0)
                         ForcePin_X = SplitByCommaStr(1)
                         
                          ForceVal_X = CDbl(val(TheExec.Flow.var(SweepIndexStr_X).value)) / 1000
                          TheExec.Datalog.WriteComment "ForcePin = " & ForcePin_X & "; ForceVal_X = " & ForceVal_X & "V"
                          'TheExec.Datalog.WriteComment "ForcePin = " & SplitByCommaStr(2) & ";  ForceVal_X  = " & ForceVal_X & "V"
                          'TheHdw.DCVS.Pins(ForcePin_X).Voltage.Value = ForceVal_X
                          'TheHdw.DCVS.Pins(SplitByCommaStr(2)).Voltage.Value = ForceVal_X
                          
                        TheHdw.Digital.Pins(ForcePin_X).Disconnect
                        
                            With TheHdw.PPMU.Pins(ForcePin_X)
                                .Gate = tlOff
                                .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_Max_InitialValue_FI_Range
                                .ForceV CDbl(ForceVal_X), 0.02
                                .Connect
                                .Gate = tlOn
                            End With
                          
                          
                        FourceV = ForceVal_X
                End If
            
            
            
                '' 20160713 - If no cpuflags in the test item modify the code to run pattern by using .test
                If (CPUA_Flag_In_Pat) Then
                    Call TheHdw.Patterns(Pat).start
                Else
                    Call TheHdw.Patterns(Pat).test(pfNever, 0)
                End If
            End If
            
            'TestSeqNum = 0
            
            'Call ProcessTestNameInputString(OutputTname, UBound(TestSequenceArray))    Remove
            
'            If PatCount > 1 Then
'                Dim ot_cnt As Long
'                For ot_cnt = 0 To UBound(OutputTname)
'                    OutputTname(ot_cnt) = OutputTname(ot_cnt) & Split(Split(Split(Pat, "\")(UBound(Split(Pat, "\"))), ":")(0), "_")(12)
'                Next ot_cnt
'            End If
            

            TestSeqNum = 0


            For Each Ts In TestSequenceArray
                Instance_Data.TestSeqNum = TestSeqNum
                ''20150907 - Only need CPUA_Flag_In_Pat to do control
                If (CPUA_Flag_In_Pat) Then
                    Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0) 'Meas during CPUA loop
                Else
                    Call TheHdw.Digital.Patgen.HaltWait 'Pattern without CPUA loop
                End If
                
                If InStr(MeasF_PinS_SingleEnd, "$") Then
                    Dim MeasF_Set() As String
                    MeasF_Set = Split(MeasF_PinS_SingleEnd, ",")
                End If
                
                ''20160923 - Add Interpose_PreMeas entry point by each sequence
                    
            '''------- Carter, 20190616
'                If Interpose_PreMeas <> "" Then
'                    If UBound(Interpose_PreMeas_Ary) = 0 Then
'                        Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
'                    ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
'                        Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
'                    End If
'                End If
            '''------- Carter, 20190616
                
                If InStr(Ts, "VDM") <> 0 Then
                    TestOptLen = 1
                    Ts = "V"
                Elseif UCase(Ts) = "VDIFF2" Then		'[20240207][All][Brian] Support Vdiff2 in universal
					TestOptLen = 1
				else
                    TestOptLen = Len(Ts)
                End If           
                
                
                For k = 1 To TestOptLen
                    Instance_Data.TestSeqSweepNum = k - 1
                    If UCase(Ts) = "VDIFF2" Then
                        TestOption = Ts
                    Else
                        TestOption = mid(Ts, k, 1)
                    End If
                    
               '''-------Start - Add per sweep feature for interpose_premeas - Carter, 20190614-------
                    If Interpose_PreMeas <> "" Then
                        If PreMeas_Ary(TestSeqNum, k - 1) <> "" Then
                            Call SetForceCondition(PreMeas_Ary(TestSeqNum, k - 1) & ";STOREPREMEAS")
                        End If
                    End If
               '''-------End - Add per sweep feature for interpose_premeas - Carter, 20190614-------
               
                    For Each site In TheExec.sites.Active
                        testnum = TheExec.sites.item(site).TestNumber
                    Next site
                    
                    '----------------0427 begin-------------------------------------
                    If InStr(MeasF_PinS_SingleEnd, "$") Then
                        MeasureF_Pin_SingleEnd = Replace(MeasF_Set(current_pat_index), "$", vbNullString)
                    End If
                    '----------------0427 end---------------------------------------
                    'TTR
                    ''glb_Disable_CurrRangeSetting_Print = TheExec.Flow.EnableWord("Enable_HardIP_FieldProcesingTTR")
                    
                    Select Case UCase(TestOption)
					
                        Case "VDIFF2"
						
                            Call HardIP_Vdiff2_PPMU		'[20240207][All][Brian] Support Vdiff2 in universal

                        Case "V"
'                                                                                                                                                                Dim VBT_LIB_HardIP_ProfileMark_1347 As Long: VBT_LIB_HardIP_ProfileMark_1347 = ProfileMarkEnter(2, instance_name & "_" & "MeasV&Module=VBT_LIB_HardIP&ProcName=Meas_FreqVoltCurr_Universal_func&LineNumber=1345")    ' Profile Mark
                            
                            Call HardIP_MeasureVolt
'                                                                                                                                                                ProfileMarkLeave VBT_LIB_HardIP_ProfileMark_1347    ' Profile Mark
                        Case "F"
'                                                                                                                                                                Dim VBT_LIB_HardIP_ProfileMark_1352 As Long: VBT_LIB_HardIP_ProfileMark_1352 = ProfileMarkEnter(2, instance_name & "_" & "MeasF&Module=VBT_LIB_HardIP&ProcName=Meas_FreqVoltCurr_Universal_func&LineNumber=1350")    ' Profile Mark
                            
                            Call HardIP_MeasureFreq
'                                                                                                                                                                ProfileMarkLeave VBT_LIB_HardIP_ProfileMark_1352    ' Profile Mark
                        Case "D"

                            Call HardIP_MeasureDuty '[20231107][All][Neil] Duty Cycle Test Methodology on Digital and DCVI instrument
							
                        Case "I"
                            If DisableFRC = True Then FreeRunClk_Disable (FRCPortName)
'                                                                                                                                                                Dim VBT_LIB_HardIP_ProfileMark_1358 As Long: VBT_LIB_HardIP_ProfileMark_1358 = ProfileMarkEnter(2, instance_name & "_" & "MeasI&Module=VBT_LIB_HardIP&ProcName=Meas_FreqVoltCurr_Universal_func&LineNumber=1356")    ' Profile Mark
                            
                            Call HardIP_MeasureCurrent
'                                                                                                                                                                ProfileMarkLeave VBT_LIB_HardIP_ProfileMark_1358    ' Profile Mark
                        Case "R"
'                                                                                                                                                                Dim VBT_LIB_HardIP_ProfileMark_1363 As Long: VBT_LIB_HardIP_ProfileMark_1363 = ProfileMarkEnter(2, instance_name & "_" & "MeasR&Module=VBT_LIB_HardIP&ProcName=Meas_FreqVoltCurr_Universal_func&LineNumber=1361")    ' Profile Mark
                            
                            Call HardIP_SetupAndMeasureR
'                                                                                                                                                                ProfileMarkLeave VBT_LIB_HardIP_ProfileMark_1363    ' Profile Mark
                        Case "Z"
'                                                                                                                                                                Dim VBT_LIB_HardIP_ProfileMark_1368 As Long: VBT_LIB_HardIP_ProfileMark_1368 = ProfileMarkEnter(2, instance_name & "_" & "MeasZ&Module=VBT_LIB_HardIP&ProcName=Meas_FreqVoltCurr_Universal_func&LineNumber=1366")    ' Profile Mark
                            
                            Call HardIP_SetupAndMeasureZ
'                                                                                                                                                                ProfileMarkLeave VBT_LIB_HardIP_ProfileMark_1368    ' Profile Mark
                        Case "P"
                            HardIP_BySeqCurrentProfile
                        Case "N"
                            restore_Flag = True
                        Case Else
                            TheExec.Datalog.WriteComment "Error: Test Option " & UCase(TestOption) & " cannot be recognized!!!"
                    End Select
                    If TheExec.sites.Active.Count = 0 Then Exit Function
                    
                '''-------Start - Add per sweep feature for interpose_premeas - Carter, 20190614-------
                    If Interpose_PreMeas <> "" Then
                        If PreMeas_Ary(TestSeqNum, k - 1) <> "" And UCase(TestOption) <> "N" Then
                            Call SetForceCondition("RESTOREPREMEAS")
                        End If
                    End If
                '''-------End - Add per sweep feature for interpose_premeas - Carter, 20190614-------
                
                '''-------Start - Add Interpose_PostMeas for Pattern Burst, 20210526-------
                    If Interpose_PostMeas <> "" Then
                        If PostMeas_Ary(TestSeqNum, k - 1) <> "" Then
                            Call SetForceCondition(PostMeas_Ary(TestSeqNum, k - 1))
                        End If
                    End If
                '''-------End - Add Interpose_PostMeas for Pattern Burst, 20210526-------
                Next k
                
                ''20161206-Restore force condiction after measurement
    ''            Call SetForceCondition("RESTORE")
    
    '''------- Carter, 20190616
'                If Interpose_PreMeas <> "" And Ts <> "N" Then
'                    If UBound(Interpose_PreMeas_Ary) = 0 Then
'                        Call SetForceCondition("RESTOREPREMEAS")
'                    ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
'                        Call SetForceCondition("RESTOREPREMEAS")
'                    End If
'                End If
   '''------- Carter, 20190616
                
                TestSeqNum = TestSeqNum + 1
                
                If (CPUA_Flag_In_Pat) Then Call TheHdw.Digital.Patgen.Continue(0, cpuA)
                Instance_Data.TestSeqNum = TestSeqNum
                
            Next Ts
            
            If DebugPrintEnable = True Then TheExec.Datalog.WriteComment "  Pattern(" & PatCount & "): " & Pat & vbNullString
            
            TheHdw.Digital.Patgen.HaltWait ' Haltwait at patten end
            
            PatCount = PatCount + 1
            
'            If gl_FlowForLoop_DigSrc_SweepCode <> "" Then         '20180509
'                gl_FlowForLoop_DigSrc_SweepCode = ""
'            End If
            
            '' 20160211 - Process DigCapData by using DSP
    ''        If b_ProcessDigCapByDSP = True Then
                If DigCap_Sample_Size <> 0 Then
                    Dim DigCapPinAry() As String, NumberPins As Long
                    Dim CUS_Str_DigCapData_temp As String
                    Call TheExec.DataManager.DecomposePinList(DigCap_Pin, DigCapPinAry(), NumberPins)
                    
                    
                    If NumberPins > 1 Then
                        Call CreateSimulateDataDSPWave_Parallel(OutDspWave, DigCap_Sample_Size)
                        Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDspWave, NumberPins)
                        Call DigCapDataProcessByDSP_Parallel(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, NumberPins, , DigCap_Pin.value)
    
                    ElseIf NumberPins = 1 Then
                        Call CreateSimulateDataDSPWave(OutDspWave, DigCap_Sample_Size, DigCap_DataWidth)
                        Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDspWave, NumberPins)
'                                                                                                                                                                Dim VBT_LIB_HardIP_ProfileMark_1429 As Long: VBT_LIB_HardIP_ProfileMark_1429 = ProfileMarkEnter(2, instance_name & "_" & "MeasC&Module=VBT_LIB_HardIP&ProcName=Meas_FreqVoltCurr_Universal_func&LineNumber=1427")    ' Profile Mark
                        
                        Call DigCapDataProcessByDSP(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, DigCap_DataWidth, CUS_Str_MainProgram, , DigCap_Pin.value, , MSB_First_Flag)
'                                                                                                                                                             ProfileMarkLeave VBT_LIB_HardIP_ProfileMark_1429    ' Profile Mark
                    Else
                    'Do nothing
                    End If
                End If
                
                '' 20160907 - Process calculate equation by dictionary.
                If Calc_Eqn <> "" And InStr(LCase(TestSequence), "p") = 0 Then
                    Call ProcessCalcEquation(Calc_Eqn)
                End If
                
		If gl_Disable_HIP_debug_log = False Then
			Call PatternBurstCheck(CStr(patt)) ''Datalog Pattern Module Name, 20240430
		End If	
				
                '' 20160713 - Call write functional result if cpu flag in pattern
                'If (CPUA_Flag_In_Pat) Then
                'For Pattern Burst BinOut, 20200527, From Sicily, Oscar
                If gl_Instpatcount_Dic.Exists(LCase(glb_TestInstance)) And gl_patflag_Dic.Exists(LCase(patt + Split(glb_TestInstance, "_")(UBound(Split(glb_TestInstance, "_"))))) Then
                    Call HardIP_WriteFuncResult(, , glb_TestInstance, gl_patflag_Dic(LCase(patt + Split(glb_TestInstance, "_")(UBound(Split(glb_TestInstance, "_"))))))
                ElseIf patt <> "" Then
                    'Added by Oscar to accept no patt instance, From JadeCdie, 20201023
                    Call HardIP_WriteFuncResult(, , glb_TestInstance)
                Else
                'Do nothing
                End If
                'End If
                
                If gl_FlowForLoop_DigSrc_SweepCode <> "" Then        '20180814
                    gl_FlowForLoop_DigSrc_SweepCode = vbNullString
                    gl_FlowForLoop_DigSrc_SweepCode_Dec = vbNullString '20190613 CT add for Decimal value printing
                End If
                
    ''        End If
    
            current_pat_index = current_pat_index + 1
            
                If Interpose_PreMeas <> "" And restore_Flag = True Then
                    Call SetForceCondition("RESTOREPREMEAS")

                End If
            
            
            
               If LCase(CUS_Str_MainProgram) Like "*x_sweep*" Then

                    With TheHdw.PPMU.Pins(ForcePin_X)
                            .ForceV pc_Def_PPMU_InitialValue_FV, pc_Def_PPMU_Max_InitialValue_FI_Range ''FVMI - Carter, 20190503
                            .Disconnect
                            .Gate = tlOff
                    End With
                    TheHdw.Digital.Pins(ForcePin_X).Connect ''Connect Digital pins after measurement - Carter, 20190503

                End If
                    
                  
            gl_Sweep_Glb_TName = vbNullString '' 20190529 - Add for sweep force V
                  
        Next patt
            
''      ''20170405-Record all functional test result from flow for loop opcode, use global string to store them
        If CUS_Str_DigSrcData <> "" And UCase(CUS_Str_DigSrcData) = UCase("BinToGray") Then
            If CPUA_Flag_In_Pat = False Then
                Call DisplayForLoopFuncResult_EndOfTest(CUS_Str_DigSrcData, Rtn_SweepTestName, CPUA_Flag_In_Pat, DigSrc_FlowForLoopIntegerName)
            End If
        End If
     If MeasureV_pin <> "" Then
         Call EndSetupForMeasureVoltPins(MeasureV_Pin_PPMU, MeasureV_Pin_UVI80)
     End If
     
     If DisableConnectPins <> "" Then TheHdw.Digital.Pins(DisableConnectPins).Connect
     If DisableComparePins <> "" Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = False
         
     If DisableFRC = True Then
         Call ReStart_FRC(FRCPortName)
     End If
     
     DebugPrintFunc patset.value, True ' print all debug information
     
     If TheExec.sites.item(site).SiteVariableValue("Flow_Shmoo_DevCharSetup") <> "" Then
     'Do Flow Shmoo
         If Flow_Shmoo_Port_Name <> "" Then Restart_All_Freerun_Clk
     End If
     
     If Interpose_PrePat <> "" Then
         Call SetForceCondition("RESTOREPREPAT")
     End If
     
     ''=============================== CharSetName ====================================
     Dim p As Variant
     If TheExec.DevChar.Setups.IsRunning = False And CharSetName <> "" Then
         Dim ApplyPins As String, Setup_mode As String, p_ary() As String, p_cnt As Long
         'If TheExec.DevChar.Setups(CharSetName).TestMethod.Value = tlDevCharTestMethod_Reburst Then TheExec.Datalog.WriteComment "[PrintCharCondition:" & PrintCharSetup(Interpose_PrePat_GLB) & ",Test]"
         Setup_mode = TheExec.DevChar.Setups(CharSetName).Shmoo.axes(tlDevCharShmooAxis_X).Parameter.name
         If (LCase(Setup_mode) <> "vid" And LCase(Setup_mode) <> "vicm") Then
             ApplyPins = TheExec.DevChar.Setups(CharSetName).Shmoo.axes(tlDevCharShmooAxis_X).ApplyTo.Pins
             TheExec.DataManager.DecomposePinList ApplyPins, p_ary, p_cnt
             For Each p In p_ary
                 TheExec.DevChar.Setups(CharSetName).Shmoo.axes(tlDevCharShmooAxis_X).ApplyTo.Pins = p
                 run_shmoo CharSetName
             Next p
             TheExec.DevChar.Setups(CharSetName).Shmoo.axes(tlDevCharShmooAxis_X).ApplyTo.Pins = ApplyPins
         Else
             ApplyPins = TheExec.DevChar.Setups(CharSetName).Shmoo.axes(tlDevCharShmooAxis_X).ApplyTo.Pins
             p_ary = Split(ApplyPins, ",")
             For Each p In p_ary
                 TheExec.DevChar.Setups(CharSetName).Shmoo.axes(tlDevCharShmooAxis_X).ApplyTo.Pins = p
                 run_shmoo CharSetName
             Next p
             TheExec.DevChar.Setups(CharSetName).Shmoo.axes(tlDevCharShmooAxis_X).ApplyTo.Pins = ApplyPins
             'run_shmoo CharSetName
         End If
     End If
    
    If CUS_Str_MainProgram <> "" And InStr(UCase(CUS_Str_MainProgram), UCase("Loop_DigSrc")) <> 0 And Loop_Step <> 1 Then
        Loop_count = Loop_count + Loop_Step - 1
    End If
 
    If InStr(UCase(CUS_Str_MainProgram), UCase("V_Sweep")) > 0 Then
        sweep_power_val_per_loop_count = vbNullString
    End If
 
	Next Loop_count2    '2nd Inner Loop level 'sweepvoltage--20240325
    Next Loop_count
    ''================================================================================
    
    ' [20230721][All][Neil] Move Interpose_PostTest to out of Pattern loop level
    Call SetForceCondition(Interpose_PostTest)
        
    If InStr(digsrc_assignment, "digsrctable") <> 0 Then
                  'Table_Decvalue = ""
                  gl_SweepNum = vbNullString
    End If

    ReDim TestConditionSeqData(0)
    Dim Instance_Data_temp() As Instance_Type
    ReDim Instance_Data_temp(0)
    Instance_Data = Instance_Data_temp(0)
    'TTR,20200423, Oscar
    Instance_Data.Meas_StoreName_Flag = False ''Carter, 20190521
    temp_CUS_String = vbNullString
    'Alarm check From Sicily,20200423, Oscar
    ' Check implicit alarms
    TheHdw.Alarms.Check
    
''    ProfileCollect_End m_InstanceName, Validating_
''    TestTimeCollect_End TestTime_D, m_InstanceName, Validating_

    If gl_EnableCurrentProfile Or gl_EnableVoltageProfile Then
        ProfileCollect_End glb_TestInstance
    End If

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP", "Meas_FreqVoltCurr_Universal_func") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function


Public Function Opt_DdrLpBkFunc2(DqsSwpPat As Pattern, DqSwpPat As Pattern, _
                            DisableComparePins As PinList, DisableConnectPins As PinList, _
                            DigCap_Pin As PinList, NoOfBists As Long, _
                            DqSwpNoOfBits As String, DqsSwpNoOfBits As String, _
                            Optional DispCaptStrm As Boolean = True, _
                            Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As Long, Optional DigSrc_Sample_Size As String, _
                            Optional DqsDigSrc_Equation As String, Optional DqDigSrc_Equation As String, _
                            Optional digsrc_assignment As String, _
                            Optional CUS_Str_DigSrcData As String, _
                            Optional DigCap_DSPWaveSetting As CalculateMethodSetup = 0, _
                            Optional EyeTestRegName As String, _
                            Optional DigCap_Sample_Size_Dqs As Long, _
                            Optional CUS_Str_DigCapData_Dqs As String, _
                            Optional DigCap_Sample_Size_Dq As Integer, _
                            Optional CUS_Str_DigCapData_Dq As String, _
                            Optional Interpose_PrePat As String, _
                            Optional SweepVtStr As String, _
                            Optional Calc_Eqn As String, _
                            Optional DigSrc_FlowForLoopIntegerName As String = vbNullString, _
                            Optional Validating_ As Boolean) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

    ''''--------------------------------------------------------------------------------------------------
    ''''    Based on TMA V04C function "Meas_FreqVoltCurr_Univeral_func" in VBT_LIB_HardIP_New Module
    ''''    rev 0, by Zheng Xiao, Apple Inc, 1/1/2016
    ''''
    ''''    Adapted for Starling DDR external loopback test.
    ''''        - MDLL code no longer captured in this test
    ''''        - DQ ELB lower limit now fixed as 1/4 of PI eye openings (DQ 64, CA 128)
    ''''        - Impedance cal settings needed to be sourced in
    ''''    rev 1, by Zheng Xiao, Apple Inc, 1/26/2016
    ''''--------------------------------------------------------------------------------------------------
    ''''    Test fucntion for DDR (AMP) loopback test, with data eye sweeping. Applicable to all buses, both
    '''         internal and external loopback tests.
    ''''    There are two sweeps, involving 2 patterns, sweeping left and right from the center point, respectively.
    ''''        The eye width will be the combination of the 2 sweeps, and being tested
    ''''
    ''''    In case of multiple eyes, the maximum eye will be tested.
    ''''
    ''''    DQ sweep : Moving DQ strobe the captured bits stream are consecutive results from the first lane to the last
    ''''    DQS sweep: moving DQS strobe
    ''''
    ''''--------------------------------------------------------------------------------------------------
    ''''    Usage
    ''''        Opt_DdrLpBkFunc2 is to be used to construct test instance directly.
    ''''--------------------------------------------------------------------------------------------------
    ''''    Function calls
    ''''        - PATT_GetPatListFromPatternSet (original)
    ''''        - DigCapSetup (original): inline codes should be used here.
    ''''        - SetupDigSrcDspWave (original): setup dssc dig source. inline be better
    ''''        - FindMaxEyeWidth: DSP fucntional call stitch 2 sweeps to a single eye diagram of each BIST,
    ''''            reporting the eyewidths
    ''''        - DebugPrintFunc (original)
    ''''--------------------------------------------------------------------------------------------------
    ''''    Modifications:
    ''''        - Completely re-written for speciallized function for DDR eye-sweep based loopback tests
    ''''        - Eliminated the need to pass the first sweep results via global variable, by including
    ''''            both sweeps in a single function
    ''''        - Instead of using VBT for waveform conversion, processing, and eye width finding, using
    ''''            DSP based function calls for efficiencies and multi-site handling
    ''''--------------------------------------------------------------------------------------------------
    ''''    Argument List
    ''''        DqSwpPat:           Pattern set for DQ sweep
    ''''        DqsSwpPat:          Pattern set for DQS sweep
    ''''        DisableComparePins: Retained from the original function. Pins to be masked
    ''''        DisableConnectPins: Retained. Pins to be disconnected. (DisconnectPins is a more suitable name)
    ''''        DigCap_Pin:         Retained. Pin group on which the digital data to be sourced
    ''''        NoOfBists:          DDR LB test consists of individual blocks, suchas lanes, byte.
    ''''        DqSwpNoOfBits:      Data points in SWQ sweep
    ''''        DqsSwpNoOfBits:     Data points in SWK sweep
    ''''        DispCaptStream:     If true the captured data would be displayed as bit stream
    ''''--------------------------------------------------------------------------------------------------
    ''''    NOTE: The impedance settings are to be sourced.
    ''''        They include zcpu, zcpd, dspu, and dspd for each instance. There are 2 instances for Starling
    ''''            Among them the first 3 are to be calibrated based on the test condition and performance mode,
    ''''            and dspd fixed.
    ''''        At the time this function is being developed, it's not clear how those calibration results as well
    ''''            dspd would be passed to this function. For the moment the an assumption is made that those
    ''''            settings are available and assigned. Will use locally defined variables with hard coded settings
    ''''            for them.
    ''''        This will be updated before get in Starling DDR flow based on the actual
    ''''

    Dim pat_count As Long
    Dim i As Long, k As Long, j As Long
    Dim patt_ary() As String
    Dim site As Variant
    Dim patt As Variant
    Dim Pat As String

    Dim EyeStrobes As Long
    Dim EyeStrobes_DQ As Long     '-- 2018_0920 HH
    Dim EyeStrobes_DQS As Long    '-- 2018_0920 HH

    Dim DqSwpWf As New DSPWave, DqsSwpWf As New DSPWave         ' captured sweeep results as well MDLL cal code if applicable
    Dim DqEyeWf As New DSPWave, DqsEyeWf As New DSPWave         ' for starting, the captured wf are Eye wf, no conversion necessary
    Dim EyeWidthWf As New DSPWave

    Dim Ddr0ImpWf As New DSPWave, Ddr1ImpWf As New DSPWave      ' impedance settings to be sourced in
    Dim DdrImpRegWidthWf As New DSPWave                         ' register bit width
    Dim Ddr0ImpDigSrcWf As New DSPWave
    Dim Ddr1ImpDigSrcWf As New DSPWave
    Dim NoOfSrcBits As Long
    Dim repeats As Long
    Dim isIndDataRepeat As Boolean
    Dim isAllDataRepeat As Boolean
    Dim DigSrc_Sample_Size_DQ As Long
    Dim DigSrc_Sample_Size_DQS As Long
    Dim SplitSize() As String
    Dim Testname_CZ_Vt As String: Testname_CZ_Vt = vbNullString
    Dim Instname_split() As String
    Dim TempStr() As String
    Dim p As Long
    Dim BistIdx As Long
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    Dim TName_Ary() As String
    Dim DigSrc_Sample_Size_Long As Long

    'speed up the first run test time
    If Validating_ Then
        Call PrLoadPattern(DqsSwpPat.value)
        Call PrLoadPattern(DqSwpPat.value)
        Exit Function    ' Exit after validation
    End If

    Dim reg_ddr0_zcpu As New SiteLong, reg_ddr0_zcpd As New SiteLong, reg_ddr0_dspu As New SiteLong, reg_ddr0_dspd As New SiteLong
    Dim reg_ddr1_zcpu As New SiteLong, reg_ddr1_zcpd As New SiteLong, reg_ddr1_dspu As New SiteLong, reg_ddr1_dspd As New SiteLong
    
        glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)
        

    'Roger New,20180510 TName
    '--------------------------------------------------------------------
    Call GetFlowTName

'    theexec.Datalog.Setup.DatalogSetup.DisableInstanceNameInPTR = True
'    theexec.Datalog.Setup.DatalogSetup.DisablePinNameInPTR = True
'    theexec.Datalog.ApplySetup
    '----------------------------20180523
'''------------------------------------------------------------------------------------------------------------------------
''' DigSrc setup
'''------------------------------------------------------------------------------------------------------------------------
    Dim DqsInDspWave As New DSPWave '''temp
    Dim DqInDspWave As New DSPWave '''temp
    Dim EyeStrobes_bywidth() As String
    Dim EyeStrobes_DQSbywidth() As String
    Dim Cus_bywidth As Boolean



'/////////////////////////////// Customize Bits width//////////////////////////  add 20180925

    If NoOfBists <> 0 Then
    
        DqSwpNoOfBits = CLng(DqSwpNoOfBits)             ' change type to long for same with original
        DqsSwpNoOfBits = CLng(DqsSwpNoOfBits)
        EyeStrobes = DqSwpNoOfBits / NoOfBists
        EyeStrobes_DQ = DqSwpNoOfBits / NoOfBists     '-- 2018_0920 HH
        EyeStrobes_DQS = DqsSwpNoOfBits / NoOfBists '-- 2018_0920 HH
        Cus_bywidth = False

    Else
        EyeStrobes_bywidth = Split(DqSwpNoOfBits, "+")
        EyeStrobes_DQSbywidth = Split(DqsSwpNoOfBits, "+")
        
        Cus_bywidth = True
        
        
    End If
    
 '///////////////////////////////////////////////////////////////////////////////


    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

    If (DisableConnectPins <> "") Then TheHdw.Digital.Pins(DisableComparePins).Disconnect
    If (DisableComparePins <> "") Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = True

    '' 20170222 - Sweep Vt from SweepVtStr
    Dim SplitByColon() As String
    Dim SourceIndexStr As String, SourceIndex As Long
    Dim StartVal As Double, StepVal As Double, FinalVal As Double
    Dim ReplaceStr() As String

    If AMP_EYE_VT_CZ_Flag = True Then
        If SweepVtStr <> "" Then
        SplitByColon = Split(SweepVtStr, ":")
        SourceIndexStr = SplitByColon(0)
        SourceIndex = TheExec.Flow.var(SourceIndexStr).value
        StartVal = SplitByColon(1)
        StepVal = SplitByColon(2)
        FinalVal = StartVal + SourceIndex * StepVal

            If InStr(UCase(Interpose_PrePat), ":VT:") <> 0 Then
                'NEW 20170731 Purpose to only update VT value and keep the other interpose setting the same
                ReplaceStr = Split(Interpose_PrePat, "VT")
                If InStr(ReplaceStr(1), ";") Then
                    TempStr = Split(ReplaceStr(1), ";")

                    For p = 0 To UBound(TempStr)
                        If p = 0 Then
                            Interpose_PrePat = ReplaceStr(0) & "VT:" & CStr(FinalVal)
                        Else
                            Interpose_PrePat = Interpose_PrePat & ";" & TempStr(p)
                        End If
                    Next p
                Else
                    Interpose_PrePat = ReplaceStr(0) & "VT:" & CStr(FinalVal)
                End If

                'NEW 20170731 For Char TestName
                FinalVal = Format(FinalVal, "0.000")
                Instname_split = Split(glb_TestInstance, "_")
                If FinalVal < 0 Then
                    Testname_CZ_Vt = Replace(CStr(FinalVal), "-", "m")
                Else
                    Testname_CZ_Vt = CStr(FinalVal)
                End If
                Testname_CZ_Vt = Replace(Testname_CZ_Vt, ".", "p")
                Testname_CZ_Vt = "_" & Instname_split(9) & "_" & Instname_split(10) & "_" & "VT" & "_" & Testname_CZ_Vt & "_" & Instname_split(UBound(Instname_split))

            End If
        End If
    End If

    If Interpose_PrePat <> "" Then
        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    End If


'    TName_Ary = Split(gl_Tname_Meas, "+")

    TheHdw.Patterns(DqsSwpPat).Load
    gl_TName_Pat = DqsSwpPat.value
    Call PATT_GetPatListFromPatternSet(DqsSwpPat.value, patt_ary, pat_count)
    ''''add src for ddr ''''''''''''SP 20180221
    Dim Rtn_SweepTestName As String
    Rtn_SweepTestName = vbNullString
    For Each patt In patt_ary
        If DigSrc_Sample_Size <> "" Then
            Dim DqsSwpPat_Str As String
            DqsSwpPat_Str = CStr(patt)
            DigSrc_Sample_Size_Long = CLng(DigSrc_Sample_Size)
            Call GeneralDigSrcSetting(DqsSwpPat_Str, DigSrc_pin, DigSrc_Sample_Size_Long, DigSrc_DataWidth, DqsDigSrc_Equation, digsrc_assignment, _
                                                       DigSrc_FlowForLoopIntegerName, CUS_Str_DigSrcData, DqsInDspWave, Rtn_SweepTestName)
        End If
    Next patt

    '' 1. To get all DigCap bits (Not only bits for eye test) -- TYCHENGG
    If CUS_Str_DigCapData_Dqs <> "" Then
        Dim ShowDec_Dqs As String
        Dim ShowOut_Dqs As String
        Dim DigCapIndex_Dqs As Integer
        Dim DqsDataWf As New DSPWave, DqsTempWf As New DSPWave
        Dim Dqs_DSSC_OUT_Wf(0) As New DSPWave
        Dim Dqs_DSSC_OUT_Full(0) As New DSPWave
        Dim CUS_Sub_Str_DigCapData_Dqs As String
        Dim DqsNewWf As New DSPWave
    End If
    ''----------------------------------------------------

    For Each patt In patt_ary
        Pat = CStr(patt)

        Dim pat_name() As String
        Dim pat_name_module() As String
        Dim Pat_name1() As String

        pat_name_module = Split(Pat, ":")
        pat_name = Split(pat_name_module(0), "\")

        pat_name(0) = pat_name(UBound(pat_name))
        pat_name(0) = Replace(pat_name(0), ".", "_")
        Pat_name1 = Split(glb_TestInstance, "_")

        Call DigCapSetup(Pat, DigCap_Pin, pat_name(0) & "_" & Pat_name1(UBound(Pat_name1)), CLng(DigCap_Sample_Size_Dqs), DqsSwpWf)      'DqsSwpWf = 288

'        Call TheHdw.Patterns(Pat).test(pfAlways, 0)
        If gl_flag_CZ_Nominal_Measured_1st_Point Then: 'Call CZ_TNum_Increment      '20180713 TER add for increaseing FTR TNum @ CZ
        Call TheHdw.Patterns(Pat).test(pfAlways, 0)

        If DebugPrintEnable = True Then TheExec.Datalog.WriteComment "  Pattern(" & pat_count & "): " & Pat & vbNullString
    Next patt

    If TheExec.TesterMode = testModeOffline Then
        For Each site In TheExec.sites
            For i = 0 To DigCap_Sample_Size_Dqs - 1
                DqsSwpWf.Element(i) = FormatNumber(Rnd())
            Next i
        Next site
    End If

    '' 2. To Split DigCapData portion and DqsEyeWf portion -- TYCHENGG
    If CUS_Str_DigCapData_Dqs <> "" Then
        Call DSSC_Special_Str_Filter(CUS_Str_DigCapData_Dqs, EyeTestRegName, DqsSwpWf, _
                                        CUS_Sub_Str_DigCapData_Dqs, DqsTempWf, DqsDataWf)

        ' DqsSwpWf = 288
        ' DqsTempWf = 256
        ' DqsDataWf = 32
        
        
        If Cus_bywidth = False Then                                ' add 20180925
         
            Else
          
                DqsSwpNoOfBits = UBound(EyeStrobes_bywidth) + 1
        
        End If
        
        
        
        For Each site In TheExec.sites

            Dqs_DSSC_OUT_Full(0) = DqsSwpWf
            DqsSwpWf.CreateConstant 0, DqsSwpNoOfBits
            DqsSwpWf = DqsTempWf.Copy          ''
            Dqs_DSSC_OUT_Wf(0) = DqsDataWf

        Next site
        
        ' DqsSwpWf = DqsTempWf = 256
        ' DqsDataWf = Dqs_DSSC_OUT_Wf(0) = 32

        ' Print Out total 288 bits
''        Call HardIP_Digcap_Print_New(CUS_Str_DigCapData_Dqs, Dqs_DSSC_OUT_Full, CLng(DigCap_Sample_Size_Dqs), 0, ShowDec_Dqs, ShowOut_Dqs, , DigCap_DSPWaveSetting)
        Call DigCapDataProcessByDSP(CUS_Str_DigCapData_Dqs, Dqs_DSSC_OUT_Full(0), CLng(DigCap_Sample_Size_Dqs), 0)

    End If
    ''----------------------------------------------------

    For Each site In TheExec.sites
        DqsEyeWf = DqsSwpWf.Copy          '''' the original captured waveform would become stile after DSP functional call
    Next site

    ''''
    '''' DQ sweep
    ''''
    If DqSwpPat <> "" Then

        TheHdw.Patterns(DqSwpPat).Load
        Call PATT_GetPatListFromPatternSet(DqSwpPat.value, patt_ary, pat_count)
        ''''add src for ddr ''''''''''''SP 20180221
        Rtn_SweepTestName = vbNullString
        For Each patt In patt_ary
            If DigSrc_Sample_Size <> "" Then
                Dim DqSwpPat_Str As String
                DqSwpPat_Str = CStr(patt)
                DigSrc_Sample_Size_Long = CLng(DigSrc_Sample_Size)
                Call GeneralDigSrcSetting(DqSwpPat_Str, DigSrc_pin, DigSrc_Sample_Size_Long, DigSrc_DataWidth, DqDigSrc_Equation, digsrc_assignment, _
                                                           DigSrc_FlowForLoopIntegerName, CUS_Str_DigSrcData, DqInDspWave, Rtn_SweepTestName)
            End If
        Next patt

        '' 3. To get all DigCap bits (Not only bits for eye test) -- TYCHENGG
        If CUS_Str_DigCapData_Dq <> "" Then
            Dim ShowDec_Dq As String
            Dim ShowOut_Dq As String
            Dim DigCapIndex_Dq As Integer
            Dim DqDataWf As New DSPWave, DqTempWf As New DSPWave
            Dim Dq_DSSC_OUT_Wf(0) As New DSPWave
            Dim Dq_DSSC_OUT_Full_Wf(0) As New DSPWave
            Dim CUS_Sub_Str_DigCapData_Dq As String
        End If
        ''----------------------------------------------------

        For Each patt In patt_ary
            Pat = CStr(patt)

            Call DigCapSetup(Pat, DigCap_Pin, "Meas_cap", CLng(DigCap_Sample_Size_Dq), DqSwpWf)  ' DqSwpWf = 288

'            Call TheHdw.Patterns(Pat).test(pfAlways, 0)
            If gl_flag_CZ_Nominal_Measured_1st_Point Then: 'Call CZ_TNum_Increment      '20180713 TER add for increaseing FTR TNum @ CZ
            Call TheHdw.Patterns(Pat).test(pfAlways, 0)

            If DebugPrintEnable = True Then TheExec.Datalog.WriteComment "  Pattern(" & pat_count & "): " & Pat & vbNullString
        Next patt


        If TheExec.TesterMode = testModeOffline Then
            For Each site In TheExec.sites
                For i = 0 To DigCap_Sample_Size_Dq - 1
                    DqSwpWf.Element(i) = FormatNumber(Rnd())
                Next i
            Next site
        End If

        '' 4. To Split DigCapData portion and DqsEyeWf portion -- TYCHENGG
        If CUS_Str_DigCapData_Dq <> "" Then
            Call DSSC_Special_Str_Filter(CUS_Str_DigCapData_Dq, EyeTestRegName, DqSwpWf, _
                                         CUS_Sub_Str_DigCapData_Dq, DqTempWf, DqDataWf)

           ' DqSwpWf = 288
           ' CUS_Sub_Str_DigCapData_Dq : Remaining DSSC_OUT String
           ' DqTempWf = 256
           ' DqDataWf = 32
           
        If Cus_bywidth = False Then
         
          Else
          
          DqSwpNoOfBits = UBound(EyeStrobes_bywidth) + 1
        
        End If

            For Each site In TheExec.sites
                Dq_DSSC_OUT_Full_Wf(0) = DqSwpWf
                DqSwpWf.CreateConstant 0, DqSwpNoOfBits
                DqSwpWf = DqTempWf.Copy          ''
                Dq_DSSC_OUT_Wf(0) = DqDataWf
            Next site

'''''            Call HardIP_Digcap_Print_New(CUS_Str_DigCapData_Dq, Dq_DSSC_OUT_Full_Wf, CLng(DigCap_Sample_Size_Dq), 0, ShowDec_Dq, ShowOut_Dq, , DigCap_DSPWaveSetting)   ''
        Call DigCapDataProcessByDSP(CUS_Str_DigCapData_Dq, Dq_DSSC_OUT_Full_Wf(0), CLng(DigCap_Sample_Size_Dq), 0)

        End If

    Else

        'If DqSwpPat no pattern
        For Each site In TheExec.sites
            DqSwpWf.CreateConstant 0, DqsSwpNoOfBits
            Dq_DSSC_OUT_Wf(0) = DqDataWf
        Next site

    End If
    ''----------------------------------------------------

    For Each site In TheExec.sites.Active
        DqEyeWf = DqSwpWf.Copy          '''' the original captured waveform would become stile after DSP functional call
    Next site


    '''' stitch the eye diagrams from 2 sweeps, find and report the eye widths.
    EyeWidthWf.CreateConstant 0, 1
    
    If Cus_bywidth = False Then
    
    Call rundsp.FindMaxEyeWidth_reverse(DqsEyeWf, DqEyeWf, NoOfBists, EyeWidthWf)
    
    Else
    
    Dim Cont_width As New DSPWave  ' add 20180925 for difference width bits
    Dim W As Integer
    
    Cont_width.CreateConstant 0, CLng(DqsSwpNoOfBits), DspLong
    
    For W = 0 To DqsSwpNoOfBits - 1
        
        Cont_width.Element(W) = EyeStrobes_bywidth(W)
        
    Next W
     
     Call rundsp.FindMaxEyeWidth_reverse_bywidth(DqsEyeWf, DqEyeWf, Cont_width, EyeWidthWf) ' add 20180925

    End If
    
   
    
    '                                        256 , 256

    ''''
    '''' Test the eye opening, per lane
    '''' eyewidth: limits from the flow table
    ''''
    
    If NoOfBists = 0 Then
    
       NoOfBists = DqSwpNoOfBits
       
     End If
     

    For BistIdx = 0 To NoOfBists - 1

        If LCase(glb_TestInstance) Like "*cacs*_ck*" Then
            TestNameInput = Report_TName_From_Instance(CalcC, vbNullString, "DDR" & CStr(BistIdx) & "_EYE_CACS_CK" & Testname_CZ_Vt, 0, BistIdx)
            TheExec.Flow.TestLimit resultVal:=EyeWidthWf.Element(BistIdx), Tname:=TestNameInput, ForceResults:=tlForceFlow
        Else
            TestNameInput = Report_TName_From_Instance(CalcC, vbNullString, "DDR" & CStr(BistIdx \ 2) & "_EYE_DQ_DQS" & CStr(BistIdx Mod 2) & Testname_CZ_Vt, 0, BistIdx)
            TheExec.Flow.TestLimit resultVal:=EyeWidthWf.Element(BistIdx), Tname:=TestNameInput, ForceResults:=tlForceFlow
        End If

    Next BistIdx

    If DisableConnectPins <> "" Then TheHdw.Digital.Pins(DisableComparePins).Connect
    If DisableComparePins <> "" Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = False

    ''''
    '''' display the captured information, as well eye diagrams in the captured order
    ''''
  
    If Cus_bywidth = True Then
       
       DqsSwpNoOfBits = CStr(DigCap_Sample_Size_Dqs)
       DqSwpNoOfBits = CStr(DigCap_Sample_Size_Dq)
       DqsSwpNoOfBits = CLng(DqsSwpNoOfBits)
       DqSwpNoOfBits = CLng(DqSwpNoOfBits)
    End If
    
    
    
    If DispCaptStrm Then

       Dim BitStrM As String
       For Each site In TheExec.sites.Active
         
            '''' Dqs sweep
            BitStrM = CStr(DqsSwpWf(site).Element(0))
            For i = 1 To DqsSwpNoOfBits - 1                          ' DqsSwpNoOfBits = 256
                BitStrM = BitStrM & CStr(DqsSwpWf(site).Element(i))  ''DqSwpWf0=>DqSwpWf
            Next i

            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site " & site & ": 1st Sweep " & DqsSwpNoOfBits & " bits(LSB->MSB) = " & BitStrM  'cw

            '''' Dq sweep
            BitStrM = CStr(DqSwpWf(site).Element(0))               ''DqsSwpWf0=>DqsSwpWf
            For i = 1 To DqSwpNoOfBits - 1
                BitStrM = BitStrM & CStr(DqSwpWf(site).Element(i)) ''DqsSwpWf0=>DqsSwpWf
            Next i
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "        2nd Sweep " & DqSwpNoOfBits & " bits(LSB->MSB) = " & BitStrM                     'cw

            '''' Swq eyediagram, per eye sweep

            For BistIdx = 0 To NoOfBists - 1
            
                If Cus_bywidth = True Then
                 
                EyeStrobes = EyeStrobes_bywidth(BistIdx)
                
                End If
                
                
'                EyeStrobes_DQ = EyeStrobes_bywidth(BistIdx)
'                EyeStrobes_DQS = EyeStrobes_bywidth(BistIdx)
                
                Dim EyeSt As Integer
                Dim BistByte As Integer, Ddr As Integer

                BistByte = BistIdx Mod 2        '''' BistByte 0 or 1
                Ddr = BistIdx \ 2                       '''' 0, 1: 0; 2, 3: 1
                
                If Cus_bywidth = True Then
                
                If BistIdx = 0 Then
                   EyeSt = 0
                Else
                    EyeSt = EyeSt + EyeStrobes_bywidth(BistIdx - 1)
                End If
'                EyeSt = BistIdx * EyeStrobes
                End If

                BitStrM = CStr(DqsEyeWf(site).Element(EyeSt))
                For i = 1 To EyeStrobes - 1
                    BitStrM = BitStrM & CStr(DqsEyeWf(site).Element(EyeSt + i))
                Next i

                If gl_Disable_HIP_debug_log = False Then

                    If LCase(glb_TestInstance) Like "*cacs*_ck*" Then
                    ''cw add for Skye
                        TheExec.Datalog.WriteComment "         cacs Eye, DDR" & BistIdx & "eye0" & ": " & BitStrM
                    Else
                        TheExec.Datalog.WriteComment "         dq    Eye, DDR" & CInt(BistIdx \ 2) & "eye" & (BistIdx Mod 2) & ": " & BitStrM
                    End If
                End If
                    
                BitStrM = CStr(DqEyeWf(site).Element(EyeSt))
                For i = 1 To EyeStrobes - 1
                    BitStrM = BitStrM & CStr(DqEyeWf(site).Element(EyeSt + i))
                Next i

                If gl_Disable_HIP_debug_log = False Then
                    If LCase(glb_TestInstance) Like "*cacs*_ck*" Then
                    ''cw add for Skye
                        TheExec.Datalog.WriteComment "         ck   Eye, DDR" & BistIdx & "eye0" & ": " & BitStrM
                    Else
                        TheExec.Datalog.WriteComment "         dqs   Eye, DDR" & CInt(BistIdx \ 2) & "eye" & (BistIdx Mod 2) & ": " & BitStrM
                    End If
                End If

            Next BistIdx
        Next site

    End If


    For Each site In TheExec.sites
        DqsSwpWf.CreateConstant 0, DqsSwpNoOfBits
    Next site

    For Each site In TheExec.sites
        DqsSwpWf.CreateConstant 0, DqsSwpNoOfBits
    Next site

    Pat = DqsSwpPat.value & "," & DqSwpPat.value
    Shmoo_Pattern = DqsSwpPat.value & "," & DqSwpPat.value
    DebugPrintFunc Pat

     '' 20170712 - Process calculate equation by dictionary.
     If Calc_Eqn <> "" Then
         Call ProcessCalcEquation(Calc_Eqn)
     End If

    If Interpose_PrePat <> "" Then
        Call SetForceCondition("RESTOREPREPAT")
    End If
    
Exit Function


errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP", "Opt_DdrLpBkFunc2") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function


Public Function Eye_Diagram(LaneNumber As Long) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
Dim i, j, k As Long
Dim Eye_Diagram_Binary_Lane_Temp(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane() As EyeDiagram
ReDim Eye_Diagram_Binary_Lane(LaneNumber - 1) As EyeDiagram
Dim First_src_code As New SiteVariant
Dim End_src_code As New SiteVariant
Dim First_src_code_Temp As New SiteVariant
Dim End_src_code_Temp As New SiteVariant
Dim WithinEye As Boolean
Dim vertical_width As New SiteVariant
Dim Zero_counter As New SiteVariant
Dim horizontal_width As New SiteVariant
Dim Temp_counter As Integer
Dim timing_res_start As New SiteVariant
Dim timing_res_end As New SiteVariant
Dim timing_res_start_temp As Long
Dim timing_res_end_temp As Long
Dim x As Integer
Dim iLen As Integer
Dim Temp_Counter_Act As Long
Dim Total_Zero_Count As Long
Dim TestNameInput As String

glb_TestInstance = vbNullString
glb_TestInstance = UCase(TheExec.DataManager.instancename)
        
Dim vertical_volt As New SiteVariant
Dim horizontal_period As New SiteVariant
Dim TmpString As String
Dim site As Variant 'Carter, 20240304
For k = 0 To LaneNumber - 1
    For Each site In TheExec.sites
        For i = -31 To 31
       ' Eye_Diagram_Binary_Lane_Temp(i + 31)(site) = ""
        TmpString = vbNullString
            For j = 1 To 32
                TmpString = TmpString & mid(Eye_Diagram_Binary(i + 31)(site), LaneNumber * j - (LaneNumber - (k + 1)), 1)
            Next j
'        If k = LaneNumber - 1 Then: Eye_Diagram_Binary_Lane_Temp(i + 31)(Site) = ""
      '  Eye_Diagram_Binary_Lane_Temp(i + 31)(site) = TmpString
        Eye_Diagram_Binary_Lane(k).value(i + 31)(site) = TmpString 'Eye_Diagram_Binary_Lane_Temp
        Next i
        
    Next site
    
    
   '  Eye_Diagram_Binary_Lane(k).Value = Eye_Diagram_Binary_Lane_Temp
Next k
For k = 0 To LaneNumber - 1
    vertical_width = 0
    First_src_code = 0
    End_src_code = 0
    horizontal_width = 0
    timing_res_start = 0
    timing_res_end = 0
    Zero_counter = 0
    
    For Each site In TheExec.sites
        iLen = Len(Eye_Diagram_Binary_Lane(k).value(0)(site)) - 1
        'process   the  Max Zero horizontal
        For i = -31 To 31
            Temp_counter = 0
            Total_Zero_Count = 0
            Temp_Counter_Act = 0
            timing_res_start_temp = 0
            timing_res_end_temp = 0
            For x = 0 To iLen
                If mid(Eye_Diagram_Binary_Lane(k).value(i + 31)(site), iLen - x + 1, 1) = 0 Then
                    Temp_counter = Temp_counter + 1
                    Total_Zero_Count = Total_Zero_Count + 1
                    If x = iLen And Temp_counter > Temp_Counter_Act Then: Temp_Counter_Act = Temp_counter
                ElseIf mid(Eye_Diagram_Binary_Lane(k).value(i + 31)(site), iLen - x + 1, 1) = 1 Then
                    If Temp_counter > Temp_Counter_Act Then
                        Temp_Counter_Act = Temp_counter
                        timing_res_end_temp = 32 - (x - Temp_Counter_Act + 1)
                        timing_res_start_temp = timing_res_end_temp - Temp_Counter_Act + 1
                    End If
                    Temp_counter = 0
                Else
                'Do nothing
                End If
            Next x
                If horizontal_width < Temp_Counter_Act Then
                   horizontal_width = Temp_Counter_Act
                   timing_res_end = timing_res_end_temp
                   timing_res_start = timing_res_start_temp
                End If
                
                
        If InStr(glb_TestInstance, "G4") <> 0 Then
            horizontal_period = horizontal_width * 2 * (1 / 16000000000#) / 64
        ElseIf InStr(glb_TestInstance, "G3") <> 0 Then
            horizontal_period = horizontal_width * 2 * (1 / 8000000000#) / 64
        ElseIf InStr(glb_TestInstance, "G2") <> 0 Then
            horizontal_period = horizontal_width * 2 * (1 / 5000000000#) / 64
        ElseIf InStr(glb_TestInstance, "G1") <> 0 Then
            horizontal_period = horizontal_width * 2 * (1 / 2500000000#) / 64
        Else
            horizontal_period = 0
        End If
        
        If InStr(glb_TestInstance, "LPDPRX") <> 0 Then
            If InStr(glb_TestInstance, "RX6F3") <> 0 Then
                horizontal_period = horizontal_width * 2 * (0.00000000011) / 64
            ElseIf InStr(glb_TestInstance, "RX6F2") <> 0 Then
                horizontal_period = horizontal_width * 2 * (1 / 6400000000#) / 64
            ElseIf InStr(glb_TestInstance, "RX6F1") <> 0 Then
                horizontal_period = horizontal_width * 2 * (1 / 2133000000#) / 64
            Else
                horizontal_period = 0
            End If
        End If

                
                If Zero_counter < Total_Zero_Count Then: Zero_counter = Total_Zero_Count
        Next i
'============================= Vertical Width Process Start (In Site Loop) ================================' ''20180918 -- TYCHENGG
        For x = 0 To iLen
            First_src_code_Temp = 31
            End_src_code_Temp = 31
            vertical_width = 0
            WithinEye = False
            For i = -31 To 31 '' -31,31 can be defined as constant
                If WithinEye = False Then
                    If mid(Eye_Diagram_Binary_Lane(k).value(i + 31)(site), x + 1, 1) = 0 Then
                        WithinEye = True
                        First_src_code_Temp = i
                    End If
                Else
                    If mid(Eye_Diagram_Binary_Lane(k).value(i + 31)(site), x + 1, 1) = 1 Then
                        WithinEye = False
                        End_src_code_Temp = i - 1
                        vertical_width = Abs(End_src_code_Temp - First_src_code_Temp)
                        If vertical_width > Abs(End_src_code - First_src_code) Then
                            First_src_code = First_src_code_Temp
                            End_src_code = End_src_code_Temp
                        End If
                    End If
                End If
            Next i
            vertical_width = Abs(End_src_code_Temp - First_src_code_Temp)
            If vertical_width > Abs(End_src_code - First_src_code) Then
                First_src_code = First_src_code_Temp
                End_src_code = End_src_code_Temp
            End If
        Next x
        vertical_width = Abs(End_src_code - First_src_code) + 1
        

        
        If InStr(glb_TestInstance, "G4") <> 0 Then
            vertical_volt = vertical_width * 6.8 * 0.75 * 1.75 * 0.8
        ElseIf InStr(glb_TestInstance, "G3") <> 0 Then
            vertical_volt = vertical_width * 6.8 * 0.75 * 1.5 * 0.8
        ElseIf InStr(glb_TestInstance, "G2") <> 0 Then
            vertical_volt = vertical_width * 6.8 * 0.75 * 1# * 0.8
        ElseIf InStr(glb_TestInstance, "G1") <> 0 Then
            vertical_volt = vertical_width * 6.8 * 0.75 * 0.5 * 0.8
        Else
            vertical_volt = 0
        End If
        
        
        If InStr(glb_TestInstance, "LPDPRX") <> 0 Then
            If InStr(glb_TestInstance, "RX6F3") <> 0 Then
                vertical_volt = vertical_width * 6.8 * 0.75 * 1.75 * 0.8
            ElseIf InStr(glb_TestInstance, "RX6F2") <> 0 Then
                vertical_volt = vertical_width * 6.8 * 0.75 * 1.75 * 0.8
            ElseIf InStr(glb_TestInstance, "RX6F1") <> 0 Then
                vertical_volt = vertical_width * 6.8 * 0.75 * 1.75 * 0.8
            Else
                vertical_volt = 0
            End If
        End If
        
        
        
        
'horizontal_Volt
'============================= Vertical Width Process End   (In Site Loop) ================================' ''20180918 -- TYCHENGG
'//////////////////////// for all 1 eye by csho/////////////////
        If vertical_width = "" Then: vertical_width = 0
        If First_src_code = "" Then: First_src_code = 0
        If End_src_code = "" Then: End_src_code = 0
        If horizontal_width = "" Then: horizontal_width = 0
        If timing_res_end = "" Then: timing_res_end = 0
        If timing_res_start = "" Then: timing_res_start = 0
        If Zero_counter = "" Then: Zero_counter = 0
'/////////////////////////////////////////////////////////////////////////
    Next site
    
    Dim PhyName As String
    If InStr(glb_TestInstance, "PCIE") <> 0 Then 'Staten @CW
        ''' Update for Crete @William 210811
        Select Case k:
            Case 0: PhyName = "LN0"
            Case 1: PhyName = "LN1"
            Case 2: PhyName = "LN2"
            Case 3: PhyName = "LN3"
            
'            Case 0: PhyName = "ST-LN0"
'            Case 1: PhyName = "ST-LN1"
'            Case 2: PhyName = "GP-LN0"
'            Case 3: PhyName = "GP-LN1"
'            Case 4: PhyName = "GP-LN2"
'            Case 5: PhyName = "GP-LN3"
            Case Else
        End Select
        
    ElseIf InStr(glb_TestInstance, "CIO") <> 0 Then 'Staten @CW
        Select Case k:
            Case 0: PhyName = "ATC0-LN0"
            Case 1: PhyName = "ATC0-LN1"
            Case 2: PhyName = "ATC1-LN0"
            Case 3: PhyName = "ATC1-LN1"
            Case Else
        End Select
    ElseIf InStr(glb_TestInstance, "CIORX") <> 0 Then
        If InStr(glb_TestInstance, "UELB") <> 0 Then
            Select Case k:
                Case 0: PhyName = "AUSB-LN0"
                Case 1: PhyName = "AUSB-LN1"
                Case Else
            End Select
        Else
            Select Case k:
                Case 0: PhyName = "ATC3-LN0"
                Case 1: PhyName = "ATC2-LN0"
                Case 2: PhyName = "ATC1-LN0"
                Case 3: PhyName = "ATC0-LN0"
                Case 4: PhyName = "ATC3-LN1"
                Case 5: PhyName = "ATC2-LN1"
                Case 6: PhyName = "ATC1-LN1"
                Case 7: PhyName = "ATC0-LN1"
                Case Else
            End Select
        End If
    ElseIf InStr(glb_TestInstance, "AUSRX") <> 0 Then
        Select Case k:
            Case 0: PhyName = "ST1-L0"
            Case 1: PhyName = "GP-L0"
            Case 2: PhyName = "ST0-L0"
            Case 3: PhyName = "ST1-L1"
            Case 4: PhyName = "GP-L1"
            Case 5: PhyName = "ST0-L1"
            Case 6: PhyName = "ST1-L2"
            Case 7: PhyName = "GP-L2"
            Case 8: PhyName = "ST0-L2"
            Case 9: PhyName = "ST1-L3"
            Case 10: PhyName = "GP-L3"
            Case 11: PhyName = "ST0-L3"
            Case Else
        End Select
    Else
        PhyName = "Lane " & k
        
    End If
        
        
        TestNameInput = Report_TName_From_Instance("C", "JTAGTDO", "vertical-width-" & PhyName, ForceResult:=tlForceNone)
        TheExec.Flow.TestLimit resultVal:=vertical_width, Tname:=TestNameInput
        TestNameInput = Report_TName_From_Instance("C", "JTAGTDO", "vertical-width-Start-" & PhyName, ForceResult:=tlForceNone)
        TheExec.Flow.TestLimit resultVal:=First_src_code, Tname:=TestNameInput
        TestNameInput = Report_TName_From_Instance("C", "JTAGTDO", "width-End-" & PhyName, ForceResult:=tlForceNone)
        TheExec.Flow.TestLimit resultVal:=End_src_code, Tname:=TestNameInput
        TestNameInput = Report_TName_From_Instance("C", "JTAGTDO", "horizontal-width-" & PhyName, ForceResult:=tlForceNone)
        TheExec.Flow.TestLimit resultVal:=horizontal_width, Tname:=TestNameInput
        TestNameInput = Report_TName_From_Instance("C", "JTAGTDO", "horizontal-width-Start-" & PhyName, ForceResult:=tlForceNone)
        TheExec.Flow.TestLimit resultVal:=timing_res_start, Tname:=TestNameInput
        TestNameInput = Report_TName_From_Instance("C", "JTAGTDO", "horizontal-width-End-" & PhyName, ForceResult:=tlForceNone)
        TheExec.Flow.TestLimit resultVal:=timing_res_end, Tname:=TestNameInput
        TestNameInput = Report_TName_From_Instance("C", "JTAGTDO", "Max-Zero-" & PhyName, ForceResult:=tlForceNone)
        TheExec.Flow.TestLimit resultVal:=Zero_counter, Tname:=TestNameInput
        
        If InStr(glb_TestInstance, "PCIE") <> 0 Then 'Staten @CW
            TestNameInput = Report_TName_From_Instance("C", "JTAGTDO", "vertical-Voltage-" & PhyName, ForceResult:=tlForceNone)
            TheExec.Flow.TestLimit resultVal:=vertical_volt.Multiply(0.001), Tname:=TestNameInput, unit:=unitVolt
            TestNameInput = Report_TName_From_Instance("C", "JTAGTDO", "horizontal-Period-" & PhyName, ForceResult:=tlForceNone)
            TheExec.Flow.TestLimit resultVal:=horizontal_period, Tname:=TestNameInput, unit:=unitTime
        ElseIf glb_TestInstance Like "*CIORX*" Then
        
            TestNameInput = Report_TName_From_Instance("C", "JTAGTDO", "vertical-Voltage-" & PhyName, ForceResult:=tlForceNone)
            TheExec.Flow.TestLimit resultVal:=vertical_volt.Multiply(0.001), lowVal:=0.15, Tname:=TestNameInput, unit:=unitVolt
            TestNameInput = Report_TName_From_Instance("C", "JTAGTDO", "horizontal-Period-" & PhyName, ForceResult:=tlForceNone)
            TheExec.Flow.TestLimit resultVal:=horizontal_period, lowVal:=0.000000000015, Tname:=TestNameInput, unit:=unitTime
        ElseIf glb_TestInstance Like "*AUSRXA*" Then
            TestNameInput = Report_TName_From_Instance("C", "JTAGTDO", "vertical-Voltage-" & PhyName, ForceResult:=tlForceNone)
            TheExec.Flow.TestLimit resultVal:=vertical_volt.Multiply(0.001), lowVal:=0.2, Tname:=TestNameInput, unit:=unitVolt
            TestNameInput = Report_TName_From_Instance("C", "JTAGTDO", "horizontal-Period-" & PhyName, ForceResult:=tlForceNone)
            TheExec.Flow.TestLimit resultVal:=horizontal_period, lowVal:=0.00000000002, Tname:=TestNameInput, unit:=unitTime
        
        ElseIf glb_TestInstance Like "*LPDPRX*" Then
            TestNameInput = Report_TName_From_Instance("C", "JTAGTDO", "vertical-Voltage-" & PhyName, ForceResult:=tlForceNone)
            TheExec.Flow.TestLimit resultVal:=vertical_volt.Multiply(0.001), lowVal:=0.25, Tname:=TestNameInput, unit:=unitVolt
            TestNameInput = Report_TName_From_Instance("C", "JTAGTDO", "horizontal-Period-" & PhyName, ForceResult:=tlForceNone)
            TheExec.Flow.TestLimit resultVal:=horizontal_period, lowVal:=0.000000000025, Tname:=TestNameInput, unit:=unitTime
        Else
            TestNameInput = Report_TName_From_Instance("C", "JTAGTDO", "vertical-Voltage-" & PhyName, ForceResult:=tlForceNone)
            TheExec.Flow.TestLimit resultVal:=vertical_volt.Multiply(0.001), Tname:=TestNameInput, unit:=unitVolt
            TestNameInput = Report_TName_From_Instance("C", "JTAGTDO", "horizontal-Period-" & PhyName, ForceResult:=tlForceNone)
            TheExec.Flow.TestLimit resultVal:=horizontal_period, Tname:=TestNameInput, unit:=unitTime
        End If
        
        
Next k
For k = 0 To LaneNumber - 1


    If InStr(glb_TestInstance, "PCIE") <> 0 Then 'Staten @CW
        ''' Update for Crete @William 210811
        Select Case k:
            Case 0: PhyName = "LN0"
            Case 1: PhyName = "LN1"
            Case 2: PhyName = "LN2"
            Case 3: PhyName = "LN3"
            Case Else
            
'            Case 0: PhyName = "ST-LN0"
'            Case 1: PhyName = "ST-LN1"
'            Case 2: PhyName = "GP-LN0"
'            Case 3: PhyName = "GP-LN1"
'            Case 4: PhyName = "GP-LN2"
'            Case 5: PhyName = "GP-LN3"
        End Select
    ElseIf InStr(glb_TestInstance, "CIO") <> 0 Then 'Staten @CW
        Select Case k:
            Case 0: PhyName = "ATC0-LN0"
            Case 1: PhyName = "ATC0-LN1"
            Case 2: PhyName = "ATC1-LN0"
            Case 3: PhyName = "ATC1-LN1"
            Case Else
        End Select
    ElseIf InStr(glb_TestInstance, "CIORX") <> 0 Then
        If InStr(glb_TestInstance, "UELB") <> 0 Then
            Select Case k:
                Case 0: PhyName = "AUSB-LN0"
                Case 1: PhyName = "AUSB-LN1"
                Case Else
            End Select
        Else
            Select Case k:
                Case 0: PhyName = "ATC3-LN0"
                Case 1: PhyName = "ATC2-LN0"
                Case 2: PhyName = "ATC1-LN0"
                Case 3: PhyName = "ATC0-LN0"
                Case 4: PhyName = "ATC3-LN1"
                Case 5: PhyName = "ATC2-LN1"
                Case 6: PhyName = "ATC1-LN1"
                Case 7: PhyName = "ATC0-LN1"
                Case Else
            End Select
        End If
    ElseIf InStr(glb_TestInstance, "AUSRX") <> 0 Then
        Select Case k:
            Case 0: PhyName = "ST1-L0"
            Case 1: PhyName = "GP -L0"
            Case 2: PhyName = "ST0-L0"
            Case 3: PhyName = "ST1-L1"
            Case 4: PhyName = "GP -L1"
            Case 5: PhyName = "ST0-L1"
            Case 6: PhyName = "ST1-L2"
            Case 7: PhyName = "GP -L2"
            Case 8: PhyName = "ST0-L2"
            Case 9: PhyName = "ST1-L3"
            Case 10: PhyName = "GP -L3"
            Case 11: PhyName = "ST0-L3"
            Case Else
        End Select
    
    Else
        PhyName = "Lane " & k
    End If

    For Each site In TheExec.sites
        For i = -31 To 31

                If i <= -10 Then
                    Call TheExec.Datalog.WriteComment("[Eye Diagram, " & HramLotId(site) & "-" & CStr(HramWaferId(site)) & ", " & CStr(XCoord(site)) & ", " & CStr(YCoord(site)) & ", " & site & ", " & Eye_Diagram_Binary_Lane(k).value(i + 31)(site) & ", h0dac_off:" & i & ", " & PhyName & ", " & glb_TestInstance & "]")
                ElseIf -10 < i And i < 0 Then
                    Call TheExec.Datalog.WriteComment("[Eye Diagram, " & HramLotId(site) & "-" & CStr(HramWaferId(site)) & ", " & CStr(XCoord(site)) & ", " & CStr(YCoord(site)) & ", " & site & ", " & Eye_Diagram_Binary_Lane(k).value(i + 31)(site) & ", h0dac_off: " & i & ", " & PhyName & ", " & glb_TestInstance & "]")
                ElseIf 0 <= i And i < 10 Then
                    Call TheExec.Datalog.WriteComment("[Eye Diagram, " & HramLotId(site) & "-" & CStr(HramWaferId(site)) & ", " & CStr(XCoord(site)) & ", " & CStr(YCoord(site)) & ", " & site & ", " & Eye_Diagram_Binary_Lane(k).value(i + 31)(site) & ", h0dac_off:  " & i & ", " & PhyName & ", " & glb_TestInstance & "]")
                ElseIf i >= 10 Then
                    Call TheExec.Datalog.WriteComment("[Eye Diagram, " & HramLotId(site) & "-" & CStr(HramWaferId(site)) & ", " & CStr(XCoord(site)) & ", " & CStr(YCoord(site)) & ", " & site & ", " & Eye_Diagram_Binary_Lane(k).value(i + 31)(site) & ", h0dac_off: " & i & ", " & PhyName & ", " & glb_TestInstance & "]")
                Else
                'Do nothing
                End If
        Next i
    Next site
Next k


    ' Check implicit alarms
    TheHdw.Alarms.Check


Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP", "Eye_Diagram") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function


Public Function TrimCodeDig(Optional patset As Pattern, Optional TestSequence As String, Optional CPUA_Flag_In_Pat As Boolean, _
    Optional MeasureF_Pin As PinList, Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As String, Optional DigSrc_Equation As String, Optional digsrc_assignment As String, _
    Optional DigCap_Pin As PinList, Optional CUS_Str_DigCapData As String, Optional DigCap_Sample_Size As Long, _
    Optional TrimTarget As Double, Optional TrimTargetTolerance As Double = 0, Optional TrimStart As String, Optional TrimFormat As String, Optional TrimBitSize As Long, _
    Optional TrimStoreName As String, Optional TrimFuseTypeName As String, Optional Interpose_PrePat As String, Optional Interpose_PostTest As String, _
    Optional TrimPrcocessAll As Boolean = False, Optional UseMinimumTrimCode As Boolean = False, Optional PreCheckMinMaxTrimCode As Boolean = False, _
    Optional MSB_First_Flag As Boolean, Optional Validating_ As Boolean) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    '[20231207][All][Neil] Update DigSrcAssignment/DigSrc_Equ for flexible usage, remove TrimFormat redunant code
    Dim DSPWF_TrimCodeBin As New DSPWave
    Dim DSPWF_TrimCodeBin_Rev As New DSPWave        'nn_test_1212
    Dim DSPWF_TrimCodeDec As New DSPWave
    DSPWF_TrimCodeDec.CreateConstant 0, 1, DspLong
    
    Dim PatCount As Long, PattArray() As String
    Dim OutDspWave As New DSPWave
    TrimTarget = 0.3 ''Just for Flag    2017/Dec
    
    If Validating_ Then
        Call PrLoadPattern(patset.value)
        Exit Function    ' Exit after validation
    End If
    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)
        
    Call HardIP_InitialSetupForPatgen
    Dim Ts As Variant, TestSequenceArray() As String
    Dim InitialDSPWave As New DSPWave, PastDSPWave As New DSPWave, InDSPWave As New DSPWave

    Dim site As Variant
    Dim Pat As String
    Dim i As Long, j As Long, k As Long

    'Dim Inst_Name_Str As String: Inst_Name_Str = glb_TestInstance
   
    Dim l_DigSrc_Sample_Size As Long
    If InStr(DigSrc_Sample_Size, "|") <> 0 Then
        l_DigSrc_Sample_Size = Split(DigSrc_Sample_Size, "|")(1)
    Else
        l_DigSrc_Sample_Size = CLng(DigSrc_Sample_Size)
    End If
    
    '=== Update DigSrcAssignment/DigSrc_Equ for flexible usage === 20231207
    Call Reg_Assign_Processing(DigSrc_Equation, digsrc_assignment, CUS_Str_DigCapData)
    Call ProcessInputToGLB(patset, TestSequence, , , , , , , , , , , , , , , , DigCap_Pin, , DigCap_Sample_Size, DigSrc_pin, CStr(DigSrc_Sample_Size), CStr(DigSrc_Sample_Size), DigSrc_Equation, digsrc_assignment, , , , , , , , , , , , , , , , , , , , Interpose_PrePat, , Interpose_PostTest)
    InDSPWave.CreateConstant 0, l_DigSrc_Sample_Size, DspLong
    '=== Update DigSrcAssignment/DigSrc_Equ for flexible usage ===
    
    Dim CapValue As New PinListData, CapValue_V1 As New PinListData, CapValue_V2 As New PinListData
    CapValue.AddPin ("CapValueString")

    Call GetFlowTName
 
    TestSequenceArray = Split(TestSequence, ",")
    TheHdw.Digital.Patgen.Halt
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    If Interpose_PrePat <> "" Then ''''180109 update
        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    End If
    
    TheHdw.Patterns(patset).Load

    'Call PATT_GetPatListFromPatternSet(PatSet.value, PattArray, PatCount)
    Dim tempPatArr() As String
    Dim tempPatCnt As Long
    Dim tempSrcSigName As String
    Dim tempSrcPatSeq As Long
    tempSrcPatSeq = 1

    Call PatternBurstCheckAndSplit(patset.value, PattArray, PatCount)
    If BurstYesPatDict.Exists(LCase(patset.value)) = True Then
        Call PATT_GetPatListFromPatternSet(patset.value, tempPatArr, tempPatCnt)
        tempSrcSigName = Split(tempPatArr(tempSrcPatSeq), ":")(1)
    End If
    
    '' 20160425 - Check format from TrimFormat
    Dim StrSeparatebyComma() As String
    Dim ExecutionMax As Long
    StrSeparatebyComma = Split(TrimFormat, ";")
    ExecutionMax = UBound(StrSeparatebyComma)
    Dim StrSeparatebyEqual() As String, StrSeparatebyColon() As String '' Get Src bit
    Dim SrcStartBit As Long, SrcEndBit As Long
    
    Dim b_HigherThanTarget As New SiteBoolean
    b_HigherThanTarget = False
    
    Dim LastSectionV1V2_Index As Long
    LastSectionV1V2_Index = 0
    Dim OutputTrimCode As String, TestNameInput As String
    Dim SourceTrimCode As String
    Dim TestLimitIndex As Long '
    
'    Dim TrimStart_1st() As String
    Dim Dec_TrimStart_1st As Long
    
    Dim StoredTargetTrimCode As New DSPWave
    Dim b_MatchTagetCap As New SiteBoolean
    Dim b_DisplayCap As New SiteBoolean
    b_MatchTagetCap = False
    b_DisplayCap = False
    
    Dim StoreEachTrimCode() As New DSPWave
    Dim DigCapData() As New PinListData
    'StoredTargetTrimCode.CreateConstant 0, l_DigSrc_Sample_Size, DspLong
    'ReDim StoreEachTrimCode(l_DigSrc_Sample_Size + 1) As New DSPWave
    'ReDim DigCapData(l_DigSrc_Sample_Size + 1) As New PinListData
    
    'Update for DigSrcAssignment/DigSrc_Equ flexible usage -- 20231207
    StoredTargetTrimCode.CreateConstant 0, l_DigSrc_Sample_Size, DspLong
    ReDim StoreEachTrimCode(l_DigSrc_Sample_Size + 1) As New DSPWave
    ReDim DigCapData(l_DigSrc_Sample_Size + 1) As New PinListData
    
    Dim StoreEachIndex As Long
        
    ''20161128-Stop trim code process
    Dim b_StopTrimCodeProcess As New SiteBoolean
    b_StopTrimCodeProcess = False
    
    For i = 0 To UBound(StoreEachTrimCode)
        'StoreEachTrimCode(i).CreateConstant 0, l_DigSrc_Sample_Size, DspLong
        StoreEachTrimCode(i).CreateConstant 0, l_DigSrc_Sample_Size, DspLong      'Update for DigSrcAssignment/DigSrc_Equ flexible usage -- 20231207
    Next i

    If TrimStart <> "" And TrimStart Like "*&*" Then
        TrimStart = Replace(TrimStart, "&", vbNullString)
    End If

    Dim InDspWave_Temp_Rev As New DSPWave
    Dim InDspWave_Temp As New DSPWave
    Dim TrimFormat_RegSize As Long
    Dim TrimFormat_StoreName As String
    Dim TrimFormat_Ary() As String
    If DigSrc_Equation <> "" Or digsrc_assignment <> "" Then
        If InStr(TrimFormat, "=") <> 0 Then
            TrimFormat_Ary = Split(TrimFormat, "=")
            TrimFormat_StoreName = TrimFormat_Ary(0)
            TrimFormat_RegSize = Split(TrimFormat_StoreName, "_")(1)
        End If
    End If
'' ===============================   TrimStart_1st = TrimStart===============================================
    'Dec_TrimStart_1st = Bin2Dec(TrimStart)
    
    'InitialDSPWave.CreateConstant Dec_TrimStart_1st, 1, DspLong

'    Call rundsp.CreateFlexibleDSPWave(InitialDSPWave, l_DigSrc_Sample_Size, InDSPWave)

    'Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeCap", DigSrc_Sample_Size, InDSPWave)
'    If BurstYesPatDict.Exists(LCase(patset.value)) = True Then
'        Call SetupDigSrcDspWave(tempPatArr(tempSrcPatSeq), DigSrc_pin, tempSrcSigName, l_DigSrc_Sample_Size, InDSPWave)
'    Else
'        Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeCap", l_DigSrc_Sample_Size, InDSPWave)
'    End If

    '=== Update for DigSrcAssignment/DigSrc_Equ flexible usage === 20231207
    DSPWF_TrimCodeBin.CreateConstant 0, TrimFormat_RegSize, DspLong
    For Each site In TheExec.sites
        For i = 1 To Len(TrimStart)
            DSPWF_TrimCodeBin.Element(i - 1) = mid(TrimStart, i, 1) ' bit order
        Next i
    Next site
'    InDSPWave.CreateConstant 0, TrimFormat_RegSize, DspLong
    'Call HardIP_Dec2Bin(DSPWF_TrimCodeBin, DSPWF_TrimCodeDec, TrimBitSize)
    If MSB_First_Flag = True Then
        Call rundsp.Split_Dspwave_Reverse(DSPWF_TrimCodeBin, DSPWF_TrimCodeBin_Rev)
        Call StoreDataAllType(TrimFormat_StoreName, DSPWF_TrimCodeBin_Rev)
    Else
        Call StoreDataAllType(TrimFormat_StoreName, DSPWF_TrimCodeBin)
    End If
    
    Call GeneralDigSrcSettingWithBurst(LCase(patset.value), DigSrc_pin, InDSPWave)
    '=== Update DigSrcAssignment/DigSrc_Equ for flexible usage
    
    TheExec.Datalog.WriteComment ("========First Time Setup========")
    
    For Each site In TheExec.sites
        SourceTrimCode = vbNullString
'        For k = 0 To InDSPWave(site).sampleSize - 1
'            SourceTrimCode = SourceTrimCode & CStr(InDSPWave(site).Element(k))
'        Next k
'        TheExec.Datalog.WriteComment ("Site_" & site & " Initial Source Trim Code = " & SourceTrimCode)
'        StoreEachTrimCode(0)(site) = InDSPWave(site).Copy
        
        'Use DigSrcAssignment/DigSrc_Equ for flexible trim usage -- 20231207
        For k = 0 To DSPWF_TrimCodeBin(site).sampleSize - 1
            SourceTrimCode = SourceTrimCode & CStr(DSPWF_TrimCodeBin(site).Element(k))
        Next k
        TheExec.Datalog.WriteComment ("Site_" & site & " Initial Source Trim Code = " & SourceTrimCode)
        StoreEachTrimCode(0)(site) = DSPWF_TrimCodeBin(site).Copy
    Next site
    
'    For Each site In TheExec.sites
'        StoreEachTrimCode(0)(site) = InDSPWave(site).Copy
'    Next site
    
'=========Set Up DigCap parameter=============================
    'Call GeneralDigCapSetting(PattArray(0), DigCap_Pin, DigCap_Sample_Size, OutDspWave)
    If BurstYesPatDict.Exists(LCase(patset.value)) = True Then
        Call GeneralDigCapSetting(tempPatArr(tempSrcPatSeq), DigCap_Pin, DigCap_Sample_Size, OutDspWave)
    Else
        Call GeneralDigCapSetting(PattArray(0), DigCap_Pin, DigCap_Sample_Size, OutDspWave)
    End If
    'Call TheHdw.Patterns(PattArray(0)).start
    Call TheHdw.Patterns(patset).start
     'For Each Site In TheExec.sites
    CapValue.value = OutDspWave.Element(0)
     'Next Site
    Call PrintDigCapSetting(DigCap_Pin, DigCap_Sample_Size, CUS_Str_DigCapData)
    ''Update Interpose_PreMeas 20170801
    Dim TestSeqNum As Integer
    TestSeqNum = 0
           
    TheHdw.Digital.Patgen.HaltWait
'    For Each Site In TheExec.sites
        DigCapData(0) = CapValue
        b_HigherThanTarget = CapValue.Math.Subtract(TrimTarget).compare(GreaterThan, 0)
        'PastDSPWave = InDSPWave
        PastDSPWave = DSPWF_TrimCodeBin  'Use DigSrcAssignment/DigSrc_Equ for flexible trim usage -- 20231207
        TestNameInput = "Cap_Value_"
        TestLimitIndex = 0
'    Next Site

    Dim CUS_Str_MainProgram As String: CUS_Str_MainProgram = vbNullString

    
    For Each site In TheExec.sites
        If b_MatchTagetCap(site) = False And b_DisplayCap(site) = False Then
            
            TheExec.Datalog.WriteComment ("Site " & site & " Output CapValue = " & OutDspWave(site).Element(0))
            
        End If
    Next site
    
    b_MatchTagetCap = CapValue.Math.Subtract(TrimTarget).Abs.compare(LessThanOrEqualTo, TrimTargetTolerance)
    b_DisplayCap = b_DisplayCap.LogicalOr(b_MatchTagetCap)
    
    For Each site In TheExec.sites
        If b_DisplayCap(site) = True And StoredTargetTrimCode(site).CalcSum = 0 Then
            'StoredTargetTrimCode(site) = InDSPWave(site).Copy
            
            'Use DigSrcAssignment/DigSrc_Equ for flexible trim usage -- 20231207
            StoredTargetTrimCode(site) = DSPWF_TrimCodeBin(site).Copy
            b_StopTrimCodeProcess(site) = True
        End If
    Next site
    
    TheExec.Datalog.WriteComment ("======================================================================================")
    
    Dim b_KeepGoing As New SiteBoolean
    
    Dim PreviousCap As New PinListData
   
    Dim b_ControlNextBit As Boolean
    b_ControlNextBit = False
    Dim b_FirstExecution As Boolean
    b_FirstExecution = False
    StoreEachIndex = 1
    
    If PreCheckMinMaxTrimCode = False Then
        b_KeepGoing = True
    End If
    
    If b_KeepGoing.All(False) Then
    Else
        For i = 0 To ExecutionMax
            If TrimPrcocessAll = False Then
                If b_StopTrimCodeProcess.All(True) Then
                    Exit For
                End If
            End If
            StrSeparatebyEqual = Split(StrSeparatebyComma(i), "=")
            StrSeparatebyColon = Split(StrSeparatebyEqual(1), ":")
            SrcStartBit = StrSeparatebyColon(0)
            SrcEndBit = StrSeparatebyColon(1)
            
            If i = 0 Then
                b_FirstExecution = True
            Else
                b_FirstExecution = False
                SrcStartBit = SrcStartBit + 1
            End If
            
            For j = SrcStartBit To (SrcEndBit + 1) Step -1
            ''===============up for src bit step
    
    '            For Each Site In TheExec.sites
    '                CapValue = CStr(OutDspWave(Site).Element(0))
    '            Next Site
                
                If b_FirstExecution = True Then
                    b_ControlNextBit = True
                    If j = SrcEndBit Then ''=============Trim from Format untill same======1124
                        b_ControlNextBit = False
                    End If
                Else
                
                ''20160716-Control next bit to 1 no matter first or last progress
                    b_ControlNextBit = True
                    If j = SrcEndBit Then
                        b_ControlNextBit = False
                    End If
                End If
    
                If b_FirstExecution = True And j = SrcEndBit Then
                    'Call rundsp.SetupTrimCodeBit(PastDSPWave, True, j, b_ControlNextBit, InDSPWave)
                    Call rundsp.SetupTrimCodeBit(PastDSPWave, True, j, b_ControlNextBit, DSPWF_TrimCodeBin) 'nn_test_1207
                Else
                    'Call rundsp.SetupTrimCodeBit(PastDSPWave, b_HigherThanTarget, j, b_ControlNextBit, InDSPWave)
                    'Call rundsp.SetupTrimCodeBit(PastDSPWave, b_HigherThanTarget.LogicalNot, j, b_ControlNextBit, InDSPWave) 'add .LogicalNot   'From T-BraC
                    Call rundsp.SetupTrimCodeBit(PastDSPWave, b_HigherThanTarget.LogicalNot, j, b_ControlNextBit, DSPWF_TrimCodeBin) 'add .LogicalNot   'nn_test_1207
                End If
                
    '            If BurstYesPatDict.Exists(LCase(patset.value)) = True Then
    '                Call SetupDigSrcDspWave(tempPatArr(tempSrcPatSeq), DigSrc_pin, tempSrcSigName, l_DigSrc_Sample_Size, InDSPWave)
    '            Else
    '                Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeCap", l_DigSrc_Sample_Size, InDSPWave)
    '            End If
                
                'Use DigSrcAssignment/DigSrc_Equ for flexible trim usage -- 20231207
                For Each site In TheExec.sites
'                    StoreEachTrimCode(StoreEachIndex)(site) = InDSPWave(site).Copy
                    StoreEachTrimCode(StoreEachIndex)(site) = DSPWF_TrimCodeBin(site).Copy
                Next site
                
                If MSB_First_Flag = True Then
                    Call rundsp.Split_Dspwave_Reverse(DSPWF_TrimCodeBin, DSPWF_TrimCodeBin_Rev)
                    Call StoreDataAllType(TrimFormat_StoreName, DSPWF_TrimCodeBin_Rev)
                Else
                    Call StoreDataAllType(TrimFormat_StoreName, DSPWF_TrimCodeBin)
                End If
                Call GeneralDigSrcSettingWithBurst(LCase(patset.value), DigSrc_pin, InDSPWave)
                
                '' Debug use
                '' ==============================================================================================
                '' 20160716 - Modify trim code rule
                If b_FirstExecution = True Then
                    If j = SrcEndBit Then
                        TheExec.Datalog.WriteComment ("Setup Bit " & j & " to 0")
                    Else
                        TheExec.Datalog.WriteComment ("Setup Bit " & j & ", Trim Code Bit " & j - 1)
                    End If
                Else
                    If j = SrcEndBit Then
                        TheExec.Datalog.WriteComment ("Setup Bit " & j)
                    Else
                        TheExec.Datalog.WriteComment ("Setup Bit " & j & ", Trim Code Bit " & j - 1)
                    End If
                End If
                    
                For Each site In TheExec.sites
                    If b_MatchTagetCap(site) = False And b_DisplayCap(site) = False Then
'                        If b_KeepGoing(site) = True Then
'                            SourceTrimCode = vbNullString
'                            For k = 0 To InDSPWave(site).sampleSize - 1
'                                SourceTrimCode = SourceTrimCode & CStr(InDSPWave(site).Element(k))
'                            Next k
'                            Dim OutputDec As String
'                            TheExec.Datalog.WriteComment ("Site_" & site & " Source Trim Code = " & SourceTrimCode)
'                        End If
                        
                        'Use DigSrcAssignment/DigSrc_Equ for flexible trim usage -- 20231207
                        If b_KeepGoing(site) = True Then
                            SourceTrimCode = vbNullString
                            For k = 0 To DSPWF_TrimCodeBin(site).sampleSize - 1
                                SourceTrimCode = SourceTrimCode & CStr(DSPWF_TrimCodeBin(site).Element(k))
                            Next k
                            Dim OutputDec As String
                            TheExec.Datalog.WriteComment ("Site_" & site & " Source Trim Code = " & SourceTrimCode)
                        End If
                    End If
                Next site
                Set OutDspWave = Nothing
                'Call GeneralDigCapSetting(PattArray(0), DigCap_Pin, DigCap_Sample_Size, OutDspWave)
                If BurstYesPatDict.Exists(LCase(patset.value)) = True Then
                    Call GeneralDigCapSetting(tempPatArr(tempSrcPatSeq), DigCap_Pin, DigCap_Sample_Size, OutDspWave)
                Else
                    Call GeneralDigCapSetting(PattArray(0), DigCap_Pin, DigCap_Sample_Size, OutDspWave)
                End If
                'Call TheHdw.Patterns(PattArray(0)).start
                Call TheHdw.Patterns(patset).start
                
                ''Update Interpose_PreMeas 20170801
                TestSeqNum = 0
                                
                
                TheHdw.Digital.Patgen.HaltWait
                
                CapValue.value = OutDspWave.Element(0)
                b_HigherThanTarget = CapValue.Math.Subtract(TrimTarget).compare(GreaterThan, 0)
                'PastDSPWave = InDSPWave
                PastDSPWave = DSPWF_TrimCodeBin   'Use DigSrcAssignment/DigSrc_Equ for flexible trim usage -- 20231207
    
                TestLimitIndex = TestLimitIndex + 1
                
                '' 20160712 - Modify to use WriteComment to display output frequency.
                For Each site In TheExec.sites
                    If b_KeepGoing(site) = True Then
    '                    TheExec.Datalog.WriteComment ("Site " & Site & " Output CapValue = " & CapValue.Pins(Site).Value(Site))
                        TheExec.Datalog.WriteComment ("Site " & site & " Output CapValue = " & OutDspWave(site).Element(0))
                    End If
                Next site
                
                b_MatchTagetCap = CapValue.Math.Subtract(TrimTarget).Abs.compare(LessThanOrEqualTo, TrimTargetTolerance)
                b_DisplayCap = b_DisplayCap.LogicalOr(b_MatchTagetCap)
                For Each site In TheExec.sites
                    If b_KeepGoing(site) = True Then
                        If b_MatchTagetCap(site) = True And StoredTargetTrimCode(site).CalcSum = 0 Then
'                            StoredTargetTrimCode(site) = InDSPWave(site).Copy
                            
                            'Use DigSrcAssignment/DigSrc_Equ for flexible trim usage -- 20231207
                            StoredTargetTrimCode(site) = DSPWF_TrimCodeBin(site).Copy
                            b_StopTrimCodeProcess(site) = True
                        End If
                    End If
                Next site
                ''20161128-Stop trim code process if found out match code of all site
                If TrimPrcocessAll = False Then
                    If b_StopTrimCodeProcess.All(True) Then
                        Exit For
                    End If
                End If
    
                TheExec.Datalog.WriteComment ("======================================================================================")
            Next j
            
            'Use DigSrcAssignment/DigSrc_Equ for flexible trim usage -- 20231207
            Dim decDSPwave As New DSPWave
                decDSPwave.CreateConstant 0, 1, DspLong			
            For Each site In TheExec.sites
                SourceTrimCode = CapValue
                For k = 1 To DSPWF_TrimCodeBin(site).sampleSize - 1
                    SourceTrimCode = SourceTrimCode & CStr(DSPWF_TrimCodeBin(site).Element(k))
                Next k
                DSPWF_TrimCodeBin(site).Element(0) = CapValue
                TheExec.Datalog.WriteComment ("Site_" & site & " Final Source Trim Code = " & SourceTrimCode)
                If gl_Disable_HIP_debug_log = False Then                        'Printing for store info
                    decDSPwave(site) = InDSPWave(site).ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspParallel, InDSPWave(site).SampleSize, 0, Bit0IsMsb)
                    TheExec.Datalog.WriteComment "Site : " & site & ", Store Value : " & decDSPwave(site).Element(0) & ", Binary Bits : " & InDSPWave(site).SampleSize & ", Store Name : " & TrimStoreName
                End If			   
            Next site
            
'            For Each site In TheExec.sites
'                If CapValue = 1 Then
'                    SourceTrimCode = vbNullString
'                    SourceTrimCode = SourceTrimCode & "0"
'    '                For k = 1 To InDSPWave(site).sampleSize - 1
'    '                    SourceTrimCode = SourceTrimCode & CStr(InDSPWave(site).Element(k))
'    '                Next k
'    '                InDSPWave(site).Element(0) = 0
'
'                    'Use DigSrcAssignment/DigSrc_Equ for flexible trim usage -- 20231207
'                    For k = 1 To DSPWF_TrimCodeBin(site).sampleSize - 1
'                        SourceTrimCode = SourceTrimCode & CStr(DSPWF_TrimCodeBin(site).Element(k))
'                    Next k
'                    DSPWF_TrimCodeBin(site).Element(0) = 0
'                Else
'                    SourceTrimCode = vbNullString
'    '                For k = 0 To InDSPWave(site).sampleSize - 1
'    '                    SourceTrimCode = SourceTrimCode & CStr(InDSPWave(site).Element(k))
'    '                Next k
'
'                    'Use DigSrcAssignment/DigSrc_Equ for flexible trim usage -- 20231207
'                    For k = 0 To DSPWF_TrimCodeBin(site).sampleSize - 1
'                        SourceTrimCode = SourceTrimCode & CStr(DSPWF_TrimCodeBin(site).Element(k))
'                    Next k
'
'                End If
'                TheExec.Datalog.WriteComment ("Site_" & site & " Final Source Trim Code = " & SourceTrimCode)
'            Next site
        Next i
        
    End If
    
    If TrimStoreName <> "" Then
        'Call Checker_StoreDigCapAllToDictionary(TrimStoreName, InDSPWave)
        'Use DigSrcAssignment/DigSrc_Equ for flexible trim usage -- 20231207
        If MSB_First_Flag = True Then
            Call rundsp.Split_Dspwave_Reverse(DSPWF_TrimCodeBin, DSPWF_TrimCodeBin_Rev)
            Call StoreDataAllType(TrimStoreName, DSPWF_TrimCodeBin_Rev)
        Else
            Call StoreDataAllType(TrimStoreName, DSPWF_TrimCodeBin)
        End If
    End If
    Instance_Data.patset = patset
    Call HardIP_WriteFuncResult(, , glb_TestInstance)

    Dim ConvertedDataWf As New DSPWave

    'rundsp.ConvertToLongAndSerialToParrel InDSPWave, l_DigSrc_Sample_Size, ConvertedDataWf
    
    rundsp.ConvertToLongAndSerialToParrel DSPWF_TrimCodeBin, l_DigSrc_Sample_Size, ConvertedDataWf
        
    'rundsp.ConvertToLongAndSerialToParrel InDSPWave, TrimBitSize, ConvertedDataWf
    
    TestNameInput = Report_TName_From_Instance("C", "X", , , , , , , tlForceFlow)   ' Update for C-CHOP 20200130
    
    'TheExec.Flow.TestLimit ConvertedDataWf.Element(0), Tname:=TheExec.DataManager.instanceName, PinName:="SEPVM_Trim_Dec", ForceResults:=tlForceFlow     '''Original 20200130 C-CHOP
    TheExec.Flow.TestLimit ConvertedDataWf.Element(0), Tname:=TestNameInput, PinName:="VM-DTB", ForceResults:=tlForceFlow     '''20200130 C-CHOP
    
    'Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeCap", DigSrc_Sample_Size, StoredTargetTrimCode)
'    If BurstYesPatDict.Exists(LCase(patset.value)) = True Then
'        Call SetupDigSrcDspWave(tempPatArr(tempSrcPatSeq), DigSrc_pin, tempSrcSigName, l_DigSrc_Sample_Size, StoredTargetTrimCode)
'    Else
'        Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeCap", l_DigSrc_Sample_Size, StoredTargetTrimCode)
'    End If
    
    'Use DigSrcAssignment/DigSrc_Equ for flexible trim usage -- 20231207
    'Call GeneralDigSrcSettingWithBurst(LCase(patset.value), DigSrc_pin, InDSPWave)
    
    'Call TheHdw.Patterns(PattArray(0)).start
    'Call thehdw.Patterns(patset).start

    ''Update Interpose_PreMeas 20170801
    TestSeqNum = 0
   
    TheHdw.Digital.Patgen.HaltWait
    
    DebugPrintFunc patset.value
    
    If Interpose_PrePat <> "" Then '''180109 update
        Call SetForceCondition("RESTOREPREPAT")
    End If
    
    Call SetForceCondition(Interpose_PostTest)
    

    
    Exit Function
    
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP", "TrimCodeDig") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function TrimCodeDig_Reverse_Ellis(Optional patset As Pattern, Optional TestSequence As String, Optional CPUA_Flag_In_Pat As Boolean, _
    Optional MeasureF_Pin As PinList, Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As Long, Optional CUS_Str_DigCapData As String, _
    Optional TrimPrcocessAll As Boolean = False, Optional UseMinimumTrimCode As Boolean = False, Optional PreCheckMinMaxTrimCode As Boolean = False, _
    Optional TrimTarget As Double, Optional TrimTargetTolerance As Double = 0, Optional TrimStart As String, Optional TrimFormat As String, _
    Optional TrimStoreName As String, Optional TrimFuseName As String, Optional TrimFuseTypeName As String, Optional Interpose_PrePat As String, Optional DigCap_Pin As PinList, Optional DigCap_Sample_Size As Long, _
    Optional Validating_ As Boolean, Optional Interpose_PostTest As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    
    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)
    
    
    
    Dim PatCount As Long, PattArray() As String
    Dim OutDspWave As New DSPWave
               TrimTarget = 0.3 ''Just for Flag    2017/Dec
    If Validating_ Then
        Call PrLoadPattern(patset.value)
        Exit Function    ' Exit after validation
    End If
    Call HardIP_InitialSetupForPatgen
    Dim Ts As Variant, TestSequenceArray() As String
    Dim InitialDSPWave As New DSPWave, PastDSPWave As New DSPWave, InDSPWave As New DSPWave

    Dim site As Variant
    Dim Pat As String
    Dim i As Long, j As Long, k As Long
    'Dim Inst_Name_Str As String: Inst_Name_Str = glb_TestInstance
    
    Dim CapValue As New PinListData, CapValue_V1 As New PinListData, CapValue_V2 As New PinListData
    CapValue.AddPin ("CapValueString")

    Call GetFlowTName
 
    TestSequenceArray = Split(TestSequence, ",")
    TheHdw.Digital.Patgen.Halt
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    If Interpose_PrePat <> "" Then ''''180109 update
        If (UCase(Interpose_PrePat) Like "*CP*=*" Or UCase(Interpose_PrePat) Like "*FT*=*") Then Interpose_PrePat = Select_forcecondition(Interpose_PrePat, CurrentJobName_U)
        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    End If
    Mease_SEPVME_Force_condition '200710
    
    TheHdw.Patterns(patset).Load

    Call PATT_GetPatListFromPatternSet(patset.value, PattArray, PatCount)

    
    '' 20160425 - Check format from TrimFormat
    Dim StrSeparatebyComma() As String
    Dim ExecutionMax As Long
    StrSeparatebyComma = Split(TrimFormat, ";")
    ExecutionMax = UBound(StrSeparatebyComma)
    Dim StrSeparatebyEqual() As String, StrSeparatebyColon() As String '' Get Src bit
    Dim SrcStartBit As Long, SrcEndBit As Long
    
    Dim b_HigherThanTarget As New SiteBoolean
    b_HigherThanTarget = False
    
    Dim LastSectionV1V2_Index As Long
    LastSectionV1V2_Index = 0
    Dim OutputTrimCode As String, TestNameInput As String
    Dim SourceTrimCode As String
    Dim TestLimitIndex As Long '
    
'    Dim TrimStart_1st() As String
    Dim Dec_TrimStart_1st As Long
    
    Dim StoredTargetTrimCode As New DSPWave
    Dim b_MatchTagetCap As New SiteBoolean
    Dim b_DisplayCap As New SiteBoolean
    b_MatchTagetCap = False
    b_DisplayCap = False
    
    StoredTargetTrimCode.CreateConstant 0, DigSrc_Sample_Size, DspLong
    Dim StoreEachTrimCode() As New DSPWave
    ReDim StoreEachTrimCode(DigSrc_Sample_Size + 1) As New DSPWave
    
    Dim DigCapData() As New PinListData
    ReDim DigCapData(DigSrc_Sample_Size + 1) As New PinListData
    
    Dim StoreEachIndex As Long
    
    ''20161128-Stop trim code process
    Dim b_StopTrimCodeProcess As New SiteBoolean
    b_StopTrimCodeProcess = False
    
    For i = 0 To UBound(StoreEachTrimCode)
        StoreEachTrimCode(i).CreateConstant 0, DigSrc_Sample_Size, DspLong
    Next i

    If TrimStart <> "" And TrimStart Like "*&*" Then
        TrimStart = Replace(TrimStart, "&", vbNullString)
    End If

'' ===============================   TrimStart_1st = TrimStart===============================================
    Dec_TrimStart_1st = Bin2Dec(TrimStart)
    
    InitialDSPWave.CreateConstant Dec_TrimStart_1st, 1, DspLong

    Call rundsp.CreateFlexibleDSPWave(InitialDSPWave, DigSrc_Sample_Size, InDSPWave)

    Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeCap", DigSrc_Sample_Size, InDSPWave)
    TheExec.Datalog.WriteComment ("========First Time Setup========")
    
    For Each site In TheExec.sites
        SourceTrimCode = vbNullString
        For k = 0 To InDSPWave(site).SampleSize - 1
            SourceTrimCode = SourceTrimCode & CStr(InDSPWave(site).Element(k))
        Next k
        TheExec.Datalog.WriteComment ("Site_" & site & " Initial Source Trim Code = " & SourceTrimCode)
    Next site
    
    For Each site In TheExec.sites
        StoreEachTrimCode(0)(site).data = InDSPWave(site).data
    Next site
    
'=========Set Up DigCap parameter=============================
    Call GeneralDigCapSetting(PattArray(0), DigCap_Pin, DigCap_Sample_Size, OutDspWave)
    Call TheHdw.Patterns(PattArray(0)).start
     'For Each Site In TheExec.sites
        CapValue.value = OutDspWave.Element(0)
     'Next Site
    Call PrintDigCapSetting(DigCap_Pin, DigCap_Sample_Size, CUS_Str_DigCapData)
    ''Update Interpose_PreMeas 20170801
    Dim TestSeqNum As Integer
    TestSeqNum = 0
           
    TheHdw.Digital.Patgen.HaltWait
'    For Each Site In TheExec.sites
        DigCapData(0) = CapValue
        b_HigherThanTarget = CapValue.Math.Subtract(TrimTarget).compare(GreaterThan, 0)
        PastDSPWave = InDSPWave
        TestNameInput = "Cap_Value_"
        TestLimitIndex = 0
'    Next Site

    Dim CUS_Str_MainProgram As String: CUS_Str_MainProgram = vbNullString

    
    For Each site In TheExec.sites
        If b_MatchTagetCap(site) = False And b_DisplayCap(site) = False Then
            
            TheExec.Datalog.WriteComment ("Site " & site & " Output CapValue = " & OutDspWave(site).Element(0))
            
        End If
    Next site
    
    b_MatchTagetCap = CapValue.Math.Subtract(TrimTarget).Abs.compare(LessThanOrEqualTo, TrimTargetTolerance)
    b_DisplayCap = b_DisplayCap.LogicalOr(b_MatchTagetCap)
    
    For Each site In TheExec.sites
        If b_DisplayCap(site) = True And StoredTargetTrimCode(site).CalcSum = 0 Then
            StoredTargetTrimCode(site).data = InDSPWave(site).data
            b_StopTrimCodeProcess(site) = True
        End If
    Next site
    
    TheExec.Datalog.WriteComment ("======================================================================================")
    
    Dim b_KeepGoing As New SiteBoolean
    
    Dim PreviousCap As New PinListData
   
    Dim b_ControlNextBit As Boolean
    b_ControlNextBit = False
    Dim b_FirstExecution As Boolean
    b_FirstExecution = False
    StoreEachIndex = 1
    
    If PreCheckMinMaxTrimCode = False Then
        b_KeepGoing = True
    End If
    
    If b_KeepGoing.All(False) Then
    Else

    For i = 0 To ExecutionMax
            If TrimPrcocessAll = False Then
                If b_StopTrimCodeProcess.All(True) Then
                    Exit For
                End If
            End If
                
            StrSeparatebyEqual = Split(StrSeparatebyComma(i), "=")
            StrSeparatebyColon = Split(StrSeparatebyEqual(1), ":")
            SrcStartBit = StrSeparatebyColon(0)
            SrcEndBit = StrSeparatebyColon(1)
            
            If i = 0 Then
                b_FirstExecution = True
            Else
                b_FirstExecution = False
                SrcStartBit = SrcStartBit + 1
            End If
        
        For j = SrcStartBit To (SrcEndBit + 1) Step -1
        ''===============up for src bit step

'            For Each Site In TheExec.sites
'                CapValue = CStr(OutDspWave(Site).Element(0))
'            Next Site
            
            If b_FirstExecution = True Then
                b_ControlNextBit = True
                If j = SrcEndBit Then ''=============Trim from Format untill same======1124
                    b_ControlNextBit = False
                End If
            Else
            
            ''20160716-Control next bit to 1 no matter first or last progress
                b_ControlNextBit = True
                If j = SrcEndBit Then
                    b_ControlNextBit = False
                End If
            End If

            If b_FirstExecution = True And j = SrcEndBit Then
                Call rundsp.SetupTrimCodeBit(PastDSPWave, True, j, b_ControlNextBit, InDSPWave)
            Else
                Call rundsp.SetupTrimCodeBit(PastDSPWave, b_HigherThanTarget.LogicalNot, j, b_ControlNextBit, InDSPWave)
            End If
            
            Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeCap", DigSrc_Sample_Size, InDSPWave)
            
                For Each site In TheExec.sites
                    StoreEachTrimCode(StoreEachIndex)(site).data = InDSPWave(site).data
                Next site
            
            '' Debug use
            '' ==============================================================================================
            '' 20160716 - Modify trim code rule
            If b_FirstExecution = True Then
                If j = SrcEndBit Then
                    TheExec.Datalog.WriteComment ("Setup Bit " & j & " to 0")
                Else
                    TheExec.Datalog.WriteComment ("Setup Bit " & j & ", Trim Code Bit " & j - 1)
                End If
            Else
                If j = SrcEndBit Then
                    TheExec.Datalog.WriteComment ("Setup Bit " & j)
                Else
                    TheExec.Datalog.WriteComment ("Setup Bit " & j & ", Trim Code Bit " & j - 1)
                End If
            End If
                
                For Each site In TheExec.sites
                    If b_MatchTagetCap(site) = False And b_DisplayCap(site) = False Then
                    If b_KeepGoing(site) = True Then
                        SourceTrimCode = vbNullString
                        For k = 0 To InDSPWave(site).SampleSize - 1
                            SourceTrimCode = SourceTrimCode & CStr(InDSPWave(site).Element(k))
                        Next k
                        Dim OutputDec As String
                        TheExec.Datalog.WriteComment ("Site_" & site & " Source Trim Code = " & SourceTrimCode)
                    End If
                    End If
                Next site
            Set OutDspWave = Nothing
            Call GeneralDigCapSetting(PattArray(0), DigCap_Pin, DigCap_Sample_Size, OutDspWave)
            Call TheHdw.Patterns(PattArray(0)).start
            
            ''Update Interpose_PreMeas 20170801
            TestSeqNum = 0
                            
            
            TheHdw.Digital.Patgen.HaltWait
            
                        CapValue.value = OutDspWave.Element(0)
            b_HigherThanTarget = CapValue.Math.Subtract(TrimTarget).compare(GreaterThan, 0)
                PastDSPWave = InDSPWave

            TestLimitIndex = TestLimitIndex + 1
            
            '' 20160712 - Modify to use WriteComment to display output frequency.
            For Each site In TheExec.sites
                If b_KeepGoing(site) = True Then
'                    TheExec.Datalog.WriteComment ("Site " & Site & " Output CapValue = " & CapValue.Pins(Site).Value(Site))
                    TheExec.Datalog.WriteComment ("Site " & site & " Output CapValue = " & OutDspWave(site).Element(0))
                End If
            Next site
            
            b_MatchTagetCap = CapValue.Math.Subtract(TrimTarget).Abs.compare(LessThanOrEqualTo, TrimTargetTolerance)
            b_DisplayCap = b_DisplayCap.LogicalOr(b_MatchTagetCap)
            For Each site In TheExec.sites
                If b_KeepGoing(site) = True Then
                    If b_MatchTagetCap(site) = True And StoredTargetTrimCode(site).CalcSum = 0 Then
                        StoredTargetTrimCode(site).data = InDSPWave(site).data
                        b_StopTrimCodeProcess(site) = True
                    End If
                End If
            Next site
            ''20161128-Stop trim code process if found out match code of all site
            If TrimPrcocessAll = False Then
                If b_StopTrimCodeProcess.All(True) Then
                    Exit For
                End If
            End If

            TheExec.Datalog.WriteComment ("======================================================================================")
        Next j
        For Each site In TheExec.sites
'            If CapValue = 1 Then
'                SourceTrimCode = ""
'                    SourceTrimCode = SourceTrimCode & "0"
'                For k = 1 To InDSPWave(site).sampleSize - 1
'                    SourceTrimCode = SourceTrimCode & CStr(InDSPWave(site).Element(k))
'                Next k
'                InDSPWave(site).Element(0) = 0
'            Else
'                SourceTrimCode = ""
'
'                For k = 0 To InDSPWave(site).sampleSize - 1
'                    SourceTrimCode = SourceTrimCode & CStr(InDSPWave(site).Element(k))
'                Next k
'
'            End If
            If CapValue = 1 Then
                InDSPWave(site).Element(0) = 1 '0
            Else
                InDSPWave(site).Element(0) = 0
            End If

            SourceTrimCode = vbNullString
            For k = 0 To InDSPWave(site).SampleSize - 1
                SourceTrimCode = SourceTrimCode & CStr(InDSPWave(site).Element(k))
            Next k
            '200708
           TheExec.Datalog.WriteComment ("Site_" & site & " Final Source Trim Code = " & SourceTrimCode)
        Next site
    
    Next i
        
    End If
    
    If TrimStoreName <> "" Then
       Call Checker_StoreDigCapAllToDictionary(TrimStoreName, InDSPWave)
    End If
    
    Call HardIP_WriteFuncResult(, , glb_TestInstance)

    Dim ConvertedDataWf As New DSPWave

    rundsp.ConvertToLongAndSerialToParrel InDSPWave, DigSrc_Sample_Size, ConvertedDataWf
    
    TestNameInput = Report_TName_From_Instance("C", "VM-DTB", , , , , , , tlForceFlow)   ' Update for C-CHOP 20200130
    
    'TheExec.Flow.TestLimit ConvertedDataWf.Element(0), Tname:=TheExec.DataManager.instanceName, PinName:="SEPVM_Trim_Dec", ForceResults:=tlForceFlow     '''Original 20200130 C-CHOP
    TheExec.Flow.TestLimit ConvertedDataWf.Element(0), Tname:=TestNameInput, PinName:="VM-DTB", ForceResults:=tlForceFlow     '''20200130 C-CHOP
    
    Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeCap", DigSrc_Sample_Size, StoredTargetTrimCode)
    Call TheHdw.Patterns(PattArray(0)).start

    ''Update Interpose_PreMeas 20170801
    TestSeqNum = 0
   
    TheHdw.Digital.Patgen.HaltWait
    
    
    Call SetForceCondition(Interpose_PostTest)
    
    If Interpose_PrePat <> "" Then '''180109 update
        Call SetForceCondition("RESTOREPREPAT")
    End If
    

    Dim sl_FUSE_Val As New SiteLong

    DebugPrintFunc patset.value
    Exit Function
    
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP", "TrimCodeDig_Reverse_Ellis") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function TrimCodeBasicDig(Optional patset As Pattern, Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As String, _
Optional TrimTarget As Double, Optional TrimStart As String, Optional TrimFormat As String, Optional TrimStoreName As String, _
Optional IncreaseFlag As Boolean = True, Optional BinarySearchFlag As Boolean = True, Optional TrimPrcocessAll As Boolean = True, _
Optional Interpose_PrePat As String, Optional DigCap_Pin As PinList, Optional DigCap_Sample_Size As Long, Optional Validating_ As Boolean, _
Optional Interpose_PostTest As String, Optional TrimOffset As String, Optional TrimBase As String, Optional DigSrc_Sample_Size_Real As String, Optional digsrc_assignment As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
'Dylan Edited 20190615

    If Validating_ Then
        Call PrLoadPattern(patset.value)
        Exit Function    ' Exit after validation
    End If
    
    Dim site As Variant
    Dim PatCount As Long

    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)
    
    Dim TrimBaseStr() As String
    Dim TrimBaseNum() As String
    Dim OffsetDelta As Integer
    'Dim Inst_Name_Str As String: Inst_Name_Str = glb_TestInstance
    Dim OutputTname_format() As String
    Dim TestNameInput As String
    
    Dim LimitCodeStart As Integer
    Dim LimitCodeEnd As Integer
    Dim i, j, k, x As Integer
    Dim SrcEndBit As Integer
    Dim RegSeparate As Integer
    Dim SrcStartBit As Integer
    Dim ExecutionMax As Integer
    Dim DecTrimStart As Integer
    Dim InDSPWave As New DSPWave
    Dim OutDspWave As New DSPWave
    Dim InDspWave_New As New DSPWave
    Dim InitialDSPWave As New DSPWave
    Dim CaptureDSPWave() As New DSPWave
    ReDim CaptureDSPWave(0)
    Dim ProcessDoneDSPWave() As New DSPWave
    ReDim ProcessDoneDSPWave(0)
    Dim ConvertedDataWf As New DSPWave
    
    Dim PattArray() As String
    Dim SourceTrimCode As String
    Dim CapValue As New PinListData
    Dim StrSeparatebyEqual() As String
    Dim StrSeparatebyColon() As String
    Dim StrSeparatebyComma() As String
    Dim RegSeparatebyComma() As String
    Dim EachRegSize As New SiteLong
    Dim InitStateCode As New SiteLong
    Dim TrimOriginalSize As New SiteLong
    Dim TrimScanPoint As New SiteLong
    Dim TrimOffsetPoint As New SiteLong
    Dim CaptureDSPFlag As New SiteBoolean
    Dim b_HigherThanTarget As New SiteBoolean
    Dim b_StopTrimCodeProcess As New SiteBoolean
    Dim DigSrc_Sample_Size_Real_Temp() As String
    Dim assignment As New DSPWave
    Dim DigSrc_Assignment_Temp() As String
    Dim AssignmentDSPWave As New DSPWave
    Dim Src_dig As New SiteBoolean
    Dim b_ControlNextBit As Boolean
    b_ControlNextBit = False
    
                
    Call HardIP_InitialSetupForPatgen
    Call GetFlowTName
    TheHdw.Patterns(patset).Load
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    Call PATT_GetPatListFromPatternSet(patset.value, PattArray, PatCount)
    
    If Interpose_PrePat <> "" Then
        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    End If
    
    CapValue.AddPin ("CapValueString")
    StrSeparatebyComma = Split(TrimFormat, ";")
    ExecutionMax = UBound(StrSeparatebyComma)
    RegSeparatebyComma = Split(DigSrc_Sample_Size_Real, ";")
    Src_dig = False
    assignment.CreateConstant 0, 1, DspLong
    AssignmentDSPWave.CreateConstant 0, 1, DspLong
    
    If digsrc_assignment <> "" Then
    
    DigSrc_Assignment_Temp = Split(digsrc_assignment, "+")
    AssignmentDSPWave.CreateConstant 0, UBound(DigSrc_Assignment_Temp) + 1, DspLong
    For i = 0 To UBound(DigSrc_Assignment_Temp)
        If LCase(DigSrc_Assignment_Temp(i)) <> "sweep" Then
          assignment = GetStoreDataAllType(DigSrc_Assignment_Temp(i))
          AssignmentDSPWave.Element(i) = 1
          Src_dig = True
        End If
    Next
    
    End If
    
    For i = 0 To ExecutionMax
        StrSeparatebyEqual = Split(StrSeparatebyComma(0), "=")
        StrSeparatebyColon = Split(StrSeparatebyEqual(1), ":")
        SrcEndBit = StrSeparatebyColon(1)
        SrcStartBit = StrSeparatebyColon(0)
        DigSrc_Sample_Size_Real_Temp = Split(RegSeparatebyComma(i), "@")
        CaptureDSPWave(0).CreateConstant 0, 1, DspLong                  ' Avoid sweep fail which any site
        RegSeparate = DigSrc_Sample_Size_Real_Temp(1) / DigSrc_Sample_Size_Real_Temp(0)
        InDSPWave.CreateConstant 0, CLng(DigSrc_Sample_Size_Real_Temp(1)), DspLong
        InDspWave_New.CreateConstant 0, CLng(DigSrc_Sample_Size_Real_Temp(1)), DspLong
        
        If BinarySearchFlag = False Then                                ' This initial method Only support linear search mode
            If TrimStart = "" Then
                If IncreaseFlag = True Then                             ' Based on IncreaseFlag state to determine start point
                    TrimStart = 0
                Else
                    TrimStart = CStr(2 ^ (SrcStartBit + 1) - 1)
                End If
            End If
        End If
        
'''        For Each site In TheExec.sites.Active
            CaptureDSPFlag = False
            If UBound(StrSeparatebyColon) > 1 Then
                InitStateCode = StrSeparatebyColon(2)
            End If
            EachRegSize = CLng(DigSrc_Sample_Size_Real_Temp(0))
            TrimScanPoint = CLng(TrimStart)
            TrimOffsetPoint = CLng(TrimOffset)
            TrimOriginalSize = CLng(SrcStartBit) + 1
'''        Next site
        
        
        
        TrimStart = CStr(CInt(TrimStart) + CInt(TrimOffset))
        InitialDSPWave.CreateConstant TrimStart, 1, DspLong                           ' Define first trim code from TrimStart
        rundsp.CreateFlexibleDSPWave InitialDSPWave, CLng(DigSrc_Sample_Size_Real_Temp(0)), InDSPWave
        rundsp.ElementTransformer InDSPWave, CLng(DigSrc_Sample_Size_Real_Temp(0)), CLng(DigSrc_Sample_Size_Real_Temp(1))
        
               
        If BinarySearchFlag = True Then                                               ' Binary Search
            For j = SrcStartBit + 1 To SrcEndBit Step -1
            
               If j = SrcEndBit Then
                   b_ControlNextBit = False
            
               Else
                    b_ControlNextBit = True
            
               End If
               
            
                rundsp.ReAssignmentDSPWave InDSPWave, RegSeparate, InDspWave_New, Src_dig, assignment, AssignmentDSPWave              ' ReAssignment DSPWave element to each register
                
                
                
                '*****************************************************For HardIP_D2D debug*****************************************************
'''''''''''                Dim Constant As Integer
'''''''''''                Dim ForConstantCode As String
'''''''''''                Dim ForConstantSplit() As String
'''''''''''                ForConstantCode = "1,0,1,0,1,0,1,0"
'''''''''''                ForConstantCode = StrReverse(ForConstantCode)
'''''''''''
'''''''''''                ForConstantSplit = Split(ForConstantCode, ",")
'''''''''''                If theexec.DataManager.instanceName = "D2D_ZCPDD2DIMPCL_PP_SHKA0_C_FULP_AN_AMXX_DLL_JTG_PRG_ALLFV_SI_D2DIMPCL_ZCPD_NV" Then
'''''''''''                    Constant = 0
'''''''''''                    For k = 0 To 7                                    '|TrimCode|Constanct|TrimCode|Constanct|Constanct|
'''''''''''                        If CStr(ForConstantSplit(Constant)) = "0" Then
'''''''''''                            InDspWave_New(0).Element(k) = 0
'''''''''''                        Else
'''''''''''                            InDspWave_New(0).Element(k) = 1
'''''''''''                        End If
'''''''''''                        Constant = Constant + 1
'''''''''''                    Next k
'''''''''''                    Constant = 0
'''''''''''                    For k = 8 To 15
'''''''''''                        If CStr(ForConstantSplit(Constant)) = "0" Then
'''''''''''                            InDspWave_New(0).Element(k) = 0
'''''''''''                        Else
'''''''''''                            InDspWave_New(0).Element(k) = 1
'''''''''''                        End If
'''''''''''                        Constant = Constant + 1
'''''''''''                    Next k
'''''''''''                    Constant = 0
'''''''''''                    For k = 24 To 31
'''''''''''                        If CStr(ForConstantSplit(Constant)) = "0" Then
'''''''''''                            InDspWave_New(0).Element(k) = 0
'''''''''''                        Else
'''''''''''                            InDspWave_New(0).Element(k) = 1
'''''''''''                        End If
'''''''''''                        Constant = Constant + 1
'''''''''''                    Next k
'''''''''''                End If
                '*****************************************************For HardIP_D2D debug*****************************************************
                
                
                
                
                
                For Each site In TheExec.sites.Active
                    SourceTrimCode = vbNullString
                    For k = InDspWave_New.SampleSize - 1 To 0 Step -1
                        SourceTrimCode = SourceTrimCode & CStr(InDspWave_New.Element(k))
                    Next k
                    
                    If j = SrcStartBit + 1 Then
                        TheExec.Datalog.WriteComment ("InitialBit , Trim Code Bit " & SourceTrimCode)
                    Else
                        TheExec.Datalog.WriteComment ("Setup Bit " & (j) & ", Trim Code Bit " & SourceTrimCode)
                    End If
                Next site
                Dim tempVarArray As Variant
                Dim Digstart_name As String
                tempVarArray = TheHdw.DSSC.Pins(DigSrc_pin).Pattern(PattArray(1)).Source.Labels.list ''20210609 temp
                Digstart_name = tempVarArray(0)
                
                Call SetupDigSrcDspWave(PattArray(1), DigSrc_pin, Digstart_name, CLng(DigSrc_Sample_Size_Real_Temp(1)), InDspWave_New)   '''' not LSB to MSB ???
                Set OutDspWave = Nothing
                Call GeneralDigCapSetting(PattArray(1), DigCap_Pin, DigCap_Sample_Size, OutDspWave)
                
                Dim DSPwave_temp As New DSPWave
                Dim W As Long
                
                
                For Each site In TheExec.sites.Active
                
                DSPwave_temp.CreateConstant 0, 1, DspLong
                
                
                
                DSPwave_temp = InDspWave_New.ConvertStreamTo(tldspParallel, CLng(DigSrc_Sample_Size_Real_Temp(0)), 0, Bit0IsMsb)
                
                For W = 0 To DSPwave_temp.SampleSize - 1
                
                TheExec.Datalog.WriteComment CStr(DSPwave_temp.Element(W))
                
                Next
                Next site
                            
                Call TheHdw.Patterns(LCase(patset.value)).start
                TheHdw.Digital.Patgen.HaltWait
                
                For Each site In TheExec.sites.Active
                    CapValue.value = OutDspWave.Element(0)
                    ' TrimTarget is your expected transfer-point
                    b_HigherThanTarget = CapValue.Math.Subtract(TrimTarget).compare(EqualTo, 0)
                    
                    '///////////////// fail stop ////////////////////
                    If TrimPrcocessAll = False Then
                        If b_HigherThanTarget = True Then
                            If CaptureDSPFlag = False Then
                                CaptureDSPWave(0) = InDspWave_New.Copy                  ' For each site arry(0) is uncalculate value
                                CaptureDSPFlag = True
                            End If
                            b_StopTrimCodeProcess(site) = True
                        End If
                    Else
                    '/////////////////  do all  ////////////////////
                        If TrimPrcocessAll = True Then
                           CaptureDSPWave(0) = InDspWave_New.Copy
                        End If
                    End If
                    '///////////////////////////////////////////////
                    TheExec.Datalog.WriteComment ("Site " & site & " Output CapValue = " & CapValue.value)
                Next site
    
                TheExec.Datalog.WriteComment ("======================================================================================")

                If j <> SrcEndBit Then '20190530 Need to include initial bit so endbit don't need to do transform
                    If j - 1 = SrcEndBit Then
                      b_ControlNextBit = False
                      
                     End If
                      
                     
                    rundsp.SetupBinaryTrimCodeBit InDspWave_New, b_HigherThanTarget, j - 1, InitStateCode, TrimOffsetPoint, TrimOriginalSize, InDSPWave, b_ControlNextBit, AssignmentDSPWave, EachRegSize
                    rundsp.ReAssignmentDSPWave InDSPWave, RegSeparate, InDspWave_New, Src_dig, assignment, AssignmentDSPWave      ' ReAssignment DSPWave element to each register
                    
         
                    
                End If
                
                If TrimPrcocessAll = False Then                                     ' Immediately stop if all site capture done
                    If b_StopTrimCodeProcess.All(True) Then
                        Exit For
                    End If
                End If
                
            Next j
           
            
        Else                                                                        ' Linear Search
            If IncreaseFlag = True Then
                LimitCodeEnd = CInt((2 ^ (SrcStartBit + 1) - 1))
                LimitCodeStart = TrimStart
            Else                                                                    ' Decrease sweep
                LimitCodeEnd = TrimStart
                LimitCodeStart = 0
            End If
            
            For j = LimitCodeStart To LimitCodeEnd
            
                rundsp.ReAssignmentDSPWave InDSPWave, RegSeparate, InDspWave_New, Src_dig, assignment, AssignmentDSPWave   ' ReAssignment DSPWave element to each register
                
                For Each site In TheExec.sites.Active
                    SourceTrimCode = vbNullString
                    For k = InDspWave_New.SampleSize - 1 To 0 Step -1
                        SourceTrimCode = SourceTrimCode & CStr(InDspWave_New.Element(k))
                    Next k
                    If j = LimitCodeStart Then
                        b_StopTrimCodeProcess = False                               ' Inititial b_StopTrimCodeProcess flag
                        TheExec.Datalog.WriteComment ("InitialCode, Trim Code Bit " & SourceTrimCode)
                    Else
                        TheExec.Datalog.WriteComment ("LinearSweep " & TrimScanPoint(0) & "," & " Trim Code Bit " & SourceTrimCode)
                    End If
                Next site
                
                Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeCap", CLng(DigSrc_Sample_Size_Real_Temp(1)), InDspWave_New)
                Call GeneralDigCapSetting(PattArray(0), DigCap_Pin, DigCap_Sample_Size, OutDspWave)
                
                Call TheHdw.Patterns(PattArray(0)).start
                TheHdw.Digital.Patgen.HaltWait
                
                For Each site In TheExec.sites.Active
                    CapValue.value = OutDspWave.Element(0)
                    ' TrimTarget is your expected transfer-point
                    b_HigherThanTarget = CapValue.Math.Subtract(TrimTarget).compare(EqualTo, 0)
                    If b_HigherThanTarget = True Then
                    
                        If CaptureDSPFlag = False Then
                            CaptureDSPWave(0) = InDspWave_New.Copy                  ' For each site arry(0) is uncalculate value
                            CaptureDSPFlag = True
                        End If
                        
                        If TrimPrcocessAll = False Then
                            b_StopTrimCodeProcess(site) = True
                        End If
                    End If
                    TheExec.Datalog.WriteComment ("Site " & site & " Output CapValue = " & OutDspWave(site).Element(0))
                Next site
    
                TheExec.Datalog.WriteComment ("======================================================================================")
                
                If j <> LimitCodeEnd Then
'                    Judgment addition "1" or "0" based on b_HigherThanTarget value
                    rundsp.SetupLinearTrimCodeBit IncreaseFlag, TrimScanPoint, b_HigherThanTarget, EachRegSize, InDSPWave, TrimPrcocessAll
                End If
                
                If TrimPrcocessAll = False Then                                     ' Immediately stop if all site capture done
                    If b_StopTrimCodeProcess.All(True) Then
                        Exit For
                    End If
                End If
            Next j
 
        End If
        
        
        Instance_Data.patset = patset
        Call HardIP_WriteFuncResult(, , glb_TestInstance)
        rundsp.ConvertToLongAndSerialToParrel InDSPWave, EachRegSize, ConvertedDataWf
        
        Call GetFlowTName
        
        If gl_UseStandardTestName_Flag = True Then
            Call Report_ALG_TName_From_Instance(OutputTname_format, "C", TrimStoreName, gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex))
            TestNameInput = Merge_TName(OutputTname_format)
                  
        Else
            TestNameInput = glb_TestInstance & "DDR_Sweep"
        End If
        
        
        TheExec.Flow.TestLimit ConvertedDataWf.Element(0), Tname:=TestNameInput, PinName:="D2D_Trim_Dec", ForceResults:=tlForceFlow
        
        'TheExec.Flow.TestLimit ConvertedDataWf.Element(0), TName:=TheExec.DataManager.instanceName, PinName:="SEPVM_Trim_Dec", ForceResults:=tlForceFlow
            
           
        If TrimBase <> "" Then
            TrimBaseStr = Split(TrimBase, ";")
            OffsetDelta = CInt(TrimOffset) + 2 ^ CInt(StrSeparatebyColon(0))
            For j = 0 To UBound(TrimBaseStr)
                TrimBaseNum = Split(TrimBaseStr(j), ":")
                ReDim Preserve CaptureDSPWave(UBound(CaptureDSPWave) + 1)
                ReDim Preserve ProcessDoneDSPWave(UBound(ProcessDoneDSPWave) + 1)                   'This dspwave will save as to Dictionary
                
                CaptureDSPWave(UBound(CaptureDSPWave)).CreateConstant 0, 1, DspLong
                ProcessDoneDSPWave(UBound(ProcessDoneDSPWave)).CreateConstant 0, 1, DspLong
                                
                If CaptureDSPFlag.Any(True) Then
                    rundsp.CalculateDSPWaveforTrimCode CaptureDSPWave(0), EachRegSize, CaptureDSPWave(UBound(CaptureDSPWave)), _
                                                       CInt(OffsetDelta), CInt(TrimBaseNum(1)), ProcessDoneDSPWave(UBound(ProcessDoneDSPWave))
                                                       
                    
                    
                    
'''''                    TheExec.Flow.TestLimit CaptureDSPWave(UBound(CaptureDSPWave)).Element(0), TName:=TheExec.DataManager.instanceName, PinName:="SEPVM_Trim_Dec", ForceResults:=tlForceFlow
                                                       
                    StoreDataAllType TrimBaseNum(0), ProcessDoneDSPWave(UBound(ProcessDoneDSPWave))
                End If
            Next j
        End If
    Next i
    
    
    If TrimStoreName <> "" Then
       Call Checker_StoreDigCapAllToDictionary(TrimStoreName, InDSPWave)
    End If
    
        
    If Interpose_PrePat <> "" Then '''180109 update
        Call SetForceCondition("RESTOREPREPAT")
    End If

    Call SetForceCondition(Interpose_PostTest)
        
    Exit Function
    
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP", "TrimCodeBasicDig") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function Enable_HIP_Datalog_Format()
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

        With TheExec.Datalog
'            .Setup.DatalogSetup.DisableInstanceNameInPTR = False
'                'TTR,20200423, Oscar
'            .Setup.DatalogSetup.DisablePinNameInPTR = False
'            .Setup.DatalogSetup.DisableChannelNumberInPTR = True
'            .Setup.DatalogSetup.PTR_InstanceNameIsTINameOnly = True
        
        ''' 20210709 for datalog format alignment
        '''    .Setup.Shared.ascii.Columns.EnableCustomWidths = True
        '''    .Setup.Shared.ascii.Columns.Parametric.TestName.Width = 150
        '''    .Setup.Shared.ascii.Columns.Parametric.Measured.Width = 16
        '''    .Setup.Shared.ascii.Columns.Functional.TestName.Width = 150
        '''    .Setup.Shared.ascii.Columns.Functional.pattern.Width = 80
            .ApplySetup
        End With
    If False Then 'TheExec.Flow.EnableWord("Dummy") = True Then
                TheHdw.DCVI.Pins("All_DCVI").Alarm(tlDCVIAlarmOverRange) = tlAlarmOff
                TheHdw.DCVI.Pins("All_DCVI").Alarm(tlDCVIAlarmMode) = tlAlarmOff
                TheHdw.DCVI.Pins("All_DCVI").Alarm(tlDCVIAlarmCapture) = tlAlarmOff
    End If
'TTR,20200423, Oscar
'        If EnableDigitalTestLimitTTR = False Then				'[20240207][All][Brian] conment for fix and optimize flag (PTR_InstanceNameIsTINameOnly issue) move to Onprogram Validate
'            EnableDigitalTestLimitTTR = TheExec.Flow.enableWord("Enable_HardIP_DigitalTestLimitTTR")
'        End If
'        If EnableFieldProcesingTTR = False Then					'[20240207][All][Brian] conment for fix and optimize flag (PTR_InstanceNameIsTINameOnly issue) move to HardIP_OnProgramStarted_Process
'            EnableFieldProcesingTTR = TheExec.Flow.enableWord("Enable_HardIP_FieldProcesingTTR")
'        End If
'        
'        If gl_Disable_HIP_debug_log = False Or glb_Disable_CurrRangeSetting_Print = False Then				'[20240207][All][Brian] conment for fix and optimize flag (PTR_InstanceNameIsTINameOnly issue) move to HardIP_OnProgramStarted_Process
'                gl_Disable_HIP_debug_log = TheExec.Flow.enableWord("Enable_HardIP_Debug_log_disable_TTR")
'                glb_Disable_CurrRangeSetting_Print = gl_Disable_HIP_debug_log
'        End If
'         
'        If EnableHardIPDigCapsdisableTTR = False Then				'[20240207][All][Brian] conment for fix and optimize flag (PTR_InstanceNameIsTINameOnly issue) move to HardIP_OnProgramStarted_Process
'                EnableHardIPDigCapsdisableTTR = TheExec.Flow.enableWord("Enable_HardIP_DigCaps_disable_TTR")
'        End If
'        
'        If EnableAnalogMuxOutTTR = False Then				'[20240207][All][Brian] conment for fix and optimize flag (PTR_InstanceNameIsTINameOnly issue) move to Onprogram Validate
'                EnableAnalogMuxOutTTR = TheExec.Flow.enableWord("Enable_HardIP_AnalogMuxOutTTR")
'        End If
'        
'        If EnableHardIPTnameConstructionTTR = False Then				'[20240207][All][Brian] conment for fix and optimize flag (PTR_InstanceNameIsTINameOnly issue) move to Onprogram Validate
'                EnableHardIPTnameConstructionTTR = TheExec.Flow.enableWord("Enable_HardIP_TnameConstructionTTR")
'        End If

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP", "Enable_HIP_Datalog_Format") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

'Public Function Disable_HIP_Datalog_Format()
'On Error GoTo errHandler 'Add ErrHandler 2023/05/29
'
'With TheExec.Datalog
'    .Setup.DatalogSetup.DisableInstanceNameInPTR = False
'    .Setup.DatalogSetup.DisablePinNameInPTR = False
'    .Setup.DatalogSetup.DisableChannelNumberInPTR = True
'    .Setup.DatalogSetup.PTR_InstanceNameIsTINameOnly = False
'    .ApplySetup
'End With
'
'Exit Function 'Add ErrHandler 2023/05/29
'errHandler: 'Add ErrHandler 2023/05/29
'    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP", "Disable_HIP_Datalog_Format") 'Add ErrHandler 2023/05/29
'    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
'End Function

'New Module From Sicily, 20200423, Oscar

Public Function LDO_Calibration(Optional Pat As Pattern, Optional TestSequence As String, Optional MeasV_PinS As String, _
            Optional WaitTime_VIRZ As String, Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As String, Optional DigSrc_Equation As String, Optional digsrc_assignment As String, _
            Optional TrimStoreName As String, Optional TrimTarget As Double, Optional TrimStart As Long, Optional TrimCodeSize As Long, Optional TrimMethod As String, Optional Trimming_Direction_Increase As String, Optional TrimStepSize As Double, _
            Optional ForceV_Val As String, Optional ForceI_Val As String, Optional MeasI_pinS As String, Optional MeasI_Range As String, Optional Interpose_PrePat As String, Optional Interpose_PreMeas As String, Optional Interpose_PostTest As String, Optional RAK_Flag As Enum_RAK, _
            Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, Optional CUS_Str_DigCapData As String, Optional CUS_Str_MainProgram As String, Optional MSB_First_Flag As Boolean, Optional b_MinimumBestVal As Boolean, _
            Optional Calc_Eqn As String, Optional InstSpecialSetting As InstrumentSpecialSetup = 0, Optional Validating_ As Boolean)
    
    On Error GoTo errHandler
    
    If Validating_ Then
        Call PrLoadPattern(Pat.value)
        Exit Function    ' Exit after validation
    End If
    
    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)
    
    Dim site As Variant
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l_TestSeq_Num As Long
    Dim l_StoreName_Num As Long
    Dim l_StoreName_Max As Long
    
	
	If LCase(TrimMethod) = "linearsearch-all-combination" Then

	ElseIf LCase(TrimMethod) = "linearsearch" Then

	ElseIf LCase(TrimMethod) = "prediction_mode" Then

	Else

		TrimMethod = "linearsearch"
		TheExec.Datalog.WriteComment "TrimMethod is wrong, change TrimMethod to linearsearch for next step!!"
		Call Print_Error_Message(Warning_Info, "VBT_LIB_HardIP", "LDO_Calibration", "TrimMethod should be linearsearch-all-combination  or  linearsearch  or  prediction_mode !!") ' Add Warning message -- 20231018

	End If	
	
    l_TestSeq_Num = UBound(Split(TestSequence, ","))
    l_StoreName_Num = UBound(Split(TrimStoreName, "+"))
    Dim StoreName_Ary() As String
    
    Dim sTestSeqPin_Ary() As String
    Dim sTestSeqPin_Ary_Temp() As String
    Dim sTestSequence_Ary() As String
    Dim sTestStoreName_Ary() As String
    
    'sTestSeqPin_Ary = Split(MeasV_PinS, "+")
    If MeasV_PinS <> "" Then
        sTestSeqPin_Ary_Temp = Split(MeasV_PinS, "+")
    End If
    If MeasI_pinS <> "" Then
        sTestSeqPin_Ary_Temp = Split(MeasI_pinS, "+")
    End If
    ReDim sTestSeqPin_Ary(UBound(sTestSeqPin_Ary_Temp))
    
    sTestSequence_Ary = Split(TestSequence, ",")
    sTestStoreName_Ary = Split(TrimStoreName, "+")
    
    l_StoreName_Max = 0
    For i = 0 To l_TestSeq_Num
        If sTestStoreName_Ary(i) <> "" Then
            StoreName_Ary = Split(sTestStoreName_Ary(i), ",")
            If l_StoreName_Max < UBound(StoreName_Ary) Then
                l_StoreName_Max = UBound(StoreName_Ary)
            End If
        End If
    Next i
    
    If l_TestSeq_Num <> l_StoreName_Num Then
        TheExec.Datalog.WriteComment "Test Sequence and TrimStoreName are not match, Please check the argument."
        Exit Function
    End If
    
    Dim pats() As String
    Dim NumberOfMeasV As Integer: NumberOfMeasV = l_TestSeq_Num + 1
    
    '''----------TrimCode Variable----------
    Dim code() As New SiteLong: ReDim code(l_TestSeq_Num, l_StoreName_Max)
    Dim BestCode() As New SiteLong: ReDim BestCode(l_TestSeq_Num, l_StoreName_Max)
    Dim sl_GrayCode() As New SiteLong: ReDim sl_GrayCode(l_TestSeq_Num, l_StoreName_Max)
    Dim Trimming_Direction_Increase_Ary() As String: ReDim Trimming_Direction_Increase_Ary(l_TestSeq_Num, l_StoreName_Max)
    
    Dim BestVal() As New PinListData: ReDim BestVal(l_TestSeq_Num)
    Dim pl_PreviousMeasValue() As New PinListData: ReDim pl_PreviousMeasValue(l_TestSeq_Num)
    
    Dim vout() As New SiteDouble: ReDim vout(l_TestSeq_Num, l_StoreName_Max)
    
    Dim DecideTrim() As New SiteBoolean: ReDim DecideTrim(l_TestSeq_Num, l_StoreName_Max)
    Dim PreviousNegative() As New SiteBoolean: ReDim PreviousNegative(l_TestSeq_Num, l_StoreName_Max)
    Dim PreviousPositive() As New SiteBoolean: ReDim PreviousPositive(l_TestSeq_Num, l_StoreName_Max)
    
    Dim FinalTrimCode() As New DSPWave: ReDim FinalTrimCode(l_TestSeq_Num, l_StoreName_Max)
    Dim FinalTrimCode_Array() As Long: ReDim FinalTrimCode_Array(TrimCodeSize - 1) As Long
    '' GrayCode
    Dim srcwave_array_graycode() As Long: ReDim srcwave_array_graycode(TrimCodeSize - 1) As Long
    
    Dim PresentMeasValue() As New SiteDouble: ReDim PresentMeasValue(l_TestSeq_Num, l_StoreName_Max)
    Dim PreviousMeasValue() As New SiteDouble: ReDim PreviousMeasValue(l_TestSeq_Num, l_StoreName_Max)
    '''----------TrimCode Variable----------
    Dim Trim_Flag As Boolean

    Dim BlockName() As String: BlockName = Split(glb_TestInstance, "_")
    Dim MeasValue() As New PinListData: ReDim MeasValue(NumberOfMeasV - 1)

    Dim StepCount As Long: StepCount = 0
    Dim TestNameInput As String
    Dim PatCount As Long, PattArray() As String
    
    Dim PinName As String
    Dim TempVal As Integer
    
    Dim bTrimTestSeq() As Boolean
    Dim sTrim_Pin As String
    Dim sTestSeq_V As String
    Dim sTestSeq_I As String
    Dim sTestSeq_V_Ary() As String
    Dim sTestSeq_I_Ary() As String
    Dim CUS_Str_DigSrcData As String
    Dim Meas_StoreName As String
    Dim DigSrc_FlowForLoopIntegerName As String
    
 '   ReDim sTestSeqPin_Ary(l_TestSeq_Num)
    ReDim sTestSeq_V_Ary(l_TestSeq_Num)
    ReDim sTestSeq_I_Ary(l_TestSeq_Num)
    For i = 0 To l_TestSeq_Num
        If UCase(sTestSequence_Ary(i)) Like "V*" Then
            If sTestSequence_Ary(i) = "" Then
                sTestSeq_V_Ary(i) = vbNullString
            Else
                If UBound(sTestSeqPin_Ary_Temp) > 0 Then
                        sTestSeq_V_Ary(i) = sTestSeqPin_Ary_Temp(i)
                        sTestSeqPin_Ary(i) = sTestSeqPin_Ary_Temp(i)
                Else
                        sTestSeq_V_Ary(i) = sTestSeqPin_Ary_Temp(0)
                        sTestSeqPin_Ary(i) = sTestSeqPin_Ary_Temp(0)
                End If
            End If
        ElseIf UCase(sTestSequence_Ary(i)) Like "I*" Or UCase(sTestSequence_Ary(i)) Like "R*" Then
            If sTestSequence_Ary(i) = "" Then
                sTestSeq_I_Ary(i) = vbNullString
            Else
                If UBound(sTestSeqPin_Ary_Temp) > 0 Then
                        sTestSeq_I_Ary(i) = sTestSeqPin_Ary_Temp(i)
                        sTestSeqPin_Ary(i) = sTestSeqPin_Ary_Temp(i)
                Else
                        sTestSeq_I_Ary(i) = sTestSeqPin_Ary_Temp(0)
                        sTestSeqPin_Ary(i) = sTestSeqPin_Ary_Temp(0)
                End If
            End If
        End If
    Next i
    
    sTestSeq_V = Join(sTestSeq_V_Ary, "+")
    sTestSeq_I = Join(sTestSeq_I_Ary, "+")

    ReDim bTrimTestSeq(l_TestSeq_Num)

    For i = 0 To l_StoreName_Num
        If sTestStoreName_Ary(i) <> "" Then
            bTrimTestSeq(i) = True
        Else
            bTrimTestSeq(i) = False
        End If
    Next i
    
    For i = 0 To l_TestSeq_Num
        If bTrimTestSeq(i) Then
            StoreName_Ary = Split(sTestStoreName_Ary(i), ",")
            For j = 0 To UBound(StoreName_Ary)
                For Each site In TheExec.sites.Active
                    code(i, j) = TrimStart
                    vout(i, j) = 0
                    sl_GrayCode(i, j) = 0
                    PreviousNegative(i, j) = False
                    PreviousPositive(i, j) = False
                    DecideTrim(i, j) = False
                    PresentMeasValue(i, j) = 0
                    PreviousMeasValue(i, j) = 0
                Next site
                If Trimming_Direction_Increase <> "" Then
                    Trimming_Direction_Increase_Ary(i, j) = Split(Split(Trimming_Direction_Increase, "+")(i), ",")(j)
                Else
                    Trimming_Direction_Increase_Ary(i, j) = "true"
                    'Call Print_Error_Message(Warning_Info, "VBT_LIB_HardIP", "LDO_Calibration", "Trimming_Direction_Increase empty,Use default direction as TRUE, Please check and add the direction!!") ' Add Warning message -- 20231018
                End If
            Next j
        End If
    Next i
    
    Call Reg_Assign_Processing(DigSrc_Equation, digsrc_assignment, CUS_Str_DigCapData, Calc_Eqn, CUS_Str_DigSrcData, MeasV_PinS, MeasI_pinS, , Interpose_PreMeas, Meas_StoreName, CUS_Str_MainProgram, DigSrc_FlowForLoopIntegerName)
    
    Call ProcessInputToGLB(Pat, TestSequence, True, , , , , sTestSeq_V, , , , , , , sTestSeq_I, MeasI_Range, , DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, DigSrc_Pin, CStr(DigSrc_Sample_Size), CStr(DigSrc_Sample_Size), DigSrc_Equation, DigSrc_Assignment, , , InstSpecialSetting, CUS_Str_MainProgram, , , , , , , , , , , , , , , , Interpose_PrePat, Interpose_PreMeas, Interpose_PostTest, , ForceV_Val, ForceI_Val, , RAK_Flag, WaitTime_VIRZ)
    
    For i = 0 To l_TestSeq_Num
        If sTestSeqPin_Ary(i) <> "" Then
            For j = 0 To UBound(Split(sTestSeqPin_Ary(i), ","))
                MeasValue(i).AddPin (Split(sTestSeqPin_Ary(i), ",")(j))
                MeasValue(i).Pins(Split(sTestSeqPin_Ary(i), ",")(j)).value = 0
                BestVal(i).AddPin (Split(sTestSeqPin_Ary(i), ",")(j))
                BestVal(i).Pins(Split(sTestSeqPin_Ary(i), ",")(j)).value = 0
                pl_PreviousMeasValue(i).AddPin (Split(sTestSeqPin_Ary(i), ",")(j))
                pl_PreviousMeasValue(i).Pins(Split(sTestSeqPin_Ary(i), ",")(j)).value = 0
            Next j
        End If
    Next i
    
    Call GetFlowTName

    If TheExec.DevChar.Setups.IsRunning Then
        If TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.axes.Contains(tlDevCharShmooAxis_Y) Then
            If gl_Flag_HardIP_Trim_Set_PrePoint And Not (gl_Flag_HardIP_Characterization_1stRun) Then
                TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
                Call TheExec.Overlays.ApplyUniformSpecToHW(XI0_Shmoo & "_Freq_VAR", TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.axes(tlDevCharShmooAxis_Y).value)
            ElseIf gl_Flag_HardIP_Trim_Set_PostPoint Then
                TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
                Call TheExec.Overlays.ApplyUniformSpecToHW(XI0_Shmoo & "_Freq_VAR", 24000000#)
            Else
                TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
            End If
        End If
    Else
        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    End If
    
    TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD
    
    If BurstYesPatDict.Exists(LCase(Pat.value)) Then
        Call PatternBurstCheckAndSplit(Pat.value, pats, PatCount) '' 220708 for palma
    Else
        PATT_GetPatListFromPatternSet Pat.value, pats, PatCount
    End If
        
    Dim TrimCodeValue_Min As Long, TrimCodeValue_Max As Long
    TrimCodeValue_Min = 0
    TrimCodeValue_Max = 2 ^ TrimCodeSize - 1
  
    Dim lTestIimitIndex As Long
    Dim lTestIimitIndex_TrimFinish As Long
    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("**************** The Current Trimming Method is: " & TrimMethod & "****************")     
    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("**************** The Current Trimming Direction is: " & Trimming_Direction_Increase & "****************")     
    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("**************** The Measurement at Trim Start Point ****************")
    
    lTestIimitIndex = TheExec.Flow.TestLimitIndex
    Call LDO_Measurement_Process_Universal(pats(0), DigSrc_pin, code, vout, TrimCodeSize, NumberOfMeasV, l_StoreName_Max, MeasValue, DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, sTestStoreName_Ary(), WaitTime_VIRZ, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, CUS_Str_MainProgram, MSB_First_Flag, bTrimTestSeq, StepCount, sl_GrayCode, False)
    TheExec.Flow.TestLimitIndex = lTestIimitIndex
    
    For Each site In TheExec.sites.Active
        For i = 0 To l_TestSeq_Num
            BestVal(i) = MeasValue(i)
        Next i
    Next site
    
    For i = 0 To l_TestSeq_Num
        If bTrimTestSeq(i) Then
            StoreName_Ary = Split(sTestStoreName_Ary(i), ",")
            For j = 0 To UBound(StoreName_Ary)
            BestCode(i, j) = code(i, j)
                If LCase(TrimMethod) = "linearsearch-all-combination" Then
                    For Each site In TheExec.sites.Active
                        code(i, j) = code(i, j) + 1
                        PreviousMeasValue(i, j) = vout(i, j).Subtract(TrimTarget)
                        PreviousMeasValue(i, j) = PreviousMeasValue(i, j).Abs
                        DecideTrim(i, j) = True
                    Next site
                ElseIf LCase(TrimMethod) = "linearsearch" Then
                
                    For Each site In TheExec.sites.Active
                        If vout(i, j).compare(LessThan, TrimTarget) Then
                            If LCase(Trimming_Direction_Increase_Ary(i, j)) = "true" Then
                                code(i, j) = code(i, j) + 1
                            ElseIf LCase(Trimming_Direction_Increase_Ary(i, j)) = "false" Then
                                code(i, j) = code(i, j) - 1
                            Else
                            End If
                        ElseIf vout(i, j).compare(GreaterThan, TrimTarget) Then
                            If LCase(Trimming_Direction_Increase_Ary(i, j)) = "true" Then
                                If code(i, j) > 0 Then
                                    code(i, j) = code(i, j) - 1
                                End If
                            ElseIf LCase(Trimming_Direction_Increase_Ary(i, j)) = "false" Then
                                code(i, j) = code(i, j) + 1
                            Else
                            End If
                        Else
                        'Do nothing
                        End If
                        DecideTrim(i, j) = True
                    Next site
                ElseIf LCase(TrimMethod) = "prediction_mode" Then
                    For Each site In TheExec.sites.Active
                        code(i, j) = code(i, j) + Fix((TrimTarget - vout(i, j)) / TrimStepSize)
						' [20230928][All][William] Fix Trim code negative issue
                        If code(i, j).compare(GreaterThan, TrimCodeValue_Max) Then
                            code(i, j) = TrimCodeValue_Max
                        ElseIf code(i, j).compare(LessThan, TrimCodeValue_Min) Then
                            code(i, j) = TrimCodeValue_Min
                        End If						
                        DecideTrim(i, j) = True
                    Next site
                Else
                    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP", "LDO_Calibration", "TrimMethod empty,Use linear search for current trimming, Please check and add the TrimMethod!!") ' Add Warning message -- 20231018
                End If
            Next j
        End If
    Next i

    For i = 0 To l_TestSeq_Num
        If i = 0 Then Trim_Flag = False
        If bTrimTestSeq(i) Then
            StoreName_Ary = Split(sTestStoreName_Ary(i), ",")
            For j = 0 To UBound(StoreName_Ary)
                Trim_Flag = Trim_Flag Or DecideTrim(i, j).Any(True)
            Next j
        End If
    Next i
    
StartTrim:
    If Trim_Flag Then
        StepCount = StepCount + 1

        If gl_Disable_HIP_debug_log = False Then
            If right(CStr(StepCount), 1) = "1" Then
                TheExec.Datalog.WriteComment ("**************** The " & StepCount & "st Trim Process ****************")
            ElseIf right(CStr(StepCount), 1) = "2" Then
                TheExec.Datalog.WriteComment ("**************** The " & StepCount & "nd Trim Process ****************")
            ElseIf right(CStr(StepCount), 1) = "3" Then
                TheExec.Datalog.WriteComment ("**************** The " & StepCount & "rd Trim Process ****************")
            Else
                TheExec.Datalog.WriteComment ("**************** The " & StepCount & "th Trim Process ****************")
            End If
        End If

        lTestIimitIndex = TheExec.Flow.TestLimitIndex
        Call LDO_Measurement_Process_Universal(pats(0), DigSrc_pin, code, vout, TrimCodeSize, NumberOfMeasV, l_StoreName_Max, MeasValue, DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, sTestStoreName_Ary(), WaitTime_VIRZ, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, CUS_Str_MainProgram, MSB_First_Flag, bTrimTestSeq, StepCount, sl_GrayCode, False)
        ''------- Test Limit Index Value for Trimming Finsh -------
        lTestIimitIndex_TrimFinish = TheExec.Flow.TestLimitIndex
        ''------- Test Limit Index Value for Trimming Finsh -------
        TheExec.Flow.TestLimitIndex = lTestIimitIndex
    End If
    
    Dim sd_TrimTarget As New SiteDouble
    Dim pl_TrimTarget As New PinListData
    For i = 0 To l_TestSeq_Num
        If bTrimTestSeq(i) Then
            StoreName_Ary = Split(sTestStoreName_Ary(i), ",")
            For j = 0 To UBound(StoreName_Ary)
                sTrim_Pin = sTestSeqPin_Ary(i)
                For Each site In TheExec.sites.Active
                    If DecideTrim(i, j) Then
                        PresentMeasValue(i, j) = vout(i, j).Subtract(TrimTarget)
                        PresentMeasValue(i, j) = PresentMeasValue(i, j).Abs
                        
                        If LCase(TrimMethod) = "linearsearch-all-combination" Then
                            If code(i, j) <= TrimCodeValue_Max Then
                                If PresentMeasValue(i, j).compare(LessThan, PreviousMeasValue(i, j)) Then
                                    BestCode(i, j) = code(i, j)
    '                                DecideTrim(i, j) = False
                                    For k = 0 To UBound(Split(sTrim_Pin, ","))
                                        If j = k Then BestVal(i).Pins(k).value = MeasValue(i).Pins(k).value
                                    Next k
                                    PreviousMeasValue(i, j) = PresentMeasValue(i, j)
                                End If
                                
                                code(i, j) = code(i, j) + 1
                            Else
                                DecideTrim(i, j) = False
                            End If
                        Else
                            If StepCount > (TrimCodeValue_Max - TrimCodeValue_Min) Then
                                BestCode = code
                                DecideTrim(i, j) = False
                                For k = 0 To UBound(Split(sTrim_Pin, ","))
                                    If j = k Then BestVal(i).Pins(k).value = MeasValue(i).Pins(k).value
                                Next k
                            ElseIf code(i, j).compare(GreaterThan, TrimCodeValue_Max) Then
                                code(i, j) = TrimCodeValue_Max
                                PreviousMeasValue(i, j) = vout(i, j)
                                DecideTrim(i, j) = True
                            ElseIf code(i, j).compare(LessThan, TrimCodeValue_Min) Then
                                code(i, j) = TrimCodeValue_Min
                                PreviousMeasValue(i, j) = vout(i, j)
                                DecideTrim(i, j) = True
							' [20230928][All][William] Fix Trim process stop when trim code = 0 issue
							ElseIf (code(i, j).compare(EqualTo, TrimCodeValue_Max) And vout(i, j).compare(LessThan, TrimTarget)) Or _
								    (code(i, j).compare(EqualTo, TrimCodeValue_Min) And vout(i, j).compare(GreaterThan, TrimTarget)) Then
                                BestCode(i, j) = code(i, j)
                                DecideTrim(i, j) = False
                                For k = 0 To UBound(Split(sTrim_Pin, ","))
                                    If j = k Then BestVal(i).Pins(k).value = MeasValue(i).Pins(k).value
                                Next k
                            ElseIf vout(i, j).compare(LessThan, TrimTarget) And PreviousPositive(i, j) Then
                                If b_MinimumBestVal Then
                                    '''----------Minimum Error Difference----------
                                    If PreviousMeasValue(i, j) > PresentMeasValue(i, j) Then
                                        BestCode(i, j) = code(i, j)
                                        For k = 0 To UBound(Split(sTrim_Pin, ","))
                                            If j = k Then BestVal(i).Pins(k).value = MeasValue(i).Pins(k).value
                                        Next k									
                                     Else
                                        If LCase(Trimming_Direction_Increase_Ary(i, j)) = "true" Then
                                            BestCode(i, j) = code(i, j) + 1
                                        ElseIf LCase(Trimming_Direction_Increase_Ary(i, j)) = "false" Then
                                            BestCode(i, j) = code(i, j) - 1
                                        Else
                                        End If										
                                        For k = 0 To UBound(Split(sTrim_Pin, ","))
											If j = k Then BestVal(i).Pins(k).value = pl_PreviousMeasValue(i).Pins(k).value
										Next k	
                                    End If
                                    '''----------Minimum Error Difference----------
                                Else
                                    If LCase(Trimming_Direction_Increase_Ary(i, j)) = "true" Then
                                        BestCode(i, j) = code(i, j) + 1
                                    ElseIf LCase(Trimming_Direction_Increase_Ary(i, j)) = "false" Then
                                        BestCode(i, j) = code(i, j) - 1
                                    End If
                                    For k = 0 To UBound(Split(sTrim_Pin, ","))
                                        If j = k Then BestVal(i).Pins(k).value = pl_PreviousMeasValue(i).Pins(k).value
                                    Next k
                                End If
                                DecideTrim(i, j) = False
                            ElseIf vout(i, j).compare(LessThan, TrimTarget) And Not (PreviousPositive(i, j)) Then
                                If LCase(Trimming_Direction_Increase_Ary(i, j)) = "true" Then
                                    code(i, j) = code(i, j) + 1
                                ElseIf LCase(Trimming_Direction_Increase_Ary(i, j)) = "false" Then
                                    code(i, j) = code(i, j) - 1
                                Else
                                End If
                                PreviousMeasValue(i, j) = vout(i, j).Subtract(TrimTarget)
                                PreviousMeasValue(i, j) = PreviousMeasValue(i, j).Abs
                                PreviousNegative(i, j) = True
                                DecideTrim(i, j) = True
                                For k = 0 To UBound(Split(sTrim_Pin, ","))
                                    If j = k Then pl_PreviousMeasValue(i).Pins(k).value = MeasValue(i).Pins(k).value
                                Next k
                            ElseIf vout(i, j).compare(GreaterThan, TrimTarget) And PreviousNegative(i, j) Then
                                If b_MinimumBestVal Then
                                        '''----------Minimum Error Difference----------
                                        If PreviousMeasValue(i, j) > PresentMeasValue(i, j) Then
                                            BestCode(i, j) = code(i, j)
                                            For k = 0 To UBound(Split(sTrim_Pin, ","))
                                                If j = k Then BestVal(i).Pins(k).value = MeasValue(i).Pins(k).value
                                            Next k
                                        Else
                                            If LCase(Trimming_Direction_Increase_Ary(i, j)) = "true" Then
                                                BestCode(i, j) = code(i, j) - 1
                                            ElseIf LCase(Trimming_Direction_Increase_Ary(i, j)) = "false" Then
                                                BestCode(i, j) = code(i, j) + 1
                                            Else
                                            End If
                                            For k = 0 To UBound(Split(sTrim_Pin, ","))
                                                If j = k Then BestVal(i).Pins(k).value = pl_PreviousMeasValue(i).Pins(k).value
                                            Next k
                                        End If
                                        '''----------Minimum Error Difference----------
                                Else
                                    BestCode(i, j) = code(i, j)
                                    For k = 0 To UBound(Split(sTrim_Pin, ","))
                                        If j = k Then BestVal(i).Pins(k).value = MeasValue(i).Pins(k).value
                                    Next k
                                End If
                                DecideTrim(i, j) = False
                            ElseIf vout(i, j).compare(GreaterThan, TrimTarget) And Not (PreviousNegative(i, j)) Then
                                If LCase(Trimming_Direction_Increase_Ary(i, j)) = "true" Then
                                    code(i, j) = code(i, j) - 1
                                ElseIf LCase(Trimming_Direction_Increase_Ary(i, j)) = "false" Then
                                    code(i, j) = code(i, j) + 1
                                Else
                                End If
                                PreviousMeasValue(i, j) = vout(i, j).Subtract(TrimTarget)
                                PreviousMeasValue(i, j) = PreviousMeasValue(i, j).Abs
                                PreviousPositive(i, j) = True
                                DecideTrim(i, j) = True
                                For k = 0 To UBound(Split(sTrim_Pin, ","))
                                    If j = k Then pl_PreviousMeasValue(i).Pins(k).value = MeasValue(i).Pins(k).value
                                Next k
                            Else
                            'Do nothing
                            End If
                        End If
                    End If
                Next site
            Next j
        End If
    Next i

    For i = 0 To l_TestSeq_Num
        If i = 0 Then Trim_Flag = False
        If bTrimTestSeq(i) Then
            StoreName_Ary = Split(sTestStoreName_Ary(i), ",")
            For j = 0 To UBound(StoreName_Ary)
                Trim_Flag = Trim_Flag Or DecideTrim(i, j).Any(True)
            Next j
        End If
    Next i

    If Trim_Flag Then
        GoTo StartTrim
    Else
        TheExec.Flow.TestLimitIndex = lTestIimitIndex_TrimFinish
    End If
    
    Dim l_FinalStepCnt As Long
    l_FinalStepCnt = StepCount
    
'''    TheExec.Datalog.WriteComment ("**************** The Final Best Code Measurement ****************")
'''    Call LDO_Measurement_Process_Universal(pats(0), DigSrc_pin, code, vout, TrimCodeSize, NumberOfMeasV, l_StoreName_Max, MeasValue, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, sTestStoreName_Ary(), MeaV_WaitTime, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, CUS_Str_MainProgram, MSB_First_Flag, bTrimTestSeq, l_FinalStepCnt, sl_GrayCode, True)
    
    TheExec.Datalog.WriteComment ("**************** The Final Best Value & Best Code ****************")
    Dim sTemp() As String
    Dim sTemp_FinalTrimCode As String
    For i = 0 To l_TestSeq_Num
        If bTrimTestSeq(i) Then
            sTrim_Pin = sTestSeqPin_Ary(i)
            StoreName_Ary = Split(sTestStoreName_Ary(i), ",")
            For j = 0 To UBound(StoreName_Ary)
                For k = 0 To UBound(Split(sTrim_Pin, ","))
                    If j = k Then
                        PinName = Split(sTrim_Pin, ",")(k)
                        TestNameInput = Report_TName_From_Instance("V", PinName, vbNullString, i, 0)
                        TheExec.Flow.TestLimit resultVal:=BestVal(i).Math.Abs, Tname:=TestNameInput, PinName:=PinName, ForceResults:=tlForceFlow
          
                    End If
                Next k
            Next j
        Else
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
        End If
    Next i

    For i = 0 To l_TestSeq_Num
        If bTrimTestSeq(i) Then
            sTrim_Pin = sTestSeqPin_Ary(i)
            StoreName_Ary = Split(sTestStoreName_Ary(i), ",")
            For j = 0 To UBound(StoreName_Ary)
                For k = 0 To UBound(Split(sTrim_Pin, ","))
                    If j = k Then
                        PinName = Split(sTrim_Pin, ",")(k)
                        TestNameInput = Report_TName_From_Instance(CalcC, vbNullString, vbNullString, i, 0)
                        If UCase(CUS_Str_MainProgram) Like UCase("*BinToGray*") Then
                            TheExec.Flow.TestLimit resultVal:=sl_GrayCode(i, j), Tname:=TestNameInput, ForceResults:=tlForceFlow
                        Else
                            TheExec.Flow.TestLimit resultVal:=BestCode(i, j), Tname:=TestNameInput, ForceResults:=tlForceFlow
                        End If
                    End If
                Next k
            Next j
        End If
    Next i

    For i = 0 To l_TestSeq_Num
        If bTrimTestSeq(i) Then
            StoreName_Ary = Split(sTestStoreName_Ary(i), ",")
            For j = 0 To UBound(StoreName_Ary)
                For Each site In TheExec.sites
                    If UCase(CUS_Str_MainProgram) Like UCase("*BinToGray*") Then
                        TempVal = sl_GrayCode(i, j)
                    Else
                        TempVal = BestCode(i, j)
                    End If
                    
                    For k = 0 To TrimCodeSize - 1
                        FinalTrimCode_Array(k) = TempVal Mod 2
                        TempVal = TempVal \ 2
                    Next k
                    FinalTrimCode(i, j).data = FinalTrimCode_Array
					
                    If gl_Disable_HIP_debug_log = False Then            'Printing for store info
                        If UCase(CUS_Str_MainProgram) Like UCase("*BinToGray*") Then
                            TheExec.Datalog.WriteComment "Site : " & site & ", Store Value(GrayCode) : " & sl_GrayCode(i, j) & ", Binary Bits : " & TrimCodeSize & ", Store Name : " & StoreName_Ary(j)
                        Else
                            TheExec.Datalog.WriteComment "Site : " & site & ", Store Value : " & BestCode(i, j) & ", Binary Bits : " & TrimCodeSize & ", Store Name : " & StoreName_Ary(j)
                        End If
                    End If
                Next site
                Call StoreDataAllType(StoreName_Ary(j), FinalTrimCode(i, j))
                
                
            Next j
        End If
    Next i
    
    DebugPrintFunc Pat.value
    
    If Calc_Eqn <> "" And InStr(LCase(TestSequence), "p") = 0 Then
        Call ProcessCalcEquation(Calc_Eqn)
    End If
    Call HardIP_WriteFuncResult(, , glb_TestInstance)
    TheHdw.Alarms.Check
    Exit Function
    
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP", "LDO_Calibration") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function





'Public Function LDO_Calibration(Optional Pat As Pattern, Optional TestSequence As String, Optional MeasV_PinS As String, Optional MeaV_WaitTime As String, Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As String, Optional DigSrc_Equation As String, Optional digsrc_assignment As String, Optional TrimStoreName As String, Optional TrimTarget As Double, Optional TrimStart As Long, Optional TrimCodeSize As Long, Optional TrimMethod As String, Optional TrimStepSize As Double, Optional Validating_ As Boolean)
'    Dim site As Variant
'    Dim i As Integer
'    Dim j As Integer
'    Dim k As Integer
'    Dim pats() As String
'    Dim code() As New SiteLong: ReDim code(UBound(Split(TestSequence, ",")))
'    Dim BestCode() As New SiteLong: ReDim BestCode(UBound(Split(TestSequence, ",")))
'    Dim vout() As New SiteDouble: ReDim vout(UBound(Split(TestSequence, ",")))
'        For i = 0 To UBound(Split(TestSequence, ","))
'            For Each site In TheExec.sites.Active
'                code(i) = TrimStart
'                vout(i) = 0
'            Next site
'        Next i
'    Dim NumberOfMeasV As Integer: NumberOfMeasV = UBound(Split(TestSequence, ",")) + 1
'    Dim PreviousNegative() As New SiteBoolean: ReDim PreviousNegative(UBound(Split(TestSequence, ",")))
'    Dim PreviousPositive() As New SiteBoolean: ReDim PreviousPositive(UBound(Split(TestSequence, ",")))
'    Dim DecideTrim() As New SiteBoolean: ReDim DecideTrim(UBound(Split(TestSequence, ",")))
'    Dim Trim_Flag As Boolean
'        For i = 0 To UBound(Split(TestSequence, ","))
'            For Each site In TheExec.sites.Active
'                PreviousNegative(i) = False
'                PreviousPositive(i) = False
'                DecideTrim(i) = False
'            Next site
'        Next i
'
'    glb_TestInstance = vbNullString
'    glb_TestInstance = UCase(TheExec.DataManager.instancename)
'
'    Dim BlockName() As String: BlockName = Split(glb_TestInstance, "_")
'    Dim MeasValue() As New PinListData: ReDim MeasValue(NumberOfMeasV - 1)
'    Dim PreviousMeasValue() As New PinListData: ReDim PreviousMeasValue(NumberOfMeasV - 1)
'    Dim BestVal() As New PinListData: ReDim BestVal(NumberOfMeasV - 1)
'    Dim StepCount As Long: StepCount = 0
'    Dim TestNameInput As String
'    Dim PatCount As Long, PattArray() As String
'    Dim PreviousTargetCompare() As New SiteDouble: ReDim PreviousTargetCompare(UBound(Split(TestSequence, ",")))
'    Dim TrimStoreName_Array() As String: TrimStoreName_Array = Split(TrimStoreName, ",")
'    Dim PinName As String
'    Dim TempVal As Integer
'    Dim FinalTrimCode() As New DSPWave: ReDim FinalTrimCode(UBound(TrimStoreName_Array))
'    Dim FinalTrimCode_Array() As Long: ReDim FinalTrimCode_Array(TrimCodeSize - 1) As Long
'
'    Call ProcessInputToGLB(Pat, TestSequence, True, , , , , MeasV_PinS, , , , , , , , , , , , , DigSrc_pin, CStr(DigSrc_Sample_Size), CStr(DigSrc_Sample_Size), DigSrc_Equation, digsrc_assignment, , , , , , , , , , , , , , , , , , , , , , , , , , MeaV_WaitTime)
'
'    For i = 0 To NumberOfMeasV - 1
'        For j = 0 To UBound(Split(MeasV_PinS, ","))
'            PreviousMeasValue(i).AddPin (Split(MeasV_PinS, ",")(j))
'            PreviousMeasValue(i).Pins(Split(MeasV_PinS, ",")(j)).value = 0
'            MeasValue(i).AddPin (Split(MeasV_PinS, ",")(j))
'            MeasValue(i).Pins(Split(MeasV_PinS, ",")(j)).value = 0
'            BestVal(i).AddPin (Split(MeasV_PinS, ",")(j))
'            BestVal(i).Pins(Split(MeasV_PinS, ",")(j)).value = 0
'        Next j
'    Next i
'
'    Call GetFlowTName
'
'    If Validating_ Then
'        Call PrLoadPattern(Pat.value)
'        Exit Function    ' Exit after validation
'    End If
'
'    On Error GoTo errHandler
'
'    'Dim Inst_Name_Str As String: Inst_Name_Str = glb_TestInstance
'
'    If TheExec.DevChar.Setups.IsRunning Then
'        If TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes.Contains(tlDevCharShmooAxis_Y) Then
'            If gl_Flag_HardIP_Trim_Set_PrePoint And Not (gl_Flag_HardIP_Characterization_1stRun) Then
'                TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
'                Call TheExec.Overlays.ApplyUniformSpecToHW("XI0_Shmoo_Freq_VAR", TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_Y).value)
'            ElseIf gl_Flag_HardIP_Trim_Set_PostPoint Then
'                TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
'                Call TheExec.Overlays.ApplyUniformSpecToHW("XI0_Shmoo_Freq_VAR", 24000000#)
'            Else
'                TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
'            End If
'        End If
'    Else
'        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
'    End If
'    TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD
'
'    PATT_GetPatListFromPatternSet Pat.value, pats, PatCount
'
'    Dim TrimCodeValue_Min As Long, TrimCodeValue_Max As Long
'    TrimCodeValue_Min = 0
'    TrimCodeValue_Max = 2 ^ TrimCodeSize - 1
'
'
'    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("**************** The Measurement at Trim Start Point ****************")
'    Call LDO_Measurement_Process(pats(0), DigSrc_pin, code(), vout(), TrimCodeSize, NumberOfMeasV, MeasValue(), DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, TrimStoreName_Array(), MeaV_WaitTime)
'
'    BestCode = code
'    For Each site In TheExec.sites.Active
'        For i = 0 To NumberOfMeasV - 1
'            BestVal(i) = MeasValue(i)
'        Next i
'    Next site
'    For i = 0 To UBound(Split(TestSequence, ","))
'        If LCase(TrimMethod) = "linearsearch" Then
'            For Each site In TheExec.sites.Active
'                If vout(i).compare(LessThan, TrimTarget) Then
'                    code(i) = code(i) + 1
'                ElseIf vout(i).compare(GreaterThan, TrimTarget) Then
'                    code(i) = code(i) - 1
'                End If
'                DecideTrim(i) = True
'            Next site
'        Else
'            For Each site In TheExec.sites.Active
'                code(i) = code(i) + Fix((TrimTarget - vout(i)) / TrimStepSize)
'                DecideTrim(i) = True
'            Next site
'        End If
'    Next i
'StartTrim:
'    For i = 0 To UBound(Split(TestSequence, ","))
'        If i = 0 Then
'            Trim_Flag = DecideTrim(i).Any(True)
'        Else
'            Trim_Flag = Trim_Flag Or DecideTrim(i).Any(True)
'        End If
'    Next i
'        If Trim_Flag Then
'            StepCount = StepCount + 1
'
'            If gl_Disable_HIP_debug_log = False Then
'                If right(CStr(StepCount), 1) = "1" Then
'                    TheExec.Datalog.WriteComment ("**************** The " & StepCount & "st Trim Process ****************")
'                ElseIf right(CStr(StepCount), 1) = "2" Then
'                    TheExec.Datalog.WriteComment ("**************** The " & StepCount & "nd Trim Process ****************")
'                ElseIf right(CStr(StepCount), 1) = "3" Then
'                    TheExec.Datalog.WriteComment ("**************** The " & StepCount & "rd Trim Process ****************")
'                Else
'                    TheExec.Datalog.WriteComment ("**************** The " & StepCount & "th Trim Process ****************")
'                End If
'            End If
'
'            Call LDO_Measurement_Process(pats(0), DigSrc_pin, code(), vout(), TrimCodeSize, NumberOfMeasV, MeasValue(), DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, TrimStoreName_Array(), MeaV_WaitTime)
'        End If
'    For k = 0 To UBound(Split(TestSequence, ","))
'        For Each site In TheExec.sites.Active
'            If DecideTrim(k) Then
'                If StepCount > (TrimCodeValue_Max - TrimCodeValue_Min) Then
'                    BestCode = code
'                    For j = 0 To UBound(Split(MeasV_PinS, ","))
'                        BestVal(k).Pins(j).value = MeasValue(k).Pins(j).value
'                    Next j
'                    DecideTrim(k) = False
'                ElseIf code(k).compare(GreaterThan, TrimCodeValue_Max) Then
'                    code(k) = TrimCodeValue_Max
'                    DecideTrim(k) = True
'                ElseIf code(k).compare(LessThan, TrimCodeValue_Min) Then
'                    code(k) = TrimCodeValue_Min
'                    DecideTrim(k) = True
'                ElseIf code(k).compare(EqualTo, TrimCodeValue_Max) Or code(k).compare(EqualTo, TrimCodeValue_Min) Then
'                    BestCode(k) = code(k)
'                        For j = 0 To UBound(Split(MeasV_PinS, ","))
'                            BestVal(k).Pins(j).value = MeasValue(k).Pins(j).value
'                        Next j
'                    DecideTrim(k) = False
'                ElseIf vout(k).compare(LessThan, TrimTarget) And PreviousPositive(k) Then
'                        BestCode(k) = code(k) + 1
'                            For j = 0 To UBound(Split(MeasV_PinS, ","))
'                                BestVal(k).Pins(j).value = PreviousMeasValue(k).Pins(j).value
'                            Next j
'                    DecideTrim(k) = False
'                ElseIf vout(k).compare(LessThan, TrimTarget) And Not (PreviousPositive(k)) Then
'                    code(k) = code(k) + 1
'                    PreviousNegative(k) = True
'                    PreviousTargetCompare(k) = vout(k).Subtract(TrimTarget).Abs
'                        For j = 0 To UBound(Split(MeasV_PinS, ","))
'                            PreviousMeasValue(k).Pins(j).value = MeasValue(k).Pins(j).value
'                        Next j
'                    DecideTrim(k) = True
'                ElseIf vout(k).compare(GreaterThan, TrimTarget) And PreviousNegative(k) Then
'                        BestCode(k) = code(k)
'                            For j = 0 To UBound(Split(MeasV_PinS, ","))
'                                BestVal(k).Pins(j).value = MeasValue(k).Pins(j).value
'                            Next j
'                    DecideTrim(k) = False
'                ElseIf vout(k).compare(GreaterThan, TrimTarget) And Not (PreviousNegative(k)) Then
'                    code(k) = code(k) - 1
'                    PreviousPositive(k) = True
'                    PreviousTargetCompare(k) = vout(k).Subtract(TrimTarget).Abs
'                        For j = 0 To UBound(Split(MeasV_PinS, ","))
'                            PreviousMeasValue(k).Pins(j).value = MeasValue(k).Pins(j).value
'                        Next j
'                    DecideTrim(k) = True
'                End If
'            End If
'        Next site
'    Next k
'
'    For i = 0 To UBound(Split(TestSequence, ","))
'        If i = 0 Then
'            Trim_Flag = DecideTrim(i).Any(True)
'        Else
'            Trim_Flag = Trim_Flag Or DecideTrim(i).Any(True)
'        End If
'    Next i
'
'    If Trim_Flag Then GoTo StartTrim
'
'
'    For i = 0 To NumberOfMeasV - 1
'        For j = 0 To UBound(Split(MeasV_PinS, ","))
'            PinName = Split(MeasV_PinS, ",")(j)
'            TestNameInput = Report_TName_From_Instance("V", PinName, vbNullString, i, 0)
'            TheExec.Flow.TestLimit resultVal:=BestVal(i), Tname:=TestNameInput, PinName:=Split(MeasV_PinS, ",")(j), ForceResults:=tlForceFlow
'        Next j
'    Next i
'    For i = 0 To UBound(Split(TestSequence, ","))
'        TestNameInput = Report_TName_From_Instance("C", vbNullString, BlockName(0) & "Trim", i, 0)
'        TheExec.Flow.TestLimit resultVal:=BestCode(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
'    Next i
'    For i = 0 To UBound(Split(TestSequence, ","))
'        For Each site In TheExec.sites
'            TempVal = BestCode(i)
'            For j = 0 To TrimCodeSize - 1
'                FinalTrimCode_Array(j) = TempVal Mod 2
'                TempVal = TempVal \ 2
'            Next j
'            FinalTrimCode(i).data = FinalTrimCode_Array
'        Next site
'        Call StoreDataAllType(TrimStoreName_Array(i), FinalTrimCode(i))
'    Next i
'    DebugPrintFunc Pat.value
'
'    Call HardIP_WriteFuncResult(, , glb_TestInstance)
'    'Alarm check From Sicily,20200423, Oscar
'    ' Check implicit alarms
'    TheHdw.Alarms.Check
'
'    Exit Function
'
'errHandler:
'    TheExec.Datalog.WriteComment "error in LDO_Calibration"
'    If AbortTest Then Exit Function Else Resume Next
'End Function
'
'
'Public Function LDO_Calibration_Universal(Optional Pat As Pattern, Optional TestSequence As String, Optional MeasV_PinS As String, _
'            Optional MeaV_WaitTime As String, Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As String, Optional DigSrc_Equation As String, Optional digsrc_assignment As String, _
'            Optional TrimStoreName As String, Optional TrimTarget As Double, Optional TrimStart As Long, Optional TrimCodeSize As Long, Optional TrimMethod As String, Optional TrimStepSize As Double, _
'            Optional ForceV_Val As String, Optional MeasI_Range As String, Optional Interpose_PrePat As String, Optional Interpose_PreMeas As String, Optional Interpose_PostTest As String, Optional RAK_Flag As Enum_RAK, _
'            Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, Optional CUS_Str_DigCapData As String, Optional CUS_Str_MainProgram As String, Optional MSB_First_Flag As Boolean, Optional b_MinimumBestVal As Boolean, Optional Validating_ As Boolean)
'
'    On Error GoTo errHandler
'
'    If Validating_ Then
'        Call PrLoadPattern(Pat.value)
'        Exit Function    ' Exit after validation
'    End If
'
'    glb_TestInstance = ""
'    glb_TestInstance = UCase(TheExec.DataManager.instancename)
'
'    Dim site As Variant
'    Dim i As Integer
'    Dim j As Integer
'    Dim k As Integer
'    Dim l_TestSeq_Num As Long
'    Dim l_StoreName_Num As Long
'    Dim l_StoreName_Max As Long
'
'    l_TestSeq_Num = UBound(Split(TestSequence, ","))
'    l_StoreName_Num = UBound(Split(TrimStoreName, "+"))
'    Dim StoreName_Ary() As String
'
'    Dim sTestSeqPin_Ary() As String
'    Dim sTestSequence_Ary() As String
'    Dim sTestStoreName_Ary() As String
'
'    sTestSeqPin_Ary = Split(MeasV_PinS, "+")
'    sTestSequence_Ary = Split(TestSequence, ",")
'    sTestStoreName_Ary = Split(TrimStoreName, "+")
'
'    l_StoreName_Max = 0
'    For i = 0 To l_TestSeq_Num
'        If sTestStoreName_Ary(i) <> "" Then
'            StoreName_Ary = Split(sTestStoreName_Ary(i), ",")
'            If l_StoreName_Max < UBound(StoreName_Ary) Then
'                l_StoreName_Max = UBound(StoreName_Ary)
'            End If
'        End If
'    Next i
'
'    If l_TestSeq_Num <> l_StoreName_Num Then
'        TheExec.Datalog.WriteComment "Test Sequence and TrimStoreName are not match, Please check the argument."
'        Exit Function
'    End If
'
'    Dim pats() As String
'    Dim NumberOfMeasV As Integer: NumberOfMeasV = l_TestSeq_Num + 1
'
'    '''----------TrimCode Variable----------
'    Dim code() As New SiteLong: ReDim code(l_TestSeq_Num, l_StoreName_Max)
'    Dim BestCode() As New SiteLong: ReDim BestCode(l_TestSeq_Num, l_StoreName_Max)
'    Dim sl_GrayCode() As New SiteLong: ReDim sl_GrayCode(l_TestSeq_Num, l_StoreName_Max)
'
'    Dim BestVal() As New PinListData: ReDim BestVal(l_TestSeq_Num)
'    Dim pl_PreviousMeasValue() As New PinListData: ReDim pl_PreviousMeasValue(l_TestSeq_Num)
'
'    Dim vout() As New SiteDouble: ReDim vout(l_TestSeq_Num, l_StoreName_Max)
'
'    Dim DecideTrim() As New SiteBoolean: ReDim DecideTrim(l_TestSeq_Num, l_StoreName_Max)
'    Dim PreviousNegative() As New SiteBoolean: ReDim PreviousNegative(l_TestSeq_Num, l_StoreName_Max)
'    Dim PreviousPositive() As New SiteBoolean: ReDim PreviousPositive(l_TestSeq_Num, l_StoreName_Max)
'
'    Dim FinalTrimCode() As New DSPWave: ReDim FinalTrimCode(l_TestSeq_Num, l_StoreName_Max)
'    Dim FinalTrimCode_Array() As Long: ReDim FinalTrimCode_Array(TrimCodeSize - 1) As Long
'    '' GrayCode
'    Dim srcwave_array_graycode() As Long: ReDim srcwave_array_graycode(TrimCodeSize - 1) As Long
'
'    Dim PresentMeasValue() As New SiteDouble: ReDim PresentMeasValue(l_TestSeq_Num, l_StoreName_Max)
'    Dim PreviousMeasValue() As New SiteDouble: ReDim PreviousMeasValue(l_TestSeq_Num, l_StoreName_Max)
'    '''----------TrimCode Variable----------
'    Dim Trim_Flag As Boolean
'
'    Dim BlockName() As String: BlockName = Split(glb_TestInstance, "_")
'    Dim MeasValue() As New PinListData: ReDim MeasValue(NumberOfMeasV - 1)
'
'    Dim StepCount As Long: StepCount = 0
'    Dim TestNameInput As String
'    Dim PatCount As Long, PattArray() As String
'
'    Dim PinName As String
'    Dim TempVal As Integer
'
'    Dim bTrimTestSeq() As Boolean
'    Dim sTrim_Pin As String
'    Dim sTestSeq_V As String
'    Dim sTestSeq_I As String
'    Dim sTestSeq_V_Ary() As String
'    Dim sTestSeq_I_Ary() As String
'    Dim Calc_Eqn As String
'    Dim CUS_Str_DigSrcData As String
'    Dim MeasI_pinS As String
'    Dim Meas_StoreName As String
'    Dim DigSrc_FlowForLoopIntegerName As String
'
'
'    ReDim sTestSeq_V_Ary(l_TestSeq_Num)
'    ReDim sTestSeq_I_Ary(l_TestSeq_Num)
'    For i = 0 To l_TestSeq_Num
'        If UCase(sTestSequence_Ary(i)) Like "V*" Then
'            If sTestSequence_Ary(i) = "" Then
'                sTestSeq_V_Ary(i) = ""
'            Else
'                sTestSeq_V_Ary(i) = sTestSeqPin_Ary(i)
'            End If
'        ElseIf UCase(sTestSequence_Ary(i)) Like "I*" Or UCase(sTestSequence_Ary(i)) Like "R*" Then
'            If sTestSequence_Ary(i) = "" Then
'                sTestSeq_I_Ary(i) = ""
'            Else
'                sTestSeq_I_Ary(i) = sTestSeqPin_Ary(i)
'            End If
'        End If
'    Next i
'
'    sTestSeq_V = Join(sTestSeq_V_Ary, "+")
'    sTestSeq_I = Join(sTestSeq_I_Ary, "+")
'
'    ReDim bTrimTestSeq(l_TestSeq_Num)
'
'    For i = 0 To l_StoreName_Num
'        If sTestStoreName_Ary(i) <> "" Then
'            bTrimTestSeq(i) = True
'        Else
'            bTrimTestSeq(i) = False
'        End If
'    Next i
'
'    For i = 0 To l_TestSeq_Num
'        If bTrimTestSeq(i) Then
'            StoreName_Ary = Split(sTestStoreName_Ary(i), ",")
'            For j = 0 To UBound(StoreName_Ary)
'                For Each site In TheExec.sites.Active
'                    code(i, j) = TrimStart
'                    vout(i, j) = 0
'                    sl_GrayCode(i, j) = 0
'                    PreviousNegative(i, j) = False
'                    PreviousPositive(i, j) = False
'                    DecideTrim(i, j) = False
'                    PresentMeasValue(i, j) = 0
'                    PreviousMeasValue(i, j) = 0
'                Next site
'            Next j
'        End If
'    Next i
'
'    Call Reg_Assign_Processing(DigSrc_Equation, digsrc_assignment, CUS_Str_DigCapData, Calc_Eqn, CUS_Str_DigSrcData, MeasV_PinS, Interpose_PreMeas, Meas_StoreName, DigSrc_FlowForLoopIntegerName)
'
'    Call ProcessInputToGLB(Pat, TestSequence, True, , , , , sTestSeq_V, , , , , , , sTestSeq_I, MeasI_Range, , DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, DigSrc_pin, CStr(DigSrc_Sample_Size), CStr(DigSrc_Sample_Size), DigSrc_Equation, digsrc_assignment, , , , CUS_Str_MainProgram, , , , , , , , , , , , , , , , Interpose_PrePat, Interpose_PreMeas, Interpose_PostTest, , ForceV_Val, , , RAK_Flag, MeaV_WaitTime)
'
'    For i = 0 To l_TestSeq_Num
'        If sTestSeqPin_Ary(i) <> "" Then
'            For j = 0 To UBound(Split(sTestSeqPin_Ary(i), ","))
'                MeasValue(i).AddPin (Split(sTestSeqPin_Ary(i), ",")(j))
'                MeasValue(i).Pins(Split(sTestSeqPin_Ary(i), ",")(j)).value = 0
'                BestVal(i).AddPin (Split(sTestSeqPin_Ary(i), ",")(j))
'                BestVal(i).Pins(Split(sTestSeqPin_Ary(i), ",")(j)).value = 0
'                pl_PreviousMeasValue(i).AddPin (Split(sTestSeqPin_Ary(i), ",")(j))
'                pl_PreviousMeasValue(i).Pins(Split(sTestSeqPin_Ary(i), ",")(j)).value = 0
'            Next j
'        End If
'    Next i
'
'    Call GetFlowTName
'
'    If TheExec.DevChar.Setups.IsRunning Then
'        If TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes.Contains(tlDevCharShmooAxis_Y) Then
'            If gl_Flag_HardIP_Trim_Set_PrePoint And Not (gl_Flag_HardIP_Characterization_1stRun) Then
'                TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
'                Call TheExec.Overlays.ApplyUniformSpecToHW("XI0_Shmoo_Freq_VAR", TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_Y).value)
'            ElseIf gl_Flag_HardIP_Trim_Set_PostPoint Then
'                TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
'                Call TheExec.Overlays.ApplyUniformSpecToHW("XI0_Shmoo_Freq_VAR", 24000000#)
'            Else
'                TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
'            End If
'        End If
'    Else
'        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
'    End If
'
'    TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD
'
'    If BurstYesPatDict.Exists(LCase(Pat.value)) Then
'        Call PatternBurstCheckAndSplit(Pat.value, pats, PatCount) '' 220708 for palma
'    Else
'        PATT_GetPatListFromPatternSet Pat.value, pats, PatCount
'    End If
'
'    Dim TrimCodeValue_Min As Long, TrimCodeValue_Max As Long
'    TrimCodeValue_Min = 0
'    TrimCodeValue_Max = 2 ^ TrimCodeSize - 1
'
'    Dim lTestIimitIndex As Long
'    Dim lTestIimitIndex_TrimFinish As Long
'    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("**************** The Measurement at Trim Start Point ****************")
'
'    lTestIimitIndex = TheExec.Flow.TestLimitIndex
'    Call LDO_Measurement_Process_Universal(pats(0), DigSrc_pin, code, vout, TrimCodeSize, NumberOfMeasV, l_StoreName_Max, MeasValue, DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, sTestStoreName_Ary(), MeaV_WaitTime, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, CUS_Str_MainProgram, MSB_First_Flag, bTrimTestSeq, StepCount, sl_GrayCode, False)
'    TheExec.Flow.TestLimitIndex = lTestIimitIndex
'
'    For Each site In TheExec.sites.Active
'        For i = 0 To l_TestSeq_Num
'            BestVal(i) = MeasValue(i)
'        Next i
'    Next site
'
'    For i = 0 To l_TestSeq_Num
'        If bTrimTestSeq(i) Then
'            StoreName_Ary = Split(sTestStoreName_Ary(i), ",")
'            For j = 0 To UBound(StoreName_Ary)
'            BestCode(i, j) = code(i, j)
'                If LCase(TrimMethod) = "linearsearch-all-combination" Then
'                    For Each site In TheExec.sites.Active
'                        code(i, j) = code(i, j) + 1
'                        PreviousMeasValue(i, j) = vout(i, j).Subtract(TrimTarget)
'                        PreviousMeasValue(i, j) = PreviousMeasValue(i, j).Abs
'                        DecideTrim(i, j) = True
'                    Next site
'                ElseIf LCase(TrimMethod) = "linearsearch" Then
'
'                    For Each site In TheExec.sites.Active
'                        If vout(i, j).compare(LessThan, TrimTarget) Then
'                            code(i, j) = code(i, j) + 1
'                        ElseIf vout(i, j).compare(GreaterThan, TrimTarget) Then
'                            If code(i, j) > 0 Then
'                            code(i, j) = code(i, j) - 1
'                            End If
'                        End If
'                        DecideTrim(i, j) = True
'                    Next site
'                Else
'                    For Each site In TheExec.sites.Active
'                        code(i, j) = code(i, j) + Fix((TrimTarget - vout(i, j)) / TrimStepSize)
'                        DecideTrim(i, j) = True
'                    Next site
'                End If
'            Next j
'        End If
'    Next i
'
'    For i = 0 To l_TestSeq_Num
'        If i = 0 Then Trim_Flag = False
'        If bTrimTestSeq(i) Then
'            StoreName_Ary = Split(sTestStoreName_Ary(i), ",")
'            For j = 0 To UBound(StoreName_Ary)
'                Trim_Flag = Trim_Flag Or DecideTrim(i, j).Any(True)
'            Next j
'        End If
'    Next i
'
'StartTrim:
'    If Trim_Flag Then
'        StepCount = StepCount + 1
'
'        If gl_Disable_HIP_debug_log = False Then
'            If right(CStr(StepCount), 1) = "1" Then
'                TheExec.Datalog.WriteComment ("**************** The " & StepCount & "st Trim Process ****************")
'            ElseIf right(CStr(StepCount), 1) = "2" Then
'                TheExec.Datalog.WriteComment ("**************** The " & StepCount & "nd Trim Process ****************")
'            ElseIf right(CStr(StepCount), 1) = "3" Then
'                TheExec.Datalog.WriteComment ("**************** The " & StepCount & "rd Trim Process ****************")
'            Else
'                TheExec.Datalog.WriteComment ("**************** The " & StepCount & "th Trim Process ****************")
'            End If
'        End If
'
'        lTestIimitIndex = TheExec.Flow.TestLimitIndex
'        Call LDO_Measurement_Process_Universal(pats(0), DigSrc_pin, code, vout, TrimCodeSize, NumberOfMeasV, l_StoreName_Max, MeasValue, DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, sTestStoreName_Ary(), MeaV_WaitTime, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, CUS_Str_MainProgram, MSB_First_Flag, bTrimTestSeq, StepCount, sl_GrayCode, False)
'        ''------- Test Limit Index Value for Trimming Finsh -------
'        lTestIimitIndex_TrimFinish = TheExec.Flow.TestLimitIndex
'        ''------- Test Limit Index Value for Trimming Finsh -------
'        TheExec.Flow.TestLimitIndex = lTestIimitIndex
'    End If
'
'    Dim sd_TrimTarget As New SiteDouble
'    Dim pl_TrimTarget As New PinListData
'    For i = 0 To l_TestSeq_Num
'        If bTrimTestSeq(i) Then
'            StoreName_Ary = Split(sTestStoreName_Ary(i), ",")
'            For j = 0 To UBound(StoreName_Ary)
'                sTrim_Pin = sTestSeqPin_Ary(i)
'                For Each site In TheExec.sites.Active
'                    If DecideTrim(i, j) Then
'                        PresentMeasValue(i, j) = vout(i, j).Subtract(TrimTarget)
'                        PresentMeasValue(i, j) = PresentMeasValue(i, j).Abs
'
'                        If LCase(TrimMethod) = "linearsearch-all-combination" Then
'                            If code(i, j) < TrimCodeValue_Max Then
'                                If PresentMeasValue(i, j).compare(LessThan, PreviousMeasValue(i, j)) Then
'                                    BestCode(i, j) = code(i, j)
'    '                                DecideTrim(i, j) = False
'                                    For k = 0 To UBound(Split(sTrim_Pin, ","))
'                                        If j = k Then BestVal(i).Pins(k).value = MeasValue(i).Pins(k).value
'                                    Next k
'                                    PreviousMeasValue(i, j) = PresentMeasValue(i, j)
'                                End If
'
'                                code(i, j) = code(i, j) + 1
'                            Else
'                                DecideTrim(i, j) = False
'                            End If
'                        Else
'                            If StepCount > (TrimCodeValue_Max - TrimCodeValue_Min) Then
'                                BestCode = code
'                                DecideTrim(i, j) = False
'                                For k = 0 To UBound(Split(sTrim_Pin, ","))
'                                    If j = k Then BestVal(i).Pins(k).value = MeasValue(i).Pins(k).value
'                                Next k
'                            ElseIf code(i, j).compare(GreaterThan, TrimCodeValue_Max) Then
'                            code(i, j) = TrimCodeValue_Max
'                            PreviousMeasValue(i, j) = vout(i, j)
'                            DecideTrim(i, j) = True
'                        ElseIf code(i, j).compare(LessThan, TrimCodeValue_Min) Then
'                            code(i, j) = TrimCodeValue_Min
'                            PreviousMeasValue(i, j) = vout(i, j)
'                            DecideTrim(i, j) = True
'                        ElseIf code(i, j).compare(EqualTo, TrimCodeValue_Max) Or code(i, j).compare(EqualTo, TrimCodeValue_Min) Then
'                            BestCode(i, j) = code(i, j)
'                            DecideTrim(i, j) = False
'                            For k = 0 To UBound(Split(sTrim_Pin, ","))
'                                If j = k Then BestVal(i).Pins(k).value = MeasValue(i).Pins(k).value
'                            Next k
'                        ElseIf vout(i, j).compare(LessThan, TrimTarget) And PreviousPositive(i, j) Then
'                            If b_MinimumBestVal Then
'                                '''----------Minimum Error Difference----------
'                                If PreviousMeasValue(i, j) > PresentMeasValue(i, j) Then
'                                    BestCode(i, j) = code(i, j) + 1
'                                Else
'                                    BestCode(i, j) = code(i, j)
'                                End If
'                                '''----------Minimum Error Difference----------
'                            Else
'                                BestCode(i, j) = code(i, j) + 1
'                            End If
'                            DecideTrim(i, j) = False
'                            For k = 0 To UBound(Split(sTrim_Pin, ","))
'                                If j = k Then BestVal(i).Pins(k).value = pl_PreviousMeasValue(i).Pins(k).value
'                            Next k
'                        ElseIf vout(i, j).compare(LessThan, TrimTarget) And Not (PreviousPositive(i, j)) Then
'                            code(i, j) = code(i, j) + 1
'                            PreviousMeasValue(i, j) = vout(i, j).Subtract(TrimTarget)
'                            PreviousMeasValue(i, j) = PreviousMeasValue(i, j).Abs
'                            PreviousNegative(i, j) = True
'                            DecideTrim(i, j) = True
'                            For k = 0 To UBound(Split(sTrim_Pin, ","))
'                                If j = k Then pl_PreviousMeasValue(i).Pins(k).value = MeasValue(i).Pins(k).value
'                            Next k
'                        ElseIf vout(i, j).compare(GreaterThan, TrimTarget) And PreviousNegative(i, j) Then
'                            If b_MinimumBestVal Then
'                                '''----------Minimum Error Difference----------
'                                If PreviousMeasValue(i, j) > PresentMeasValue(i, j) Then
'                                    BestCode(i, j) = code(i, j)
'                                Else
'                                    BestCode(i, j) = code(i, j) - 1
'                                End If
'                                '''----------Minimum Error Difference----------
'                            Else
'                                BestCode(i, j) = code(i, j)
'                            End If
'                            DecideTrim(i, j) = False
'                            For k = 0 To UBound(Split(sTrim_Pin, ","))
'                                If j = k Then BestVal(i).Pins(k).value = MeasValue(i).Pins(k).value
'                                Next k
'                            ElseIf vout(i, j).compare(GreaterThan, TrimTarget) And Not (PreviousNegative(i, j)) Then
'                                code(i, j) = code(i, j) - 1
'                                PreviousMeasValue(i, j) = vout(i, j).Subtract(TrimTarget)
'                                PreviousMeasValue(i, j) = PreviousMeasValue(i, j).Abs
'                                PreviousPositive(i, j) = True
'                                DecideTrim(i, j) = True
'                                For k = 0 To UBound(Split(sTrim_Pin, ","))
'                                    If j = k Then pl_PreviousMeasValue(i).Pins(k).value = MeasValue(i).Pins(k).value
'                                Next k
'                            End If
'                        End If
'                    End If
'                Next site
'            Next j
'        End If
'    Next i
'
'    For i = 0 To l_TestSeq_Num
'        If i = 0 Then Trim_Flag = False
'        If bTrimTestSeq(i) Then
'            StoreName_Ary = Split(sTestStoreName_Ary(i), ",")
'            For j = 0 To UBound(StoreName_Ary)
'                Trim_Flag = Trim_Flag Or DecideTrim(i, j).Any(True)
'            Next j
'        End If
'    Next i
'
'    If Trim_Flag Then
'        GoTo StartTrim
'    Else
'        TheExec.Flow.TestLimitIndex = lTestIimitIndex_TrimFinish
'    End If
'
'    Dim l_FinalStepCnt As Long
'    l_FinalStepCnt = StepCount
'
''''    TheExec.Datalog.WriteComment ("**************** The Final Best Code Measurement ****************")
''''    Call LDO_Measurement_Process_Universal(pats(0), DigSrc_pin, code, vout, TrimCodeSize, NumberOfMeasV, l_StoreName_Max, MeasValue, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, sTestStoreName_Ary(), MeaV_WaitTime, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, CUS_Str_MainProgram, MSB_First_Flag, bTrimTestSeq, l_FinalStepCnt, sl_GrayCode, True)
'
'    TheExec.Datalog.WriteComment ("**************** The Final Best Value & Best Code ****************")
'    Dim sTemp() As String
'    Dim sTemp_FinalTrimCode As String
'    For i = 0 To l_TestSeq_Num
'        If bTrimTestSeq(i) Then
'            sTrim_Pin = sTestSeqPin_Ary(i)
'            StoreName_Ary = Split(sTestStoreName_Ary(i), ",")
'            For j = 0 To UBound(StoreName_Ary)
'                For k = 0 To UBound(Split(sTrim_Pin, ","))
'                    If j = k Then
'                        PinName = Split(sTrim_Pin, ",")(k)
'                        TestNameInput = Report_TName_From_Instance("V", PinName, "", i, 0)
'                        TheExec.Flow.TestLimit resultVal:=BestVal(i).Math.Abs, Tname:=TestNameInput, PinName:=PinName, ForceResults:=tlForceFlow
'
'                    End If
'                Next k
'            Next j
'        Else
'            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
'        End If
'    Next i
'
'    For i = 0 To l_TestSeq_Num
'        If bTrimTestSeq(i) Then
'            sTrim_Pin = sTestSeqPin_Ary(i)
'            StoreName_Ary = Split(sTestStoreName_Ary(i), ",")
'            For j = 0 To UBound(StoreName_Ary)
'                For k = 0 To UBound(Split(sTrim_Pin, ","))
'                    If j = k And gl_Disable_HIP_debug_log = False Then
'                        PinName = Split(sTrim_Pin, ",")(k)
'                        TestNameInput = Report_TName_From_Instance(CalcC, "", "", i, 0)
'                        If UCase(CUS_Str_MainProgram) Like UCase("*BinToGray*") Then
'                            TheExec.Flow.TestLimit resultVal:=sl_GrayCode(i, j), Tname:=TestNameInput, ForceResults:=tlForceFlow
'                        Else
'                            TheExec.Flow.TestLimit resultVal:=BestCode(i, j), Tname:=TestNameInput, ForceResults:=tlForceFlow
'                        End If
'                    End If
'                Next k
'            Next j
'        End If
'    Next i
'
'    For i = 0 To l_TestSeq_Num
'        If bTrimTestSeq(i) Then
'            StoreName_Ary = Split(sTestStoreName_Ary(i), ",")
'            For j = 0 To UBound(StoreName_Ary)
'                For Each site In TheExec.sites
'                    TempVal = BestCode(i, j)
'                    For k = 0 To TrimCodeSize - 1
'                        FinalTrimCode_Array(k) = TempVal Mod 2
'                        TempVal = TempVal \ 2
'                    Next k
'                    FinalTrimCode(i, j).data = FinalTrimCode_Array
'
'                    If UCase(CUS_Str_MainProgram) Like UCase("*BinToGray*") Then
'                        ''' GrayCode
'                        For k = 0 To TrimCodeSize - 1
'                            If k = UBound(FinalTrimCode_Array) Then
'                                srcwave_array_graycode(k) = FinalTrimCode_Array(k)
'                            Else
'                                If FinalTrimCode_Array(k + 1) = FinalTrimCode_Array(k) Then
'                                    srcwave_array_graycode(k) = 0
'                                Else
'                                    srcwave_array_graycode(k) = 1
'                                End If
'                            End If
'                            'sl_GrayCode(z) = sl_GrayCode(z).Add(srcwave_array_graycode(j) * 2 ^ j)
'                        Next k
'                        ''' GrayCode
'                        FinalTrimCode(i, j).data = srcwave_array_graycode
'                    End If
'                Next site
'                Call StoreDataAllType(StoreName_Ary(j), FinalTrimCode(i, j))
'
'
'            Next j
'        End If
'    Next i
'
'    DebugPrintFunc Pat.value
'
'    Call HardIP_WriteFuncResult(, , glb_TestInstance)
'    TheHdw.Alarms.Check
'    Exit Function
'
'errHandler:
'    TheExec.Datalog.WriteComment "error in LDO_Calibration_Universal"
'    If AbortTest Then Exit Function Else Resume Next
'End Function


Public Function HIP_TTR_Enable_Control()
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
'[20240207][All][Brian] conment for fix and optimize flag (PTR_InstanceNameIsTINameOnly issue) move to HIP_TTR_Enable_Control
    If HIP_TTR_Enable_Control_Flag = False Then
        EnableDigitalTestLimitTTR = TheExec.Flow.enableWord("Enable_HardIP_DigitalTestLimitTTR")            
        EnableAnalogMuxOutTTR = TheExec.Flow.enableWord("Enable_HardIP_AnalogMuxOutTTR")                     
        EnableHardIPTnameConstructionTTR = TheExec.Flow.enableWord("Enable_HardIP_TnameConstructionTTR")        
        EnableFieldProcesingTTR = TheExec.Flow.enableWord("Enable_HardIP_FieldProcesingTTR")
        gl_Disable_HIP_debug_log = TheExec.Flow.enableWord("Enable_HardIP_Debug_log_disable_TTR")
        glb_Disable_CurrRangeSetting_Print = gl_Disable_HIP_debug_log
        EnableHardIPDigCapsdisableTTR = TheExec.Flow.enableWord("Enable_HardIP_DigCaps_disable_TTR")
        HIP_TTR_Enable_Control_Flag = True
    End If
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP", "HIP_TTR_Enable_Control") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function VTHSENSOR_Calibration(Optional patset As Pattern, Optional TestSequence As String, Optional CPUA_Flag_In_Pat As Boolean, _
Optional MeasI_PinS As String, Optional MeasI_Range As String, Optional MeasI_WaitTime As String, Optional Sweep_V_pin As String, _
Optional Start_V As String, Optional End_V As String, Optional Step_V As String, Optional TrimTarget As String, Optional Increase_TrimV As Boolean, Optional TrimV_mode As String, _
Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As String, Optional DigSrc_Sample_Size As String, _
Optional DigSrc_Equation As String, Optional DigSrc_Assignment As String, Optional DigSrc_FlowForLoopIntegerName As String = vbNullString, _
Optional CUS_Str_MainProgram As String = vbNullString, Optional CUS_Str_DigSrcData As String = vbNullString, _
Optional Meas_StoreName As String, Optional Interpose_PrePat As String, Optional Interpose_PreMeas As String, _
Optional Interpose_PostMeas As String, Optional Interpose_PostTest As String, Optional Validating_ As Boolean) As Long
        
    If Validating_ Then
        Call PrLoadPattern(patset.value)
        Exit Function    ' Exit after validation
    End If

    gl_Sweep_V_pin = Sweep_V_pin 'Tname UD(8) using
    'From T-Don
    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instanceName)
    
    ByPassTestLimit = True

    m_InstanceName = LCase(glb_TestInstance)

    Dim PatCount As Long
    Dim PattArray() As String
    
    Call HardIP_InitialSetupForPatgen

    Dim i As Long, j As Long, k As Long
    Dim TestOptLen As Integer
    Dim TestSequenceArray() As String, MeasPinAry_I() As String, MeasPinAry_IRange() As String
    Dim Ts As Variant, TestOption As Variant, site As Variant
    Dim TestSeqNum As Integer
    Dim MeasureI_pin As New PinList
    Dim MeasureI_Pin_CurrentRange As String
    Dim TestNum As Long
    Dim InDSPWave As New DSPWave, OutDspWave As New DSPWave
    Dim patt As Variant
    Dim Pat As String
    Dim restore_Flag As Boolean
    
    ''20160906 - Return measurement to directionary if needed
    Dim Interpose_PreMeas_Ary() As String
    Dim Interpose_PostMeas_Ary() As String
''    Dim RTN_InterposeString As String
    
    On Error GoTo errHandler

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    Dim Loop_Idx As Long
    Dim Loop_count As Long
    Dim Loop_Init As Long
    Dim Loop_Max As Long
    Dim Loop_Step As Long


    Loop_Idx = 0
    Loop_Init = 0
    Loop_Max = 0
    Loop_Step = 1
    
    Call ProcessInputToGLB(patset, TestSequence, CPUA_Flag_In_Pat, , , , , , , , , , , , MeasI_PinS, MeasI_Range, _
    MeasI_WaitTime, , , , DigSrc_pin, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, DigSrc_FlowForLoopIntegerName, , , CUS_Str_MainProgram, , _
    CUS_Str_DigSrcData, , , , , , , , , , , , Meas_StoreName, , Interpose_PrePat, Interpose_PreMeas, Interpose_PostTest, , , , , , , Interpose_PostMeas)
                                                                                                                                                               
    
    If TestSequence = "" Then                       '20170714
        ReDim TestSequenceArray(0) As String
        TestSequenceArray(0) = TestSequence
    Else
        TestSequenceArray = Split(TestSequence, ",")
    End If
    
    Interpose_PreMeas_Ary = SplitInputCondition(Interpose_PreMeas, "|") ''Carter, 20190616
    Dim PreMeas_Ary() As String
    If Interpose_PreMeas <> "" Then
        PreMeas_Ary = ParseData_InterPose(Interpose_PreMeas_Ary, TestSequenceArray)
    End If
    Dim PostMeas_Ary() As String

    Interpose_PostMeas_Ary = SplitInputCondition(Interpose_PostMeas, "|")
    If Interpose_PostMeas <> "" Then
        PostMeas_Ary = ParseData_InterPose(Interpose_PostMeas_Ary, TestSequenceArray)
    End If
    
    Dim loop_i As Long, Loop_j As Long
    Dim Temp_Equal_Str As String
    Dim Final_Comma_Str As String
    Temp_Equal_Str = vbNullString
    Final_Comma_Str = vbNullString
    Dim Split_CUS_Str_MainProgram_Str() As String
    Dim srcsweeparray() As String
    Dim Scaling As String
    Dim Scaling_index As Long
    temp_CUS_String = vbNullString

    'New Add for Dig sweep Src report tname with Decimal format -- 20221018
    Dim List_code_splitbyEql() As String
    Dim List_code_splitbyAt() As String
    Dim List_code_splitbyCondon() As String
    Dim List_DictName As String
    Dim List_Code_DSPWF As New DSPWave
    Dim vsite As Variant
    Dim List_Code_Index As Long
    
    If InStr(UCase(CUS_Str_MainProgram), UCase("sweepsrc:")) <> 0 Then
        'Dim srcsweeparray() As String
        ''temp_CUS_String = CUS_Str_MainProgram
        Split_CUS_Str_MainProgram_Str = Split(CUS_Str_MainProgram, ";")
        For i = 0 To UBound(Split_CUS_Str_MainProgram_Str)
            If InStr(Split_CUS_Str_MainProgram_Str(i), "=") <> 0 Then
                If temp_CUS_String = "" Then
                    temp_CUS_String = Split_CUS_Str_MainProgram_Str(i)
                Else
                    temp_CUS_String = temp_CUS_String & ";" & Split_CUS_Str_MainProgram_Str(i)
                End If
            End If
        Next i
        glb_s_sweepsrc_DigSrcAssignment = LCase(DigSrc_Assignment)
        Call ProcessSweepString(DigSrc_Assignment, temp_CUS_String, srcsweeparray, Loop_Max)
        srcnameindex = 0
        Sweepnameforsweep = srcsweeparray
        Loop_Max = Loop_Max - 1
    End If
    
    If InStr(UCase(CUS_Str_MainProgram), UCase("sweep_all_src:")) <> 0 Then
        'Dim srcsweeparray() As String
        ''temp_CUS_String = CUS_Str_MainProgram
        Split_CUS_Str_MainProgram_Str = Split(Split(CUS_Str_MainProgram, "sweep_all_src:")(1), ";")
        Dim src_size As Long
        Dim ss As Long
        For i = 0 To UBound(Split_CUS_Str_MainProgram_Str)
            If InStr(Split_CUS_Str_MainProgram_Str(i), "=") <> 0 Then
                src_size = UBound(Split(Split_CUS_Str_MainProgram_Str(i), ","))
                If i = 0 Then
                    Loop_Max = src_size
                Else
                    If Loop_Max <> src_size Then
                        TheExec.Datalog.WriteComment "Error Sweep Size are not the same!"
                        GoTo errHandler
                    End If
                End If
            End If
        Next i
        glb_s_sweepsrc_DigSrcAssignment = LCase(DigSrc_Assignment)
        ReDim srcsweeparray(Loop_Max)
        For i = 0 To Loop_Max
            srcsweeparray(i) = LCase(DigSrc_Assignment)
            For ss = 0 To UBound(Split_CUS_Str_MainProgram_Str)
                srcsweeparray(i) = Replace(srcsweeparray(i), Split(Split_CUS_Str_MainProgram_Str(ss), "=")(0), Split(Split(Split_CUS_Str_MainProgram_Str(ss), "=")(1), ",")(i))
            Next ss
        Next i
        srcnameindex = 0
        Sweepnameforsweep = srcsweeparray
        Loop_Max = Loop_Max
        CUS_Str_MainProgram = Replace(CUS_Str_MainProgram, "sweep_all_src:", "sweepsrc:")
        Instance_Data.CUS_Str_MainProgram = CUS_Str_MainProgram
    End If
    
    Scaling = vbNullString
    Scaling_index = (Len(Step_V)) - 2

    For i = 0 To Scaling_index
        If i = 0 Then
            Scaling = "0"
        ElseIf i = 1 Then
            Scaling = Scaling & ".0"
        Else
            Scaling = Scaling & "0"
        End If
    Next i

    For Loop_count = Loop_Init To Loop_Max
        Dim sweep_stop As New SiteBoolean
        Dim Final_High_Voltage_get As New SiteBoolean
        Dim Final_Low_Voltage_get As New SiteBoolean
        Dim temp_High As New SiteDouble
        Dim temp_Low  As New SiteDouble
'        Dim temp_minDeltaValue  As New SiteDouble
        Dim Final_High_Voltage As New SiteDouble
        Dim Final_Low_Voltage As New SiteDouble
        Dim Final_High_Current As New SiteDouble
        Dim Final_Low_Current As New SiteDouble
        
            temp_High = 9999
            temp_Low = 9999
            'temp_minDeltaValue = 9999
            Final_High_Voltage = 0
            Final_Low_Voltage = 0
            Final_High_Current = 0
            Final_Low_Current = 0
        
            If InStr(UCase(CUS_Str_MainProgram), UCase("sweepsrc:")) Or InStr(UCase(CUS_Str_MainProgram), UCase("specialsorce")) <> 0 Then
                TheExec.Flow.TestLimitIndex = 0
                DigSrc_Assignment = srcsweeparray(Loop_count)
                Instance_Data.DigSrc_Assignment = DigSrc_Assignment
            End If
        

            '' 20160923 - Add Interpose_PrePat entry point
            If Interpose_PrePat <> "" Then
                Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
            End If
    
        
            If patset.value <> "" Then
                Shmoo_Pattern = patset.value '' 20170808 add for shmoo pattern name print
                TheHdw.Patterns(patset).Load
                Call PatternBurstCheckAndSplit(patset.value, PattArray, PatCount)
    '''            Call PATT_GetPatListFromPatternSet(PatSet.value, PattArray, PatCount)
            Else
                ReDim PattArray(0)
                PattArray(0) = vbNullString
            End If

        
            ''20161107-Return sweep test name
            Dim Rtn_SweepTestName As String
            Rtn_SweepTestName = vbNullString
            gl_TName_Pat = patset.value
            
            Dim current_pat_index As Integer
            current_pat_index = 0
            
            '20191003 add for CPM with Multi_Init(DigSrc)_PL
            Dim DigSrc_Equation_temp_array() As String
            Dim DigSrc_Assignment_temp_array() As String
            Dim DigSec_Multi_Init_PL__Seq_index As Long: DigSec_Multi_Init_PL__Seq_index = 0

        
            For Each patt In PattArray

                If patt <> "" Then
                    TheExec.Flow.TestLimitIndex = 0
                    Pat = CStr(patt)
                    TheHdw.Patterns(Pat).Load

                    Set InDSPWave = Nothing
                
                    If LCase(DigSrc_Assignment) Like "*table*" Then DigSrc_Assignment = DigSrc_Assignment & "_" & CStr(TheExec.Flow.var("SrcCodeIndx").value)

                    Call GeneralDigSrcSettingWithBurst(LCase(patt), DigSrc_pin, InDSPWave, Rtn_SweepTestName)

                    Set OutDspWave = Nothing

                    If (CPUA_Flag_In_Pat) Then
                        Call TheHdw.Patterns(Pat).start
                    Else
                        Call TheHdw.Patterns(Pat).test(pfNever, 0)
                    End If

                End If
            
                TestSeqNum = 0

                For Each Ts In TestSequenceArray
                    Instance_Data.TestSeqNum = TestSeqNum

                    If (CPUA_Flag_In_Pat) Then
                        Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0) 'Meas during CPUA loop
                    Else
                        Call TheHdw.Digital.Patgen.HaltWait 'Pattern without CPUA loop
                    End If

                    If UCase(Ts) = "VDIFF2" Then
                        TestOptLen = 1
                    Else
                        TestOptLen = Len(Ts)
                    End If
                    For k = 1 To TestOptLen
                        Instance_Data.TestSeqSweepNum = k - 1
                        If Ts = "VDIFF2" Then
                            TestOption = Ts
                        Else
                            TestOption = mid(Ts, k, 1)
                        End If
                        
                   '''-------Start - Add per sweep feature for interpose_premeas - Carter, 20190614-------
                        If Interpose_PreMeas <> "" Then
                            If PreMeas_Ary(TestSeqNum, k - 1) <> "" Then
                                Call SetForceCondition(PreMeas_Ary(TestSeqNum, k - 1) & ";STOREPREMEAS")
                            End If
                        End If
                   '''-------End - Add per sweep feature for interpose_premeas - Carter, 20190614-------
                    Dim ForceVlotagesweep As New SiteDouble
                    Dim Vsitevaluetmp As New SiteDouble
                    Dim StepForLimit As Double
                    Dim Start_V_temp As New SiteDouble
                    Dim End_V_temp As New SiteDouble
                    StepForLimit = (CDbl(Step_V) / 2) * 0.8
                    For Each site In TheExec.sites.Active
                        TestNum = TheExec.sites.item(site).TestNumber
                        If InStr(TrimV_mode, "Binary") Then
                            ForceVlotagesweep(site) = Format((CDbl(Start_V) + CDbl(End_V)) / 2, Scaling)
                            'ForceVlotagesweep(site) = (CDbl(Start_V) + CDbl(End_V)) / 2
                        Else
                            ForceVlotagesweep(site) = CDbl(Start_V)
                        End If
                        sweep_stop(site) = False
                        Final_Low_Voltage_get = False
                        Final_High_Voltage_get = False
                        Start_V_temp(site) = CDbl(Start_V)
                        End_V_temp(site) = CDbl(End_V)
                    Next site
                    If TheExec.TesterMode = testModeOffline Then
                       TrimTarget = "1.5E-4"
                    End If

                                                                                                                                                                
                        
Nextvoltage:
                        
                        Vsitevaluetmp = 0#
                        If UCase(gl_GetInstrument_Dic(LCase(Sweep_V_pin))) = "VS-800MA" And glb_TesterType = "UltraFLEXplus" Then
                            With TheHdw.DCVS.Pins(Sweep_V_pin)
    
                                If .Gate = False Then
                                    .Disconnect tlDCVSConnectDefault
                                    .mode = tlDCVSModeVoltage
                                    .Voltage.ValuePerSite = Vsitevaluetmp
                                    .Connect tlDCVSConnectDefault
                                    .Gate = True
                                    If TheHdw.Tester.type = "Jaguar" Then
                                        Call Print_Error_Message(Warning_Info, "VBT_LIB_HardIP", "VTHSENSOR_Calibration", "Please check tester type")
                                    Else
                                        .Voltage.ValuePerSite = ForceVlotagesweep
                                    End If
                                Else
                                    .Voltage.ValuePerSite = ForceVlotagesweep
                                End If
    
                            End With
                        ElseIf UCase(gl_GetInstrument_Dic(LCase(Sweep_V_pin))) = "HSD-U" Then
                            For Each site In TheExec.sites
                                    '=============== UP1600===============
                                    If TheHdw.PPMU.Pins(Sweep_V_pin).Gate = tlOff Then
                                        TheHdw.Digital.Pins(Sweep_V_pin).Disconnect
                                        TheHdw.PPMU.Pins(Sweep_V_pin).Gate = tlOff
                                        TheHdw.PPMU.Pins(Sweep_V_pin).ForceV (CDbl(ForceVlotagesweep(site))), 0.02
                                        TheHdw.PPMU.Pins(Sweep_V_pin).Connect
                                        TheHdw.PPMU.Pins(Sweep_V_pin).Gate = tlOn
                                    Else
                                        TheHdw.PPMU.Pins(Sweep_V_pin).ForceV (CDbl(ForceVlotagesweep(site))), 0.02
                                    End If
                                    gl_ForceVlotagesweep_VTHSENSOR = ForceVlotagesweep 'Tname UD(8) using
                            Next site
                        End If
                        
                        For Each site In TheExec.sites.Active
                            TheExec.Datalog.WriteComment "Site " & site & "   Force Condtion: " & Sweep_V_pin & ":V:" & ForceVlotagesweep
                        Next site
                        
                        Dim pl_MeasureDict As New PinListData
                        Dim sd_MeasureData As Double
                        Dim sd_DeltaValue As New SiteDouble
                        Dim TestNameSourceCode As String: TestNameSourceCode = vbNullString
                        Dim sourceIndex As Long
                        
                        For sourceIndex = 0 To InDSPWave(0).sampleSize - 1
                            TestNameSourceCode = TestNameSourceCode & CStr(InDSPWave(0).Element(sourceIndex))
                        Next sourceIndex
                        
                        Select Case UCase(TestOption)
                            Case "I"
                                pl_MeasureDict = HardIP_MeasureCurrent
                            If gl_Disable_HIP_debug_log = False Then
                                Dim p As Long
                                Dim TestNameInput As String
                                For p = 0 To pl_MeasureDict.Pins.Count - 1
                                    For Each site In TheExec.sites
                                        'TestNameInput = "Site " & site & "     " & UCase(theexec.DataManager.instanceName) & " (LSB => MSB) Source code:" & TestNameSourceCode & "                   Measured Pins: " & pl_MeasureDict.Pins(p) & ", Measured data: " & Format(pl_MeasureDict.Pins(p).value * 10000000#, ".000") & " uA."
                                    
                                        TheExec.Datalog.WriteComment "Site " & site & "     " & UCase(TheExec.DataManager.instanceName) & " (LSB => MSB) Source code:" & TestNameSourceCode & "                   Measured Pins: " & pl_MeasureDict.Pins(p) & ", Measured data: " & Format(pl_MeasureDict.Pins(p).value * 10000000#, "0.000") & " uA."
                                    Next site
                                Next p
                            End If
                        End Select
                                          
                    
                        pl_MeasureDict = GetStoredMeasurement(Meas_StoreName)      'Get Stored Current Data
                        
                        For Each site In TheExec.sites.Active
                            
                            If Final_High_Voltage_get = False Or Final_Low_Voltage_get = False Then
                            
                                sd_MeasureData = pl_MeasureDict.Pins(0).value
                                    'High Search (<0uA) --> Small than -1.92uA (Ex: -2.07uA)
                                'If sd_MeasureData < 0 And sd_MeasureData < CDbl(TrimTarget) Then
                                If sd_MeasureData < CDbl(TrimTarget) Then
                                    sd_DeltaValue = Abs(sd_MeasureData - (CDbl(TrimTarget)))
                                    If sd_DeltaValue < temp_High Then
                                        temp_High = sd_DeltaValue
'                                    If sd_DeltaValue < temp_minDeltaValue Then
'                                        temp_minDeltaValue = sd_DeltaValue
                                        Final_High_Current = sd_MeasureData
                                        Final_High_Voltage = ForceVlotagesweep
                                        If InStr(TrimV_mode, "Binary") Or LCase(TrimV_mode) = "linear-all" Then
                                        ElseIf LCase(TrimV_mode) = "linear" And sd_MeasureData < 0 Then 'linear
                                            Final_High_Voltage_get = True
                                        Else
                                        End If
                                    End If
                                    If Final_High_Voltage_get = True And Final_Low_Voltage_get = True Then
                                    Else
                                        If Increase_TrimV Then
                                            If InStr(TrimV_mode, "Binary") Then
                                                End_V_temp(site) = ForceVlotagesweep(site)
                                                ForceVlotagesweep(site) = Format((Start_V_temp(site) + End_V_temp(site)) / 2, Scaling)
                                                'If Round((ForceVlotagesweep(site) + CDbl(Start_V)) / 2, 4) <= CDbl(Start_V) Then
                                                If ForceVlotagesweep(site) = Start_V_temp(site) Or ForceVlotagesweep(site) = End_V_temp(site) Then
                                                    Final_High_Voltage_get = True
                                                    Final_Low_Voltage_get = True
                                                End If
                                            Else 'linear
                                                ForceVlotagesweep = ForceVlotagesweep.Subtract(CDbl(Step_V))
                                                If ForceVlotagesweep < (CDbl(Start_V) - StepForLimit) Then 'original Donan
                                                'If ForceVlotagesweep < CDbl(Start_V) Then
                                                    If Not LCase(TrimV_mode) = "linear-all" Then Call Print_Error_Message(Warning_Info, "VBT_LIB_HardIP", "VTHSENSOR_Calibration", "Can't search the target, out of the range !!!")
                                                    Final_Low_Voltage_get = True
                                                    Final_High_Voltage_get = True
                                                End If
                                            End If
                                        Else
                                            If InStr(TrimV_mode, "Binary") Then
                                                Start_V_temp(site) = ForceVlotagesweep(site)
                                                ForceVlotagesweep(site) = Format((Start_V_temp(site) + End_V_temp(site)) / 2, Scaling)
                                                'If Round((ForceVlotagesweep(site) + CDbl(End_V)) / 2, 4) >= CDbl(End_V) Then
                                                If ForceVlotagesweep(site) = Start_V_temp(site) Or ForceVlotagesweep(site) = End_V_temp(site) Then
                                                    Final_High_Voltage_get = True
                                                    Final_Low_Voltage_get = True
                                                End If
                                            Else 'linear
                                                ForceVlotagesweep = ForceVlotagesweep.Add(CDbl(Step_V))
                                                If ForceVlotagesweep > (CDbl(End_V) + StepForLimit) Then 'original Donan
                                                'If ForceVlotagesweep > CDbl(End_V) Then
                                                    If Not LCase(TrimV_mode) = "linear-all" Then Call Print_Error_Message(Warning_Info, "VBT_LIB_HardIP", "VTHSENSOR_Calibration", "Can't search the target, out of the range !!!")
                                                    Final_Low_Voltage_get = True
                                                    Final_High_Voltage_get = True
                                                End If
                                            End If
                                        End If
                                    End If
                                'ElseIf sd_MeasureData < 0 And sd_MeasureData >= CDbl(TrimTarget) Then   'Low Search (<0uA) --> Larger than -1.92uA (Ex: -1.05uA)
                                ElseIf sd_MeasureData >= CDbl(TrimTarget) Then   'Low Search (<0uA) --> Larger than -1.92uA (Ex: -1.05uA)
                                    sd_DeltaValue = Abs(sd_MeasureData - (CDbl(TrimTarget)))
                                    If sd_DeltaValue < temp_Low Then
                                        temp_Low = sd_DeltaValue
'                                    If sd_DeltaValue < temp_minDeltaValue Then
'                                        temp_minDeltaValue = sd_DeltaValue
                                        Final_Low_Current = sd_MeasureData
                                        Final_Low_Voltage = ForceVlotagesweep
                                        If InStr(TrimV_mode, "Binary") Or LCase(TrimV_mode) = "linear-all" Then
                                        ElseIf LCase(TrimV_mode) = "linear" And sd_MeasureData < 0 Then 'linear
                                            Final_Low_Voltage_get = True
                                        Else
                                        End If
                                    End If
                                        If Final_High_Voltage_get = True And Final_Low_Voltage_get = True Then
                                        Else
                                            If Increase_TrimV Then
                                                If InStr(TrimV_mode, "Binary") Then
                                                    Start_V_temp(site) = ForceVlotagesweep(site)
                                                    ForceVlotagesweep(site) = Format((Start_V_temp(site) + End_V_temp(site)) / 2, Scaling)
                                                    'If Round((ForceVlotagesweep(site) + CDbl(End_V)) / 2, 4) >= CDbl(End_V) Then
                                                    If ForceVlotagesweep(site) = Start_V_temp(site) Or ForceVlotagesweep(site) = End_V_temp(site) Then
                                                        Final_High_Voltage_get = True
                                                        Final_Low_Voltage_get = True
                                                    End If
                                                Else 'linear
                                                    ForceVlotagesweep = ForceVlotagesweep.Add(CDbl(Step_V))
                                                    If ForceVlotagesweep > (CDbl(End_V) + StepForLimit) Then 'original Donan
                                                    'If ForceVlotagesweep > CDbl(End_V) Then
                                                        If Not LCase(TrimV_mode) = "linear-all" Then Call Print_Error_Message(Warning_Info, "VBT_LIB_HardIP", "VTHSENSOR_Calibration", "Can't search the target, out of the range !!!")
                                                        Final_High_Voltage_get = True
                                                        Final_Low_Voltage_get = True
                                                    End If
                                                End If
                                            Else
                                                If InStr(TrimV_mode, "Binary") Then
                                                    End_V_temp(site) = ForceVlotagesweep(site)
                                                    ForceVlotagesweep(site) = Format((Start_V_temp(site) + End_V_temp(site)) / 2, Scaling)
                                                    'If Round((ForceVlotagesweep(site) + CDbl(End_V)) / 2, 4) <= CDbl(Start_V) Then
                                                    If ForceVlotagesweep(site) = Start_V_temp(site) Or ForceVlotagesweep(site) = End_V_temp(site) Then
                                                        Final_High_Voltage_get = True
                                                        Final_Low_Voltage_get = True
                                                    End If
                                                Else 'linear
                                                    If sd_MeasureData > 0 Then
                                                        ForceVlotagesweep = ForceVlotagesweep.Add(CDbl(Step_V))
                                                        Call Print_Error_Message(Warning_Info, "VBT_LIB_HardIP", "VTHSENSOR_Calibration", "Measure current data " & sd_MeasureData & " > 0 !!!")
                                                    ElseIf Final_High_Voltage_get = False Then
                                                        ForceVlotagesweep = ForceVlotagesweep.Add(CDbl(Step_V))
                                                        Call Print_Error_Message(Warning_Info, "VBT_LIB_HardIP", "VTHSENSOR_Calibration", "Measure current data " & sd_MeasureData & " not expected, expected > target: " & TrimTarget & " !!!")
                                                        If ForceVlotagesweep > (CDbl(End_V) + StepForLimit) Then
                                                            Final_High_Voltage_get = True
                                                            Final_Low_Voltage_get = True
                                                        End If
                                                    Else
                                                        ForceVlotagesweep = ForceVlotagesweep.Subtract(CDbl(Step_V))
                                                        If ForceVlotagesweep < (CDbl(Start_V) - StepForLimit) Then 'original Donan
                                                        'If ForceVlotagesweep < CDbl(Start_V) Then
                                                            If Not LCase(TrimV_mode) = "linear-all" Then Call Print_Error_Message(Warning_Info, "VBT_LIB_HardIP", "VTHSENSOR_Calibration", "Can't search the target, out of the range !!!")
                                                            Final_High_Voltage_get = True
                                                            Final_Low_Voltage_get = True
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                Else
'                                    Final_Low_Voltage_get = True
'                                    Final_High_Voltage_get = True
'                                    Call Print_Error_Message(Warning_Info, "VBT_LIB_HardIP", "VTHSENSOR_Calibration", "Please check measured value!!!")
                                End If
                                
                                If Final_Low_Voltage_get = True And Final_High_Voltage_get = True Then
                                    sweep_stop = True
                                End If
                                If TheExec.TesterMode = testModeOffline Then
                                   Final_High_Current(site) = 5
                                   Final_Low_Current(site) = 4
                                End If
    
                             End If
                        Next site
                    
                        If Not sweep_stop.All(True) Then
                            TheExec.Flow.TestLimitIndex = 0
                            GoTo Nextvoltage
                        Else
                             If UCase(gl_GetInstrument_Dic(LCase(Sweep_V_pin))) = "VS-800MA" And glb_TesterType = "UltraFLEXplus" Then
                                With TheHdw.DCVS.Pins(Sweep_V_pin)
                                    .Voltage.ValuePerSite = Vsitevaluetmp
                                    .Gate = False
                                    .Disconnect
                                End With
                            ElseIf UCase(gl_GetInstrument_Dic(LCase(Sweep_V_pin))) = "HSD-U" Then
                                    TheHdw.PPMU.Pins(Sweep_V_pin).ForceV 0
                                    TheHdw.PPMU.Pins(Sweep_V_pin).Gate = tlOff
                                    TheHdw.PPMU.Pins(Sweep_V_pin).Disconnect
                            End If
                        End If

                        Dim TnameInput As String
                        Dim m As New SiteDouble
                        Dim B As New SiteDouble
                        Dim Vt As New SiteDouble
                        Dim T2NMOSVt As New SiteDouble
                        Dim T3PMOSVt As New SiteDouble
                        If Not LCase(TrimV_mode) = "linear-all" Then
                                For Each site In TheExec.sites
                                    If Final_High_Current = 0 Then
                                        Final_High_Current = 0.0000001
                                        TheExec.Datalog.WriteComment "Final_High_Current > 0, measurement value is not correct!"
                                    End If
                                Next site
                        
                                m = Final_High_Voltage.Subtract(Final_Low_Voltage).divide(Final_High_Current.Multiply(1000000#).Abs.Subtract(Final_Low_Current.Multiply(1000000#).Abs))
                                B = Final_Low_Voltage.Subtract(Final_Low_Current.Abs.Multiply(m).Multiply(1000000#))
                                Vt = m.Multiply(1.93).Add(B)
                                T2NMOSVt = Vt
                                T3PMOSVt = Vt.Negate.Add(0.75)

                                                                                                                                                                
        
                        
                            Dim TName_Ary() As String
                            ''' Fix test number for VthSensor @William 230628
                            For Each site In TheExec.sites
                                TheExec.sites.item(site).TestNumber = TestNum + 500
                            Next site
                            'High Voltage Result
                            TnameInput = Report_TName_From_Instance(CalcV, "X", , , , , , , tlForceNone)
                            TName_Ary = Split(TnameInput, "_")
                            TName_Ary(8) = "High"
                            TnameInput = Join(TName_Ary, "_")   'EX: HAC_CalcV_N_CZT2PCORE_VTHSENSOR_X_X_dpcsocchpcmwrapctrlvthsnscfgV00010100&dpcsocchpcmwrapctrlvthsnsnostressV0_High_0
                            TheExec.Flow.TestLimit Final_High_Voltage, 0, , , , , unitVolt, , TnameInput, , Sweep_V_pin, , , , , tlForceNone       'Search Nothing would '0V' --> Lolimit
                            'High Current Result
                            TName_Ary(1) = "CalcI"
                            TnameInput = Join(TName_Ary, "_")   'EX: HAC_CalcI_N_CZT2PCORE_VTHSENSOR_X_X_dpcsocchpcmwrapctrlvthsnscfgV00010100&dpcsocchpcmwrapctrlvthsnsnostressV0_High_0
                            TheExec.Flow.TestLimit Final_High_Current, , 0, , , , unitAmp, , TnameInput, , pl_MeasureDict.Pins(0).name, , , , , tlForceNone        'Search Nothing would '0A' --> Hilimit
                              'Low Voltage Result
                            TName_Ary(1) = "CalcV"
                            TName_Ary(8) = "Low"
                            TnameInput = Join(TName_Ary, "_")   'EX: HAC_CalcV_N_CZT2PCORE_VTHSENSOR_X_X_dpcsocchpcmwrapctrlvthsnscfgV00010100&dpcsocchpcmwrapctrlvthsnsnostressV0_Low_0
                            TheExec.Flow.TestLimit Final_Low_Voltage, 0, , , , , unitVolt, , TnameInput, , Sweep_V_pin, , , , , tlForceNone
                              'Low Current Result
                            TName_Ary(1) = "CalcI"
                            TnameInput = Join(TName_Ary, "_")   'EX: HAC_CalcI_N_CZT2PCORE_VTHSENSOR_X_X_dpcsocchpcmwrapctrlvthsnscfgV00010100&dpcsocchpcmwrapctrlvthsnsnostressV0_Low_0
                            TheExec.Flow.TestLimit Final_Low_Current, , 0, , , , unitAmp, , TnameInput, , pl_MeasureDict.Pins(0).name, , , , , tlForceNone
                            For Each site In TheExec.sites.Active
                                TheExec.Datalog.WriteComment "***site(" & site & "),m= " & m & " Volt ***"
                                TheExec.Datalog.WriteComment "***site(" & site & "),b= " & B & " Volt ***"
                            Next site
                          
                        
                            'T2 Voltage Result
                            If Split(TheExec.DataManager.instanceName, "_")(1) Like "*T2*" Then
                                TName_Ary(1) = "CalcV"
                                TName_Ary(8) = "T2-NMOS-Vt"
                                TnameInput = Join(TName_Ary, "_")   'EX: HAC_CalcV_N_CZT2PCORE_VTHSENSOR_X_X_dpcsocchpcmwrapctrlvthsnscfgV00010100&dpcsocchpcmwrapctrlvthsnsnostressV0_Low_0
                                TheExec.Flow.TestLimit T2NMOSVt, 0, , , , , , , TnameInput, , Sweep_V_pin, , , , , tlForceNone
                            'T3 Voltage Result
                            ElseIf Split(TheExec.DataManager.instanceName, "_")(1) Like "*T3*" Then
                                TName_Ary(1) = "CalcV"
                                TName_Ary(8) = "T3-PMOS-Vt"
                                TnameInput = Join(TName_Ary, "_")   'EX: HAC_CalcV_N_CZT2PCORE_VTHSENSOR_X_X_dpcsocchpcmwrapctrlvthsnscfgV00010100&dpcsocchpcmwrapctrlvthsnsnostressV0_Low_0
                                TheExec.Flow.TestLimit T3PMOSVt, 0, , , , , , , TnameInput, , Sweep_V_pin, , , , , tlForceNone
                            End If
                        End If
                        
                        
    
                        '''-------Start - Add per sweep feature for interpose_premeas - Carter, 20190614-------
                            If Interpose_PreMeas <> "" Then
                                If PreMeas_Ary(TestSeqNum, k - 1) <> "" And UCase(TestOption) <> "N" Then
                                    Call SetForceCondition("RESTOREPREMEAS")
                                End If
                            End If
                        '''-------End - Add per sweep feature for interpose_premeas - Carter, 20190614-------
                        
                        '''-------Start - Add Interpose_PostMeas for Pattern Burst, 20210526-------
                            If Interpose_PostMeas <> "" Then
                                If PostMeas_Ary(TestSeqNum, k - 1) <> "" Then
                                    Call SetForceCondition(PostMeas_Ary(TestSeqNum, k - 1))
                                End If
                            End If
                        '''-------End - Add Interpose_PostMeas for Pattern Burst, 20210526-------
                        Next k
                        
                        TestSeqNum = TestSeqNum + 1
                        
                        If (CPUA_Flag_In_Pat) Then Call TheHdw.Digital.Patgen.Continue(0, cpuA)
                        Instance_Data.TestSeqNum = TestSeqNum
                    
                Next Ts
            
                If DebugPrintEnable = True Then TheExec.Datalog.WriteComment "  Pattern(" & PatCount & "): " & Pat & vbNullString
                
                TheHdw.Digital.Patgen.HaltWait ' Haltwait at patten end
                
                PatCount = PatCount + 1
                
                '' 20160923 - Add Interpose_PostTest entry point
                If Loop_count = Loop_Max Then   'brian
                    Call SetForceCondition(Interpose_PostTest)
                ElseIf Not (TheExec.DataManager.instanceName Like "*VTHSENSOR*") Then
                    Call SetForceCondition(Interpose_PostTest)
                End If
                    
                    If patt <> "" Then
                        Call HardIP_WriteFuncResult(, , glb_TestInstance)
                    End If
                    'End If
                    
                    If gl_FlowForLoop_DigSrc_SweepCode <> "" Then        '20180814
                        gl_FlowForLoop_DigSrc_SweepCode = vbNullString
                        gl_FlowForLoop_DigSrc_SweepCode_Dec = vbNullString '20190613 CT add for Decimal value printing
                    End If
                    
        ''        End If
        
                current_pat_index = current_pat_index + 1
            
                If Interpose_PreMeas <> "" And restore_Flag = True Then
                    Call SetForceCondition("RESTOREPREMEAS")

                End If
                
                gl_Sweep_Glb_TName = vbNullString '' 20190529 - Add for sweep force V
                  
            Next patt
        
             
             DebugPrintFunc patset.value, True ' print all debug information
             

             
             If Interpose_PrePat <> "" Then
                 Call SetForceCondition("RESTOREPREPAT")
             End If
             
             ''=============================== CharSetName ====================================
             'Dim p As Variant
            
            If InStr(UCase(CUS_Str_MainProgram), UCase("sweepsrc:")) <> 0 Then: srcnameindex = srcnameindex + 1 '@220104 updated by Walker
    Next Loop_count
    ''================================================================================
    gl_Sweep_V_pin = vbNullString 'avoid UD(8) inherit to next instance
    
    ReDim TestConditionSeqData(0)
    Dim Instance_Data_temp() As Instance_Type
    ReDim Instance_Data_temp(0)
    Instance_Data = Instance_Data_temp(0)
    'TTR,20200423, Oscar
    Instance_Data.Meas_StoreName_Flag = False ''Carter, 20190521
    temp_CUS_String = vbNullString
    TheHdw.Alarms.Check
    ByPassTestLimit = False
    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in VTHSENSOR_Calibration"
'    Resume Next
    If AbortTest Then Exit Function Else Resume Next
  
End Function
