Attribute VB_Name = "LIB_Digital_Shmoo"
Option Explicit

Public pseudo_result_index As Long
'Revision History
' 1.3 add support for edge shmoo with skip tests
' 1.4 add support for "retest" in characterization
' 1.5 add shmoo pattern gloabal variable for Char setup retest function
' 1.6 add Char key word Shmoo_header for hard IP used
' 1.7 add IG-XL 8.10.12 coding for pattern list
' 1.8 change lotId and wafer Id to site variable
' 1.9 add Char_map.txt for in job characterization
' 1.10 add function Shmoo_Test_Pattern for in round-1 characterization
' 1.11 make VBT_shmoo.bas more independent

Public Set_Pin_NV As Double
Public Const VBT_Shmoo_Version = "1.11"
Public Const MAX_CHAR_ENABLE_ROW = 30000
Public Const MAX_CHAR_SETUP_ROW = 100
Public Const Char_Flow_Enable_Sheet = ".\Setup\Char_Flow_Enable.txt"
Public Const Shmoo_Setup_Sheet_Enable = ".\Setup\Char_Enable.txt"
Public Const Shmoo_Setup_Sheet_Setup = ".\Setup\Char_Setup.txt"
Public Const char_map_Sheet = ".\Common\Char_Map.txt"
Public wb As Workbook
Public Shmoo_Pattern As String
Public Shmoo_Pattern_Payload As String
Public Shmoo_header As String
Public Shmoo_Vcc_Min As New SiteDouble
Public Shmoo_Vcc_Max As New SiteDouble
Public ShmooPowerName As String
Public Shmoo_Instance_Name() As String
Public Shmoo_setup_name() As String
Public Shmoo_Setup_Name_New() As String
Public Shmoo_Setup_idx As Integer
''For AI use 20150715
Public Voltage_fail_point As Long
Public Voltage_fail_point_request As Long
Public Voltage_fail_collect(10) As String
Public ReportHVCC As Boolean
Public ReportLVCC As Boolean
Public ShmResult As New SiteVariant
Public RTOSPatResult As New SiteBoolean
Public AnalogMeas As Boolean
Public SweepGuardBand As Boolean
Public SweepGuardBandVal As Double
Public CalcSymbol As String
Public ReviseSetup As Boolean
Public Enable_CZsetup_Chk As Boolean

Type Char_Enable
    Enable As String
    TestInstance As String
    charSetup As String
    Pattern As String
    Count As Long
End Type
Type Char_setup
    Setup_name As String
    Test_Method As String
    Step_Name As String
    mode As String
    Parameter_Type As String
    Parameter_Name As String
    Range_Calc_Field As String
    Range_From As String
    Range_To As String
    Range_Steps As String
    Range_Step_Size As String
    Perform_Test  As String
    Test_Limits_Low As String
    Test_Limits_High As String
    Algorithm_Name As String
    Algorithm_Arguments As String
    Algorith_Results_Check As String
    Algorithm_Transition As String
    Apply_To_Pins As String
    Apply_To_Pin_Exec_Mode As String
    Apply_To_Time_Sets As String
    Adjust_Backoff As String
    Adjust_Spec_Name As String
    Adjust_From_Setup As String
    Function As String
    Function_Arguments As String
    Interpose_Functions_Pre_Setup As String
    Interpose_Functions_Pre_Setup_Arguments As String
    Interpose_Functions_Pre_Step As String
    Interpose_Functions_Pre_Step_Arguments As String
    Interpose_Functions_Pre_Point As String
    Interpose_Functions_Pre_Point_Arguments As String
    Interpose_Functions_Post_Point As String
    Interpose_Functions_Post_Point_Arguments As String
    Interpose_Functions_Post_Step As String
    Interpose_Functions_Post_Step_Arguments As String
    Interpose_Functions_Post_Setup As String
    Interpose_Functions_Post_Setup_Arguments As String
    Output_Format As String
    Output_Text_File As String
    Output_Sheet As String
    Output_Destinations_Text_File As String
    Output_Destinations_Sheet As String
    Output_Destinations_Datalog As String
    Output_Destinations_Immediate_Win As String
    Output_Destinations_Output_Win As String
    comment As String
    Count As Long
End Type
Public Const MaxCharSetup = 15
Public Const MaxCharCorePower = 7
Public Const MaxCharInitPatt = 4
Public Const MaxFuncBlock = 100
'Enum Char_Enable_Enum
'    Fail_Enable = 1
'    Disable = 2
'    Enable = 3
'End Enum
Type Char_map
     TestNum(MaxCharSetup) As String
     Func_Block As String
     PowerCondition(MaxCharSetup) As String
     Enable(MaxCharSetup) As String
     Char_setup(MaxCharSetup) As String
     NV_Power(MaxCharSetup) As String
     Core_power(MaxCharSetup, MaxCharCorePower) As Double
     Init_Patt(MaxCharSetup, MaxCharInitPatt) As String
     Count As Long
End Type
Type Current_Shmoo_Setup
    TestNum As Long
    Enable As String
    Func_Block As String
    Func_block_index As Long
    PowerCondition As String
    Char_Setup_Index As Long 'index of  char setup within a function block
    Char_Setup_Name As String
    Pins_Apply As String
End Type
Public Curr_Shmoo_Condition As Current_Shmoo_Setup
Public char_map_entry(MaxFuncBlock) As Char_map
Public Char_Setup_Collection_Index As New Collection
Public count_func_block As Long
Public ShmooSweepPower(100) As New SiteDouble
Public Power_Level_Last As New SiteVariant
Public Shmoo_Apply_Pin As String

Dim char_flow_enable_entry(MAX_CHAR_ENABLE_ROW) As Char_Enable
Dim char_enable_entry(MAX_CHAR_ENABLE_ROW) As Char_Enable
Dim char_setup_entry(MAX_CHAR_SETUP_ROW) As Char_setup
Dim char_flow_enable_key As New Collection
Dim char_enable_key As New Collection
Dim char_setup_key As New Collection
Dim char_setup_count As Long
Dim char_enable_count As Long
Dim shmoo_mode As tlDevCharShmooAxis
Dim shmoo_algorithm As tlDevCharShmooPGA
Dim shmoo_Calc_Field As tlDevCharRangeField
Dim shmoo_Apply_To_Pin_Exec_Mode As tlDevCharPinExecMode
Dim shmoo_Destination_DataLog As tlDevCharOutputDestinationState
Dim shmoo_Destination_OutputWindow As tlDevCharOutputDestinationState
Dim shmoo_Destination_Sheet As tlDevCharOutputDestinationState
Dim shmoo_Destination_TextFile As tlDevCharOutputDestinationState
Dim shmoo_Destination_ImmediateWindow As tlDevCharOutputDestinationState

Public Flow_Shmoo_Axis(20) As String
Public Flow_Shmoo_Axis_Count As Long
Public Flow_Shmoo_X_Step As Long
Public Flow_Shmoo_Y_Step As Long
Public Flow_Shmoo_X_Current_Step As Long
Public Flow_Shmoo_Y_Current_Step As Long
Public Flow_Shmoo_X_Last_Value As Long
Public Flow_Shmoo_Y_Last_Value As Long
Public Flow_Shmoo_X_Start As Long
Public Flow_Shmoo_Y_Start As Long
Public Flow_Shmoo_X_Fast As Boolean
Public Flow_Shmoo_Force_Condition As String
Public Shmoo_setup_str As String
Public Shmoo_End As Boolean
Public Flow_Shmoo_Port_Name As String
Public FlowShmooString_GLB As String
Public shmoohole_count As New SiteLong
Public shmooallfail_count As New SiteLong
Public shmooalarm_count As New SiteLong
Public included_shmoo_count As New SiteLong
Public excluded_shmoo_count As New SiteLong
Public total_shmoo_count As New SiteLong
Public F_shmoo_abnormal_counter As Boolean

'---------------------------------
'--------- for EMA DigSrc --------
'---------------------------------
Public Type testCondition
    DigSrc_BitStr As String
    DigSrc_BitStrAry() As String
    ConditionName As String
    DigSrc_BitCount As Double
End Type

Public Type DynamicSrc
    PatternName As String
    TestCase() As testCondition
    TestCase_index As New Dictionary
End Type
Public SrcStock() As DynamicSrc

Public Dic_SrcStockIndex As New Dictionary
Public Dic_TestCaseIndex As New Dictionary

Public SELSRM_MappingTableIsRead As Boolean
Public DSSCMappingTableIsRead As Boolean
Public g_Retention_Start As Boolean
Public g_Retention_Shmoo As Boolean
Public g_ForceCond_VDD As String
Public g_ForceDCVS As String
Public g_ForceDCVI As String
Public g_ShmooDCVS As String
Public g_ShmooDCVI As String
Public g_Retention_FC As String ' Retention pin/Voltage parsed from force condition "RETV", Eg. VDD1:RETV:1.0;VDD3,VDD4:RETV:1.1 => VDD1=1.0;VDD3,VDD4=1.1
Public g_Retention_VDD As String 'Retention pin parsed from force condition "RETV"
Public g_Retention_ForceV As String 'Retention Voltage parsed from force condition "RETV"

'=================================================================
' 201810 add these parameters for Select Sram START
''''''''''''''''''''''''''''''''''''''''''''
Public Type Sub_Info
    BITS As Integer
    logicPin As String
    SramPin As String
    SelSram1 As Integer
    SelSram0 As Integer
End Type
Public Type Domain
    DomainName As String
    Pattern() As String
    DomainBits() As Sub_Info
End Type
Public Type mapping_table
    Block() As Domain
End Type
Public GetSelSram As mapping_table
''''''''''''''''''''''''''''''''''''''''''''
Public PrintDSSCSwitchVoltage As New PinListData
Public PrintSwitchDspWave As New DSPWave
Public g_BlockType As String
Public digSrc_EQ_GB As String
Public BlockType_GB As String
Public TestType_GB As String
Public DigSrc_pin_GB As New PinList
Public DigSrcSize_GB As String
Public dssc_pat_init_GB As String
Public g_shmoo_ret As Boolean
Public g_InitSeq As String
Public g_dyanmicDSSCbits As String
Public RTOS_Shmoo_Start As Boolean
Public g_VminBoundary_selsrm As Boolean

Public g_FirstSetp As Boolean
Public g_Vbump_function As Boolean
Public g_Print_SELSRM_Def As Boolean
Public ShmooSweepPowerDict As New PinListData
Public Power_Level_Vmode_Last As String
Public g_ApplyLevelTimingVmain As New PinListData
Public g_ApplyLevelTimingValt As New PinListData
Public g_ContextVmainValue As New PinListData
Public g_ContextVAltValue As New PinListData
Public DICT_DCVS_PIN_INDEX As New Dictionary
Public g_CharInputString_Voltage_Dict As New Dictionary
Public g_Globalpointval As New PinListData
Public g_RetntionVal As New PinListData
Public g_RestoreMainOrAlt As New PinListData
Public g_VDDForce As String
Public g_PLSWEEP As Boolean
' 201810 add these parameters for Select Sram END
'=================================================================

Public g_ShmooPin As New PinListData
Public gL_AllPattAry() As New Pattern
Public gL_DynamicSourceBitAry() As String
Public gL_TestType() As String
''//////// Put outside the function call/////////
Public Type Error_type
    Pmode_Naming_error As Boolean
    Ret_Scenario_error As Boolean
    Ret_WaitTime_error As Boolean
    Patsets_error As Boolean
    Patsets_error_print As Boolean
End Type
Public Type Error_Inst
    Error_status As Error_type
    instance_name As String
    ErrorPmode As String
    ErrorPatsets As String
    ErrorPatsetsPrint As String
End Type

Public Glb_RetentionMeasurement As Boolean
Public Type MeasureCond
    PinNameAry() As String
    InstrumentType() As String
    ForceValueAry() As String
    MeasureRangeAry() As String
    WaitTimeAry() As String
    RestoreGateOffAry() As Boolean
End Type
Public Type MeasureInfo
    DCVI_Pins As MeasureCond
    DCVS_Pins As MeasureCond
    digital_pins As MeasureCond
End Type
Public Enum MeasureSeq
    RET_MeasV = 0
    RET_MeasI = 1
    RET_MeasF = 2
End Enum
Public glb_RET_MeasAry() As MeasureInfo
Public Enable_AutoRange As Boolean
Public Selsram_print() As String
Public SrcBitFromArgument As Boolean
Public Dict_INI_NV As New Dictionary
Public Dict_INI_HV As New Dictionary
Public Dict_INI_LV As New Dictionary
Public Dict_PL_NV As New Dictionary
Public Dict_PL_HV As New Dictionary
Public Dict_PL_LV As New Dictionary
'' //////////////////////////////////////////////
Public do_shmoo_hole_flag As Boolean '20240531shmoo hole fail log


Public Function Shmoo_Test_Pattern_bk(ByVal patt As Pattern, ReportResult As PFType, ResultMode As tlResultMode, ConcurrentMode As tlPatConcurrentMode, Power_Run_Scenario As String, powerPin As String, set_init As Boolean, seq As Long, wait_time As String, _
                                    Optional DigSrc_BitSize As String, Optional DigSrc_Seg As String, Optional DigSrc_DigSrcPin As String, Optional digSrc_EQ As String, _
                                    Optional RTOSRelaySwith As Boolean, _
                                    Optional allPowerPins As PinList, _
                                    Optional DecideSPIMatchLoopFlag As Boolean, _
                                    Optional SPIMatchLoopCountValue As Long, _
                                    Optional CharInputString As String, _
                                    Optional RTOSPatIndex As Integer, _
                                    Optional BlockType As String, Optional DynamicSelSrmBits As String, Optional Vbump As Boolean = False, _
                                    Optional testType As String)
On Error GoTo errHandler

    Dim External_Retention As Boolean
    Dim test_name_ary() As String
    Dim SRV_type As String
    Dim block_name As String
    Dim site As Variant
    Dim lPatternCount As Long
    Dim astrPattTemp() As String
    Dim bstrPattTemp() As String
    Dim TestCase As String
    Dim DigSrc_Size As Double
    Dim DigSrc_flag As Boolean
    Dim digcap_flag As Boolean 'add for DigCap function
    Dim DigSrc_wav As New DSPWave
    Dim DigSrc_pin As New PinList
    Dim PattArray() As String
    Dim PatCount As Long
    Dim Seg_Arr() As String
    Dim Pin_Ary() As String
    Dim pin_count As Long
    Dim pattstr As String

    Dim i As Integer
    '========================== 'add for Multi Pat function ==========================
    Dim MultiPatAry() As String
    Dim MultiPat As Boolean
    Dim MultiPatCount As Long
    Dim CountMultiPat As Long
    '========================== 'add for Multi Pat function ==========================
    
    '========================== 'add for DigCap function ============================
    Dim DigCapName() As String
    Dim DigSrcPin As String, DigCapPin As String, DigSrcSize As String, DigCapSize As String
    Dim DigCap_Info_Dict As New Dictionary
    Dim DigCap_Pin As New PinList
    Dim OutDspWave As New DSPWave
    Dim DSSC_Capture_Out As String
    '========================== 'add for DigCap function ============================
    
    '========================== 'add for SELSRM function ============================
    Dim SELSRM_Fun As Boolean
    '========================== 'add for SELSRM function ============================
    Dim T0, T1 As Double
    
    pattstr = patt.value
    '' Add for Pattern loop ,20160607, KS
    MultiPat = False 'add for Multi Pat function
    digcap_flag = False 'add for DigCap function
    g_Retention_Shmoo = False 'add for SelSram function
    DigSrc_flag = False 'init flag
    SELSRM_Fun = False 'init SELSRM flag
    lPatternCount = 0 'initial PatternCount for pat loops
    MultiPatCount = 0 'initial Multi patterns count
    
    If patt.value = "" Then Exit Function   '#16_ELSE_CASE_CHK
    '' auto convert T_update to T_char, need to modified VBT for each project
'''    If Vbump = True Then ' SELSM function for debug use
'''        If TheExec.EnableWord("BringUp_Shmoo") = True Then
'''                SELSRM_Fun = True ''' to avoid VBT error while BringUp_Shmoo enable word is opening
'''           If InStr(UCase(patt), "DSSC") > 0 Then
'''              digSrcPin = "JTAG_TDI"
'''              DigSrc_flag = True
'''               If UCase(BlockType) Like "*SOC*" Then
'''                  digSrc_EQ = "SSSSSSSSSSSSSSSSSSSSS"
'''                  DigSrcSize = "21"
'''               ElseIf UCase(BlockType) Like "*CPU*" Then
'''                  digSrc_EQ = "SSSSSSSSSSSSSSSSS"
'''                  DigSrcSize = "17"
'''               ElseIf UCase(BlockType) Like "*GPU*" Or UCase(BlockType) Like "*GFX*" Then
'''                  digSrc_EQ = "SSSSSSS"
'''                  DigSrcSize = "7"
'''               End If
'''               GoTo BringUp_Shmoo
'''            End If
'''        End If
'''    End If
    
    If wait_time <> "" Then g_Retention_Shmoo = True ' use for non SELSRM Function
    
    If InStr(patt, ":") > 0 Then    '#16_ELSE_CASE_CHK
        astrPattTemp = Split(patt, ":")
        bstrPattTemp = Split(astrPattTemp(1), "_")
        
        '========================================================================Process SELSRM format=====================================================================
        If InStr(LCase(bstrPattTemp(0)), "selsrm") > 0 Then '#16_ELSE_CASE_CHK
            If Vbump = True Then
                SELSRM_Fun = True
                If DigSrc_BitSize <> "" And DigSrc_DigSrcPin <> "" And digSrc_EQ <> "" Then
                    Call Char_Process_DigString(DigSrc_BitSize, DigSrc_Seg, DigSrc_DigSrcPin, DigCapName, DigSrcPin, DigCapPin, DigSrcSize, DigCapSize, DigCap_Info_Dict)
                    If DynamicSelSrmBits <> "" Then '#16_ELSE_CASE_CHK
                        If Not UCase(digSrc_EQ) = UCase(DynamicSelSrmBits) Then '#16_ELSE_CASE_CHK
                           Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Test_Pattern", "DynamicSelSrmBits Count Not Match !!!")
    '                       TheExec.ErrorLogMessage "DynamicSelSrmBits"
                           GoTo errHandler
                        Else
                        End If
                    Else
                    End If
                ElseIf DynamicSelSrmBits <> "" And DigSrc_BitSize = "" And DigSrc_DigSrcPin = "" Then
                    DigSrcSize = Len(DynamicSelSrmBits)
                    DigSrcPin = "JTAG_TDI"
                    digSrc_EQ = DynamicSelSrmBits
                ElseIf DynamicSelSrmBits <> "" And DigSrc_BitSize <> "" And DigSrc_DigSrcPin <> "" And digSrc_EQ = "" Then
                    Call Char_Process_DigString(DigSrc_BitSize, DigSrc_Seg, DigSrc_DigSrcPin, DigCapName, DigSrcPin, DigCapPin, DigSrcSize, DigCapSize, DigCap_Info_Dict)
                    digSrc_EQ = DynamicSelSrmBits
                Else
                    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Test_Pattern", "No Digital source for SELSRM Char")
    '                TheExec.ErrorLogMessage "No Digital source for SELSRM Char"
                End If
            Else
                Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Test_Pattern", "Please enable Vbump Function")
    '            TheExec.ErrorLogMessage "Please enable Vbump function"
                GoTo errHandler
            End If
        End If
        '=======================================================================Process SELSRM format=======================================================================
        
        '============================================================Process DSSC string, merge DigSrc/DigCap===============================================================
        If InStr(LCase(bstrPattTemp(0)), "digsrc") > 0 Then '#16_ELSE_CASE_CHK
            Call Char_Process_DigString(DigSrc_BitSize, DigSrc_Seg, DigSrc_DigSrcPin, DigCapName, DigSrcPin, DigCapPin, DigSrcSize, DigCapSize, DigCap_Info_Dict)
        End If
        '============================================================Process DSSC string, merge DigSrc/DigCap===============================================================
        
        '========================================================================Mapping Table Method=======================================================================
        If (UBound(bstrPattTemp()) = 1) Then    '#16_ELSE_CASE_CHK
            Dim SrcBitAry() As String
            TestCase = bstrPattTemp(1)
            Call GetSrcString_fromEMAArray(astrPattTemp(0), TestCase, digSrc_EQ, DigSrc_Size, SrcBitAry)
            DigSrc_BitSize = CStr(DigSrc_Size)
        End If
        '========================================================================Mapping Table Method=======================================================================
        
        
        If (LCase(bstrPattTemp(0)) <> "digsrc") And (LCase(bstrPattTemp(0)) <> "selsrm") Then 'Pattern loops
            lPatternCount = CLng(astrPattTemp(1)) - 1
            patt = astrPattTemp(0)
            theexec.Datalog.WriteComment "Loop Pattern :" & patt & "_" & "Repeat count :" & lPatternCount + 1
        Else ' Create DSPWave signal
            If DigSrcPin <> "" Then DigSrc_flag = True
            If DigCapPin <> "" Then digcap_flag = True
            patt = astrPattTemp(0)
            
'''BringUp_Shmoo:
            Call PATT_GetPatListFromPatternSet(patt.value, PattArray, PatCount)
            If DigSrc_flag = True Then
                Set DigSrc_wav = Nothing
                DigSrc_wav.CreateConstant 0, DigSrcSize
               
               '===================================================DSSC Switching for SELSRM Function====================================================='
                If SELSRM_Fun = True Then
                    Dim DC_Spec_Level As New PinListData
                    Dim DecodeingString As String
                    If theexec.enableWord("Shmoo_TTR") = True Then
                        If InStr(UCase(digSrc_EQ), "S") > 0 Then
                            If set_init = True And g_InitSeq = "" Then g_InitSeq = CStr(seq)
                        Else
                            If g_InitSeq = "" Then g_InitSeq = "Payload1"
                        End If
                    End If
                    Decide_DC_Level DC_Spec_Level, g_ApplyLevelTimingValt, g_ApplyLevelTimingVmain, BlockType, testType
                    digSrc_EQ = Decide_Switching_Bit(digSrc_EQ, DigSrc_wav, DC_Spec_Level, BlockType, DecodeingString, powerPin, g_Globalpointval, g_ForceCond_VDD, g_CharInputString_Voltage_Dict, patt.value)
            '===================================================DSSC Switching for SELSRM Function====================================================='
                Else ' Without DSSC Switching
                  'Decide source bit if DigSrc with "S"
                    If InStr(UCase(digSrc_EQ), "S") > 0 Then    '#16_ELSE_CASE_CHK
                        Decide_DC_Level DC_Spec_Level, g_ApplyLevelTimingValt, g_ApplyLevelTimingVmain, BlockType, testType
                        digSrc_EQ = Decide_Switching_Bit(digSrc_EQ, DigSrc_wav, DC_Spec_Level, BlockType, DecodeingString, powerPin, g_Globalpointval, g_ForceCond_VDD, g_CharInputString_Voltage_Dict, patt.value)
                    End If
                    For i = 0 To Len(digSrc_EQ) - 1
                       For Each site In theexec.sites.Active
                           DigSrc_wav.Element(i) = CDbl(mid(digSrc_EQ, i + 1, 1))
                       Next site
                    Next i
                End If
               
                DigSrc_pin.value = DigSrcPin
                Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "FUNC_SRC", CLng(DigSrcSize), DigSrc_wav)
               
               '===========================================================Debug LVCC/HVCC/Diagnostic Char=========================================================='
                If (SELSRM_Fun = True And theexec.Flow.enableWord("Debug_LVCC_VminBoundary") = True) Or (SELSRM_Fun = True And theexec.Flow.enableWord("Debug_HVCC_VminBoundary") = True) Then
                    digSrc_EQ_GB = digSrc_EQ:: BlockType_GB = BlockType:: DigSrcSize_GB = DigSrcSize:: dssc_pat_init_GB = PattArray(0):: DigSrc_pin_GB = DigSrc_pin:: TestType_GB = testType
    '                 g_VminBoundary_selsrm = True
                End If
               '===========================================================Debug LVCC/HVCC/Diagnostic Char=========================================================='
                If SELSRM_Fun = True Then
                    If set_init Then
                        theexec.Datalog.WriteComment "DigSrc pattern = " & "Init" & seq & ": " & patt & "," & "Src Bits = " & Len(digSrc_EQ) & "," & "Output String [ LSB(L) ==> MSB(R) ]:" & "SourceCode:" & digSrc_EQ '& "," & "DSelSrm:" & DecodeingString
                    Else
                        theexec.Datalog.WriteComment "DigSrc pattern = " & "Payload" & seq & ": " & patt & "," & "Src Bits = " & Len(digSrc_EQ) & "," & "Output String [ LSB(L) ==> MSB(R) ]:" & "SourceCode:" & digSrc_EQ '& "," & "DSelSrm:" & DecodeingString
                    End If
                Else
                    If set_init Then
                        theexec.Datalog.WriteComment "DigSrc pattern = " & "Init" & seq & ": " & patt & "," & "Src Bits = " & Len(digSrc_EQ) & "," & "Output String [ LSB(L) ==> MSB(R) ]:" & digSrc_EQ
                    Else
                        theexec.Datalog.WriteComment "DigSrc pattern = " & "Payload" & seq & ": " & patt & "," & "Src Bits = " & Len(digSrc_EQ) & "," & "Output String [ LSB(L) ==> MSB(R) ]:" & digSrc_EQ
                    End If
                End If
            End If
            ' ==============================================================Creat DSP wave for DigCap=============================================================
            If digcap_flag = True Then  '#16_ELSE_CASE_CHK
                Set OutDspWave = Nothing
                DigCap_Pin.value = DigCapPin
                Call GeneralDigCapSetting(PattArray(0), DigCap_Pin, CLng(DigCapSize), OutDspWave)
                theexec.Datalog.WriteComment ("Cap Bits = " & CLng(DigCapSize))
                theexec.Datalog.WriteComment ("Cap Pin = " & DigCap_Pin)
                theexec.Datalog.WriteComment ("======== Setup Dig Cap Test End   ========")
            End If
            ' ==============================================================Creat DSP wave for DigCap=============================================================
        End If
 
    ElseIf InStr(patt, ",") > 0 Then 'Multi Pattern function
        MultiPatAry = Split(patt, ",")
        MultiPatCount = UBound(MultiPatAry)
        MultiPat = True
    Else
    End If
    

''===========================================================SET RUN LEVEL=========================================================
    If Vbump = True Then  'add for SelSram function
        Set_Run_Level_Vbump Power_Run_Scenario, powerPin, set_init, seq 'add for Vbump function
    Else
        If Not UCase(Power_Run_Scenario) Like "INIT_SWEEP_PL_SWEEP" Then '#16_ELSE_CASE_CHK
       ''no need to change voltage conditions if init_sweep_pl_sweep (it apply to correct sweep condition by IG-XL)
            Set_Run_Level Power_Run_Scenario, powerPin, set_init, seq
        Else
        End If
    End If
''===========================================================SET RUN LEVEL=========================================================

    Dim InDSPWave As New DSPWave
    Dim Count As Long
    Dim TestNumber As Long
            

    For CountMultiPat = 0 To MultiPatCount  'Multi pat function
    
        If MultiPat = True Then
           Call thehdw.Patterns(MultiPatAry(CountMultiPat)).Load
        Else
           Call thehdw.Patterns(patt).Load
        End If
                    
     ''-------------------------------------------
     '' HRAM setup capture on first fail 20170425
        thehdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
        thehdw.Digital.hram.size = 512
        thehdw.Digital.hram.CaptureType = captFail
        thehdw.Digital.hram.SetTrigger trigFirst, True, 0, True
     ''-------------------------------------------
     
        For Count = 0 To lPatternCount
            If MultiPat = True Then
                Call thehdw.Patterns(MultiPatAry(CountMultiPat)).start
            Else
                Call thehdw.Patterns(patt).start ' make sure to jump out  the cpu loop before halt
            End If
            While thehdw.Digital.Patgen.IsRunning = True
                thehdw.Digital.Patgen.Continue 0, cpuA
            Wend
            thehdw.Digital.Patgen.HaltWait
        Next Count
        '------------------
        '------------------
    '===============================================================================
    '20190319 update
    'if SuspendDatalog=false the Tname need to include all information from X,Y,Z current point
        Dim Suspend_Flag As Boolean
        Dim TnameCombShmooInfo As String
        Dim ConditionString As String
        If theexec.DevChar.Setups.IsRunning = True Then Suspend_Flag = theexec.DevChar.Setups.item(theexec.DevChar.Setups.ActiveSetupName).Output.SuspendDatalog    '#16_ELSE_CASE_CHK
    '===============================================================================

        If TPModeAsCharz_GLB = True Then
            For Each site In theexec.sites
                TnameCombShmooInfo = vbNullString
                TestNumber = theexec.sites.item(site).TestNumber
                If thehdw.Digital.Patgen.PatternBurstPassed(site) Then
                    theexec.sites.item(site).testResult = sitePass
                    If theexec.DevChar.Setups.IsRunning = True Then
                        If Suspend_Flag = False Then
                            Call PrintEachPoint_TestName(TnameCombShmooInfo)
                            ConditionString = TnameCombShmooInfo
                            If Glb_RetentionMeasurement = True Then
                                TnameCombShmooInfo = theexec.DataManager.instancename
                            Else
                                TnameCombShmooInfo = theexec.DataManager.instancename & TnameCombShmooInfo
                            End If
                            If g_TestNum <> -1 Then
                                Call theexec.Datalog.WriteFunctionalResult(site, g_TestNum, logTestPass, , TnameCombShmooInfo)
                            Else
                                Call theexec.Datalog.WriteFunctionalResult(site, TestNumber, logTestPass, , TnameCombShmooInfo)
                            End If
                        Else
                            Call theexec.Datalog.WriteFunctionalResult(site, TestNumber, logTestPass)
                        End If
                    Else
                        Call theexec.Datalog.WriteFunctionalResult(site, TestNumber, logTestPass)
                    End If
                Else
                    theexec.sites.item(site).testResult = siteFail
                    If theexec.DevChar.Setups.IsRunning = True Then
                        If Suspend_Flag = False Then
                            Call PrintEachPoint_TestName(TnameCombShmooInfo)
                            TnameCombShmooInfo = theexec.DataManager.instancename & TnameCombShmooInfo
                            ConditionString = TnameCombShmooInfo
                            If Glb_RetentionMeasurement = True Then
                                TnameCombShmooInfo = theexec.DataManager.instancename
                            End If
                            If g_TestNum <> -1 Then
                               Call theexec.Datalog.WriteFunctionalResult(site, g_TestNum, logTestFail, , TnameCombShmooInfo)
                            Else
                               Call theexec.Datalog.WriteFunctionalResult(site, TestNumber, logTestFail, , TnameCombShmooInfo)
                            End If
                        Else
                            Call theexec.Datalog.WriteFunctionalResult(site, TestNumber, logTestFail)
                        End If
                    Else
                        Call theexec.Datalog.WriteFunctionalResult(site, TestNumber, logTestFail)
                    End If
                   ''-----------------------------------------------------------------------------------------------
    '                    If LCase(patt) Like "*rtos*" Then Call RTOS_BCS(patt, RTOSPatIndex)
                    If LCase(theexec.DataManager.instancename) Like "*rtos*" Then Call RTOS_BCS(patt, site, RTOSPatIndex)   '#16_ELSE_CASE_CHK
                   ''------------------------------------------------------------------------------------------------
                End If
                If theexec.DevChar.Setups.IsRunning = False Then theexec.sites.item(site).TestNumber = TestNumber + 1   '#16_ELSE_CASE_CHK
                If theexec.DevChar.Setups.IsRunning = True And Suspend_Flag = False And g_TestNum = -1 Then theexec.sites.item(site).TestNumber = TestNumber + 1    '#16_ELSE_CASE_CHK
            Next site
        Else
            HardIP_WriteFuncResult
        End If
        
        If Suspend_Flag = False And g_TestNum <> -1 Then g_TestNum = g_TestNum + 1

    Next CountMultiPat
    
    '=============================================================Process DSP Capture out =================================================================
    If digcap_flag = True Then
        Call CreateSimulateDataDSPWave(OutDspWave, CLng(DigCapSize), CLng(DigCapSize))
        Call Char_Process_DSP_Capture(DigCapName, OutDspWave, DigCap_Info_Dict, CStr(DigCap_Pin))
    End If
     '======================================================================================================================================================
    For Each site In theexec.sites
        '20170213 prevent over write shmoo pattern
        DebugPrintFunc patt.value
    Next site
    'add for retention
    
    If Vbump = True Then 'Vbump function
        If wait_time <> "" And g_PLSWEEP = False Then   '#16_ELSE_CASE_CHK
            g_shmoo_ret = True
            If theexec.DevChar.Setups.IsRunning = True Then
                If g_Retention_Info = "" Then
                    g_Retention_Info = patt.value & "|" & wait_time
                Else
'                    g_Retention_Info = g_Retention_Info & "," & patt.value & "|" & Wait_Time
                End If
                Shmoo_Restore_Power_per_site_Vbump_Retention powerPin, True
            Else
                Shmoo_Restore_Power_per_site_Vbump_Retention powerPin, False
            End If
            Power_Level_Last = ""
            If set_init = True Then
                theexec.Datalog.WriteComment "wait " & wait_time & " after init pattern " & seq
            Else
                theexec.Datalog.WriteComment "wait " & wait_time & " after payload pattern " & seq
            End If
            thehdw.Wait CDbl(wait_time) / 2
            T0 = theexec.Timer(0)
            ''221116 Retention_meas
            'If TheExec.DevChar.Setups.IsRunning = True Then
            If Glb_RetentionMeasurement = True Then
                Pins_Measure_Case ConditionString
            End If
            'End If
            T1 = theexec.Timer(T0)
            If T1 >= (CDbl(wait_time) / 2) Then
            Else
                thehdw.Wait (CDbl(wait_time) / 2) - T1
            End If
           '
            If theexec.Flow.enableWord("Enable_RET_RampDownUp") = False And theexec.DevChar.Setups.IsRunning = True Then
                Retention_RampdownUp Shmoo_Apply_Pin, "UP"
            End If
           
        ElseIf wait_time <> "" And g_PLSWEEP = True Then ' Disrtub retention function
           If set_init = True Then
              theexec.Datalog.WriteComment "wait " & wait_time & " after init pattern " & seq
           Else
              theexec.Datalog.WriteComment "wait " & wait_time & " after payload pattern " & seq
           End If
           thehdw.Wait CDbl(wait_time)
        End If
        
    Else ' without Vbump function
        If theexec.Flow.enableWord("Enable_RET_RampDownUp") = False Then
    
            If wait_time <> "" Then         ' add for wait time between patterns    #16_ELSE_CASE_CHK
                Power_Level_Last = "Sweep" '20181101 move out varint from site-loop
                For Each site In theexec.sites
                    Shmoo_Restore_Power_per_site Shmoo_Apply_Pin, ShmooSweepPower, "Restore to Sweep V"
                    If theexec.DevChar.Setups.IsRunning = False Then
                        Shmoo_Set_Retention_Power False ' for functional test
                    Else
                        Shmoo_Set_Retention_Power True  ' Skip set retention power for shmoo pin
                    End If
                    print_core_power "Retention Power", Shmoo_Apply_Pin
                
                    '20170213 prevent over write shmoo pattern
                    DebugPrintFunc patt.value
                Next site
                If set_init = True Then
                    theexec.Datalog.WriteComment "wait " & wait_time & " after init pattern " & seq
                Else
                    theexec.Datalog.WriteComment "wait " & wait_time & " after payload pattern " & seq
                End If
                thehdw.Wait CDbl(wait_time)
            End If
        Else
            If wait_time <> "" Then         ' add for wait time between patterns    #16_ELSE_CASE_CHK
                Dim RetPowers As Double
                Dim RetPins As New PinList
                Dim Retention_V(100) As New SiteDouble
                For Each site In theexec.sites: For i = 0 To 99: Retention_V(i)(site) = 0: Next i: Next site ' initialize Retention_V array #16_ELSE_CASE_CHK
                Decide_retetntion_power Retention_V(), RetPins
                If RetPins <> "" Then   '#16_ELSE_CASE_CHK
                    Call MbistRetentionLevelWait_ForChar(CDbl(wait_time) * 1000, Retention_V(), RetPins, 10, 0)
                Else
                End If
                If set_init = True Then
                    theexec.Datalog.WriteComment "wait " & wait_time & "after init pattern " & seq
                Else
                    theexec.Datalog.WriteComment "wait " & wait_time & "after payload pattern " & seq
                End If
            End If
        End If
    End If
    patt.value = pattstr
    'On Error GoTo 0
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Test_Pattern")
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function RTOS_BCS(patt As Pattern, site As Variant, Optional RTOSPatIndex As Integer)
On Error GoTo errHandler

    Dim w_CurrFailingPat As String
    Dim w_CurrFailingVector As Integer
    '20170428 add case "C", judge TestDone srm
    Dim r_TestDoneIdx As Long
    Dim PattArray() As String
    Dim PatCount As Long
    Dim i As Integer
    Dim w_BootIndex As Integer, w_BistDownIndex As Integer, w_HaltIndex As Integer, w_cmdIndex As Integer
    Dim VectorStr As String
    Dim w_CmdStrFlag As Boolean
    Dim pattmp() As String
    
    
''-----------------------------------------------------------------------------------------------
'                    ''20170425
      w_CurrFailingPat = thehdw.Digital.hram.PatGenInfo(0, pgPattern)
      w_CurrFailingVector = thehdw.Digital.hram.PatGenInfo(0, pgVector)
      w_CmdStrFlag = True

      Call PATT_GetPatListFromPatternSet(patt.value, PattArray, PatCount)
      pattmp = Split(PattArray(0), ":")
      PattArray(0) = pattmp(0)
'      If LCase(PattArray(0)) Like "*rtos*" Then ' only RTOS pattern entry
          For i = 0 To 1000
              VectorStr = thehdw.Digital.Patterns(PattArray(0)).GetCommandString("", i)
'              If LCase(VectorStr) Like "*ready_wait_loop*" Then 'Keyword from boot done
              If LCase(VectorStr) Like "*rdywait*" Then 'Keyword from boot done
                 w_BootIndex = i + RTOSPatIndex
'                  w_BootIndex = i + 0
'              ElseIf LCase(VectorStr) Like "*cmd_done*" Then 'Keyword from command done
              ElseIf LCase(VectorStr) Like "*cmddone*" Then 'Keyword from command done
                  If w_CmdStrFlag = True Then
                      w_cmdIndex = i - 35  ' Cyprus 20170902 pat
                      w_CmdStrFlag = False
                  End If
'              ElseIf LCase(VectorStr) Like "*test_done*" Then 'Keyword from Scenrio done
              ElseIf LCase(VectorStr) Like "*tstdone*" Then 'Keyword from Scenrio done
                  w_BistDownIndex = i - 1
              ElseIf LCase(VectorStr) Like "*halt*" Then
                  w_HaltIndex = i
                  Exit For
              End If
          Next i
          
          If w_BootIndex - RTOSPatIndex = 0 And w_cmdIndex = 0 And w_BistDownIndex = 0 Then
            ShmResult(site) = ShmResult(site) & "-"
            GoTo bypass
          End If
          
          If LCase(w_CurrFailingPat) Like "*rdywait*" Then
              ShmResult(site) = ShmResult(site) & "B"
          ElseIf LCase(w_CurrFailingPat) Like "*cmddone*" Then
             ShmResult(site) = ShmResult(site) & "C"
          ElseIf LCase(w_CurrFailingPat) Like "*tstdone*" Then
             ShmResult(site) = ShmResult(site) & "S"
          Else 'VM
              If w_CurrFailingVector <= w_BootIndex Then
                 ShmResult(site) = ShmResult(site) & "B"
              ElseIf w_CurrFailingVector > w_BootIndex And w_CurrFailingVector < w_BistDownIndex Then
                 ShmResult(site) = ShmResult(site) & "C"
              ElseIf w_CurrFailingVector > w_BistDownIndex Then
                 ShmResult(site) = ShmResult(site) & "S"
              End If
          End If
'      End If
'----------------------------------------------------------------------------------------------
bypass:
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "RTOS_BCS")
    If AbortTest Then Exit Function Else Resume Next
End Function


'20170104 Roy modified
'Replace module from "Select case" to "If-else"

Public Function Set_Run_Level_bk(Power_Run_Scenario As String, powerPin As String, set_init As Boolean, seq As Long)
On Error GoTo errHandler
    Dim VoltageLevel As String, Scenario As String
    Dim i As Long
    Dim init_level As String
    Dim pl_level As String
    Dim Power_Run_Scenario_ary() As String
    Dim inst_name As String
    Dim inst_level As String
    Dim site As Variant
    
    Power_Run_Scenario_ary = Split(Power_Run_Scenario, "_")
    inst_name = theexec.DataManager.instancename
    inst_level = right(theexec.DataManager.instancename, 2)
    init_level = "-99"
    pl_level = "-99"

    If set_init = True Then
            
        If LCase(Power_Run_Scenario) Like LCase("*Init_Sweep*") Then
            init_level = "Sweep"
            If Not (Power_Level_Last Like init_level) Then
                For Each site In theexec.sites
                    Shmoo_Restore_Power_per_site Shmoo_Apply_Pin, ShmooSweepPower, "*** Char_Init" & seq & "_" & inst_level & "_Sweep ***"
                Next site
            End If
        ElseIf LCase(Power_Run_Scenario) Like LCase("*Init_[NHL]V*") Then
            init_level = mid(Power_Run_Scenario, InStr(LCase(Power_Run_Scenario), "init_") + 5, 2)
            If Not (Power_Level_Last Like init_level) Then Shmoo_Set_Power Shmoo_Apply_Pin, init_level, "*** Char_Init" & seq & "_" & init_level & " ***", True
        ElseIf LCase(Power_Run_Scenario) Like LCase("*init" & seq & "_Sweep*") Then
            init_level = "Sweep"
            If Not (Power_Level_Last Like init_level) Then
                For Each site In theexec.sites
                    Shmoo_Restore_Power_per_site Shmoo_Apply_Pin, ShmooSweepPower, "*** Char_Init" & seq & "_" & inst_level & "_Sweep ***"
                Next site
            End If
        ElseIf LCase(Power_Run_Scenario) Like LCase("*init" & seq & "_[NHL]V*") Then
            init_level = mid(Power_Run_Scenario, InStr(LCase(Power_Run_Scenario), "init" & seq & "_") + 6, 2)
            If Not (Power_Level_Last Like init_level) Then Shmoo_Set_Power Shmoo_Apply_Pin, init_level, "*** Char_Init" & seq & "_" & init_level & " ***", True
        End If
        Power_Level_Last = init_level
        If init_level Like "-99" Then
            Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Set_Run_Level", "Power Run Scenario " & Power_Run_Scenario & " is not Supported")
'            TheExec.ErrorLogMessage "Power Run Scenario " & Power_Run_Scenario & " is not supported"
        Else
        End If
    Else
            
        If LCase(Power_Run_Scenario) Like LCase("*pl_Sweep*") Then
            pl_level = "Sweep"
            If Not (Power_Level_Last Like pl_level) Then
                For Each site In theexec.sites
                    Shmoo_Restore_Power_per_site Shmoo_Apply_Pin, ShmooSweepPower, "*** PL" & seq & "_Sweep ***"
                Next site
            Else
                For Each site In theexec.sites
                    print_core_power "*** PL" & seq & "_Sweep ***", Shmoo_Apply_Pin
                Next site
            End If
        ElseIf LCase(Power_Run_Scenario) Like LCase("*pl_[NHL]V*") Then
            pl_level = mid(Power_Run_Scenario, InStr(LCase(Power_Run_Scenario), "pl_") + 3, 2)
            If g_Retention_Shmoo = True Then
               'For retention payload, use force condition instead of N/L/HV for force pin
                'Modify for force condition "VRET" 20171213
                    If g_ForceCond_VDD <> "" Or g_Retention_VDD <> "" Then
                        Shmoo_Restore_Power_per_site Shmoo_Apply_Pin, ShmooSweepPower, "*** PL" & seq & "_" & pl_level & " ***" & pl_level & " Force***", g_ForceCond_VDD
                    End If
                Shmoo_Set_Power Shmoo_Apply_Pin, pl_level, "*** PL" & seq & "_" & pl_level & " ***", True, g_ForceCond_VDD
            Else
                If Not (Power_Level_Last Like pl_level) Then
                    Shmoo_Set_Power Shmoo_Apply_Pin, pl_level, "*** PL" & seq & "_" & pl_level & " ***", True
'                Else
'                    print_core_power "*** PL" & seq & "_" & pl_level & " ***", Shmoo_Apply_Pin
                End If
            End If
        ElseIf LCase(Power_Run_Scenario) Like LCase("*pl" & seq & "_Sweep*") Then
            pl_level = "Sweep"
            If Not (Power_Level_Last Like pl_level) Then
                For Each site In theexec.sites
                    Shmoo_Restore_Power_per_site Shmoo_Apply_Pin, ShmooSweepPower, "*** PL" & seq & "_Sweep ***"
                Next site
            Else
                For Each site In theexec.sites
                    print_core_power "*** PL" & seq & "_Sweep ***", Shmoo_Apply_Pin
                Next site
            End If
        ElseIf LCase(Power_Run_Scenario) Like LCase("*pl" & seq & "_[NHL]V*") Then
            pl_level = mid(Power_Run_Scenario, InStr(LCase(Power_Run_Scenario), "pl" & seq & "_") + 4, 2)
            If g_Retention_Shmoo = True Then
               'For retention payload, use force condition instead of N/L/HV for force pin
                'Modify for force condition "VRET" 20171213
                If g_ForceCond_VDD <> "" Or g_Retention_VDD <> "" Then
                    Shmoo_Restore_Power_per_site Shmoo_Apply_Pin, ShmooSweepPower, "*** PL" & seq & "_" & pl_level & " Force***", g_ForceCond_VDD
                End If
                Shmoo_Set_Power Shmoo_Apply_Pin, pl_level, "*** PL" & seq & "_" & pl_level & " ***", True, g_ForceCond_VDD
             Else
                If Not (Power_Level_Last Like pl_level) Then
                    Shmoo_Set_Power Shmoo_Apply_Pin, pl_level, "*** PL" & seq & "_" & pl_level & " ***", True
'                Else
'                    print_core_power "*** PL" & seq & "_" & pl_level & " ***", Shmoo_Apply_Pin
                End If
            End If
        End If
           
        Power_Level_Last = pl_level
        If pl_level Like "-99" Then
            Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Set_Run_Level", "Power Run Scenario " & Power_Run_Scenario & " is not Supported")
'            TheExec.ErrorLogMessage "Power Run Scenario " & Power_Run_Scenario & " is not supported"
        Else
        End If
    End If
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Set_Run_Level")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function gen_search_string(SetupName As String, ByRef Search_String As String, axis_type As tlDevCharShmooAxis, Optional RangeFrom As Double, Optional RangeTo As Double, Optional RangeStepSize As Double, Optional RangeSteps As Long)
On Error GoTo errHandler
''
    Search_String = vbNullString
    ShmooPowerName = vbNullString
    Dim axis_header As String, p As Variant
    Dim RangeFromTracking As Variant, RangeToTracking As Variant, RangeStepSizeTracking As Variant, RangeStepsTracking As Variant
    Dim StepName As Variant, Pin_Ary() As String, shmoo_pin_string As String, PinName As Variant
    Dim StepNameTrack As Variant
    Dim Search_String_Main As String
    Dim Search_String_Tracking As String
    
    Select Case axis_type
            Case tlDevCharShmooAxis_X: axis_header = "X@"
            Case tlDevCharShmooAxis_Y: axis_header = "Y@"
    End Select
    
    If theexec.DevChar.Setups(SetupName).Output.Format Like "SwapXY" Then
        Select Case axis_type
                Case tlDevCharShmooAxis_X: axis_header = "Y@"
                Case tlDevCharShmooAxis_Y: axis_header = "X@"
        End Select
    Else
        Select Case axis_type
                Case tlDevCharShmooAxis_X: axis_header = "X@"
                Case tlDevCharShmooAxis_Y: axis_header = "Y@"
        End Select
    End If
    
    With theexec.DevChar

        RangeStepsTracking = RangeSteps - 1 'tracking steps is the same as main
        If .Setups(SetupName).Shmoo.Axes(axis_type).ApplyTo.Pins <> vbNullString Then   '#16_ELSE_CASE_CHK
            Pin_Ary = Split(.Setups(SetupName).Shmoo.Axes(axis_type).ApplyTo.Pins, ",")
            shmoo_pin_string = .Setups(SetupName).Shmoo.Axes(axis_type).ApplyTo.Pins
            For Each PinName In Pin_Ary
            ShmooPowerName = ShmooPowerName & "_" & PinName
                 Search_String_Main = Search_String_Main & axis_header & PinName & "="                                      ' need to modify 0.0000
                 Search_String_Main = Search_String_Main & Format(RangeFrom, "0.0000########") & ":"                                ' need to modify 0.0000
                 Search_String_Main = Search_String_Main & Format(RangeTo, "0.0000########") & ":"                                  ' need to modify 0.0000
                 Search_String_Main = Search_String_Main & Format(RangeStepSize, "0.0000########") & ","                            ' need to modify 0.0000
            Next PinName
        ElseIf LCase(.Setups.item(SetupName).Shmoo.Axes.item(axis_type).Parameter.type) Like "*spec" Then
        ShmooPowerName = ShmooPowerName & "_" & PinName
            PinName = .Setups.item(SetupName).Shmoo.Axes.item(axis_type).Parameter.name
            Search_String_Main = Search_String_Main & axis_header & PinName & "="
            Search_String_Main = Search_String_Main & Format(RangeFrom, "0.0000") & ":"                                     ' need to modify 0.0000
            Search_String_Main = Search_String_Main & Format(RangeTo, "0.0000") & ":"                                       ' need to modify 0.0000
            Search_String_Main = Search_String_Main & Format(RangeStepSize, "0.0000") & ","                                 ' need to modify 0.0000
        Else
        End If
        With .Setups.item(SetupName).Shmoo.Axes.item(axis_type).TrackingParameters
            For Each StepNameTrack In .list
                RangeFromTracking = .item(StepNameTrack).range.from
                RangeToTracking = .item(StepNameTrack).range.to
                RangeStepSizeTracking = (RangeToTracking - RangeFromTracking) / RangeStepsTracking
                If .item(StepNameTrack).ApplyTo.Pins <> "" Then '#16_ELSE_CASE_CHK
                       Pin_Ary = Split(.item(StepNameTrack).ApplyTo.Pins, ",")
                       shmoo_pin_string = shmoo_pin_string & "," & .item(StepNameTrack).ApplyTo.Pins
                       For Each p In Pin_Ary
                       ShmooPowerName = ShmooPowerName & "_" & p
                          Search_String_Tracking = Search_String_Tracking & axis_header & p & "="
                          Search_String_Tracking = Search_String_Tracking & Format(RangeFromTracking, "0.0000########") & ":"       ' need to modify 0.0000
                          Search_String_Tracking = Search_String_Tracking & Format(RangeToTracking, "0.0000########") & ":"         ' need to modify 0.0000
                          Search_String_Tracking = Search_String_Tracking & Format(RangeStepSizeTracking, "0.0000########") & ","   ' need to modify 0.0000
                       Next p
                ElseIf .item(StepNameTrack).type Like "*Spec" Then
                    PinName = .item(StepNameTrack).name
                    ShmooPowerName = ShmooPowerName & "_" & PinName
                    Search_String_Tracking = Search_String_Tracking & axis_header & PinName & "="
                    Search_String_Tracking = Search_String_Tracking & Format(RangeFromTracking, "0.0000") & ":"             ' need to modify 0.0000
                    Search_String_Tracking = Search_String_Tracking & Format(RangeToTracking, "0.0000") & ":"               ' need to modify 0.0000
                    Search_String_Tracking = Search_String_Tracking & Format(RangeStepSizeTracking, "0.0000") & ","         ' need to modify 0.0000
                Else
                End If
            Next StepNameTrack
        End With
        Search_String = Search_String_Tracking & Search_String_Main
     End With
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "gen_search_string")
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function ShmooPostStep2Dto1D(argc As Long, argv() As String)
On Error GoTo errHandler

    Dim SetupName As String
    Dim i As Long
    Dim OutputString As String
    Dim instancename As String
    Dim TestNum As Long
    Dim lvccf As Integer
    Dim Lvcc As Double
    Dim site As Variant
'    Dim v_Xi0 As Double
'    Dim xio_spec As String
    Dim TestVoltage As String
    Dim StartVoltage As Double, EndVoltage As Double, StepSize As Double
    Dim Patt_String As String, Shmoo_Result As String
    Dim Pat As Variant
    Dim PinName As Variant
    Dim StepName As Variant
    Dim RangeFrom As Double, RangeTo As Double, RangeStepSize As Double, RangeSteps As Long
    Dim RangeLow As Double
    Dim RangeCalcType As tlDevCharRangeField
    Dim allPowerPins As String
    Dim PowerPinCnt As Long, PowerPinAry() As String
    Dim FlagFirstPass As Boolean
    Dim last_point_result As tlDevCharResult, current_point_result As tlDevCharResult
    Dim min_point As Long, max_point As Long, current_point As Long
    Dim Vcc_min As String, Vcc_max As String
    Dim patt_ary() As String, pat_count As Long, p As Variant
    Dim Pin_Ary() As String, Pin_Cnt As Long
    Dim shmoo_pin_string As String
    Dim tmp As String
    Dim Search_String As String
    Dim FlagHole As Boolean
    Dim Shmoo_hole As String
    Dim FlagPF(1000) As Boolean
    Dim FlagFP(1000) As Boolean
    Dim FlagPF_Count As Long
    Dim FlagFP_Count As Long
    Dim ch As String
    Dim Group As Boolean
    Dim Label As String
    Dim step_Start As Long
    Dim step_Stop As Long
    Dim Step_x As Long
    Dim Range_temp As Double
    Dim range_plus As Long
    Dim Shmoo_Pattset As New Pattern
    Dim CharShmooAxis_Inter As tlDevCharShmooAxis
    Dim CharShmooAxis_Outer As tlDevCharShmooAxis
    
    Dim CharShmooType_X As String
    Dim CharShmooType_Y As String
    Dim patset As Variant, patset1 As Variant, j As Long
    Dim Outer_StepName As Variant
    Dim Outer_RangeFrom As Double, Outer_RangeTo As Double, Outer_RangeStepSize As Double, Outer_RangeSteps As Long
    Dim Outer_ParameterName As String   '20180716 Auto parsing FRC info
    Dim Outer_Step_start As Long
    Dim Outer_Step_stop As Long
    Dim Outer_Step_Index As Long
    Dim j_Outer  As Long
    
    Dim HIO_PinName_Updated As Boolean      '20180515 TER
    
    Dim index As Long
    Dim vbump_value As String
        
    
    Shmoo_hole = "NH"
    
    instancename = theexec.DataManager.instancename     '20180616 TER
    Call Get_Tname_FromFlowSheet(instancename, HIO_PinName_Updated)      '20180515 TER
    
    For Each site In theexec.sites
        OutputString = vbNullString
        lvccf = 0

        
        If theexec.DevChar.Setups.item(theexec.DevChar.Setups.ActiveSetupName).Output.SuspendDatalog = False Then    '20180718 add
            Call theexec.sites(site).IncrementTestNumber
        End If
        
        TestNum = theexec.sites(site).TestNumber
        
        If theexec.DevChar.Setups.item(theexec.DevChar.Setups.ActiveSetupName).Output.SuspendDatalog = True Then    '20180718 add
            Call theexec.sites(site).IncrementTestNumber
        End If
        
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        SetupName = theexec.DevChar.Setups.ActiveSetupName
        With theexec.DevChar
            StepName = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).StepName
            RangeFrom = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.from
            RangeTo = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.to
            RangeSteps = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.Steps + 1
            RangeStepSize = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.StepSize
            RangeCalcType = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.CalculatedField
            

            If RangeFrom < RangeTo Then ' always start from lower Value
                step_Start = 0
                step_Stop = RangeSteps - 1
                Step_x = 1
                RangeLow = RangeFrom
               range_plus = -1
            Else
                step_Start = RangeSteps - 1
                step_Stop = 0
                Step_x = -1
                If RangeCalcType = tlDevCharRangeField_Steps Then 'calculate step
                    RangeTo = RangeFrom - (RangeSteps - 1) * RangeStepSize
                End If
                RangeLow = RangeTo
                range_plus = 1
            End If
        End With
        
        Patt_String = vbNullString
        
        With theexec.DevChar.Results(SetupName).Shmoo
            FlagPF_Count = 1
            FlagFP_Count = 1
            For i = 0 To 9
                FlagPF(i) = False
                FlagFP(i) = False
            Next i
        End With
        j = 0
            
        Shmoo_Pattset.value = Shmoo_Pattern
        Patt_String = PatSetToPat(Shmoo_Pattset)
    
        CharShmooType_X = theexec.DevChar.Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.type.value
        CharShmooType_Y = theexec.DevChar.Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.type.value
        
        If CharShmooType_X = "Level" Or CharShmooType_X = "DC Spec" Then
            CharShmooAxis_Inter = tlDevCharShmooAxis_X
            CharShmooAxis_Outer = tlDevCharShmooAxis_Y
        ElseIf CharShmooType_Y = "Level" Or CharShmooType_Y = "DC Spec" Then
            CharShmooAxis_Inter = tlDevCharShmooAxis_Y
            CharShmooAxis_Outer = tlDevCharShmooAxis_X
        Else '' 20180710 to avoid wrong result in output string when 2D shmoo is not include power pin
            CharShmooAxis_Inter = tlDevCharShmooAxis_X
            CharShmooAxis_Outer = tlDevCharShmooAxis_Y
        End If
        
        With theexec.DevChar
            Outer_StepName = .Setups(SetupName).Shmoo.Axes(CharShmooAxis_Outer).StepName
            Outer_RangeFrom = .Setups(SetupName).Shmoo.Axes(CharShmooAxis_Outer).Parameter.range.from
            Outer_RangeTo = .Setups(SetupName).Shmoo.Axes(CharShmooAxis_Outer).Parameter.range.to
            Outer_RangeSteps = .Setups(SetupName).Shmoo.Axes(CharShmooAxis_Outer).Parameter.range.Steps + 1
            Outer_ParameterName = .Setups(SetupName).Shmoo.Axes(CharShmooAxis_Outer).Parameter.name.value '20180716 Auto parsing FRC info
            Outer_RangeStepSize = .Setups(SetupName).Shmoo.Axes(CharShmooAxis_Outer).Parameter.range.StepSize
        End With
        
        If Outer_RangeFrom < Outer_RangeTo Then
            Outer_Step_start = 0
            Outer_Step_stop = Outer_RangeSteps - 1
            Outer_Step_Index = 1
        Else
            Outer_Step_start = Outer_RangeSteps - 1
            Outer_Step_stop = 0
            Outer_Step_Index = -1
        End If
         
         
        For j_Outer = Outer_Step_start To Outer_Step_stop Step Outer_Step_Index
        
            Search_String = vbNullString
                                            ''tlDevCharShmooAxis_X
            With theexec.DevChar
                StepName = .Setups(SetupName).Shmoo.Axes(CharShmooAxis_Inter).StepName                                ''tlDevCharShmooAxis_X
                RangeFrom = .Setups(SetupName).Shmoo.Axes(CharShmooAxis_Inter).Parameter.range.from            ''tlDevCharShmooAxis_X
                RangeTo = .Setups(SetupName).Shmoo.Axes(CharShmooAxis_Inter).Parameter.range.to                    ''tlDevCharShmooAxis_X
                RangeSteps = .Setups(SetupName).Shmoo.Axes(CharShmooAxis_Inter).Parameter.range.Steps + 1     ''tlDevCharShmooAxis_X
                               
                If RangeFrom < RangeTo Then
                    step_Start = 0
                    step_Stop = RangeSteps - 1
                    Step_x = 1
                    range_plus = -1
                Else
                     step_Start = RangeSteps - 1
                    step_Stop = 0
                    Step_x = -1
                    range_plus = 1
                End If
                
                If RangeSteps = 0 Then RangeSteps = 1
                If RangeSteps > 0 Then
                   RangeStepSize = (RangeTo - RangeFrom) / (RangeSteps - 1)
                Else
                    If RangeStepSize <> 0 Then
                        RangeSteps = (RangeTo - RangeFrom) / RangeStepSize + 1
                    Else
                        RangeSteps = 1
                    End If
                End If
                gen_search_string SetupName, Search_String, CharShmooAxis_Inter, RangeFrom, RangeTo, RangeStepSize, RangeSteps
                shmoo_pin_string = .Setups(SetupName).Shmoo.Axes(CharShmooAxis_Inter).ApplyTo.Pins    ''tlDevCharShmooAxis_X
                Shmoo_Result = vbNullString
               
               With theexec.DevChar.Results(SetupName).Shmoo
                    min_point = 999
                    max_point = 999
                    current_point_result = tlDevCharResult_Fail
                    last_point_result = tlDevCharResult_Fail
                    FlagFirstPass = False
                       
                      ' For i = 0 To RangeSteps - 1
                      
                        For i = step_Start To step_Stop Step Step_x
                            If CharShmooType_X = "Level" Or CharShmooType_X = "DC Spec" Then
                                current_point_result = .Points(i, j_Outer).ExecutionResult
                            ElseIf CharShmooType_Y = "Level" Or CharShmooType_Y = "DC Spec" Then
                                current_point_result = .Points(j_Outer, i).ExecutionResult
                            Else '' 20180710 to avoid wrong result in output string when 2D shmoo is not include power pin
                                current_point_result = .Points(i, j_Outer).ExecutionResult
                            End If
                            
                            Select Case current_point_result
                            Case tlDevCharResult_Pass:
                                        Shmoo_Result = Shmoo_Result & "+"
                            
                            Case tlDevCharResult_Fail:
                                        Shmoo_Result = Shmoo_Result & "-"
                            
                            Case tlDevCharResult_NoTest:
                                        Shmoo_Result = Shmoo_Result & "_"
                                        current_point_result = last_point_result
                            
                            Case tlDevCharResult_AssumedPass:
                                        Shmoo_Result = Shmoo_Result & "*"
                                        current_point_result = last_point_result
                            
                            Case tlDevCharResult_AssumedFail:
                                        Shmoo_Result = Shmoo_Result & "~"
                                        current_point_result = last_point_result
                            Case Else:
                                        Shmoo_Result = Shmoo_Result & "?"
        
                            End Select
        
                        If last_point_result = tlDevCharResult_Fail And current_point_result = tlDevCharResult_Pass Then
                            FlagFP(FlagFP_Count) = True
                            FlagFP_Count = FlagFP_Count + 1
                        Else
                        End If
                        
                        If last_point_result = tlDevCharResult_Pass And current_point_result = tlDevCharResult_Fail Then
                            FlagPF(FlagPF_Count) = True
                            FlagPF_Count = FlagPF_Count + 1
                        Else
                        End If
                  
                        If current_point_result = tlDevCharResult_Pass And FlagFirstPass = False Then  'find first pass point   #16_ELSE_CASE_CHK
                        'If current_point_result = tlDevCharResult_Pass And last_point_result = tlDevCharResult_Fail Then  'find last F-> P
                            min_point = i
                            FlagFirstPass = True 'always take the first pass point
                        Else
                        End If
                  
                    'If current_point_result = tlDevCharResult_Pass And FlagFirstPass = False Then  'find first pass point
                        If current_point_result = tlDevCharResult_Pass And last_point_result = tlDevCharResult_Fail Then  'find last F-> P  #16_ELSE_CASE_CHK
                            min_point = i
        '                       FlagFirstPass = True 'always take the first pass point
                        Else
                        End If
                  
                        If current_point_result = tlDevCharResult_Fail And last_point_result = tlDevCharResult_Pass Then       'find last pass point    #16_ELSE_CASE_CHK
                            max_point = i + range_plus 'always take the last pass point
                        Else
                        End If
                       
                        last_point_result = current_point_result
                    Next i
                End With
                
                If FlagFP(1) = True And FlagFP(2) = False Then  '#16_ELSE_CASE_CHK
                    Shmoo_hole = "NH"
                Else
                End If
                
                If FlagFP(1) = True And FlagFP(2) = True Then   '#16_ELSE_CASE_CHK
                    Shmoo_hole = "LH"
                Else
                End If
                
                If FlagFP(1) = True And FlagFP(2) = True And FlagFP(3) = True Then  '#16_ELSE_CASE_CHK
                    Shmoo_hole = "BH"
                Else
                End If
                
                If min_point <> 999 Then
                    Vcc_min = CStr(RangeFrom + min_point * RangeStepSize)
                    Shmoo_Vcc_Min(site) = Vcc_min
                Else
                    Vcc_min = "N/A"
                    Shmoo_Vcc_Min(site) = -0.1
                End If
                
                If max_point <> 999 Then
                    Vcc_max = CStr(RangeFrom + max_point * RangeStepSize)
                Else
                    If Vcc_min <> "N/A" Then
                        If range_plus = 1 Then
                            Vcc_max = Format(RangeFrom, "0.000")
                        Else
                            Vcc_max = Format(RangeTo, "0.000")
                        End If
                    Else
                        Vcc_max = "N/A"
                    End If
                End If
                
                If last_point_result = tlDevCharResult_Pass Then
                    If range_plus = 1 Then
                        Vcc_max = Format(RangeFrom, "0.000")
                    Else
                        Vcc_max = Format(RangeTo, "0.000")
                    End If
                    '  Vcc_max = CStr(RangeTo)
                End If
                
                '  If RangeFrom > RangeTo Then
                '     tmp = Vcc_max
                '     Vcc_max = Vcc_min
                '     Vcc_min = tmp
                '  End If
            End With
            If InStr(theexec.DataManager.instancename, "_NV") Then TestVoltage = "NV"   '#16_ELSE_CASE_CHK
            If InStr(theexec.DataManager.instancename, "_HV") Then TestVoltage = "HV"   '#16_ELSE_CASE_CHK
            If InStr(theexec.DataManager.instancename, "_LV") Then TestVoltage = "LV"   '#16_ELSE_CASE_CHK

    '[Char,N99G19-1,16,7,V,0,XI0=24000000,CpuBira_P0001_IN02_BIR_SI_PL00_CL16_BIR_59N_SI_PP_NV,CPU_BIST_CPU_Domain_CPU_SRAM_Domain_P1_Full_Range,1069,
    '.\pattern\CpuMbist\PP_FIJA0_C_IN00_XX_CLXX_XXX_XXX_XXX_P0001_1308131609_SI_mod.pat,.\pattern\CpuMbist\PP_FIJA0_C_IN02_BI_CLXX_BIR_JTG_XXX_ALLFV_1306250000_SI.pat,.\pattern\CpuMbist\PP_FIJA0_C_PL00_BI_CL16_BIR_JTG_59N_ALLFV_1306250000_SI.pat,
    'NV,VDD_FIXED=0.528:1.404:0.005,VDD_VAR_SOC_VAR=0.500:1.330:0.005,
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++,NH,0.5,1.260]
            
            If Shmoo_header = "" Then Shmoo_header = "Char" '#16_ELSE_CASE_CHK
            If (LCase(theexec.CurrentJob) Like "*cp*") Then
                OutputString = OutputString & "[" & Shmoo_header & "," & HramLotId(site) & "-" & CStr(HramWaferId(site)) & "," & CStr(XCoord(site)) & "," & CStr(YCoord(site))
            Else
                OutputString = OutputString & "[" & Shmoo_header & "," & HramLotId(site) & "-" & CStr(HramWaferId(site)) & "," & CStr(HramXCoord(site)) & "," & CStr(HramYCoord(site))
            End If
            

            Dim SetupName_New As String, k As Integer
            Dim InstanceName_New As String
            
            SetupName_New = SetupName
    
            'Shmoo_header
            Dim VIL_Flag As Boolean
            
            VIL_Flag = False
            ShmooPowerName = ShmooPowerName
            'v_Xi0 = thehdw.
            
                    


        '20180716 Auto parsing FRC info
        Dim nWire_port_ary() As String
        Dim nwp As Variant
        Dim port_pa As String, ac_spec_pa As String, pin_pa As String, global_spec_pa As String
        Dim FRC_Name As String, FRC_Value As Double, All_FRC_Status As String
        All_FRC_Status = vbNullString
        nWire_port_ary = Split(nWire_Ports_GLB, ",")
        For Each nwp In nWire_port_ary
            Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
            If LCase(glb_TesterType) = "jaguar" Then
                If thehdw.Protocol.ports(port_pa).Enabled = True Then
                    FRC_Name = Replace(UCase(ac_spec_pa), "_FREQ_VAR", vbNullString)
                    If UCase(Outer_ParameterName) = UCase(ac_spec_pa) Then
                        FRC_Value = Outer_RangeFrom + j_Outer * Outer_Step_Index * Outer_RangeStepSize
                    Else
                        FRC_Value = theexec.Specs.AC(ac_spec_pa).CurrentValue
                    End If
                End If
            Else
                FRC_Name = UCase(pin_pa)
                ' Fixed Y domain is XI0 print problem
                If UCase(Outer_ParameterName) = UCase(ac_spec_pa) Then
                    FRC_Value = Outer_RangeFrom + j_Outer * Outer_Step_Index * theexec.DevChar.Setups(SetupName).Shmoo.Axes(CharShmooAxis_Outer).Parameter.range.StepSize
                Else
                    If thehdw.Digital.Pins(pin_pa).FreeRunningClock.IsRunning Then
                        FRC_Value = thehdw.Digital.Pins(pin_pa).FreeRunningClock.Frequency
                    Else
                        FRC_Value = theexec.Specs.AC(ac_spec_pa).CurrentValue
                    End If
                End If
            End If
            If All_FRC_Status = "" Then
                All_FRC_Status = FRC_Name & "=" & FRC_Value
            Else
                All_FRC_Status = All_FRC_Status & ";" & FRC_Name & "=" & FRC_Value
            End If
        Next nwp
        If FRC_Name = "" Then ' Default use XI0, if no input of FRC info
            FRC_Name = XI0
            FRC_Value = TheExec.Specs.AC(XI0_Diff & "_Freq_VAR").CurrentValue
            All_FRC_Status = FRC_Name & "=" & FRC_Value
        End If
        ' Add for 2D Debug LVCC Boundary
        If theexec.Flow.enableWord("Debug_LVCC_VminBoundary") = True Then
            Dim y_value As Double
            y_value = Outer_RangeFrom + j_Outer * Outer_Step_Index * Outer_RangeStepSize
            If Outer_ParameterName Like "*VDD*" Then
                Dim Y_Pin As String
                Dim Y_Pins() As String
                Dim Y_Tracking_Pins() As String
                Y_Pin = theexec.DevChar.Setups(SetupName).Shmoo.Axes(CharShmooAxis_Outer).ApplyTo.Pins
                Y_Tracking_Pins = theexec.DevChar.Setups(SetupName).Shmoo.Axes(CharShmooAxis_Outer).TrackingParameters.list
                Y_Pins = Split(Y_Pin, ",")
                For i = 0 To UBound(Y_Pins)
                    g_CharInputString_Voltage_Dict(Y_Pins(i)) = y_value
                Next i
                For i = 0 To UBound(Y_Tracking_Pins)
                    g_CharInputString_Voltage_Dict(Y_Tracking_Pins(i)) = y_value
                Next i
            ElseIf Outer_ParameterName Like "XI0*" Then
                Call VaryFreq(XI0_Diff_Port, y_value, Outer_ParameterName)
            Else
                Call theexec.Overlays.ApplyUniformSpecToHW(Outer_ParameterName, y_value)
            End If
        End If
            OutputString = OutputString & ",V," & site & "," & All_FRC_Status & ","    '20180716 Auto parsing FRC info

            OutputString = OutputString & theexec.DataManager.instancename & ShmooPowerName & "," & SetupName_New & "," & CStr(TestNum) & ","       '20180616 TER

            OutputString = OutputString & Patt_String & ","
            OutputString = OutputString & TestVoltage & ","
            
            If argv(0) <> Empty Then
                theexec.DataManager.DecomposePinList argv(0), Pin_Ary, Pin_Cnt
                PinName = argv(0) 'setup voltage
            Else
            End If
            
            If Vbump_for_Interpose = True Then
                Dim PL_DC_conditions_str As String
                PL_DC_conditions_str = Replace(PL_DC_conditions_GLB, ":V:", "=")
                PL_DC_conditions_str = Replace(PL_DC_conditions_str, ";", ",")
                OutputString = OutputString & PL_DC_conditions_str
            
            Else
                For j = 0 To Pin_Cnt - 1
                    PinName = Pin_Ary(j)
                    If theexec.DataManager.ChannelType(PinName) <> "N/C" Then   '#16_ELSE_CASE_CHK
                        If j = 0 Then
                            OutputString = OutputString & PinName & "=" & Format(thehdw.DCVS.Pins(PinName).Voltage.Main.value, "0.000")
                        Else
                            OutputString = OutputString & "," & PinName & "=" & Format(thehdw.DCVS.Pins(PinName).Voltage.Main.value, "0.000")
                        End If
                    Else
                    End If
                Next j
            End If
            
            For i = 1 To argc - 1
                If UCase(argv(i)) = "VIL" Or UCase(argv(i)) = "VOL" Then
                    VIL_Flag = True
                Else
                    theexec.DataManager.DecomposePinList argv(i), Pin_Ary, Pin_Cnt
                    For j = 0 To Pin_Cnt - 1
                        PinName = Pin_Ary(j)
                        If theexec.DataManager.ChannelType(PinName) <> "N/C" Then
                            If Vbump_for_Interpose = True Then
                                index = InStr(LCase(PL_DC_conditions_str), LCase(PinName) & "=")
                                vbump_value = mid(LCase(PL_DC_conditions_str), index + Len(PinName) + 1, 5)
                                OutputString = OutputString & "," & PinName & "=" & vbump_value
                            Else
                                OutputString = OutputString & "," & PinName & "=" & Format(thehdw.DCVS.Pins(PinName).Voltage.Main.value, "0.000")
                            End If
                        End If
                    Next j
                End If
            Next
            PL_DC_conditions_str = vbNullString
            OutputString = OutputString & ","
            Search_String = mid(Search_String, 1, Len(Search_String) - 1) 'take out last ","
            Search_String = Replace(Search_String, "X@", vbNullString)
            OutputString = OutputString & Search_String
            OutputString = OutputString & ","
            OutputString = OutputString & Shmoo_Result & ","
            
            
            ''''''''****20180709  adding for printing Vcc_min/Vcman for specail case ****'''''''''''''''''''''''''''''''''''''''''''
            ''''''''Vcc_min/Vcman = -9999/9999(all fail), -5555/5555(shmoo hole), -7777/7777(alarm/error/unknown)'''''''''''''''''''''
            If Vcc_min = "N/A" And Vcc_max = "N/A" Then  ' shmoo points all fail    #16_ELSE_CASE_CHK
                Vcc_min = "-9999"
                Vcc_max = "9999"
            Else
            End If
            
            If FlagFP(2) = True Or FlagPF(2) = True Then  ' shmoo holes #16_ELSE_CASE_CHK
                Vcc_min = "-5555"
                Vcc_max = "5555"
            Else
            End If
            
            If InStr(Shmoo_Result, "?") Then ' any unknown situations, like "alarm" or "error"  #16_ELSE_CASE_CHK
                Vcc_min = "-7777"
                Vcc_max = "7777"
            Else
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            If VIL_Flag = True Then
                OutputString = OutputString & Shmoo_hole & "," & Vcc_max & "," & Vcc_min & "]"
            Else
                OutputString = OutputString & Shmoo_hole & "," & Vcc_min & "," & Vcc_max & "]"
            End If
            
            theexec.Datalog.WriteComment OutputString
            
            '' Reset to default
            OutputString = vbNullString
            FlagPF_Count = 1
            FlagFP_Count = 1
            For i = 0 To 9
                FlagPF(i) = False
                FlagFP(i) = False
            Next i
            
            If Vcc_min = "N/A" Then
                Shmoo_Vcc_Min(site) = -0.1
            Else
                Shmoo_Vcc_Min(site) = Vcc_min
            End If
                
                If Vcc_max = "N/A" Then
                    If RangeFrom > RangeTo Then
                        Shmoo_Vcc_Max(site) = RangeFrom + 0.1
                    Else
                        Shmoo_Vcc_Max(site) = RangeTo + 0.1
                    End If
                    
                Else
                    Shmoo_Vcc_Max(site) = Vcc_max
                End If
            
            Dim dfc As LVCC_VminBoundary_List
            Set dfc = New LVCC_VminBoundary_List
            If theexec.Flow.enableWord("Debug_LVCC_VminBoundary") = True Then   '#16_ELSE_CASE_CHK
                If (dfc.IsEnableDFC Or dfc.IsEnableFAILLOG) And Not dfc.IsItemEnabled Then Exit For ''' if enable dfc and not list in dfc then exit #16_ELSE_CASE_CHK
                theexec.Datalog.WriteComment Outer_ParameterName & ": " & Outer_RangeFrom + Outer_RangeStepSize * j_Outer
                If Shmoo_Vcc_Min(site) > 0 Then
                    If (LCase(theexec.CurrentJob) Like "*cp*") Then
                        'If g_VminBoundary_selsrm = True Then
                            FailingDatalog_HLvcc_Boundary_SELSRM Search_String, lotId, CStr(WaferID), CStr(XCoord(site)), _
                            CStr(YCoord(site)), Shmoo_Pattern, "Shmoo LVCC", High_to_Low, site, RangeFrom, RangeTo, RangeSteps, Abs(RangeStepSize)
                        'Else
                        '    FailingDatalog_Lvcc_Boundary Search_String, lotid, CStr(WaferID), CStr(XCoord(site)), _
                        '    CStr(YCoord(site)), Shmoo_Pattern, "Shmoo LVCC", High_to_Low, site, RangeFrom, RangeTo, RangeSteps, RangeStepSize
                        'End If
                    Else
                        'If g_VminBoundary_selsrm = True Then
                            FailingDatalog_HLvcc_Boundary_SELSRM Search_String, HramLotId(site), CStr(HramWaferId(site)), CStr(HramXCoord(site)), _
                            CStr(HramYCoord(site)), Shmoo_Pattern, "Shmoo LVCC", High_to_Low, site, RangeFrom, RangeTo, RangeSteps, RangeStepSize
                        'Else
                        '    FailingDatalog_Lvcc_Boundary Search_String, HramLotId(site), CStr(HramWaferId(site)), CStr(HramXCoord(site)), _
                        '    CStr(HramYCoord(site)), Shmoo_Pattern, "Shmoo LVCC", High_to_Low, site, RangeFrom, RangeTo, RangeSteps, RangeStepSize
                        'End If
                    End If
                Else
                End If
            Else
            End If
        Next j_Outer
        
        '20180716 add for 2D shmoo to print force cnodition
        theexec.Datalog.WriteComment "[Force_condition_during_shmoo:" & Charz_Force_Power_condition & "]"
        
        
    Next site
    
    If Vcc_min = "N/A" Then Vcc_min = 9999  '#16_ELSE_CASE_CHK
    
    '20170126 Add Limit judgement
    Dim print_all As Boolean
    Dim print_lvcc As Boolean
    Dim print_hvcc As Boolean

    Dim DFTH_Testname As String
    Dim DFTL_Testname As String
    print_all = False
    print_lvcc = False
    print_hvcc = False
    
    If InStr(instancename, "DFTLH_") <> 0 Or InStr(instancename, "DFTHL_") <> 0 Then print_all = True   '#16_ELSE_CASE_CHK
    If InStr(instancename, "HFLH_") <> 0 Or InStr(instancename, "HFHL_") <> 0 Then print_all = True '#16_ELSE_CASE_CHK
    If InStr(instancename, "MCLH_") <> 0 Or InStr(instancename, "MCHL_") <> 0 Then print_all = True '#16_ELSE_CASE_CHK
    
    If InStr(instancename, "DFTL_") <> 0 Then print_lvcc = True '#16_ELSE_CASE_CHK
    If InStr(instancename, "HFL_") <> 0 Then print_lvcc = True  '#16_ELSE_CASE_CHK
    If InStr(instancename, "MCL_") <> 0 Then print_lvcc = True  '#16_ELSE_CASE_CHK
    
    If InStr(instancename, "DFTH_") <> 0 Then print_hvcc = True '#16_ELSE_CASE_CHK
    If InStr(instancename, "HFH_") <> 0 Then print_hvcc = True  '#16_ELSE_CASE_CHK
    If InStr(instancename, "MCH_") <> 0 Then print_hvcc = True  '#16_ELSE_CASE_CHK
    
    If theexec.DevChar.Setups.item(theexec.DevChar.Setups.ActiveSetupName).Output.SuspendDatalog = True Then
    If RangeFrom < RangeTo Then
        If print_all Or print_lvcc Then
            theexec.Datalog.WriteComment ("Test name :" & instancename & "_" & SetupName & "_Vmin")
            theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, lowVal:=RangeFrom, hiVal:=RangeTo, ForceResults:=tlForceNone, Tname:=instancename & "_" & SetupName & "_Vmin"
        End If
        If print_all Or print_hvcc Then
            theexec.Datalog.WriteComment ("Test name :" & instancename & "_" & SetupName & "_Vmax")
            theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, lowVal:=RangeFrom, hiVal:=RangeTo, ForceResults:=tlForceNone, Tname:=instancename & "_" & SetupName & "_Vmax"
        End If
    Else
        If print_all Or print_lvcc Then
            DFTL_Testname = Replace(instancename, "_CZ_NV", "_")
            DFTL_Testname = Replace(DFTL_Testname, "DFTLH", "DFTL")
            theexec.Datalog.WriteComment ("Test name :" & instancename & "_" & SetupName & "_Vmin")
            theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, lowVal:=RangeTo, hiVal:=RangeFrom, ForceResults:=tlForceNone, Tname:=DFTL_Testname & "_" & SetupName & "_Vmin"
        End If
        If print_all Or print_hvcc Then
            DFTH_Testname = Replace(instancename, "_CZ_NV", "_")
            DFTH_Testname = Replace(DFTH_Testname, "DFTLH", "DFTH")
            theexec.Datalog.WriteComment ("Test name :" & instancename & "_" & SetupName & "_Vmax")
            theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, lowVal:=RangeTo, hiVal:=RangeFrom, ForceResults:=tlForceNone, Tname:=DFTH_Testname & "_" & SetupName & "_Vmax"
        End If
    End If
    Else '20190321 update: Suspend Datalog =False, add g_TestNum
        If RangeFrom < RangeTo Then
            If print_all Or print_lvcc Then '#16_ELSE_CASE_CHK
                theexec.Datalog.WriteComment ("Test name :" & instancename & "_" & SetupName & "_Vmin")
                theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, lowVal:=RangeFrom, hiVal:=RangeTo, ForceResults:=tlForceNone, Tname:=instancename & "_" & SetupName & "_Vmin", tNum:=g_TestNum
                g_TestNum = g_TestNum + 1
            Else
            End If
            If print_all Or print_hvcc Then '#16_ELSE_CASE_CHK
                theexec.Datalog.WriteComment ("Test name :" & instancename & "_" & SetupName & "_Vmax")
                theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, lowVal:=RangeFrom, hiVal:=RangeTo, ForceResults:=tlForceNone, Tname:=instancename & "_" & SetupName & "_Vmax", tNum:=g_TestNum
                g_TestNum = g_TestNum + 1
            Else
            End If
        Else
            If print_all Or print_lvcc Then '#16_ELSE_CASE_CHK
                DFTL_Testname = Replace(instancename, "_CZ_NV", "_")
                DFTL_Testname = Replace(DFTL_Testname, "DFTLH", "DFTL")
                theexec.Datalog.WriteComment ("Test name :" & instancename & "_" & SetupName & "_Vmin")
                theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, lowVal:=RangeTo, hiVal:=RangeFrom, ForceResults:=tlForceNone, Tname:=DFTL_Testname & "_" & SetupName & "_Vmin", tNum:=g_TestNum
                g_TestNum = g_TestNum + 1
            Else
            End If
            If print_all Or print_hvcc Then '#16_ELSE_CASE_CHK
                DFTH_Testname = Replace(instancename, "_CZ_NV", "_")
                DFTH_Testname = Replace(DFTH_Testname, "DFTLH", "DFTH")
                theexec.Datalog.WriteComment ("Test name :" & instancename & "_" & SetupName & "_Vmax")
                theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, lowVal:=RangeTo, hiVal:=RangeFrom, ForceResults:=tlForceNone, Tname:=DFTH_Testname & "_" & SetupName & "_Vmax", tNum:=g_TestNum
                g_TestNum = g_TestNum + 1
            Else
            End If
        End If
    End If

 
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "ShmooPostStep2Dto1D")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function ShmooPostStep2D(argc As Long, argv() As String)
On Error GoTo errHandler
    Dim SetupName As String
    Dim i As Long
    Dim OutputString As String
    Dim instancename As String
    Dim TestNum As Long
    Dim lvccf As Integer
    Dim Lvcc As Double
    Dim site As Variant
    Dim v_Xi0 As Double
    Dim TestVoltage As String
    Dim StartVoltage As Double, EndVoltage As Double, StepSize As Double
    Dim Patt_String As String, Shmoo_Result As String
    Dim Pat As Variant
    Dim PinName As Variant
    Dim StepName As Variant
    Dim RangeFrom As Double, RangeTo As Double, RangeStepSize As Double, RangeSteps As Long
    Dim allPowerPins As String
    Dim PowerPinCnt As Long, PowerPinAry() As String
    Dim FlagFirstPass As Boolean
    Dim last_point_result As tlDevCharResult, current_point_result As tlDevCharResult
    Dim min_point As Long, max_point As Long, current_point As Long
    Dim Vcc_min As String, Vcc_max As String
    Dim patt_ary() As String, pat_count As Long, p As Variant
    Dim Pin_Ary() As String, Pin_Cnt As Long
    Dim shmoo_pin_string As String
    Dim tmp As String
    Dim Search_String As String, Search_String_X As String, Search_String_Y As String
    Dim Group As Boolean
    Dim Label As String
    Dim Shmoo_Pattset As New Pattern
    Dim VIL_Flag As Boolean
    Dim step_Start As Long
    Dim step_Stop As Long
    Dim Step_x As Long
    Dim RangeLow As Double, RangeStart As Double
    Dim Shmoo_hole As String
    Dim RangeCalcType As tlDevCharRangeField
    Dim xio_spec As String
    Dim Range_temp As Double
    Dim range_plus As Long
    
    Dim HIO_PinName_Updated As Boolean      '20180515 TER
    
    Dim index As Long
    Dim vbump_value As String
    Dim v_Shiftin As Double
    Dim Context As String: Context = vbNullString
    Dim TimeSet_Str As String: TimeSet_Str = vbNullString
    
    
    instancename = theexec.DataManager.instancename     '20180616 TER add
    Call Get_Tname_FromFlowSheet(instancename, HIO_PinName_Updated)      '20180515 TER add
    
    For Each site In theexec.sites
        OutputString = vbNullString
        lvccf = 0

        If theexec.DevChar.Setups.item(theexec.DevChar.Setups.ActiveSetupName).Output.SuspendDatalog = False Then    '20180718 add  #16_ELSE_CASE_CHK
            Call theexec.sites(site).IncrementTestNumber
        End If
                
        v_Shiftin = theexec.Specs.AC("ShiftIn_Freq_VAR").CurrentValue
        
        TestNum = theexec.sites(site).TestNumber
        
        If theexec.DevChar.Setups.item(theexec.DevChar.Setups.ActiveSetupName).Output.SuspendDatalog = True Then    '20180718 add   #16_ELSE_CASE_CHK
            Call theexec.sites(site).IncrementTestNumber
        End If
        
        
        '20180716 Auto parsing FRC info
        Dim nWire_port_ary() As String
        Dim nwp As Variant
        Dim port_pa As String, ac_spec_pa As String, pin_pa As String, global_spec_pa As String
        Dim FRC_Name As String, FRC_Value As Double, All_FRC_Status As String
        All_FRC_Status = vbNullString
        nWire_port_ary = Split(nWire_Ports_GLB, ",")
        For Each nwp In nWire_port_ary
            Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
            If LCase(glb_TesterType) = "jaguar" Then    '#16_ELSE_CASE_CHK
                If thehdw.Protocol.ports(port_pa).Enabled = True Then
                    FRC_Name = Replace(UCase(ac_spec_pa), "_FREQ_VAR", vbNullString)
                    FRC_Value = theexec.Specs.AC(ac_spec_pa).CurrentValue
                    If All_FRC_Status = "" Then
                        All_FRC_Status = FRC_Name & "=" & FRC_Value
                    Else
                        All_FRC_Status = All_FRC_Status & ";" & FRC_Name & "=" & FRC_Value
                    End If
                Else
                End If
            Else
            End If
        Next nwp
        If FRC_Name = "" Then ' Default use XI0, if no input of FRC info    #16_ELSE_CASE_CHK
            FRC_Name = XI0
            FRC_Value = TheExec.Specs.AC(XI0_Diff & "_Freq_VAR").CurrentValue
            All_FRC_Status = FRC_Name & "=" & FRC_Value
        Else
        End If
    
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        
        ''Read X axis setup information
        SetupName = theexec.DevChar.Setups.ActiveSetupName
        
        With theexec.DevChar
            StepName = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).StepName
            RangeFrom = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.from
            RangeTo = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.to
            RangeSteps = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.Steps + 1
            RangeStepSize = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.StepSize
            RangeCalcType = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.CalculatedField

            If RangeFrom < RangeTo Then ' always start from lower Value
                step_Start = 0
                step_Stop = RangeSteps - 1
                Step_x = 1
                RangeLow = RangeFrom
                range_plus = -1
            Else
                step_Start = RangeSteps - 1
                step_Stop = 0
                Step_x = -1
                If RangeCalcType = tlDevCharRangeField_Steps Then 'calculate step   #16_ELSE_CASE_CHK
                    RangeTo = RangeFrom - (RangeSteps - 1) * RangeStepSize
                End If
                RangeLow = RangeTo
                range_plus = 1
            End If
        End With
        
        Patt_String = vbNullString
        Dim patset As Variant, j As Long
        Shmoo_Pattset.value = Shmoo_Pattern
        Patt_String = PatSetToPat(Shmoo_Pattset)
        gen_search_string SetupName, Search_String_X, tlDevCharShmooAxis_X, RangeFrom, RangeTo, RangeStepSize, RangeSteps
        
        ''Read Y axis setup information
        With theexec.DevChar
            StepName = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).StepName
            RangeFrom = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.range.from
            RangeTo = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.range.to
            RangeSteps = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.range.Steps + 1
            RangeStepSize = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.range.StepSize
            RangeCalcType = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.range.CalculatedField

            If RangeFrom < RangeTo Then ' always start from lower Value
                step_Start = 0
                step_Stop = RangeSteps - 1
                Step_x = 1
                RangeLow = RangeFrom
                range_plus = -1
            Else
                step_Start = RangeSteps - 1
                step_Stop = 0
                Step_x = -1
                If RangeCalcType = tlDevCharRangeField_Steps Then 'calculate step
                    RangeTo = RangeFrom - (RangeSteps - 1) * RangeStepSize
                End If
                RangeLow = RangeTo
                range_plus = 1
            End If
        End With
        
        gen_search_string SetupName, Search_String_Y, tlDevCharShmooAxis_Y, RangeFrom, RangeTo, RangeStepSize, RangeSteps
        Search_String = Search_String_X & Search_String_Y
        
        If InStr(theexec.DataManager.instancename, "_NV") Then TestVoltage = "NV"   '#16_ELSE_CASE_CHK
        If InStr(theexec.DataManager.instancename, "_HV") Then TestVoltage = "HV"   '#16_ELSE_CASE_CHK
        If InStr(theexec.DataManager.instancename, "_LV") Then TestVoltage = "LV"   '#16_ELSE_CASE_CHK

   
        OutputString = OutputString & "[V," & site & "," & All_FRC_Status & "," & HramLotId(site) & "-" & CStr(HramWaferId(site)) & "," & CStr(XCoord(site)) & "," & CStr(YCoord(site)) & ","  '20180716 Auto parsing FRC info

        OutputString = OutputString & theexec.DataManager.instancename & "," & SetupName & "," & CStr(TestNum) & ","

        OutputString = OutputString & Patt_String & ","
        OutputString = OutputString & TestVoltage & ","
        PinName = argv(0) 'setup voltage
        If argv(0) <> Empty Then    '#16_ELSE_CASE_CHK
            theexec.DataManager.DecomposePinList argv(0), Pin_Ary, Pin_Cnt
            PinName = argv(0) 'setup voltage
        End If
        
         
        If Vbump_for_Interpose = True Then
            Dim PL_DC_conditions_str As String
            PL_DC_conditions_str = Replace(PL_DC_conditions_GLB, ":V:", "=")
            PL_DC_conditions_str = Replace(PL_DC_conditions_str, ";", ",")
            OutputString = OutputString & PL_DC_conditions_str
            
        Else
            For j = 0 To Pin_Cnt - 1
                PinName = Pin_Ary(j)
                If theexec.DataManager.ChannelType(PinName) <> "N/C" Then
                    If j = 0 Then
                        OutputString = OutputString & PinName & "=" & Format(thehdw.DCVS.Pins(PinName).Voltage.Main.value, "0.000")
                    Else
                         OutputString = OutputString & "," & PinName & "=" & Format(thehdw.DCVS.Pins(PinName).Voltage.Main.value, "0.000")
                    End If
                Else
                End If
            Next j
        End If
        
        For i = 1 To argc - 1
          If UCase(argv(i)) = "VIL" Or UCase(argv(i)) = "VOL" Then
            VIL_Flag = True
          Else
            theexec.DataManager.DecomposePinList argv(i), Pin_Ary, Pin_Cnt
            For j = 0 To Pin_Cnt - 1
                PinName = Pin_Ary(j)
                If theexec.DataManager.ChannelType(PinName) <> "N/C" Then
                    If Vbump_for_Interpose = True Then
                        index = InStr(LCase(PL_DC_conditions_str), PinName & "=")
                        vbump_value = mid(LCase(PL_DC_conditions_str), index + Len(PinName) + 1, 5)
                        OutputString = OutputString & "," & PinName & "=" & vbump_value
                        
                    Else
                        OutputString = OutputString & "," & PinName & "=" & Format(thehdw.DCVS.Pins(PinName).Voltage.Main.value, "0.000")
                    End If
                End If
            Next j
          End If
        Next
        PL_DC_conditions_str = vbNullString
        OutputString = OutputString & ","
       Search_String = mid(Search_String, 1, Len(Search_String) - 1) 'take out last ","
        OutputString = OutputString & Search_String
        OutputString = OutputString & "]"
        theexec.Datalog.WriteComment OutputString
                
        ' Add for print 2D ShiftIn Value
        Context = theexec.Contexts.ActiveSelection
        TimeSet_Str = theexec.Contexts(Context).Sheets.Timesets
        theexec.Datalog.WriteComment "[Activity_Timing_Sheet:" & UCase(TimeSet_Str) & "," & "Shiftin_Freq=" & CStr(v_Shiftin) & "]"
    Next site

          Charz_Force_Power_condition = vbNullString

    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "ShmooPostStep2D")
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function ShmooPostStep1D(argc As Long, argv() As String)
On Error GoTo errHandler

    '
    ' Assume IO are all with NV value in the level sheet
    '
    Dim SetupName As String
    Dim i As Long
    Dim OutputString As String
    Dim instancename As String
    Dim TestNum As Long
    Dim lvccf As Integer
    Dim Lvcc As Double
    Dim site As Variant
    Dim v_Shiftin As Double
    Dim LVCCResult As New SiteDouble '''''''''''jade-S use
    
    Dim TestVoltage As String
    Dim StartVoltage As Double, EndVoltage As Double, StepSize As Double
    Dim Patt_String As String, Shmoo_Result As String, Shmoo_result_PF As String
    Dim Pat As Variant
    Dim PinName As Variant
    Dim StepName As Variant
    Dim RangeFrom As Double, RangeTo As Double, RangeStepSize As Double, RangeSteps As Long
    Dim allPowerPins As String
    Dim PowerPinCnt As Long, PowerPinAry() As String
    Dim Vcc_min As String, Vcc_max As String
    Dim patt_ary() As String, pat_count As Long, p As Variant
    Dim Pin_Ary() As String, p_cnt As Long, Pin_Cnt As Long
    Dim shmoo_pin_string As String
    Dim tmp As String
    Dim Search_String As String
    Dim ch As String
    Dim Group As Boolean
    Dim Label As String
    Dim step_Start As Long
    Dim step_Stop As Long
    Dim Step_x As Long
    Dim Step_NV As Long
    Dim Range_temp As Double
    Dim range_plus As Long
    Dim Shmoo_Pattset As New Pattern
    Dim Shiftin_spec As String
    
    Dim SetupName_New As String, k As Integer
    Dim InstanceName_New As String
    Dim patset As Variant, patset1 As Variant, j As Long
    Dim RangeLow As Double, RangeStart As Double
    Dim Shmoo_hole As String
    Dim RangeCalcType As tlDevCharRangeField

    Dim RangeHigh As Double     '20180515 TER add
    Dim HIO_PinName_Updated As Boolean      '20180515 TER
    
    Dim index As Long
    Dim vbump_value As String
    Dim TestNameInput As String

    
    

    ReportHVCC = True
    ReportLVCC = True
    Shmoo_hole = "NH"
    Patt_String = vbNullString
    
    instancename = theexec.DataManager.instancename     '20180616 TER add
    Call Get_Tname_FromFlowSheet(instancename, HIO_PinName_Updated)      '20180515 TER add
    
    For Each site In theexec.sites
        
        
        OutputString = vbNullString
        lvccf = 0

        '20170125 Modify TestName width show in datalog
        theexec.Datalog.Setup.Shared.ascii.Columns.EnableCustomWidths = True
        If Len(instancename) < 235 Then
            theexec.Datalog.Setup.Shared.ascii.Columns.Parametric.TestName.Width = Len(instancename) + 20
        Else
            theexec.Datalog.Setup.Shared.ascii.Columns.Parametric.TestName.Width = theexec.Datalog.Setup.Shared.ascii.Columns.Parametric.TestName.MaximumWidth
        End If
        theexec.Datalog.ApplySetup
        
        If theexec.DevChar.Setups.item(theexec.DevChar.Setups.ActiveSetupName).Output.SuspendDatalog = False Then    '20180718 add  #16_ELSE_CASE_CHK
            Call theexec.sites(site).IncrementTestNumber
        Else
        End If
        
        TestNum = theexec.sites(site).TestNumber
        
        If theexec.DevChar.Setups.item(theexec.DevChar.Setups.ActiveSetupName).Output.SuspendDatalog = True Then    '20180718 add   #16_ELSE_CASE_CHK
            Call theexec.sites(site).IncrementTestNumber
        Else
        End If
        
        '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////  Read XI0 Nwire Setup value
        'xio_spec = "XI0_Freq_VAR"
         Shiftin_spec = "ShiftIn_Freq_VAR"
        

        v_Shiftin = theexec.Specs.AC(Shiftin_spec).CurrentValue
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

         '20180716 Auto parsing FRC info
        Dim nWire_port_ary() As String
        Dim nwp As Variant
        Dim port_pa As String, ac_spec_pa As String, pin_pa As String, global_spec_pa As String
        Dim FRC_Name As String, FRC_Value As Double, All_FRC_Status As String
        All_FRC_Status = vbNullString
        nWire_port_ary = Split(nWire_Ports_GLB, ",")
        For Each nwp In nWire_port_ary
            Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
            If LCase(glb_TesterType) = "jaguar" Then
                If thehdw.Protocol.ports(port_pa).Enabled = True Then
                    FRC_Name = Replace(UCase(ac_spec_pa), "_FREQ_VAR", vbNullString)
                    'If UCase(Outer_ParameterName) = UCase(ac_spec_pa) Then
                    '    FRC_Value = Outer_RangeFrom + j_Outer * Outer_Step_Index * Outer_RangeStepSize
                    'Else
                        FRC_Value = theexec.Specs.AC(ac_spec_pa).CurrentValue
                    'End If
                End If
            Else
                FRC_Name = UCase(pin_pa)
                ' Fixed Y domain is XI0 print problem
                'If UCase(Outer_ParameterName) = UCase(ac_spec_pa) Then
                '    FRC_Value = Outer_RangeFrom + j_Outer * Outer_Step_Index * TheExec.DevChar.Setups(SetupName).Shmoo.Axes(CharShmooAxis_Outer).Parameter.range.StepSize
                'Else
                    If thehdw.Digital.Pins(pin_pa).FreeRunningClock.IsRunning Then
                        FRC_Value = thehdw.Digital.Pins(pin_pa).FreeRunningClock.Frequency
                    Else
                    FRC_Value = theexec.Specs.AC(ac_spec_pa).CurrentValue
                    End If
                'End If
            End If
                    If All_FRC_Status = "" Then
                        All_FRC_Status = FRC_Name & "=" & FRC_Value
                    Else
                        All_FRC_Status = All_FRC_Status & ";" & FRC_Name & "=" & FRC_Value
                    End If
        Next nwp
        If FRC_Name = "" Then ' Default use XI0, if no input of FRC info    #16_ELSE_CASE_CHK
            FRC_Name = XI0
            FRC_Value = TheExec.Specs.AC(XI0_Diff & "_Freq_VAR").CurrentValue
            All_FRC_Status = FRC_Name & "=" & FRC_Value
        End If
        
        SetupName = theexec.DevChar.Setups.ActiveSetupName
        
        Shmoo_Pattset.value = Shmoo_Pattern
        Patt_String = PatSetToPat(Shmoo_Pattset)
        
        Search_String = vbNullString
        With theexec.DevChar
            StepName = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).StepName
            RangeFrom = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.from
            RangeTo = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.to
            RangeSteps = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.Steps + 1
            RangeStepSize = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.StepSize
            RangeCalcType = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.CalculatedField
            
            
            '20170210 Added to check Endpoint
            Dim RangeTo_New As Double
            Dim RangeFrom_New As Double
            

            If RangeFrom < RangeTo Then ' always start from lower Value
                step_Start = 0
                step_Stop = RangeSteps - 1
                Step_x = 1
                RangeLow = RangeFrom
'                range_plus = -1
            '20170210 Added to check Endpoint
                RangeFrom_New = RangeFrom
                RangeTo_New = RangeFrom + (RangeSteps - 1) * RangeStepSize
            Else
                step_Start = RangeSteps - 1
                step_Stop = 0
                Step_x = -1

                '20170210 Added to check Endpoint
                RangeLow = Format((RangeFrom - (RangeSteps - 1) * RangeStepSize), "0.000#########")
                
                RangeFrom_New = RangeFrom
                RangeTo_New = RangeFrom - (RangeSteps - 1) * RangeStepSize
            End If
            If Abs(RangeTo) < 0.000000000001 Then RangeTo = 0   '#16_ELSE_CASE_CHK
            If Abs(RangeFrom) < 0.000000000001 Then RangeFrom = 0   '#16_ELSE_CASE_CHK
            '20170210 Added to check Endpoint
            gen_search_string SetupName, Search_String, tlDevCharShmooAxis_X, RangeFrom_New, RangeTo_New, RangeStepSize, RangeSteps
            
            shmoo_pin_string = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).ApplyTo.Pins
            If theexec.enableWord("ShmooMakePseudoData") = True Then Call ShmooMakePseudoData(SetupName, step_Start, step_Stop, Step_x)
           Call CreateShmooResultString(Shmoo_Result, Shmoo_result_PF, SetupName, step_Start, step_Stop, Step_x, site)
            
'20161229 Roy Modified,Prevent Step_NV out of range
            Step_NV = -1
			'20240531shmoo hole fail log
            Call Decide_LVCC_HVCC(Vcc_min, Vcc_max, Shmoo_hole, Step_NV, RangeLow, RangeStepSize, Shmoo_result_PF, SetupName, step_Start, step_Stop, Step_x, site)
			'Call Decide_LVCC_HVCC(Vcc_min, Vcc_max, Shmoo_hole, Step_NV, RangeLow, RangeStepSize, Shmoo_result_PF, SetupName, step_Start, step_Stop, Step_x)
            
            
        End With
        If InStr(TheExec.DataManager.instancename, "_NV") Then TestVoltage = "NV"   '#16_ELSE_CASE_CHK
        If InStr(TheExec.DataManager.instancename, "_HV") Then TestVoltage = "HV"   '#16_ELSE_CASE_CHK
        If InStr(TheExec.DataManager.instancename, "_LV") Then TestVoltage = "LV"   '#16_ELSE_CASE_CHK
        
'    [Char,N99G19-1,16,7,V,0,XI0=24000000,CpuBira_P0001_IN02_BIR_SI_PL00_CL16_BIR_59N_SI_PP_NV,CPU_BIST_CPU_Domain_CPU_SRAM_Domain_P1_Full_Range,1069,
'.\pattern\CpuMbist\PP_FIJA0_C_IN00_XX_CLXX_XXX_XXX_XXX_P0001_1308131609_SI_mod.pat,.\pattern\CpuMbist\PP_FIJA0_C_IN02_BI_CLXX_BIR_JTG_XXX_ALLFV_1306250000_SI.pat,.\pattern\CpuMbist\PP_FIJA0_C_PL00_BI_CL16_BIR_JTG_59N_ALLFV_1306250000_SI.pat,
'NV,VDD_FIXED=0.528:1.404:0.005,VDD_VAR_SOC_VAR=0.500:1.330:0.005,
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++,NH,0.5,1.260]
        If Shmoo_header = "" Then Shmoo_header = "Char" '#16_ELSE_CASE_CHK
        
        If (LCase(theexec.CurrentJob) Like "*cp*") Then
                OutputString = OutputString & "[" & Shmoo_header & "," & HramLotId(site) & "-" & CStr(HramWaferId(site)) & "," & CStr(XCoord(site)) & "," & CStr(YCoord(site))
        Else
                OutputString = OutputString & "[" & Shmoo_header & "," & HramLotId(site) & "-" & CStr(HramWaferId(site)) & "," & CStr(HramXCoord(site)) & "," & CStr(HramYCoord(site))
        End If
        
        'OutputString = OutputString & "[" & Shmoo_header & "," & HramLotId(Site) & "," & CStr(Xcoord(Site)) & "," & CStr(Ycoord(Site))
        SetupName_New = SetupName
        
        'Shmoo_header
        Dim VIL_Flag As Boolean
        VIL_Flag = False
        ShmooPowerName = ShmooPowerName

        
       
        OutputString = OutputString & ",V," & site & "," & All_FRC_Status & ","    '20180716 Auto parsing FRC info
        OutputString = OutputString & theexec.DataManager.instancename & ShmooPowerName & "," & SetupName_New & "," & CStr(TestNum) & ","


        OutputString = OutputString & Patt_String & ","
        OutputString = OutputString & TestVoltage & ","

        If argv(0) <> Empty Then    '#16_ELSE_CASE_CHK
            theexec.DataManager.DecomposePinList argv(0), Pin_Ary, Pin_Cnt
            PinName = argv(0) 'setup voltage
        Else
        End If
        
        
        If Vbump_for_Interpose = True Then
            Dim PL_DC_conditions_str As String
            PL_DC_conditions_str = Replace(PL_DC_conditions_GLB, ":V:", "=")
            PL_DC_conditions_str = Replace(PL_DC_conditions_str, ";", ",")
            OutputString = OutputString & PL_DC_conditions_str
        
        Else
            For j = 0 To Pin_Cnt - 1
                PinName = Pin_Ary(j)
                If theexec.DataManager.ChannelType(PinName) <> "N/C" Then   '#16_ELSE_CASE_CHK
                    If j = 0 Then
                        OutputString = OutputString & PinName & "=" & Format(thehdw.DCVS.Pins(PinName).Voltage.Main.value, "0.000")
                    Else
                        OutputString = OutputString & "," & PinName & "=" & Format(thehdw.DCVS.Pins(PinName).Voltage.Main.value, "0.000")
                    End If
                End If
            Next j
        End If
        
        For i = 1 To argc - 1
            If UCase(argv(i)) = "VIL" Or UCase(argv(i)) = "VOL" Then
                VIL_Flag = True
            Else
                theexec.DataManager.DecomposePinList argv(i), Pin_Ary, Pin_Cnt
                For j = 0 To Pin_Cnt - 1
                    PinName = LCase(Pin_Ary(j))
                    If theexec.DataManager.ChannelType(PinName) <> "N/C" Then
                        If Vbump_for_Interpose = True Then
                            index = InStr(LCase(PL_DC_conditions_str), PinName & "=")
                            vbump_value = mid(LCase(PL_DC_conditions_str), index + Len(PinName) + 1, 5)
                            OutputString = OutputString & "," & PinName & "=" & vbump_value
                        Else
                            OutputString = OutputString & "," & PinName & "=" & Format(thehdw.DCVS.Pins(PinName).Voltage.Main.value, "0.000")
                        End If
                    End If
                Next j
            End If
        Next
        PL_DC_conditions_str = vbNullString
        
        OutputString = OutputString & ","
        Search_String = mid(Search_String, 1, Len(Search_String) - 1) 'take out last ","
        Search_String = Replace(Search_String, "X@", vbNullString)
        OutputString = OutputString & Search_String
        OutputString = OutputString & ","
        
'///////////////////////////////////////////////////////// check hole
        If Vcc_max = "5555" And Vcc_min <> "-5555" Then Shmoo_hole = "HH"
        If Vcc_max <> "5555" And Vcc_min = "-5555" Then Shmoo_hole = "LH"
        If Vcc_max = "5555" And Vcc_min = "-5555" Then Shmoo_hole = "BH"
        If Vcc_max <> "5555" And Vcc_min <> "-5555" Then Shmoo_hole = "NH"
'/////////////////////////////////////////////////////////
        OutputString = OutputString & Shmoo_Result & ","
        
        If VIL_Flag = True Then
            OutputString = OutputString & Shmoo_hole & "," & Vcc_max & "," & Vcc_min & "]"
        Else
            OutputString = OutputString & Shmoo_hole & "," & Vcc_min & "," & Vcc_max & "]"
        End If
        
        ''get current Timing set sheet''
        Dim Context As String: Context = vbNullString
        Dim TimeSet_Str As String: TimeSet_Str = vbNullString
        Context = theexec.Contexts.ActiveSelection
        TimeSet_Str = theexec.Contexts(Context).Sheets.Timesets
        
'        Debug.Print outputString
        theexec.Datalog.WriteComment OutputString
        theexec.Datalog.WriteComment "[Force_condition_during_shmoo:" & Charz_Force_Power_condition & "]"
        theexec.Datalog.WriteComment "[Activity_Timing_Sheet:" & UCase(TimeSet_Str) & "," & "Shiftin_Freq=" & CStr(v_Shiftin) & "]"
        
        If Vcc_min = "N/A" Then
            Shmoo_Vcc_Min(site) = -0.1
        Else
            If Vcc_min = "" Then Vcc_min = 0
            Shmoo_Vcc_Min(site) = Vcc_min
        End If
        
        If Vcc_max = "N/A" Then
            If RangeFrom > RangeTo Then
                Shmoo_Vcc_Max(site) = RangeFrom + 0.1
            Else
                Shmoo_Vcc_Max(site) = RangeTo + 0.1
            End If
        Else
            If Vcc_max = "" Then Vcc_max = 0
            Shmoo_Vcc_Max(site) = Vcc_max
        End If
        
        '**************************************************AI**********************************************
        
        If theexec.enableWord("AI_Fail_Log") = True And Voltage_fail_point <> 0 Then
        Dim Setpower As String
        Dim x As Integer
        Dim y As Integer
        'Voltage_fail_point
        'Voltage_fail_collect
        Setpower = vbNullString
        For x = 0 To Voltage_fail_point - 1
                Setpower = vbNullString
                Setpower = Replace(shmoo_pin_string, ",", "+") & ",VDD," & Voltage_fail_collect(x)
                theexec.Datalog.WriteComment Setpower
                Call SetForceCondition(Setpower)
                thehdw.Patterns(Patt_String).test pfAlways, 0
                thehdw.Digital.Patgen.HaltWait
                If thehdw.Digital.Patgen.PatternBurstPassed(site) = False Then y = y + 1    '#16_ELSE_CASE_CHK
                If y = Voltage_fail_point_request Then GoTo contine1
            Next x
contine1:
    End If
        '*************************************************AI**********************************************
    
    Next site
    
        
    If Vcc_min = "N/A" Then Vcc_min = 9999  '#16_ELSE_CASE_CHK
        'job char flag
    If UCase(currentJobName) Like "*CP*" Or UCase(currentJobName) Like "*FT*" Then

        Dim TestNameLVCC As String, TestNameHVCC As String
        Dim TestName As String
        Dim GPIO_Char_Shmoo_Pin As String
        Dim Shmoo_setup_name As String
        Shmoo_setup_name = theexec.DevChar.Setups.ActiveSetupName
        GPIO_Char_Shmoo_Pin = theexec.DevChar.Setups(Shmoo_setup_name).Shmoo.Axes(tlDevCharShmooAxis_X).ApplyTo.Pins

            'ZYCHOUA 211203, C651 request
            theexec.Datalog.Setup.Shared.ascii.Columns.EnableCustomWidths = True
            If (Len(instancename) + Len(shmoo_pin_string) + 10) < 235 Then
                theexec.Datalog.Setup.Shared.ascii.Columns.Parametric.TestName.Width = Len(instancename) + Len(shmoo_pin_string) + 10
            Else
                theexec.Datalog.Setup.Shared.ascii.Columns.Parametric.TestName.Width = theexec.Datalog.Setup.Shared.ascii.Columns.Parametric.TestName.MaximumWidth
            End If
            theexec.Datalog.ApplySetup
            shmoo_pin_string = Replace(shmoo_pin_string, ",", "-")
            'ZYCHOUA 211203, C651 request
        
        
            If UCase(instancename) Like "*_CHAR_CP*" Then
                TestName = Replace(instancename, "_Char_CP", vbNullString)
            ElseIf UCase(instancename) Like "*__H*" Then
                TestName = Replace(instancename, "__H", vbNullString)
            ElseIf UCase(instancename) Like "*__L*" Then
                TestName = Replace(instancename, "__L", vbNullString)
            Else
                TestName = instancename
            End If
            
            If UCase(TestName) Like "*_NV*" Then
                TestName = Replace(TestName, "_CZ_NV", "_")
            ElseIf UCase(TestName) Like "*_LV*" Then
                TestName = Replace(TestName, "_CZ_LV", "_")
            ElseIf UCase(TestName) Like "*_HV*" Then
                TestName = Replace(TestName, "_CZ_HV", "_")
            End If
            
            Dim HVCC_DFTLH As String
            Dim LVCC_DFTLH As String
         '20160925 Multi_USL/LSL
            Dim HVCC_MCLH As String
            Dim LVCC_MCLH As String
         If UCase(TestName) Like "*DFTLH*" Then
                HVCC_DFTLH = Replace(TestName, "DFTLH", "DFTH")
                LVCC_DFTLH = Replace(TestName, "DFTLH", "DFTL")
         End If
         '20160925 Multi_USL/LSL
         If UCase(TestName) Like "*MCLH*" Then
                HVCC_MCLH = Replace(TestName, "MCLH", "MCH")
                LVCC_MCLH = Replace(TestName, "MCLH", "MCL")
         End If
         
            
'        End If


        theexec.Datalog.Setup.DatalogSetup.DisableInstanceNameInPTR = True
        theexec.Datalog.Setup.DatalogSetup.DisablePinNameInPTR = True
        theexec.Datalog.ApplySetup
'--------------------------------------------------------------------------------------------

        
        Dim HF_HVCC_TESTNAME, HF_LVCC_TESTNAME As String
        Dim hi_limit, Low_limit As Double
        
        If RangeFrom < RangeTo Then
            hi_limit = RangeTo: Low_limit = RangeFrom
        Else
            hi_limit = RangeFrom: Low_limit = RangeTo
        End If
        
        If (CHAR_USL_HVCC = 9999) Then CHAR_USL_HVCC = hi_limit '#16_ELSE_CASE_CHK
        If (CHAR_LSL_HVCC = 9999) Then CHAR_LSL_HVCC = Low_limit    '#16_ELSE_CASE_CHK
        If (CHAR_USL_LVCC = 9999) Then CHAR_USL_LVCC = hi_limit '#16_ELSE_CASE_CHK
        If (CHAR_LSL_LVCC = 9999) Then CHAR_LSL_LVCC = Low_limit    '#16_ELSE_CASE_CHK

        If (CHAR_USL_HVCC < CHAR_LSL_HVCC) And isDebugMode = True Then theexec.AddOutput theexec.DataManager.instancename & " : Limit Error ! " & "HVCC_USL=" & CStr(CHAR_USL_HVCC) & ",HVCC_LSL=" & CStr(CHAR_LSL_HVCC): CHAR_USL_HVCC = hi_limit: CHAR_LSL_HVCC = Low_limit   '#16_ELSE_CASE_CHK
        If (CHAR_USL_LVCC < CHAR_LSL_LVCC) And isDebugMode = True Then theexec.AddOutput theexec.DataManager.instancename & " : Limit Error ! " & "LVCC_USL=" & CStr(CHAR_USL_LVCC) & ",LVCC_LSL=" & CStr(CHAR_LSL_LVCC): CHAR_USL_LVCC = hi_limit: CHAR_LSL_LVCC = Low_limit   '#16_ELSE_CASE_CHK

        If UCase(instancename) Like "HFHL*" Or UCase(instancename) Like "HFLH*" Then    '#16_ELSE_CASE_CHK
        
            If UCase(instanceName) Like "HFHL*" Then
                HF_HVCC_TESTNAME = Replace(TestName, "HFHL", "HFH") 
                HF_LVCC_TESTNAME = Replace(TestName, "HFHL", "HFL")	
            Else
                HF_HVCC_TESTNAME = Replace(TestName, "HFLH", "HFH") 
                HF_LVCC_TESTNAME = Replace(TestName, "HFLH", "HFL") 
            End If

            
            theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, ForceResults:=tlForceNone, Tname:=HF_HVCC_TESTNAME & " " & shmoo_pin_string & " <> " & HF_HVCC_TESTNAME, lowVal:=Low_limit, hiVal:=hi_limit
            '20170120 chnage print format
            theexec.Datalog.WriteComment "[Force_condition_during_shmoo_HW:" & ReadHWPowerValue_GLB & "]"
        
            theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, ForceResults:=tlForceNone, Tname:=HF_LVCC_TESTNAME & " " & shmoo_pin_string & " <> " & HF_HVCC_TESTNAME, lowVal:=Low_limit, hiVal:=hi_limit
            theexec.Datalog.WriteComment "[Force_condition_during_shmoo_HW:" & ReadHWPowerValue_GLB & "]"
            
            theexec.Datalog.Setup.DatalogSetup.DisableInstanceNameInPTR = False
            theexec.Datalog.Setup.DatalogSetup.DisablePinNameInPTR = False
            theexec.Datalog.ApplySetup
            Exit Function
        End If
        
        If UCase(instancename) Like "HFH*" Then '#16_ELSE_CASE_CHK
        
            theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, ForceResults:=tlForceNone, Tname:=TestName & " " & shmoo_pin_string & " <> " & TestName, lowVal:=CHAR_LSL_HVCC, hiVal:=CHAR_USL_HVCC
            theexec.Datalog.WriteComment "[Force_condition_during_shmoo_HW:" & ReadHWPowerValue_GLB & "]"
            
            theexec.Datalog.Setup.DatalogSetup.DisableInstanceNameInPTR = False
            theexec.Datalog.Setup.DatalogSetup.DisablePinNameInPTR = False
            theexec.Datalog.ApplySetup
            Exit Function
        End If
        
        If UCase(instancename) Like "HFL*" Then '#16_ELSE_CASE_CHK
        
            theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, ForceResults:=tlForceNone, Tname:=TestName & " " & shmoo_pin_string & " <> " & TestName, lowVal:=CHAR_LSL_LVCC, hiVal:=CHAR_USL_LVCC
            theexec.Datalog.WriteComment "[Force_condition_during_shmoo_HW:" & ReadHWPowerValue_GLB & "]"
            
            theexec.Datalog.Setup.DatalogSetup.DisableInstanceNameInPTR = False
            theexec.Datalog.Setup.DatalogSetup.DisablePinNameInPTR = False
            theexec.Datalog.ApplySetup
            Exit Function
        End If
        
        '20210309 @CW for HIP AMPLP5 LVCC/HVCC limit
        If UCase(theexec.DataManager.instancename) Like "LP5CZ_*" Then  '#16_ELSE_CASE_CHK
            theexec.Datalog.Setup.DatalogSetup.DisableInstanceNameInPTR = False
            theexec.Datalog.Setup.DatalogSetup.DisablePinNameInPTR = False
            theexec.Datalog.Setup.DatalogSetup.DisableChannelNumberInPTR = True
            theexec.Datalog.Setup.DatalogSetup.PTR_InstanceNameIsTINameOnly = True
            theexec.Datalog.ApplySetup
        
            If instancename Like "*HIO*" And instancename Like "*VIL_*" Then    '#16_ELSE_CASE_CHK
                theexec.Datalog.WriteComment ("Test name :" & instancename & "_" & SetupName & "_Vmax")
                'TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, lowVal:=CHAR_LSL_HVCC, hiVal:=CHAR_USL_HVCC, ForceResults:=tlForceNone, Tname:=TestName & " " & shmoo_pin_string '& " <> " & TestName
                TestName = mid(TestName, 1, Len(TestName) - 1)
                theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, lowVal:=CHAR_LSL_HVCC, hiVal:=CHAR_USL_HVCC, ForceResults:=tlForceNone, PinName:="X", Tname:=TestName  '& " <> " & TestName
                Exit Function
            
            ElseIf instancename Like "*HIO*" And instancename Like "*VIH_*" Then
                theexec.Datalog.WriteComment ("Test name :" & instancename & "_" & SetupName & "_Vmin")
                'TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, lowVal:=CHAR_LSL_LVCC, hiVal:=CHAR_USL_LVCC, ForceResults:=tlForceNone, Tname:=TestName & " " & shmoo_pin_string ' & " <> " & TestName
                TestName = mid(TestName, 1, Len(TestName) - 1)
                theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, lowVal:=CHAR_LSL_LVCC, hiVal:=CHAR_USL_LVCC, ForceResults:=tlForceNone, PinName:="X", Tname:=TestName ' & " <> " & TestName
                Exit Function
            End If
        End If
        
        If UCase(instancename) Like "HIO*" And UCase(instancename) Like "*VCM*" And UCase(instancename) Like "*USBPICO*" Then   '#16_ELSE_CASE_CHK
        
            HF_HVCC_TESTNAME = TestName
            
            theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, ForceResults:=tlForceNone, Tname:=HF_HVCC_TESTNAME & "   " & shmoo_pin_string & " <> " & TestName, lowVal:=CHAR_LSL_HVCC, hiVal:=CHAR_USL_HVCC

            theexec.Datalog.Setup.DatalogSetup.DisableInstanceNameInPTR = False
            theexec.Datalog.Setup.DatalogSetup.DisablePinNameInPTR = False
            theexec.Datalog.ApplySetup
            Exit Function
        End If
        
        
        
        If UCase(instancename) Like "*DIFF*" And UCase(instancename) Like "HIO*" Then   '#16_ELSE_CASE_CHK
            If (ReportHVCC And ReportLVCC) Then '#16_ELSE_CASE_CHK
                theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, ForceResults:=tlForceNone, Tname:=TestName & "_H_" & " " & shmoo_pin_string & " <> " & TestName, lowVal:=CHAR_LSL_HVCC, hiVal:=CHAR_USL_HVCC
                'TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, ForceResults:=tlForceNone, Tname:=testName & "_L_" & " & shmoo_pin_string & " <> " & testName, lowVal:=CHAR_LSL_LVCC, hiVal:=CHAR_USL_LVCC"
            ElseIf (ReportLVCC) Then
                theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, ForceResults:=tlForceNone, Tname:=TestName & " " & shmoo_pin_string & " <> " & TestName, lowVal:=CHAR_LSL_LVCC, hiVal:=CHAR_USL_LVCC
            ElseIf (ReportHVCC) Then
                theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, ForceResults:=tlForceNone, Tname:=TestName & " " & shmoo_pin_string & " <> " & TestName, lowVal:=CHAR_LSL_HVCC, hiVal:=CHAR_USL_HVCC
            End If

            theexec.Datalog.Setup.DatalogSetup.DisableInstanceNameInPTR = False
            theexec.Datalog.Setup.DatalogSetup.DisablePinNameInPTR = False
            theexec.Datalog.ApplySetup
            Exit Function
        End If
        
        If UCase(instancename) Like "*VID*" Or UCase(instancename) Like "*VICM*" Then   '#16_ELSE_CASE_CHK
            'Cyprus USB2 20170823
            theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, ForceResults:=tlForceNone, Tname:=TestName & " " & shmoo_pin_string & " <> " & TestName, lowVal:=CHAR_LSL_HVCC, hiVal:=CHAR_USL_HVCC
            theexec.Datalog.Setup.DatalogSetup.DisableInstanceNameInPTR = False
            theexec.Datalog.Setup.DatalogSetup.DisablePinNameInPTR = False
            theexec.Datalog.ApplySetup
            Exit Function
        End If
        
        If instancename Like "*HAC*" And SetupName Like "*VIL*" Then    '#16_ELSE_CASE_CHK
            theexec.Datalog.WriteComment ("Test name :" & instancename & SetupName & "_Vmax")
            theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, lowVal:=CHAR_LSL_LVCC, hiVal:=CHAR_USL_LVCC, ForceResults:=tlForceNone, Tname:=TestName & " " & shmoo_pin_string '& " <> " & TestName
            theexec.Datalog.WriteComment "[Force_condition_during_shmoo_HW:" & ReadHWPowerValue_GLB & "]"
        ElseIf instancename Like "*HAC*" And SetupName Like "*VIH*" Then
                theexec.Datalog.WriteComment ("Test name :" & instancename & SetupName & "_Vmin")
                theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, lowVal:=CHAR_LSL_HVCC, hiVal:=CHAR_USL_HVCC, ForceResults:=tlForceNone, Tname:=TestName & " " & shmoo_pin_string '& " <> " & TestName
                theexec.Datalog.WriteComment "[Force_condition_during_shmoo_HW:" & ReadHWPowerValue_GLB & "]"
        
        ElseIf instancename Like "DFTLH*" Then
                theexec.Datalog.WriteComment ("Test name :" & instancename & "_" & SetupName & "_Vmax")
                theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, lowVal:=CHAR_LSL_HVCC, hiVal:=CHAR_USL_HVCC, ForceResults:=tlForceNone, Tname:=HVCC_DFTLH & " " & shmoo_pin_string ' & " <> " & HVCC_DFTLH
                theexec.Datalog.WriteComment "[Force_condition_during_shmoo_HW:" & ReadHWPowerValue_GLB & "]"
                theexec.Datalog.WriteComment ("Test name :" & instancename & "_" & SetupName & "_Vmin")
                theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, lowVal:=CHAR_LSL_LVCC, hiVal:=CHAR_USL_LVCC, ForceResults:=tlForceNone, Tname:=LVCC_DFTLH & " " & shmoo_pin_string ' & " <> " & LVCC_DFTLH
                theexec.Datalog.WriteComment "[Force_condition_during_shmoo_HW:" & ReadHWPowerValue_GLB & "]"
        
        ElseIf instancename Like "DFTL_*" Or instancename Like "MCL_*" Then
                theexec.Datalog.WriteComment ("Test name :" & instancename & "_" & SetupName & "_Vmin")
                theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, lowVal:=CHAR_LSL_LVCC, hiVal:=CHAR_USL_LVCC, ForceResults:=tlForceNone, Tname:=TestName & " " & shmoo_pin_string ' & " <> " & TestName
                theexec.Datalog.WriteComment "[Force_condition_during_shmoo_HW:" & ReadHWPowerValue_GLB & "]"
        ElseIf instancename Like "DFTH_*" Or instancename Like "MCH_*" Then
                theexec.Datalog.WriteComment ("Test name :" & instancename & "_" & SetupName & "_Vmax")
                theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, lowVal:=CHAR_LSL_HVCC, hiVal:=CHAR_USL_HVCC, ForceResults:=tlForceNone, Tname:=TestName & " " & shmoo_pin_string ' & " <> " & TestName
                theexec.Datalog.WriteComment "[Force_condition_during_shmoo_HW:" & ReadHWPowerValue_GLB & "]"
        
        
        
        '20160925 Multi_USL/LSL
        ElseIf instancename Like "MCLH*" Then
                theexec.Datalog.WriteComment ("Test name :" & instancename & "_" & SetupName & "_Vmax")
                theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, lowVal:=CHAR_LSL_HVCC, hiVal:=CHAR_USL_HVCC, ForceResults:=tlForceNone, Tname:=HVCC_MCLH & " " & shmoo_pin_string '& " <> " & HVCC_MCLH
                theexec.Datalog.WriteComment "[Force_condition_during_shmoo_HW:" & ReadHWPowerValue_GLB & "]"
                theexec.Datalog.WriteComment ("Test name :" & instancename & "_" & SetupName & "_Vmin")
                theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, lowVal:=CHAR_LSL_LVCC, hiVal:=CHAR_USL_LVCC, ForceResults:=tlForceNone, Tname:=LVCC_MCLH & " " & shmoo_pin_string '& " <> " & LVCC_MCLH
                theexec.Datalog.WriteComment "[Force_condition_during_shmoo_HW:" & ReadHWPowerValue_GLB & "]"
        Else
        End If
        
        If UCase(theexec.DataManager.instancename) Like "*ALLPINSGPIO1_X_VI*" Then
            
            If instancename Like "*HIO*" And instancename Like "*_VIL_*" Then   '#16_ELSE_CASE_CHK
                    theexec.Datalog.WriteComment ("Test name :" & instancename & "_" & SetupName & "_Vmax")
                    theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, lowVal:=CHAR_LSL_HVCC, hiVal:=CHAR_USL_HVCC, ForceResults:=tlForceNone, Tname:=TestName & " " & GPIO_Char_Shmoo_Pin ' & " <> " & TestName
            ElseIf instancename Like "*HIO*" And instancename Like "*_VIH_*" Then
                    theexec.Datalog.WriteComment ("Test name :" & instancename & "_" & SetupName & "_Vmin")
                    theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, PinName:=GPIO_Char_Shmoo_Pin, lowVal:=CHAR_LSL_LVCC, hiVal:=CHAR_USL_LVCC, ForceResults:=tlForceNone, Tname:=TestName & " " & GPIO_Char_Shmoo_Pin '& " <> " & TestName
            Else
            End If
            
        ElseIf UCase(theexec.DataManager.instancename) Like "*AMP_*" Then
            
            If instancename Like "*HIO_VIL*" Then   '#16_ELSE_CASE_CHK
                    theexec.Datalog.WriteComment ("Test name :" & instancename & "_" & SetupName & "_Vmax")
                    theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, lowVal:=CHAR_LSL_HVCC, hiVal:=CHAR_USL_HVCC, ForceResults:=tlForceNone, Tname:=TestName & " " & shmoo_pin_string '& " <> " & TestName
            ElseIf instancename Like "*HIO_VIH*" Then
                    theexec.Datalog.WriteComment ("Test name :" & instancename & "_" & SetupName & "_Vmin")
                    theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, lowVal:=CHAR_LSL_LVCC, hiVal:=CHAR_USL_LVCC, ForceResults:=tlForceNone, Tname:=TestName & " " & shmoo_pin_string '& " <> " & TestName
            ElseIf instancename Like "*AMP*_*VIL*" Then
                    theexec.Datalog.WriteComment ("Test name :" & instancename & "_" & SetupName & "_Vmax")
                    theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, lowVal:=CHAR_LSL_HVCC, hiVal:=CHAR_USL_HVCC, ForceResults:=tlForceNone, Tname:=TestName & " " & shmoo_pin_string '& " <> " & TestName
            ElseIf instancename Like "*AMP*_*VIH*" Then
                    theexec.Datalog.WriteComment ("Test name :" & instancename & "_" & SetupName & "_Vmin")
                    theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, lowVal:=CHAR_LSL_LVCC, hiVal:=CHAR_USL_LVCC, ForceResults:=tlForceNone, Tname:=TestName & " " & shmoo_pin_string '& " <> " & TestName
            Else
            End If
        Else
            If instancename Like "*HIO*" And instancename Like "*VIL_*" Then    '#16_ELSE_CASE_CHK
                    theexec.Datalog.WriteComment ("Test name :" & instancename & "_" & SetupName & "_Vmax")
                    theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, lowVal:=CHAR_LSL_HVCC, hiVal:=CHAR_USL_HVCC, ForceResults:=tlForceNone, Tname:=TestName & " " & shmoo_pin_string '& " <> " & TestName
            ElseIf instancename Like "*HIO*" And instancename Like "*VIH_*" Then
                    theexec.Datalog.WriteComment ("Test name :" & instancename & "_" & SetupName & "_Vmin")
                    theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, lowVal:=CHAR_LSL_LVCC, hiVal:=CHAR_USL_LVCC, ForceResults:=tlForceNone, Tname:=TestName & " " & shmoo_pin_string ' & " <> " & TestName
            Else
            End If
        End If
        
'--------------------------------------------------------------------------------------------
    
        
    Else
        If RangeFrom < RangeTo Then
            theexec.Datalog.WriteComment ("Test name :" & instancename & "_" & SetupName & "_Vmin")
            theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, lowVal:=RangeFrom, hiVal:=RangeTo, ForceResults:=tlForceNone, Tname:=instancename & "_" & SetupName & "_Vmin"
            theexec.Datalog.WriteComment ("Test name :" & instancename & "_" & SetupName & "_Vmax")
            theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, lowVal:=RangeFrom, hiVal:=RangeTo, ForceResults:=tlForceNone, Tname:=instancename & "_" & SetupName & "_Vmax"
        Else
            theexec.Datalog.WriteComment ("Test name :" & instancename & "_" & SetupName & "_Vmin")
            theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, lowVal:=RangeTo, hiVal:=RangeFrom, ForceResults:=tlForceNone, Tname:=instancename & "_" & SetupName & "_Vmin"
            theexec.Datalog.WriteComment ("Test name :" & instancename & "_" & SetupName & "_Vmax")
            theexec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, lowVal:=RangeTo, hiVal:=RangeFrom, ForceResults:=tlForceNone, Tname:=instancename & "_" & SetupName & "_Vmax"
        End If
    End If
    
 '''''-------------  CHWUD 11/2 for print LVCC get 3 fail log -----------------------------------------------
    ''' dfc enable
Dim dfc As LVCC_VminBoundary_List
Set dfc = New LVCC_VminBoundary_List
For Each site In theexec.sites


         If theexec.Flow.enableWord("CaptureFaillog") = True Then
         
                If Shmoo_hole = "BH" Or Shmoo_hole = "LH" Then
                
                    If (LCase(theexec.CurrentJob) Like "*cp*") Then
                        FailingBoundaryDatalog_Func_Multi_Power Search_String, lotId, CStr(WaferID), CStr(XCoord(site)), _
                        CStr(YCoord(site)), Shmoo_Pattern, "Shmoo hole", High_to_Low, site
                         
                    Else
                        ''''0605 update to use HRAM data
                        FailingBoundaryDatalog_Func_Multi_Power Search_String, HramLotId(site), CStr(HramWaferId(site)), CStr(HramXCoord(site)), _
                        CStr(HramYCoord(site)), Shmoo_Pattern, "Shmoo hole", High_to_Low, site
                    
                    End If

                End If
        End If
        If theexec.Flow.enableWord("Debug_LVCC") = True Then    '#16_ELSE_CASE_CHK
            If (LCase(theexec.CurrentJob) Like "*cp*") Then
                FailingBoundaryDatalog_Func_Multi_Power Search_String, lotId, CStr(WaferID), CStr(XCoord(site)), _
                CStr(YCoord(site)), Shmoo_Pattern, "Shmoo LVCC", High_to_Low, site
                
            Else
                ''''0605 update to use HRAM data
                FailingBoundaryDatalog_Func_Multi_Power Search_String, HramLotId(site), CStr(HramWaferId(site)), CStr(HramXCoord(site)), _
                CStr(HramYCoord(site)), Shmoo_Pattern, "Shmoo LVCC", High_to_Low, site
            End If
        Else
        End If

        If theexec.Flow.enableWord("Debug_HVCC") = True Then    '#16_ELSE_CASE_CHK
            If (LCase(theexec.CurrentJob) Like "*cp*") Then
                FailingBoundaryDatalog_Func_Multi_Power Search_String, lotId, CStr(WaferID), CStr(XCoord(site)), _
                CStr(YCoord(site)), Shmoo_Pattern, "Shmoo HVCC", Low_to_High, site
            Else
                ''''0605 update to use HRAM data
                FailingBoundaryDatalog_Func_Multi_Power Search_String, HramLotId(site), CStr(HramWaferId(site)), CStr(HramXCoord(site)), _
                CStr(HramYCoord(site)), Shmoo_Pattern, "Shmoo HVCC", Low_to_High, site
            End If
        Else
        End If

         If theexec.Flow.enableWord("Debug_LVCC_VminBoundary") = True Then
            If (dfc.IsEnableDFC Or dfc.IsEnableFAILLOG) And Not dfc.IsItemEnabled Then Exit For ''' if enable dfc and not list in dfc then exit
			
			'20240531shmoo hole fail log
			If Shmoo_Vcc_Min(site) > 0 Or (TheExec.enableWord("Find_shmoo_hole_low_power_twice") = True And Shmoo_Vcc_Min(site) = -5555) Then
			'If Shmoo_Vcc_Min(site) > 0 Then
				 If (LCase(theexec.CurrentJob) Like "*cp*") Then
					'If g_VminBoundary_selsrm = True Then
						FailingDatalog_HLvcc_Boundary_SELSRM Search_String, lotid, CStr(WaferID), CStr(XCoord(site)), _
						CStr(YCoord(site)), Shmoo_Pattern, "Shmoo LVCC", High_to_Low, site, RangeFrom, RangeTo, RangeSteps, RangeStepSize
					'Else
					'    FailingDatalog_Lvcc_Boundary Search_String, lotid, CStr(WaferID), CStr(XCoord(site)), _
					'    CStr(YCoord(site)), Shmoo_Pattern, "Shmoo LVCC", High_to_Low, site, RangeFrom, RangeTo, RangeSteps, RangeStepSize
					'End If
					
				 Else
					'If g_VminBoundary_selsrm = True Then
						FailingDatalog_HLvcc_Boundary_SELSRM Search_String, HramLotId(site), CStr(HramWaferId(site)), CStr(HramXCoord(site)), _
						CStr(HramYCoord(site)), Shmoo_Pattern, "Shmoo LVCC", High_to_Low, site, RangeFrom, RangeTo, RangeSteps, RangeStepSize
					'Else
					'    FailingDatalog_Lvcc_Boundary Search_String, HramLotId(site), CStr(HramWaferId(site)), CStr(HramXCoord(site)), _
					'    CStr(HramYCoord(site)), Shmoo_Pattern, "Shmoo LVCC", High_to_Low, site, RangeFrom, RangeTo, RangeSteps, RangeStepSize
					'End If
				 End If
			  Else
			  End If
		ElseIf (TheExec.enableWord("Find_shmoo_hole_low_power_twice") = True And Shmoo_Vcc_Min(site) = -5555) Then '20231129 only need to enable "Find_shmoo_hole_low_power_twice" enableword to capture Shmoo hole fail log
            If (LCase(theexec.CurrentJob) Like "*cp*") Then
                'If g_VminBoundary_selsrm = True Then
                    FailingDatalog_HLvcc_Boundary_SELSRM Search_String, lotId, CStr(WaferID), CStr(XCoord(site)), _
                    CStr(YCoord(site)), Shmoo_Pattern, "Shmoo LVCC", High_to_Low, site, RangeFrom, RangeTo, RangeSteps, RangeStepSize
                'Else
                '    FailingDatalog_Lvcc_Boundary Search_String, lotid, CStr(WaferID), CStr(XCoord(site)), _
                '    CStr(YCoord(site)), Shmoo_Pattern, "Shmoo LVCC", High_to_Low, site, RangeFrom, RangeTo, RangeSteps, RangeStepSize
                'End If
                
             Else
                'If g_VminBoundary_selsrm = True Then
                    FailingDatalog_HLvcc_Boundary_SELSRM Search_String, HramLotId(site), CStr(HramWaferId(site)), CStr(HramXCoord(site)), _
                    CStr(HramYCoord(site)), Shmoo_Pattern, "Shmoo LVCC", High_to_Low, site, RangeFrom, RangeTo, RangeSteps, RangeStepSize
                'Else
                '    FailingDatalog_Lvcc_Boundary Search_String, HramLotId(site), CStr(HramWaferId(site)), CStr(HramXCoord(site)), _
                '    CStr(HramYCoord(site)), Shmoo_Pattern, "Shmoo LVCC", High_to_Low, site, RangeFrom, RangeTo, RangeSteps, RangeStepSize
                'End If
             End If
        End If
        
        If theexec.Flow.enableWord("Debug_HVCC_VminBoundary") = True Then
        Dim Vcc_max_Limit As Double
        If RangeFrom > RangeTo Then
            Vcc_max_Limit = RangeFrom
        Else
            Vcc_max_Limit = RangeTo
        End If
      If Shmoo_Vcc_Max(site) <= Vcc_max_Limit Then
         If (LCase(theexec.CurrentJob) Like "*cp*") Then
           'If g_VminBoundary_selsrm = True Then
               FailingDatalog_HLvcc_Boundary_SELSRM Search_String, lotId, CStr(WaferID), CStr(XCoord(site)), _
               CStr(YCoord(site)), Shmoo_Pattern, "Shmoo HVCC", High_to_Low, site, RangeFrom, RangeTo, RangeSteps, RangeStepSize
           'Else
           '    FailingDatalog_Hvcc_Boundary Search_String, lotid, CStr(WaferID), CStr(XCoord(site)), _
           '    CStr(YCoord(site)), Shmoo_Pattern, "Shmoo HVCC", High_to_Low, site, RangeFrom, RangeTo, RangeSteps, RangeStepSize
           'End If
         Else
           'If g_VminBoundary_selsrm = True Then
               FailingDatalog_HLvcc_Boundary_SELSRM Search_String, HramLotId(site), CStr(HramWaferId(site)), CStr(HramXCoord(site)), _
               CStr(HramYCoord(site)), Shmoo_Pattern, "Shmoo HVCC", High_to_Low, site, RangeFrom, RangeTo, RangeSteps, RangeStepSize
           'Else
           '   FailingDatalog_Hvcc_Boundary Search_String, HramLotId(site), CStr(HramWaferId(site)), CStr(HramXCoord(site)), _
           '    CStr(HramYCoord(site)), Shmoo_Pattern, "Shmoo HVCC", High_to_Low, site, RangeFrom, RangeTo, RangeSteps, RangeStepSize
           'End If
         End If
      End If
    End If
        
Next site
   Set dfc = Nothing
'-----------------------------------------------------------------------

    
    theexec.Datalog.Setup.DatalogSetup.DisableInstanceNameInPTR = False
    theexec.Datalog.Setup.DatalogSetup.DisablePinNameInPTR = False
    '20170125 Modify TestName width show in datalog
    theexec.Datalog.Setup.Shared.ascii.Columns.EnableCustomWidths = True
    theexec.Datalog.Setup.Shared.ascii.Columns.Functional.TestName.Width = 150
    theexec.Datalog.Setup.Shared.ascii.Columns.Functional.Pattern.Width = 80
    theexec.Datalog.Setup.Shared.ascii.Columns.Parametric.TestName.Width = 150
    theexec.Datalog.ApplySetup
    '20170126 Initialize GLlobal power condition
    ReadHWPowerValue_GLB = vbNullString
    Charz_Force_Power_condition = vbNullString
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "ShmooPostStep1D")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function PrintEachPoint_TestName(TestName As String)
On Error GoTo errHandler
     Dim DevSetupName As String
     Dim curr_axis As Variant
     Dim Tracking_Item As Variant
     Dim Tracking_axis_Pin As String
     Dim Tracking_axis_val As Variant
     Dim axis_val As Variant
     Dim axis_pin As String
     Dim axis_pin_Arr() As String
     Dim VarCount As Integer

     DevSetupName = theexec.DevChar.Setups.ActiveSetupName
     
     For Each curr_axis In theexec.DevChar.Setups(DevSetupName).Shmoo.Axes.list
         axis_val = theexec.DevChar.Results(DevSetupName).Shmoo.CurrentPoint.Axes(curr_axis).value ''20190319 update
         If axis_val >= 1000000 Then    '#16_ELSE_CASE_CHK
            axis_val = (axis_val / 1000000) & "M"
         ElseIf axis_val >= 1000 And axis_val < 1000000 Then
            axis_val = (axis_val / 1000) & "K"
         ElseIf axis_val >= 1 And axis_val < 1000 Then
            axis_val = (axis_val * 1000) & "m"
         ElseIf axis_val >= 0.001 And axis_val < 1 Then
            axis_val = (axis_val * 1000) & "m"
         ElseIf axis_val >= 0.000001 And axis_val < 0.001 Then
            axis_val = (axis_val * 1000000) & "u"
         Else
         End If
         axis_pin = Replace((theexec.DevChar.Setups(DevSetupName).Shmoo.Axes((curr_axis)).ApplyTo.Pins), "_", vbNullString)
         If axis_pin = "" Then axis_pin = Replace((theexec.DevChar.Setups(DevSetupName).Shmoo.Axes((curr_axis)).Parameter.name), "_", vbNullString) '#16_ELSE_CASE_CHK
         If InStr(axis_pin, ",") <> 0 Then
            axis_pin_Arr = Split(axis_pin, ",")
            For VarCount = 0 To UBound(axis_pin_Arr)
              TestName = TestName & "_" & axis_pin_Arr(VarCount) & CStr(axis_val)
            Next VarCount
         Else
            TestName = TestName & "_" & axis_pin & CStr(axis_val)
         End If
                 
         If theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Count <> 0 Then ' Tracking case   #16_ELSE_CASE_CHK
             With theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters
                  For Each Tracking_Item In .list
                      Tracking_axis_val = theexec.DevChar.Results(DevSetupName).Shmoo.CurrentPoint.Axes(curr_axis).TrackingParameters(Tracking_Item).value
                      If Tracking_axis_val >= 1000000 Then
                         Tracking_axis_val = (Tracking_axis_val / 1000000) & "M"
                      ElseIf Tracking_axis_val >= 1000 And Tracking_axis_val < 1000000 Then
                         Tracking_axis_val = (Tracking_axis_val / 1000) & "K"
                      ElseIf Tracking_axis_val >= 1 And Tracking_axis_val < 1000 Then
                         Tracking_axis_val = (Tracking_axis_val * 1000) & "m"
                      ElseIf Tracking_axis_val >= 0.001 And Tracking_axis_val < 1 Then
                         Tracking_axis_val = (Tracking_axis_val * 1000) & "m"
                      ElseIf Tracking_axis_val >= 0.000001 And Tracking_axis_val < 0.001 Then
                         Tracking_axis_val = (Tracking_axis_val * 1000000) & "u"
                      End If
                      Tracking_axis_Pin = Replace((theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).ApplyTo.Pins), "_", vbNullString)
                      If Tracking_axis_Pin = "" Then Tracking_axis_Pin = Replace((theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).name), "_", vbNullString)   '#16_ELSE_CASE_CHK
                      TestName = TestName & "_" & Tracking_axis_Pin & CStr(Tracking_axis_val)
                  Next Tracking_Item
             End With
         End If
     Next curr_axis
       
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "PrintEachPoint_TestName")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function GuardBandCondition(Power_Run_Scenario As String, seq As Long)
On Error GoTo errHandler
    Dim ScenarioAry() As String
    Dim GuardBandAry() As String
    Dim GuardBandStr As String
    Dim i As Long
    ScenarioAry = Split(Power_Run_Scenario, "_")
            For i = 1 To UBound(ScenarioAry)
                If InStr(LCase(ScenarioAry(i)), ":") <> 0 And (LCase(ScenarioAry(i - 1)) Like "*pl" & seq Or LCase(ScenarioAry(i - 1)) Like "*pl" & seq) Then
                    GuardBandAry = Split(ScenarioAry(i), ":")
                    GuardBandStr = LCase(GuardBandAry(1))
                    Select Case mid(GuardBandStr, 1, 1)
                    Case "+"
                        CalcSymbol = "+"
                        GuardBandStr = Replace(GuardBandStr, "+", vbNullString)
                        If InStr(GuardBandStr, "mv") <> 0 Then
                            SweepGuardBandVal = Format((CDbl(Replace(GuardBandStr, "mv", vbNullString)) / 1000), "0.000")
                        ElseIf InStr(GuardBandStr, "v") <> 0 Then
                            SweepGuardBandVal = Format(CDbl(Replace(GuardBandStr, "v", vbNullString)), "0.000")
                        Else
                            theexec.Datalog.WriteComment "Please Check the Scenario Formate: " & GuardBandStr
                        End If
                    Case "-"
                        CalcSymbol = "-"
                        GuardBandStr = Replace(GuardBandStr, "-", vbNullString)
                        If InStr(GuardBandStr, "mv") <> 0 Then
                            SweepGuardBandVal = Format((CDbl(Replace(GuardBandStr, "mv", vbNullString)) / 1000), "0.000")
                        ElseIf InStr(GuardBandStr, "v") <> 0 Then
                            SweepGuardBandVal = Format(CDbl(Replace(GuardBandStr, "v", vbNullString)), "0.000")
                        Else
                            theexec.Datalog.WriteComment "Please Check the Scenario Formate: " & GuardBandStr
                        End If
                    Case "*"
                        CalcSymbol = "*"
                        GuardBandStr = Replace(GuardBandStr, "*", vbNullString)
                        SweepGuardBandVal = CDbl(GuardBandStr)
                    Case "/"
                        CalcSymbol = "/"
                        GuardBandStr = Replace(GuardBandStr, "/", vbNullString)
                        SweepGuardBandVal = CDbl(GuardBandStr)
                    Case Else
                        theexec.Datalog.WriteComment "Please Check the Calculate Symbol !!"
                    End Select
                End If
            Next
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "GuardBandCondition")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function CalcPower(DictPowerVal As Double, symbol As String, GuardBand As Double) As Double
On Error GoTo errHandler
    Select Case symbol
    Case "+"
        CalcPower = DictPowerVal + GuardBand
    Case "-"
        CalcPower = DictPowerVal - GuardBand
    Case "*"
        CalcPower = DictPowerVal * GuardBand
    Case "/"
        CalcPower = DictPowerVal / GuardBand
    Case Else
        theexec.Datalog.WriteComment "Calculate Symbol Not Support !!"
    End Select
    
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "CalcPower")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ValueResolution(x_value As Variant, Optional y_value As Variant = vbNullString, Optional z_value As Variant = vbNullString) As Long
On Error GoTo errHandler
'Public Function ValueResolution(Val As Variant, asix_index As Long) As Long


    If x_value >= 1000000 Then  '#16_ELSE_CASE_CHK
        x_value = (x_value / 1000000) & "M"
    ElseIf x_value >= 1000 And x_value < 1000000 Then
        x_value = (x_value / 1000) & "K"
    ElseIf x_value >= 1 And x_value < 1000 Then
        ' do nothing
    ElseIf x_value >= 0.001 And x_value < 1 Then
        x_value = CStr((x_value * 1000)) & "m"
    ElseIf x_value >= 0.000001 And x_value < 0.001 Then
        x_value = (x_value * 1000000) & "u"
    Else
    End If
    
    If y_value <> "" Then   '#16_ELSE_CASE_CHK
        If y_value >= 1000000 Then  '#16_ELSE_CASE_CHK
            y_value = (y_value / 1000000) & "M"
        ElseIf y_value >= 1000 And y_value < 1000000 Then
            y_value = (y_value / 1000) & "K"
        ElseIf y_value >= 1 And y_value < 1000 Then
            ' do nothing
        ElseIf y_value >= 0.001 And y_value < 1 Then
            y_value = (y_value * 1000) & "m"
        ElseIf y_value >= 0.000001 And y_value < 0.001 Then
            y_value = (y_value * 1000000) & "u"
        Else
        End If
    Else
    End If
    
    If z_value <> "" Then   '#16_ELSE_CASE_CHK
        If z_value >= 1000000 Then  '#16_ELSE_CASE_CHK
            z_value = (z_value / 1000000) & "M"
        ElseIf z_value >= 1000 And z_value < 1000000 Then
            z_value = (z_value / 1000) & "K"
        ElseIf z_value >= 1 And z_value < 1000 Then
            ' do nothing
        ElseIf z_value >= 0.001 And z_value < 1 Then
            z_value = (z_value * 1000) & "m"
        ElseIf z_value >= 0.000001 And z_value < 0.001 Then
            z_value = (z_value * 1000000) & "u"
        Else
        End If
    Else
    End If
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "ValueResolution")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub VaryFreq(ClockPort As String, ClkFreq As Double, ACSpec As String)
On Error GoTo errHandler

Dim site As Variant
If LCase(glb_TesterType) = "jaguar" Then
    For Each site In theexec.sites
        thehdw.Protocol.ports(ClockPort).Halt
        thehdw.Protocol.ports(ClockPort).Enabled = False
    Next site

    Call theexec.Overlays.ApplyUniformSpecToHW(ACSpec, ClkFreq)


    thehdw.Wait 0.003
    thehdw.Protocol.ports(ClockPort).Enabled = True
    
    'Judge for RF Project 20200528
    Select Case UCase(thehdw.Protocol.ports(ClockPort).Family)
    Case "FRC"
        thehdw.Protocol.ports(ClockPort).FRC.ResetPLL
        thehdw.Wait 0.001
        thehdw.Protocol.ports(ClockPort).FRC.start
        thehdw.Wait 0.1
    Case "NWIRE"
        thehdw.Protocol.ports(ClockPort).NWire.ResetPLL
        thehdw.Wait 0.001
        Call thehdw.Protocol.ports(ClockPort).NWire.Frames("RunFreeClock").Execute
        thehdw.Protocol.ports(ClockPort).IdleWait
    Case Else
        theexec.Datalog.WriteComment "***The Familt Type in the PortMap Sheet is not support***"
    End Select
    
'''    If Flag_RF_Program = True Then
'''        TheHdw.Protocol.ports(ClockPort).FRC.ResetPLL
'''        TheHdw.Wait 0.001
'''        TheHdw.Protocol.ports(ClockPort).FRC.Start
'''        TheHdw.Wait 0.1
'''    Else
'''        TheHdw.Protocol.ports(ClockPort).NWire.ResetPLL
'''        TheHdw.Wait 0.001
'''        Call TheHdw.Protocol.ports(ClockPort).NWire.Frames("RunFreeClock").Execute
'''        TheHdw.Protocol.ports(ClockPort).IdleWait
'''    End If
Else
        Dim pin As String
        pin = Replace(LCase(ClockPort), "_port", vbNullString)
        If thehdw.Digital.Pins(pin).FreeRunningClock.IsRunning Then
            thehdw.Digital.Pins(pin).FreeRunningClock.stop
            thehdw.Digital.Pins(pin).FreeRunningClock.Enabled = False
        
        End If
        
            With thehdw.Digital.Pins(pin)
'                If Pin Like "*32768*" Then
'                    .FreeRunningClock.Frequency = ClkFreq / 2
'                Else
                    .FreeRunningClock.Frequency = ClkFreq
'                End If
                .FreeRunningClock.Enabled = True
                '' for UFP_Corr fix 200409
                '.FreeRunningClock.Frequency = TheExec.specs.AC.Item(ac_spec_pa).ContextValue
                .Connect
                .FreeRunningClock.start
            End With
    
End If
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "VaryFreq")
    If AbortTest Then Exit Sub Else Resume Next
End Sub



Public Sub MeasureFreq(MeasPin As String, ByRef result As PinListData)
On Error GoTo errHandler
    
    With thehdw.Digital.Pins(MeasPin).FreqCtr
        .Clear
        .EventSlope = Positive
        .EventSource = VOH
        .Interval = 0.01
        .Enable = IntervalEnable
        .start
    End With
    
    thehdw.Wait 10 * ms
    
    result = thehdw.Digital.Pins(MeasPin).FreqCtr.Read
    result = result.Math.divide(thehdw.Digital.Pins(MeasPin).FreqCtr.Interval)
    
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "MeasureFreq")
    If AbortTest Then Exit Sub Else Resume Next
End Sub
Public Function Decide_shmoo_patt(Init_Patt1 As Pattern, Init_Patt2 As Pattern, Init_Patt3 As Pattern, Init_Patt4 As Pattern, Init_Patt5 As Pattern, _
            Init_Patt6 As Pattern, Init_Patt7 As Pattern, Init_Patt8 As Pattern, Init_Patt9 As Pattern, Init_Patt10 As Pattern, _
            PayLoad_Patt1 As Pattern, PayLoad_Patt2 As Pattern, PayLoad_Patt3 As Pattern, PayLoad_Patt4 As Pattern, PayLoad_Patt5 As Pattern)
On Error GoTo errHandler
    
    Dim TempAry() As String
    Dim i As Integer
    Dim TempAry2() As String
    Dim tempStr As String
    
    Shmoo_Pattern = vbNullString
    
    If Init_Patt1 <> "" Then Shmoo_Pattern = Init_Patt1 '#16_ELSE_CASE_CHK
    If Shmoo_Pattern <> "" Then
        If Init_Patt2 <> "" Then Shmoo_Pattern = Shmoo_Pattern & "," & Init_Patt2
    Else
        Shmoo_Pattern = Init_Patt2
    End If
    
    If Shmoo_Pattern <> "" Then
        If Init_Patt3 <> "" Then Shmoo_Pattern = Shmoo_Pattern & "," & Init_Patt3
    Else
        Shmoo_Pattern = Init_Patt3
    End If
    
    If Shmoo_Pattern <> "" Then
        If Init_Patt4 <> "" Then Shmoo_Pattern = Shmoo_Pattern & "," & Init_Patt4
    Else
        Shmoo_Pattern = Init_Patt4
    End If
    
    If Shmoo_Pattern <> "" Then
        If Init_Patt5 <> "" Then Shmoo_Pattern = Shmoo_Pattern & "," & Init_Patt5
    Else
        Shmoo_Pattern = Init_Patt5
    End If
    
    If Shmoo_Pattern <> "" Then
        If Init_Patt6 <> "" Then Shmoo_Pattern = Shmoo_Pattern & "," & Init_Patt6
    Else
        Shmoo_Pattern = Init_Patt6
    End If
    If Shmoo_Pattern <> "" Then
        If Init_Patt7 <> "" Then Shmoo_Pattern = Shmoo_Pattern & "," & Init_Patt7
    Else
        Shmoo_Pattern = Init_Patt7
    End If
    If Shmoo_Pattern <> "" Then
        If Init_Patt8 <> "" Then Shmoo_Pattern = Shmoo_Pattern & "," & Init_Patt8
    Else
        Shmoo_Pattern = Init_Patt8
    End If
    If Shmoo_Pattern <> "" Then
        If Init_Patt9 <> "" Then Shmoo_Pattern = Shmoo_Pattern & "," & Init_Patt9
    Else
        Shmoo_Pattern = Init_Patt9
    End If
    If Shmoo_Pattern <> "" Then
        If Init_Patt10 <> "" Then Shmoo_Pattern = Shmoo_Pattern & "," & Init_Patt10
    Else
        Shmoo_Pattern = Init_Patt10
    End If
    If Shmoo_Pattern <> "" Then
        If PayLoad_Patt1 <> "" Then Shmoo_Pattern = Shmoo_Pattern & "," & PayLoad_Patt1
    Else
        Shmoo_Pattern = PayLoad_Patt1
    End If
    If Shmoo_Pattern <> "" Then
        If PayLoad_Patt2 <> "" Then Shmoo_Pattern = Shmoo_Pattern & "," & PayLoad_Patt2
    Else
        Shmoo_Pattern = PayLoad_Patt2
    End If
    If Shmoo_Pattern <> "" Then
        If PayLoad_Patt3 <> "" Then Shmoo_Pattern = Shmoo_Pattern & "," & PayLoad_Patt3
    Else
        Shmoo_Pattern = PayLoad_Patt3
    End If
    If Shmoo_Pattern <> "" Then
        If PayLoad_Patt4 <> "" Then Shmoo_Pattern = Shmoo_Pattern & "," & PayLoad_Patt4
    Else
        Shmoo_Pattern = PayLoad_Patt4
    End If
    If Shmoo_Pattern <> "" Then
        If PayLoad_Patt5 <> "" Then Shmoo_Pattern = Shmoo_Pattern & "," & PayLoad_Patt5
    Else
        Shmoo_Pattern = PayLoad_Patt5
    End If
    
    
    TempAry() = Split(Shmoo_Pattern, ",")
    For i = 0 To UBound(TempAry())
        TempAry2() = Split(TempAry(i), ":")
        If i = 0 Then
            tempStr = TempAry2(0)
        Else
            tempStr = tempStr & "," & TempAry2(0)
        End If
    Next i
    Shmoo_Pattern = tempStr
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Decide_shmoo_patt")
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function print_core_power(log_str As String, Power_pins As String)
On Error GoTo errHandler
    Dim p_ary() As String, p_cnt As Long, i As Long, j As Long
    Dim out_str As String, InstName As String, ShmooPower As Double
    Dim instancename As String
    If Power_pins = "" Then Exit Function   '#16_ELSE_CASE_CHK
    
        instancename = theexec.DataManager.instancename

    theexec.DataManager.DecomposePinList Power_pins, p_ary, p_cnt
    For i = 0 To p_cnt - 1
        If Not (theexec.DataManager.ChannelType(p_ary(i)) Like "N/C") Then  '#16_ELSE_CASE_CHK
            InstName = GetInstrument(p_ary(i), 0)
            Select Case InstName
            Case "DC-07"
                ShmooPower = thehdw.DCVI.Pins(p_ary(i)).Voltage
            Case "VHDVS", "HexVS", "VSM", "VS-5A", "VS-800mA"
                ShmooPower = thehdw.DCVS.Pins(p_ary(i)).Voltage.Main.value
            Case "HSD-U"
            Case Else
                Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "print_core_power", "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in print_core_power")
'                TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in print_core_power"
            End Select
            If i = 0 Then
      
                    out_str = instancename & "(Site" & theexec.sites.siteNumber & ")," & Curr_Shmoo_Condition.Char_Setup_Name & "," & left(log_str & Space(100), 20) & "," & p_ary(i) & "=" & Format(ShmooPower, "0.000")
            
            Else
                out_str = out_str & "," & p_ary(i) & "=" & Format(thehdw.DCVS.Pins(p_ary(i)).Voltage.Main, "0.000")
            End If
        Else
        End If
    Next i
    If theexec.Flow.enableWord("Datalog_Verbose") = True Then   '#16_ELSE_CASE_CHK
        theexec.Datalog.WriteComment out_str
        If isDebugMode = True Then theexec.AddOutput out_str    '#16_ELSE_CASE_CHK
    Else
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "print_core_power")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Shmoo_Save_core_power_per_site(Power_pins As String, ShmooPower() As SiteDouble)
On Error GoTo errHandler
    Dim p_ary() As String, p_cnt As Long, i As Long, InstName As String
    theexec.DataManager.DecomposePinList Power_pins, p_ary, p_cnt
    For i = 0 To p_cnt - 1
        If theexec.DataManager.ChannelType(p_ary(i)) <> "N/C" Then
            InstName = GetInstrument(p_ary(i), 0)
            Select Case InstName
               Case "DC-07"
                  ShmooPower(i) = thehdw.DCVI.Pins(p_ary(i)).Voltage
               Case "VHDVS", "HexVS", "VSM", "VS-5A", "VS-800mA"
                  ShmooPower(i) = thehdw.DCVS.Pins(p_ary(i)).Voltage.Main.value
               Case "HexVS", "VSM"
                   ShmooPower(i) = thehdw.DCVS.Pins(p_ary(i)).Voltage.Main.value
               Case "HSD-U"
               Case Else
                    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Save_core_power_per_site", "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Save_core_power_per_site")
'                    TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Save_core_power_per_site"
            End Select
        End If
    Next i
    
   Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Save_core_power_per_site")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Get_Current_Apply_Pin(Power_pins As String)
On Error GoTo errHandler
' Only get Power pins needed  for shmoo
' Ignore any IO pins and FreeRun Freq pins
    Dim active_setup As String, curr_axis As Variant, curr_track As Variant, apply_Pin As String, apply_Pin_arry() As String, pin_count As Long, i As Long
    Dim p_ary() As String, p_cnt As Long
    Dim InstName As String
    Dim DCVI_pin As String
    Dim DCVS_pin As String
    Dim site As Variant 'Carter, 20240304
    Power_pins = vbNullString
    Set g_Globalpointval = Nothing
    active_setup = theexec.DevChar.Setups.ActiveSetupName
    For Each curr_axis In theexec.DevChar.Setups(active_setup).Shmoo.Axes.list
        ''exit for if any axis is not power pin -by SY
        If theexec.DevChar.Setups(active_setup).Shmoo.Axes(curr_axis).ApplyTo.Pins = "" Then GoTo NextAxis
        apply_Pin = theexec.DevChar.Setups(active_setup).Shmoo.Axes(curr_axis).ApplyTo.Pins
        'Add for store shmoo global spec to avoid direct to apply Vmain used for Vbump function
        If g_Vbump_function = True Then
            Call theexec.DataManager.DecomposePinList(apply_Pin, apply_Pin_arry, pin_count)
            For i = 0 To pin_count - 1
                g_Globalpointval.AddPin (apply_Pin_arry(i))
                For Each site In theexec.sites
                    g_Globalpointval.Pins(apply_Pin_arry(i)).value = theexec.DevChar.Results(active_setup).Shmoo.CurrentPoint.Axes(curr_axis).value
                Next site
                'Check_DCVSorDCVI LCase(apply_Pin_arry(i)), DCVS_pin, DCVI_pin
                SortAllPinInstrumentType LCase(apply_Pin_arry(i)), DCVS_pin, DCVI_pin
            Next i
        Else
        End If

        For Each curr_track In theexec.DevChar.Setups(active_setup).Shmoo.Axes(curr_axis).TrackingParameters.list
            apply_Pin = theexec.DevChar.Setups(active_setup).Shmoo.Axes(curr_axis).TrackingParameters.item(curr_track).ApplyTo.Pins
            If g_Vbump_function = True Then
               Call theexec.DataManager.DecomposePinList(apply_Pin, apply_Pin_arry, pin_count)
               For i = 0 To pin_count - 1
                   g_Globalpointval.AddPin (apply_Pin_arry(i))
                   For Each site In theexec.sites
                       g_Globalpointval.Pins(apply_Pin_arry(i)).value = theexec.DevChar.Results(active_setup).Shmoo.CurrentPoint.Axes(curr_axis).TrackingParameters(curr_track).value
                   Next site
                   'Check_DCVSorDCVI LCase(apply_Pin_arry(i)), DCVS_pin, DCVI_pin
                   SortAllPinInstrumentType LCase(apply_Pin_arry(i)), DCVS_pin, DCVI_pin
                   Power_pins = Power_pins & IIf(Power_pins = "", "", ",") & apply_Pin_arry(i)
               Next i
            Else
            End If
        Next curr_track
        Power_pins = DCVS_pin & IIf(DCVI_pin = "", "", ",") & DCVI_pin
        g_ShmooDCVS = DCVS_pin
        g_ShmooDCVI = DCVI_pin
NextAxis:
   Next curr_axis
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Get_Current_Apply_Pin")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Shmoo_Restore_Power_per_site(ShmooPowerStored_Pins As String, ShmooPowerStored() As SiteDouble, log_header As String, Optional Restore_Pins As String = vbNullString)
On Error GoTo errHandler
    'if Restore_Pins="" then restore all ShmooPowerStored_Pins
    Dim p_ary() As String, p_cnt As Long, i As Long
    Dim rp_ary() As String, rp_cnt As Long
    Dim InstName As String
    Dim tmp_ShmooPowerStored_Pins() As String
    Dim p As Variant, pn As String
    Dim Need_ReStore_Pin As Boolean
    Dim Restore_Pins_Dict As New Dictionary
    Dim Restore_Pin_str As String
    Dim ShmooPowerStored_Pins_str  As String
    
    If ShmooPowerStored_Pins = "" Then Exit Function
    
    If Restore_Pins = "" Then
        Restore_Pin_str = ShmooPowerStored_Pins
    Else
        Restore_Pin_str = Restore_Pins
    End If
    theexec.DataManager.DecomposePinList ShmooPowerStored_Pins, p_ary, p_cnt
    
    theexec.DataManager.DecomposePinList Restore_Pin_str, rp_ary, rp_cnt
    Restore_Pin_str = Join(rp_ary, ",")
   Create_Pin_Dic Restore_Pin_str, Restore_Pins_Dict

    For i = 0 To p_cnt - 1
        If theexec.DataManager.ChannelType(p_ary(i)) <> "N/C" And Restore_Pins_Dict.Exists(LCase(p_ary(i))) = True Then '#16_ELSE_CASE_CHK
            InstName = GetInstrument(p_ary(i), 0)
            Select Case InstName
            Case "DC-07"
                thehdw.DCVI.Pins(p_ary(i)).Voltage = ShmooPowerStored(i)
            Case "VHDVS", "HexVS", "VSM", "VS-5A", "VS-800mA"
                thehdw.DCVS.Pins(p_ary(i)).Voltage.Main.value = ShmooPowerStored(i)
            Case "HSD-U"
            Case Else
                Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Restore_Power_per_site", "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Restore_Power_per_site")
'                TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Restore_Power_per_site"
            End Select
        Else
        End If
    Next i
    'print_core_power log_header, ShmooPowerStored_Pins
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Restore_Power_per_site")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Create_Pin_Dic(Pins As String, Pin_Dict As Dictionary)
On Error GoTo errHandler
    Dim p_ary() As String, p_cnt As Long, pn As String, p As Variant
    theexec.DataManager.DecomposePinList Pins, p_ary, p_cnt
    Pin_Dict.RemoveAll
    For Each p In p_ary
        pn = LCase(CStr(p))
        Pin_Dict.Add pn, True
    Next p
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Create_Pin_Dic")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Shmoo_Set_Power(Power_pins As String, level As String, log_header As String, Optional Use_Performance_Mode As Boolean = False, Optional skip_pin As String = vbNullString)
On Error GoTo errHandler
    Dim p_ary() As String, p_cnt As Long, i As Long, j As Long
    Dim Core_p_ary() As String, Core_p_cnt As Long
    Dim main_power As String, main_spec_name As String
    Dim ratio As Double
    Dim Flag_core_power_found As Boolean
    Dim p_mode As String, p_mode_code As Long, block_name As String
    Dim p_mode_code_str As String
    Dim tmp_ary() As String
    Dim shmoo_pin As String
    Dim Active_Test_inst_name As String
    Dim DC_cat As String, Dc_spec_type As String
    Dim SP As Variant, t As String
    Dim InstName As String
    Dim Skip_Pin_Dic As New Dictionary
    Dim Need_Skip_Pin As Boolean
    ' Assumption:
    ' 1. Only use Selector :Typ,Max,Min
    ' 2. DC spec name is  VDD_CPU_VAR_C/S/G/H
    ' 3. DC spec will not be changed
    If Power_pins = "" Then Exit Function
    
    If skip_pin <> "" Then Create_Pin_Dic skip_pin, Skip_Pin_Dic

    theexec.DataManager.DecomposePinList Power_pins, p_ary, p_cnt
    theexec.DataManager.GetInstanceContext DC_cat, t, t, t, t, t, t, t
    For Each SP In theexec.Specs.DC.Categories(DC_cat).SpecList
        SP = LCase(SP)
        
        If SP Like "*_var_c" Then
            Dc_spec_type = "C"
            Exit For
        ElseIf SP Like "*_var_g" Then
            Dc_spec_type = "G"
            Exit For
        ElseIf SP Like "*_var_s" Then
            Dc_spec_type = "S"
            Exit For
        ElseIf SP Like "*_var_h" Then
            Dc_spec_type = "H"
            Exit For
        '===================================================================================for dc spec BI/ SC
        ElseIf SP Like "*_var_bi" Then
            Dc_spec_type = "BI"
            Exit For
        ElseIf SP Like "*_var_sc" Then
            Dc_spec_type = "SC"
            Exit For
        ElseIf SP Like "*_var" Then ''added case for new DC spec sheets method
            Dc_spec_type = vbNullString
            Exit For
        '===================================================================================for dc spec BI/ SC
        Else
            Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Set_Power", "DC spec " & SP & " is not ended with _VAR_C/S/G/H in " & theexec.DataManager.instancename)
'            TheExec.ErrorLogMessage "DC spec " & SP & " is not ended with _VAR_C/S/G/H in " & TheExec.DataManager.instancename
        End If
        Exit For
    Next SP
    
    If Dc_spec_type = "" Then
        Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Set_Power", "DC spec " & SP & " is not ended with _VAR_C/S/G/H in " & theexec.DataManager.instancename)
'        TheExec.ErrorLogMessage "DC spec " & SP & " is not ended with _VAR_C/S/G/H in " & TheExec.DataManager.instancename
    End If
    
    For i = 0 To p_cnt - 1
        p_ary(i) = LCase(p_ary(i))
        Need_Skip_Pin = False
        If skip_pin <> "" Then
            If Skip_Pin_Dic.Exists(p_ary(i)) = True Then Need_Skip_Pin = True
        End If
        If Not (theexec.DataManager.ChannelType(p_ary(i)) Like "N/C") And Need_Skip_Pin = False Then
            InstName = GetInstrument(p_ary(i), 0)
            Select Case InstName
               Case "DC-07":
                        Select Case level
                            Case "NV":  thehdw.DCVI.Pins(p_ary(i)).Voltage = theexec.Specs.DC.item(p_ary(i) & "_" & "VAR" & "_" & Dc_spec_type).Categories(DC_cat).Typ.value
                            Case "LV":  thehdw.DCVI.Pins(p_ary(i)).Voltage = theexec.Specs.DC.item(p_ary(i) & "_" & "VAR" & "_" & Dc_spec_type).Categories(DC_cat).min.value
                            Case "HV":  thehdw.DCVI.Pins(p_ary(i)).Voltage = theexec.Specs.DC.item(p_ary(i) & "_" & "VAR" & "_" & Dc_spec_type).Categories(DC_cat).max.value
                            Case Else
                                Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Set_Power", level & " is not supported in " & theexec.DataManager.instancename)
'                                TheExec.ErrorLogMessage level & " is not supported in " & TheExec.DataManager.instancename
                        End Select
               Case "VHDVS", "HexVS", "VSM", "VS-5A", "VS-800mA":
                        Select Case level
                            Case "NV":  thehdw.DCVS.Pins(p_ary(i)).Voltage.Main.value = theexec.Specs.DC.item(p_ary(i) & "_" & "VAR" & "_" & Dc_spec_type).Categories(DC_cat).Typ.value
                            Case "LV":  thehdw.DCVS.Pins(p_ary(i)).Voltage.Main.value = theexec.Specs.DC.item(p_ary(i) & "_" & "VAR" & "_" & Dc_spec_type).Categories(DC_cat).min.value
                            Case "HV":  thehdw.DCVS.Pins(p_ary(i)).Voltage.Main.value = theexec.Specs.DC.item(p_ary(i) & "_" & "VAR" & "_" & Dc_spec_type).Categories(DC_cat).max.value
                            Case Else
                                Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Set_Power", level & " is not supported in " & theexec.DataManager.instancename)
'                                TheExec.ErrorLogMessage level & " is not supported in " & TheExec.DataManager.instancename
                        End Select
               Case "HSD-U"
               Case Else
                    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Set_Power", "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Set_Power")
'                        TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Set_Power"
            End Select

        End If
    Next i
    thehdw.Wait 0.001
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Set_Power")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Decide_LVCC_HVCC(Vcc_min As String, Vcc_max As String, Shmoo_hole As String, Step_NV As Long, RangeLow As Double, RangeStepSize As Double, Shmoo_result_PF As String, SetupName As String, step_Start As Long, step_Stop As Long, Step_x As Long, sites As Variant)
'Public Function Decide_LVCC_HVCC(Vcc_min As String, Vcc_max As String, Shmoo_hole As String, Step_NV As Long, RangeLow As Double, RangeStepSize As Double, Shmoo_result_PF As String, SetupName As String, step_Start As Long, step_Stop As Long, Step_x As Long)
On Error GoTo errHandler
    Dim FlagFirstPass As Boolean, FlagFirstFail As Boolean
    Dim last_point_result As String, current_point_result As String, char_pt As String
    Dim AllFail As Boolean
    Dim min_point As Long, max_point As Long, current_point As Long
    Dim FlagHole As Boolean
    Dim FlagPF(1000) As Boolean
    Dim FlagFP(1000) As Boolean
    Dim FlagPF_Count As Long
    Dim FlagFP_Count As Long
    Dim i As Long, j As Long
    Dim test_name As String
    Dim step_p As Long
    Dim x_pra As String
    Dim show_vcc As String
    Dim shmoo_form, shmoo_stop, shmoo_step As String
    Dim AllPass As Boolean
    Dim lvcc_point As Integer
    Dim hvcc_point As Integer
    Dim str_temp() As String
    Dim mode_type As String
    Dim Point_Volt() As Double
    Dim site As Variant 'Carter, 20240304
	
	'20240531shmoo hole fail log
	Dim find_shmoo_hole As Boolean
    Dim str_len, start_point As Integer
    Dim search_dif As Boolean
    Dim report_point As Integer
    Dim z As Long
	
    Vcc_min = vbNullString
    Vcc_max = vbNullString
    
    show_vcc = "[Shmoo,"
    
    Dim Shmoo_setup_name, Shmoo_TestInst_Name As String
    
    step_p = Len(Shmoo_result_PF)
    find_shmoo_hole = False	'20240531shmoo hole fail log
    Shmoo_setup_name = TheExec.DevChar.Setups.ActiveSetupName
    Shmoo_TestInst_Name = TheExec.DevChar.ActiveDataObject.TestName
    shmoo_form = CStr(RangeLow)
    shmoo_stop = CStr(RangeLow + RangeStepSize * step_p)
    shmoo_step = CStr(step_p + 1)
    x_pra = TheExec.DevChar.ActiveDataObject.XParameter
    str_temp = Split(Shmoo_TestInst_Name, "_")
    mode_type = str_temp(0)
'    If (TheExec.EnableWord("One_transition") = True) Then
        If (LCase(Shmoo_TestInst_Name) Like "dfth*" Or LCase(Shmoo_TestInst_Name) Like "hfh*" Or LCase(Shmoo_TestInst_Name) Like "mch*") Then   '#16_ELSE_CASE_CHK
            If step_Start > step_Stop Then
                Step_NV = step_Stop
            Else
                Step_NV = step_Start
            End If
        ElseIf (LCase(Shmoo_TestInst_Name) Like "dftl*" Or LCase(Shmoo_TestInst_Name) Like "hfl*" Or LCase(Shmoo_TestInst_Name) Like "mcl*") Then
            If step_Start > step_Stop Then
                Step_NV = step_Start
            Else
                Step_NV = step_Stop
            End If
        Else
        End If
'    End If
    
    If (Step_NV > step_p Or Step_NV < 0) Then   '#16_ELSE_CASE_CHK
        If (LCase(Shmoo_TestInst_Name) Like "dfth*" Or LCase(Shmoo_TestInst_Name) Like "hfh*" Or LCase(Shmoo_TestInst_Name) Like "mch*") Then
            Step_NV = step_Start
        
        ElseIf (LCase(Shmoo_TestInst_Name) Like "dftl*" Or LCase(Shmoo_TestInst_Name) Like "hfl*" Or LCase(Shmoo_TestInst_Name) Like "mcl*") Then
            Step_NV = step_Stop
        Else
        End If
    Else
    End If

    test_name = theexec.DevChar.ActiveDataObject.TestName
    
    
    Shmoo_hole = "NH"

    'Early exit if
    '   NV fails
    '   All Fail
    '   Not pass or fail
'-----------------------------------------------------------------------------------------------
    AllPass = True
    ' if fails at NV
    ' step_start is the lowest char value
    ' Check if all points fail
    
    For i = step_Start To step_Stop Step Step_x
        char_pt = mid(Shmoo_result_PF, i + 1, 1)
        If char_pt = "F" Then   '#16_ELSE_CASE_CHK
            AllPass = False
        Else
        End If
    Next i
    
    If AllPass = True Then  '#16_ELSE_CASE_CHK
        Vcc_min = CStr(RangeLow): Vcc_max = CStr(RangeLow + RangeStepSize * Abs(step_Stop - step_Start))
		'20240531shmoo hole fail log
		If TheExec.enableWord("Find_shmoo_hole_low_power_twice") = True Then
			ReDim Preserve Find_shmoo_hole_low_power_twice_hvcc_voltage(0) As New SiteDouble
            Find_shmoo_hole_low_power_twice_hvcc_voltage(0)(sites) = 0
        End If
        GoTo end_lvcc_hvcc
    Else
    End If
'-----------------------------------------------------------------------------------------------
    AllFail = True
    ' if fails at NV
    ' step_start is the lowest char value
    ' Check if all points fail
    
    For i = step_Start To step_Stop Step Step_x
        char_pt = mid(Shmoo_result_PF, i + 1, 1)
        If char_pt = "P" Then   '#16_ELSE_CASE_CHK
            AllFail = False
        Else
        End If
        If Not (char_pt = "P" Or char_pt = "F") Then    '#16_ELSE_CASE_CHK
           Vcc_min = "-7777": Vcc_max = "7777"
           GoTo end_lvcc_hvcc
        Else
        End If
    Next i
    
    If AllFail = True Then  '#16_ELSE_CASE_CHK
        Vcc_max = "9999": Vcc_min = "-9999"
        GoTo end_lvcc_hvcc
    Else
    End If
    
    
'%%%%%%%%%%%%%%%%%%%%%%%% NV Fail ,Report 5555  (Open while Check HH,LH)%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'20170213 add boolean to control NV hole
    Dim NVffff As Boolean
    NVffff = False
    '20170105 Roy added
    If NVffff = True Then
        If Step_NV > 0 Then '#16_ELSE_CASE_CHK
            If mid(Shmoo_result_PF, Step_NV, 1) = "F" Then  '#16_ELSE_CASE_CHK
                Vcc_max = "5555":  Vcc_min = "-5555"
                GoTo end_lvcc_hvcc
            Else
            End If
        Else
        End If
    Else
    
    '    If Step_NV > 0 Then
    '        If Mid(Shmoo_result_PF, Step_NV, 1) = "F" Then
    '            Vcc_max = "5555":  Vcc_min = "-5555"
    '            GoTo end_lvcc_hvcc
    '        End If
    '    End If
    
    End If
'%%%%%%%%%%%%%%%%%%%%%%%%%%%  HF VID,VICM%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
       'hvcc_point or lvcc_point is -1,shmoo hole
     If (UCase(mode_type) Like "HF*") Or (UCase(x_pra) = "VID") Or (UCase(x_pra) = "VICM") Then  '#16_ELSE_CASE_CHK
         ReDim Point_Volt(step_p) As Double
         For i = 1 To step_p
             Point_Volt(i) = RangeLow + RangeStepSize * (i - 1)
         Next i
         
         hvcc_point = Search_HVCC(Shmoo_result_PF)

		lvcc_point = Search_LVCC(Shmoo_result_PF)
         If hvcc_point > step_p Then hvcc_point = step_p '#16_ELSE_CASE_CHK
         If (hvcc_point = -1) Then
             Vcc_max = "5555"
         Else
             Vcc_max = CStr(Point_Volt(hvcc_point))
         End If
         
         If (lvcc_point = -1) Then
             Vcc_min = "-5555"
         Else
             Vcc_min = CStr(Point_Volt(lvcc_point))
         End If
         
         GoTo end_lvcc_hvcc
    Else
    End If

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%  VIH,VIL %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
       'hvcc_point or lvcc_point is -1,shmoo hole
       'hvcc_point or lvcc_point is -2,first point fail
If UCase(theexec.DataManager.instancename) Like "*USBPICO*" Or UCase(theexec.DataManager.instancename) Like "*LPDPRX*" Then
Else
        If UCase(mode_type) Like "HIO*" Then
            If (InStr(LCase(x_pra), "_vih_") > 0) Or (InStr(LCase(x_pra), "_vil_") > 0) Then
                ReDim Point_Volt(step_p) As Double
                For i = 1 To step_p
                    Point_Volt(i) = RangeLow + RangeStepSize * (i - 1)
                Next i
                If (LCase(x_pra) = "vih") Then  '#16_ELSE_CASE_CHK
                    lvcc_point = Search_VIH_LVCC(Shmoo_result_PF)
                    If (lvcc_point = -1) Then
                        Vcc_min = "-5555"
                    ElseIf (lvcc_point = -2) Then
                        Vcc_min = "-8888"
                    Else
                        Vcc_min = CStr(Point_Volt(lvcc_point))
                    End If
                Else
                End If
                If (LCase(x_pra) = "vil") Then  '#16_ELSE_CASE_CHK
                    hvcc_point = Search_VIL_HVCC(Shmoo_result_PF)
                    If (hvcc_point = -1) Then
                        Vcc_max = "5555"
                    ElseIf (lvcc_point = -2) Then
                        Vcc_min = "8888"
                    Else
                        Vcc_max = CStr(Point_Volt(hvcc_point))
                    End If
                Else
                End If
                GoTo end_lvcc_hvcc
            End If
        End If
    End If
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


    ReDim Point_Volt(step_p) As Double
    For i = 1 To step_p
        Point_Volt(i) = RangeLow + RangeStepSize * (i - 1)
    Next i
    
    hvcc_point = Search_HVCC(Shmoo_result_PF)
    '//////////////////////////Find_shmoo_hole_low_power_twice Start////////////////////////////////////////
    
    If TheExec.enableWord("Find_shmoo_hole_low_power_twice") = True Then
        str_len = Len(Shmoo_result_PF)
        start_point = Search_Low2High_First_Pass(Shmoo_result_PF)
        search_dif = False
        report_point = 0
        For z = start_point To str_len
            char_pt = mid(Shmoo_result_PF, z, 1)
            
            If Not (search_dif) Then
                If (char_pt = "F") Then
                    report_point = z - 1
                    search_dif = True
                    'If TheExec.enableWord("Find_shmoo_hole_low_power_twice") = True Then Exit For 'add by Wyatt 20231018
                End If
            Else
                If (char_pt = "P") Then
                    find_shmoo_hole = True
                    z = str_len
                End If
            End If
        Next z
        If find_shmoo_hole = True Then
            ReDim Preserve Find_shmoo_hole_low_power_twice_hvcc_step(0) As New SiteLong
            Find_shmoo_hole_low_power_twice_hvcc_step(0)(sites) = report_point + 1
            If mid(Shmoo_result_PF, report_point + 2, 1) = "F" Then
                ReDim Preserve Find_shmoo_hole_low_power_twice_hvcc_step(UBound(Find_shmoo_hole_low_power_twice_hvcc_step) + 1) As New SiteLong
                Find_shmoo_hole_low_power_twice_hvcc_step(1)(sites) = report_point + 2
            End If
        Else
            ReDim Preserve Find_shmoo_hole_low_power_twice_hvcc_step(0) As New SiteLong
            ReDim Preserve Find_shmoo_hole_low_power_twice_hvcc_voltage(0) As New SiteDouble
            'ReDim Preserve Find_shmoo_hole_low_power_twice_hvcc_step(0) As New SiteLong
            'ReDim Preserve Find_shmoo_hole_low_power_twice_hvcc_step(1) As New SiteLong
            'ReDim Preserve Find_shmoo_hole_low_power_twice_hvcc_voltage(1) As New SiteDouble
            Find_shmoo_hole_low_power_twice_hvcc_step(0)(sites) = 0
            Find_shmoo_hole_low_power_twice_hvcc_voltage(0)(sites) = 0
        End If
    End If
    
    '//////////////////////////Find_shmoo_hole_low_power_twice End////////////////////////////////////////
    lvcc_point = Search_LVCC(Shmoo_result_PF)
    If hvcc_point > step_p Then hvcc_point = step_p '#16_ELSE_CASE_CHK
    If TheExec.enableWord("Find_shmoo_hole_low_power_twice") = True And find_shmoo_hole = True Then '/20231018
        For Each site In TheExec.sites
            If site = sites Then
                ReDim Preserve Find_shmoo_hole_low_power_twice_hvcc_voltage(UBound(Find_shmoo_hole_low_power_twice_hvcc_step)) As New SiteDouble
            End If
        Next site
        For z = 0 To UBound(Find_shmoo_hole_low_power_twice_hvcc_step)
            Find_shmoo_hole_low_power_twice_hvcc_voltage(z)(sites) = CStr(Point_Volt(Find_shmoo_hole_low_power_twice_hvcc_step(z)(sites)))
        Next z
    End If
    If (hvcc_point = -1) Then
        Vcc_max = "5555"
    Else
        Vcc_max = CStr(Point_Volt(hvcc_point))
    End If
    
    If (lvcc_point = -1) Then
        Vcc_min = "-5555"
    Else
        Vcc_min = CStr(Point_Volt(lvcc_point))
    End If
    
    GoTo end_lvcc_hvcc
    
    If (Step_NV = -1) Then  '#16_ELSE_CASE_CHK
        
        If step_Start > step_Stop Then
            Step_NV = step_Stop
        Else
            Step_NV = step_Start
        End If
    Else
    End If
   
    
    
'----------------------------------------------------------------------------------------------
    If Not (InStr(LCase(test_name), "vih") > 0 Or InStr(LCase(test_name), "vil") > 0 Or InStr(LCase(x_pra), "vih") > 0 Or InStr(LCase(x_pra), "vil") > 0 Or InStr(LCase(x_pra), "vid") > 0) Then    '#16_ELSE_CASE_CHK
        If mid(Shmoo_result_PF, Step_NV + 1, 1) = "F" Then  '#16_ELSE_CASE_CHK
        
            For i = Step_NV To (step_Stop - step_Start) / Step_x    'search low to high voltage
                char_pt = mid(Shmoo_result_PF, i + 1, 1)
                If char_pt = "P" Then
                    Vcc_max = 8888
                    i = (step_Stop - step_Start) / Step_x
                End If
            Next i
            If Vcc_max = "" Then Vcc_max = 9999 '#16_ELSE_CASE_CHK
            
            
            For i = Step_NV To 0 Step -1    'search low to high voltage
                char_pt = mid(Shmoo_result_PF, i + 1, 1)
                If char_pt = "P" Then
                    Vcc_min = -8888
                    i = 0
                End If
            Next i
            If Vcc_min = "" Then Vcc_min = -9999    '#16_ELSE_CASE_CHK

            Exit Function
        Else
        End If
    Else
    End If
'-------------------------------------------------------------------------------------------------
    ' Find LVCC: NV value to Low value
    ' 01         Step
    ' FFPPPpPPFF
    ' PPPPPpPPFF
    
    lvcc_point = 0
    
    If Not (InStr(LCase(test_name), "vih") > 0 Or InStr(LCase(test_name), "vil") > 0 Or InStr(LCase(x_pra), "vih") > 0 Or InStr(LCase(x_pra), "vil") > 0 Or InStr(LCase(x_pra), "vid") > 0) Then

        For i = Step_NV To 0 Step -1
            If mid(Shmoo_result_PF, i + 1, 1) = "F" Then    '#16_ELSE_CASE_CHK

                Vcc_min = CStr(RangeLow + RangeStepSize * (i + 1))
'                TheExec.Datalog.WriteComment "[LVCC=" & Vcc_min & "]"
                lvcc_point = i
                i = 0
            Else
            End If
        Next i

        If Vcc_min = "" Then Vcc_min = CStr(RangeLow)   '#16_ELSE_CASE_CHK
    Else
        If InStr(LCase(test_name), "vih") > 0 Or InStr(LCase(x_pra), "vih") > 0 Or InStr(LCase(x_pra), "vid") > 0 Then  '#16_ELSE_CASE_CHK

            For i = (step_Stop - step_Start) / Step_x To 0 Step -1 'search high to low voltage
                char_pt = mid(Shmoo_result_PF, i + 1, 1)
                If char_pt = "F" Then
                    Vcc_min = CStr(RangeLow + RangeStepSize * (i + 1))
                    lvcc_point = i
                    i = 0
                End If
            Next i
            If Vcc_min = "" Then Vcc_min = CStr(RangeLow)   '#16_ELSE_CASE_CHK
        Else
        End If
        
    End If
 
    If lvcc_point <> 0 Then '#16_ELSE_CASE_CHK
        For i = lvcc_point - 1 To 0 Step -1
            If mid(Shmoo_result_PF, i + 1, 1) = "P" Then    '#16_ELSE_CASE_CHK
                Vcc_min = "-5555"
            End If
        Next i
    Else
    End If
    
    Dim Fail_index As Integer
    ''*******************************************AI***********************************************
    If theexec.enableWord("AI_Fail_Log") = True And LCase(theexec.DataManager.instancename) Like "*lvcc*" Then  '#16_ELSE_CASE_CHK
        If Vcc_min <> "-5555" Then
            Fail_index = 0
        
            For i = Step_NV To 0 Step -1
                If mid(Shmoo_result_PF, i + 1, 1) = "F" Then    '#16_ELSE_CASE_CHK
                    Voltage_fail_collect(Fail_index) = CStr(RangeLow + RangeStepSize * (i))
                    Fail_index = Fail_index + 1
                    If Fail_index = 5 Then i = 0    '#16_ELSE_CASE_CHK
                Else
                End If
            Next i
        Else
        ''For shmoo hole collect 10 point fail cycle
            Fail_index = 0
        
            For i = Step_NV To 0 Step -1
                If mid(Shmoo_result_PF, i + 1, 1) = "F" Then    '#16_ELSE_CASE_CHK
                    Voltage_fail_collect(Fail_index) = CStr(RangeLow + RangeStepSize * (i))
                    Fail_index = Fail_index + 1
                    If Fail_index = 10 Then i = 0   '#16_ELSE_CASE_CHK
                Else
                End If
            Next i
            
        End If
        Voltage_fail_point = Fail_index
    End If
      ''*******************************************AI***********************************************
'--------------------------------------------------------------------------------------------------------------
    ' Find HVCC: NV value to Hi value
    ' FFPPPpPPFF
    ' PPPPPpPPPP
    
    hvcc_point = 0
    
    If Not (InStr(LCase(test_name), "vih") > 0 Or InStr(LCase(test_name), "vil") > 0 Or InStr(LCase(x_pra), "vih") > 0 Or InStr(LCase(x_pra), "vil") > 0) Then
        
        For i = Step_NV To step_p Step 1
            If mid(Shmoo_result_PF, i + 1, 1) = "F" Then    '#16_ELSE_CASE_CHK
                Vcc_max = CStr(RangeLow + RangeStepSize * (i - 1))
                hvcc_point = i
                i = step_p
            Else
            End If
        Next i
        
        If Vcc_max = "" Then Vcc_max = CStr(RangeLow + RangeStepSize * Abs(step_Stop - step_Start)) '#16_ELSE_CASE_CHK
    Else
        If InStr(LCase(test_name), "vil") > 0 Or InStr(LCase(x_pra), "vil") > 0 Then    '#16_ELSE_CASE_CHK
            
            For i = 0 To (step_Stop - step_Start) / Step_x    'search low to high voltage
                char_pt = mid(Shmoo_result_PF, i + 1, 1)
                If char_pt = "F" Then   '#16_ELSE_CASE_CHK
                    Vcc_max = CStr(RangeLow + RangeStepSize * (i - 1))
                    hvcc_point = i
                    
                    i = (step_Stop - step_Start) / Step_x
                Else
                End If
            Next i
            If Vcc_max = "" Then Vcc_max = CStr(RangeLow + RangeStepSize * Abs(step_Stop - step_Start)) '#16_ELSE_CASE_CHK
        End If
    End If
    
    'show_vcc = show_vcc & Vcc_max
        
    If hvcc_point <> 0 Then '#16_ELSE_CASE_CHK
        For i = hvcc_point + 1 To (step_Stop - step_Start) / Step_x Step 1
            If mid(Shmoo_result_PF, i + 1, 1) = "P" Then    '#16_ELSE_CASE_CHK
                Vcc_max = "5555"
            Else
            End If
        Next i
    Else
    End If
    
    ''*******************************************AI***********************************************
    If theexec.enableWord("AI_Fail_Log") = True And LCase(theexec.DataManager.instancename) Like "*hvcc*" Then  '#16_ELSE_CASE_CHK
        If Vcc_min <> "5555" Then
            Fail_index = 0

            For i = Step_NV To step_p Step 1
                If mid(Shmoo_result_PF, i + 1, 1) = "F" Then    '#16_ELSE_CASE_CHK
                    Voltage_fail_collect(Fail_index) = CStr(RangeLow + RangeStepSize * (i))
                    Fail_index = Fail_index + 1
                    If Fail_index = 5 Then i = step_p
                Else
                End If
            Next i
        Else
        ''For shmoo hole collect 10 point fail cycle
            Fail_index = 0
        
            For i = Step_NV To step_p Step 1
                If mid(Shmoo_result_PF, i + 1, 1) = "F" Then    '#16_ELSE_CASE_CHK
                    Voltage_fail_collect(Fail_index) = CStr(RangeLow + RangeStepSize * (i))
                    Fail_index = Fail_index + 1
                    If Fail_index = 10 Then i = step_p
                Else
                End If
            Next i
            
        End If
        Voltage_fail_point = Fail_index
    End If
    ''*******************************************AI***********************************************
end_lvcc_hvcc:


    If Abs(Vcc_min) < 0.000000000001 Then Vcc_min = 0   '#16_ELSE_CASE_CHK
    If Abs(Vcc_max) < 0.000000000001 Then Vcc_max = 0   '#16_ELSE_CASE_CHK
    
    ''======170425 Char shmoo error code count start=====''
    
     For Each site In theexec.sites
            total_shmoo_count = total_shmoo_count + 1
     Next site
    
     If F_shmoo_abnormal_counter = True Then

        For Each site In theexec.sites
            If Trim(Vcc_max) = "5555" Or Trim(Vcc_min) = "-5555" Then   '#16_ELSE_CASE_CHK
                shmoohole_count = shmoohole_count + 1
            Else
            End If

            If Trim(Vcc_max) = "9999" Or Trim(Vcc_min) = "-9999" Then   '#16_ELSE_CASE_CHK
                shmooallfail_count = shmooallfail_count + 1
            Else
            End If

            If Trim(Vcc_max) = "7777" Or Trim(Vcc_min) = "-7777" Then   '#16_ELSE_CASE_CHK
                shmooalarm_count = shmooalarm_count + 1
            Else
            End If

            included_shmoo_count = included_shmoo_count + 1
        Next site

      Else

        For Each site In theexec.sites
            excluded_shmoo_count = excluded_shmoo_count + 1
        Next site

      End If
    ''======170425 Char shmoo error code count end=====''
    
    show_vcc = show_vcc & Vcc_min
    show_vcc = show_vcc & "," & Vcc_max
    show_vcc = show_vcc & "," & shmoo_form & "," & shmoo_stop & "," & shmoo_step & "," & CStr(RangeStepSize) & "]"
 
    theexec.Datalog.WriteComment show_vcc
    'Call Print_power_condition
    'Debug.Print show_vcc
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Decide_LVCC_HVCC")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ShmooMakePseudoData(SetupName As String, step_Start As Long, step_Stop As Long, Step_x As Long)
On Error GoTo errHandler
    Dim pseudo_result_str(200) As String, i As Long, ch As String
    Dim cnt As Long
    cnt = 0
    
    
    'Alg L2H
'    pseudo_result_str(cnt) = "+++++++++++++++": cnt = cnt + 1
'    pseudo_result_str(cnt) = "---------------": cnt = cnt + 1
    pseudo_result_str(cnt) = "++++++++++++---": cnt = cnt + 1
    pseudo_result_str(cnt) = "---++++++++++++": cnt = cnt + 1
    pseudo_result_str(cnt) = "++++++++++++--+": cnt = cnt + 1
    pseudo_result_str(cnt) = "+--++++++++++++": cnt = cnt + 1
    pseudo_result_str(cnt) = "+--+++++++++--+": cnt = cnt + 1
    pseudo_result_str(cnt) = "-------++++++++": cnt = cnt + 1
    pseudo_result_str(cnt) = "++++++---------": cnt = cnt + 1
    pseudo_result_str(cnt) = "--------+++++++": cnt = cnt + 1
    pseudo_result_str(cnt) = "-----+++++-++++": cnt = cnt + 1 'HH
    pseudo_result_str(cnt) = "---+++++-++++--": cnt = cnt + 1
    pseudo_result_str(cnt) = "---+++-++++++--": cnt = cnt + 1 'LH
    pseudo_result_str(cnt) = "-+++-+++++-----": cnt = cnt + 1 'LH
    pseudo_result_str(cnt) = "-+++-+++++-++--": cnt = cnt + 1 'BH
    pseudo_result_str(cnt) = "---++++-+++-+--": cnt = cnt + 1
    pseudo_result_str(cnt) = "---++++--++++--": cnt = cnt + 1
    pseudo_result_str(cnt) = "---+++--+++++--": cnt = cnt + 1
    pseudo_result_str(cnt) = "-++-+++-+++++--": cnt = cnt + 1
    'Alg H2L
    pseudo_result_str(cnt) = "+++++++++++++++": cnt = cnt + 1
    pseudo_result_str(cnt) = "---------------": cnt = cnt + 1
    pseudo_result_str(cnt) = "++++++++++++---": cnt = cnt + 1
    pseudo_result_str(cnt) = "---++++++++++++": cnt = cnt + 1
    pseudo_result_str(cnt) = "+++++++++------": cnt = cnt + 1
    pseudo_result_str(cnt) = "-------++++++++": cnt = cnt + 1
    pseudo_result_str(cnt) = "++++++---------": cnt = cnt + 1
    pseudo_result_str(cnt) = "--------+++++++": cnt = cnt + 1
    pseudo_result_str(cnt) = "-----+++++-++++": cnt = cnt + 1 'HH
    pseudo_result_str(cnt) = "---+++++-++++--": cnt = cnt + 1
    pseudo_result_str(cnt) = "---+++-++++++--": cnt = cnt + 1 'LH
    pseudo_result_str(cnt) = "-+++-+++++-----": cnt = cnt + 1 'LH
    pseudo_result_str(cnt) = "-+++-+++++-++--": cnt = cnt + 1 'BH
    pseudo_result_str(cnt) = "---++++-+++-+--": cnt = cnt + 1
    pseudo_result_str(cnt) = "---++++--++++--": cnt = cnt + 1
    pseudo_result_str(cnt) = "---+++--+++++--": cnt = cnt + 1
    pseudo_result_str(cnt) = "-++-+++-+++++--": cnt = cnt + 1
    
    'Misc
    pseudo_result_str(cnt) = "---+++++++++---": cnt = cnt + 1     'CpuTd_VDD_CPU_LVCConly
    pseudo_result_str(cnt) = "---+++++++++---": cnt = cnt + 1     'CpuTd_VDD_CPU_LVCConly_Over_NV
    pseudo_result_str(cnt) = "+++++++--------": cnt = cnt + 1     'CpuTd_VDD_CPU_HVCConly
    pseudo_result_str(cnt) = "+++++++--------": cnt = cnt + 1     'CpuTd_VDD_CPU_HVCConly_Under_NV
    pseudo_result_str(cnt) = "---+++++++++---": cnt = cnt + 1     'CpuTd_VDD_CPU_High_to_Low
    pseudo_result_str(cnt) = "---+++++++++---": cnt = cnt + 1     'CpuTd_VDD_CPU_Low_to_High
    pseudo_result_str(cnt) = "---+++++++++---": cnt = cnt + 1     'CpuTd_VDD_CPU_T_VDD_GPU_High_to_Low
    pseudo_result_str(cnt) = "---+++++++++---": cnt = cnt + 1     'CpuTd_VDD_CPU_T_VDD_GPU_Low_to_High
    pseudo_result_str(cnt) = "---+++++++++---": cnt = cnt + 1     'CpuTd_VDD_CPU_High_to_Low_CalcStepSize
    pseudo_result_str(cnt) = "---+++++++++---": cnt = cnt + 1     'CpuTd_VDD_CPU_Low_to_High_CalcStepSize
    pseudo_result_str(cnt) = "---+++++++++---": cnt = cnt + 1     'CpuTd_VDD_CPU_High_to_Low__StepSizeNotExact
    pseudo_result_str(cnt) = "---+++++++++---": cnt = cnt + 1     'CpuTd_VDD_CPU_Low_to_High__StepSizeNotExact
    pseudo_result_str(cnt) = "---+++++++++---": cnt = cnt + 1     'CpuTd_VIH_Pins_1p8v_IO
    pseudo_result_str(cnt) = "---+++++++++---": cnt = cnt + 1     'CpuTd_VT_Pins_1p8v_IO
    pseudo_result_str(cnt) = "---+++++++++---": cnt = cnt + 1     'CpuTd_VIH_SWD_TMS2
    pseudo_result_str(cnt) = "---+++++++++---": cnt = cnt + 1     'CpuTd_VT_SWD_TMS2
    pseudo_result_str(cnt) = "---+++++++++---": cnt = cnt + 1     'CpuTd_XI0_Freq_C


    With theexec.DevChar.Results(SetupName).Shmoo

        For i = step_Start To step_Stop Step Step_x
            ch = mid(pseudo_result_str(pseudo_result_index), i + 1, 1)
            Select Case ch
                Case "+":
                    .Points(i).ExecutionResult = tlDevCharResult_Pass
                Case "-":
                    .Points(i).ExecutionResult = tlDevCharResult_Fail
                Case "*": 'assume pass
                    .Points(i).ExecutionResult = tlDevCharResult_AssumedPass
                Case "~": 'assume fail
                    .Points(i).ExecutionResult = tlDevCharResult_AssumedFail
                Case Else:
                    .Points(i).ExecutionResult = tlDevCharResult_Error
            End Select
        Next i
    End With
    pseudo_result_index = pseudo_result_index + 1
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "ShmooMakePseudoData")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function CreateShmooResultString(Shmoo_Result, Shmoo_result_PF As String, SetupName As String, step_Start As Long, step_Stop As Long, Step_x As Long, Optional site As Variant)
On Error GoTo errHandler
    Dim i As Long
    Dim current_point_result As String
    Shmoo_Result = vbNullString: Shmoo_result_PF = vbNullString
    Dim j As Long
    If Step_x > 0 Then
        j = 1
    Else
        j = Len(ShmResult(site))
    End If
    'Always from low value to hi value
    For i = step_Start To step_Stop Step Step_x
        current_point_result = theexec.DevChar.Results(SetupName).Shmoo.Points(i).ExecutionResult
        Select Case current_point_result
            Case tlDevCharResult_Pass:
                    Shmoo_Result = Shmoo_Result & "+": Shmoo_result_PF = Shmoo_result_PF & "P"
            Case tlDevCharResult_Fail:
                    Shmoo_Result = Shmoo_Result & "-": Shmoo_result_PF = Shmoo_result_PF & "F"
            Case tlDevCharResult_NoTest:
                    Shmoo_Result = Shmoo_Result & "_": Shmoo_result_PF = Shmoo_result_PF & "_"
            Case tlDevCharResult_AssumedPass:
                    Shmoo_Result = Shmoo_Result & "*": Shmoo_result_PF = Shmoo_result_PF & "P"
            Case tlDevCharResult_AssumedFail:
                    Shmoo_Result = Shmoo_Result & "~":: Shmoo_result_PF = Shmoo_result_PF & "F"
            Case Else:
                    Shmoo_Result = Shmoo_Result & "?":: Shmoo_result_PF = Shmoo_result_PF & "?"
        End Select
    Next i
    ShmResult(site) = vbNullString
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "CreateShmooResultString")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Disable_Inst_pinname_in_PTR()
On Error GoTo errHandler

    theexec.Datalog.Setup.DatalogSetup.DisableInstanceNameInPTR = True
    theexec.Datalog.Setup.DatalogSetup.DisablePinNameInPTR = True
    theexec.Datalog.ApplySetup

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Disable_Inst_pinname_in_PTR")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Enable_Inst_pinname_in_PTR()
On Error GoTo errHandler

    theexec.Datalog.Setup.DatalogSetup.DisableInstanceNameInPTR = False
    theexec.Datalog.Setup.DatalogSetup.DisablePinNameInPTR = False
    theexec.Datalog.ApplySetup

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Enable_Inst_pinname_in_PTR")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Set_Level_Timing_Spec(Shmoo_Param_Type As String, Shmoo_Param_Name As String, shmoo_pin As String, Shmoo_TimeSets As String, Shmoo_value As Double, Port_name As String)
On Error GoTo errHandler
'Set instrument hardware
    Dim InstName As String
    Dim FRC_pin_name As String, Shmoo_Spec As String

    Select Case Shmoo_Param_Type
        Case "AC Spec", "DC Spec":
            theexec.Overlays.ApplyUniformSpecToHW Shmoo_Param_Name, Shmoo_value
            Shmoo_Spec = Shmoo_Param_Name
        Case "Level":
        '20160925 Force to Ucase
            Select Case UCase(Shmoo_Param_Name)
                Case "VMAIN":
                    InstName = GetInstrument(shmoo_pin, 0)
                    Select Case InstName
                       Case "DC-07"
                            thehdw.DCVI.Pins(shmoo_pin).Voltage = Shmoo_value
                       Case "VHDVS", "HexVS", "VSM", "VS-5A", "VS-800mA"
                            thehdw.DCVS.Pins(shmoo_pin).Voltage.Main.value = Shmoo_value
                       Case Else
                    End Select
                Case "VT":
                   thehdw.Digital.Pins(shmoo_pin).Levels.DriverMode = tlDriverModeVt
                   thehdw.Digital.Pins(shmoo_pin).Levels.value(chVt) = Shmoo_value
                Case "VALT":
                   InstName = GetInstrument(shmoo_pin, 0)
                    Select Case InstName
                       Case "DC-07"
                            thehdw.DCVI.Pins(shmoo_pin).Voltage = Shmoo_value
                       Case "VHDVS", "HexVS", "VSM", "VS-5A", "VS-800mA"
                            thehdw.DCVS.Pins(shmoo_pin).Voltage.Alt.value = Shmoo_value
                       Case Else
                    End Select
                Case "VIH": thehdw.Digital.Pins(shmoo_pin).Levels.value(chVih) = Shmoo_value
                Case "VIL": thehdw.Digital.Pins(shmoo_pin).Levels.value(chVil) = Shmoo_value
                Case "VOH": thehdw.Digital.Pins(shmoo_pin).Levels.value(chVoh) = Shmoo_value
                Case "VOL": thehdw.Digital.Pins(shmoo_pin).Levels.value(chVol) = Shmoo_value
                Case "VID": thehdw.Digital.Pins(shmoo_pin).DifferentialLevels.value(chVid) = Shmoo_value
                Case "VOD": thehdw.Digital.Pins(shmoo_pin).DifferentialLevels.value(chVod) = Shmoo_value
                Case "VICM": thehdw.Digital.Pins(shmoo_pin).DifferentialLevels.value(chVicm) = Shmoo_value
                Case "IOL": thehdw.Digital.Pins(shmoo_pin).Levels.value(chIol) = Shmoo_value
                Case "IOH": thehdw.Digital.Pins(shmoo_pin).Levels.value(chIoh) = Shmoo_value
                Case Else:
                    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Set_Level_Timing_Spec", "Not supported Shmoo Parameter Name: " & Shmoo_Param_Name)
'                    TheExec.ErrorLogMessage "Not supported Shmoo Parameter Name: " & Shmoo_Param_Name
            End Select
            Shmoo_Spec = shmoo_pin & "(" & Shmoo_Param_Name & ")"
        Case "Global Spec":
            If Port_name <> "" Then ' Shmoo pin with value from characterization loop and non-shmoo clock with AC context value
                Dim nWires_ary() As String
                Dim nwp As Variant, all_ports As String, all_pins As String
                Dim port_pa As String, ac_spec_pa As String, pin_pa As String, global_spec_pa As String
                nWires_ary = Split(nWire_Ports_GLB, ",")
                For Each nwp In nWires_ary
                    ' Convert nWires to all_ports and all_pins
                    Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
                    If Port_name Like nwp Then
                        Call VaryFreq(port_pa, Shmoo_value, ac_spec_pa)
                    Else
                        Call VaryFreq(port_pa, theexec.Specs.AC(ac_spec_pa).ContextValue, ac_spec_pa)
                    End If
'                    FreqMeasDebug pin_pa, 0.5, 0.01, 0.1             'Debug to print out freq in datalog
                Next nwp
            Else
                theexec.Overlays.ApplyUniformSpecToHW Shmoo_Param_Name, Shmoo_value
                Shmoo_Spec = Shmoo_Param_Name
            End If
       '20180702 TER add for changeing "Edge"
        Case "Edge":
            Select Case UCase(Shmoo_Param_Name)
                Case "ON": thehdw.Digital.Pins(shmoo_pin).Timing.EdgeTime(Shmoo_TimeSets, chEdgeD0) = Shmoo_value
                Case "DATA": thehdw.Digital.Pins(shmoo_pin).Timing.EdgeTime(Shmoo_TimeSets, chEdgeD1) = Shmoo_value
                Case "RETURN": thehdw.Digital.Pins(shmoo_pin).Timing.EdgeTime(Shmoo_TimeSets, chEdgeD2) = Shmoo_value
                Case "OFF": thehdw.Digital.Pins(shmoo_pin).Timing.EdgeTime(Shmoo_TimeSets, chEdgeD3) = Shmoo_value
                Case "OPEN": thehdw.Digital.Pins(shmoo_pin).Timing.EdgeTime(Shmoo_TimeSets, chEdgeR0) = Shmoo_value
                Case "CLOSE": thehdw.Digital.Pins(shmoo_pin).Timing.EdgeTime(Shmoo_TimeSets, chEdgeR1) = Shmoo_value
                Case Else:
                    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Set_Level_Timing_Spec", "Not supported To Set up Timing set Shmoo Parameter Name: " & Shmoo_Param_Name)
'                    TheExec.ErrorLogMessage "Not supported To Set up Timing set Shmoo Parameter Name: " & Shmoo_Param_Name
                    Exit Function
            End Select
        Case Else:
            Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Set_Level_Timing_Spec", "Not supported Shmoo Parameter Name: " & Shmoo_Param_Type)
'            TheExec.ErrorLogMessage "Not supported Shmoo Parameter Name: " & Shmoo_Param_Type
    End Select
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Set_Level_Timing_Spec")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Shmoo_Set_Current_Point()
On Error GoTo errHandler
' Set up shmoo condition for current shmoo point (including tracking)
' Use Set_Level_Timing_Specto set hardware
    Dim Shmoo_Pin_Str As String
    Dim Shmoo_Tracking_Item As Variant, shmoo_axis As Variant
    Dim DevChar_Setup As String
    Dim Shmoo_Param_Type As String, Shmoo_Param_Name As String, shmoo_pin As String, Shmoo_value As Double, Port_name As String
    Dim Shmoo_Step_Name As String, Shmoo_TimeSets As String
    Dim arg_ary() As String
    Dim site As Variant
    If theexec.DevChar.Setups.IsRunning = False Then
        Shmoo_End = False
    Else
        DevChar_Setup = theexec.DevChar.Setups.ActiveSetupName
       ' If Shmoo_End = True Then Exit Function  ' Prevent from setting  to last shmoo point; set Shmoo_End at the end of   PrintShmooInfo
        If theexec.DevChar.Results(DevChar_Setup).StartTime Like "1/1/0001*" Or theexec.DevChar.Results(DevChar_Setup).StartTime Like "0001/1/1*" Then Exit Function  ' initial run of shmoo, not the first point
        With theexec.DevChar.Setups(DevChar_Setup).Shmoo
            For Each shmoo_axis In .Axes.list
                If LCase(.Axes(shmoo_axis).InterposeFunctions.PrePoint.name) Like "freerunclk_set_xy" Then
                    arg_ary = Split(.Axes(shmoo_axis).InterposeFunctions.PrePoint.Arguments, ",")
                    Port_name = arg_ary(1)
                End If
                Shmoo_Param_Type = .Axes.item(shmoo_axis).Parameter.type
                Shmoo_Param_Name = .Axes.item(shmoo_axis).Parameter.name
                shmoo_pin = .Axes.item(shmoo_axis).ApplyTo.Pins
                Shmoo_TimeSets = .Axes.item(shmoo_axis).ApplyTo.Timesets
                For Each site In theexec.sites
                    Shmoo_value = theexec.DevChar.Results(DevChar_Setup).Shmoo.CurrentPoint.Axes(shmoo_axis).value
                    Set_Level_Timing_Spec Shmoo_Param_Type, Shmoo_Param_Name, shmoo_pin, Shmoo_TimeSets, Shmoo_value, Port_name
                Next site
                With theexec.DevChar.Setups(DevChar_Setup).Shmoo.Axes(shmoo_axis).TrackingParameters
                    For Each Shmoo_Tracking_Item In .list
                            Shmoo_Param_Type = .item(Shmoo_Tracking_Item).type
                            Shmoo_Param_Name = .item(Shmoo_Tracking_Item).name
                            shmoo_pin = .item(Shmoo_Tracking_Item).ApplyTo.Pins
                            Shmoo_TimeSets = .item(Shmoo_Tracking_Item).ApplyTo.Timesets
                            For Each site In theexec.sites
                                Shmoo_value = theexec.DevChar.Results(DevChar_Setup).Shmoo.CurrentPoint.Axes(shmoo_axis).TrackingParameters(Shmoo_Tracking_Item).value
                                Set_Level_Timing_Spec Shmoo_Param_Type, Shmoo_Param_Name, shmoo_pin, Shmoo_TimeSets, Shmoo_value, Port_name
                            Next site
                    Next Shmoo_Tracking_Item
                End With
            Next shmoo_axis
        End With
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Set_Current_Point")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Force_Flow_Shmoo_Condition()
On Error GoTo errHandler
    Dim x As Double, y As Double
    Dim X_axis As String, Y_axis As String
    Dim force_ary As AddIns
    
    Dim DevChar_Setup As String
    Dim shmoo_axis As Variant, Shmoo_Tracking_Item As Variant
    Dim axis_name As Variant, shmoo_val As Double, Shmoo_type As Double, Shmoo_Name As Double, Shmoo_Spec As String
    Dim i As Long, Shmoo_start As Double, shmoo_stop As Double, Shmoo_StepSize As Double, Shmoo_Current_Step As Long, shmoo_step As Long
    Dim X_pt As Double, Y_pt As Double
    Dim Port_name As String
    Dim shmoo_pin As String, Shmoo_TimeSet As String
    Dim arg_ary() As String, axis_type As String
    Dim site As Variant
    Flow_Shmoo_Axis_Count = 0
    Flow_Shmoo_Force_Condition = vbNullString
    Shmoo_setup_str = vbNullString
    Flow_Shmoo_Port_Name = vbNullString
    For Each site In theexec.sites
        DevChar_Setup = theexec.sites.item(site).SiteVariableValue("Flow_Shmoo_DevCharSetup")
        X_pt = theexec.sites.item(site).SiteVariableValue("Flow_Shmoo_X")
        Y_pt = theexec.sites.item(site).SiteVariableValue("Flow_Shmoo_Y")
        Exit For
    Next site
    If DevChar_Setup <> "" Then
        With theexec.DevChar.Setups(DevChar_Setup).Shmoo
            Shmoo_Tracking_Item = -99
            For Each shmoo_axis In .Axes.list
                If (Flow_Shmoo_X_Last_Value <> X_pt _
                    Or Flow_Shmoo_X_Last_Value <> -99) _
                    And (Flow_Shmoo_Y_Last_Value <> Y_pt _
                    Or Flow_Shmoo_Y_Last_Value <> -99) Then
                    Flow_Shmoo_X_Fast = True
                End If
                If LCase(.Axes(shmoo_axis).InterposeFunctions.PrePoint.name) Like "freerunclk_set_xy" Then
                    arg_ary = Split(.Axes(shmoo_axis).InterposeFunctions.PrePoint.Arguments, ",")
                    Port_name = arg_ary(1)
                    Flow_Shmoo_Port_Name = Port_name
                End If
                Select Case shmoo_axis
                    Case tlDevCharShmooAxis_X:
                        axis_type = "X"
                        If Flow_Shmoo_X_Current_Step <= Flow_Shmoo_X_Step _
                            And (Flow_Shmoo_X_Last_Value <> X_pt _
                            Or Flow_Shmoo_X_Last_Value = -99) Then
                            If Flow_Shmoo_X_Last_Value <> X_pt Then
                                If Flow_Shmoo_X_Current_Step = Flow_Shmoo_X_Step Then
                                    Flow_Shmoo_X_Current_Step = 0
                                Else
                                    Flow_Shmoo_X_Current_Step = Flow_Shmoo_X_Current_Step + 1
                                End If
                            End If
                        End If
                        Shmoo_Current_Step = Flow_Shmoo_X_Current_Step
                        shmoo_step = Flow_Shmoo_X_Step
                    Case tlDevCharShmooAxis_Y:
                        axis_type = "Y"
                        If Flow_Shmoo_Y_Current_Step < Flow_Shmoo_Y_Step _
                            And (Flow_Shmoo_Y_Last_Value <> Y_pt _
                            Or Flow_Shmoo_Y_Last_Value = -99) Then
                            If Flow_Shmoo_X_Fast = True Then
                                If Flow_Shmoo_Y_Last_Value <> Y_pt _
                                And Flow_Shmoo_X_Current_Step = 0 Then
                                    Flow_Shmoo_Y_Current_Step = Flow_Shmoo_Y_Current_Step + 1
                                End If
                            Else
                                Flow_Shmoo_Y_Current_Step = Flow_Shmoo_Y_Current_Step + 1
                            End If
                        End If
                        Shmoo_Current_Step = Flow_Shmoo_Y_Current_Step
                        shmoo_step = Flow_Shmoo_Y_Step
                   Case Else
                        theexec.Datalog.WriteComment "the axis isn't X or Y!"
                End Select
            
                If .Axes(shmoo_axis).Parameter.range.StepSize <> Empty Then
                    Shmoo_StepSize = .Axes(shmoo_axis).Parameter.range.StepSize
                Else
                    Shmoo_StepSize = (.Axes(shmoo_axis).Parameter.range.to - .Axes(shmoo_axis).Parameter.range.from) / .Axes(shmoo_axis).Parameter.range.Steps
                End If
                shmoo_val = .Axes(shmoo_axis).Parameter.range.from + Shmoo_Current_Step * .Axes(shmoo_axis).Parameter.range.StepSize
                Set_Level_Timing_Spec .Axes(shmoo_axis).Parameter.type.value, .Axes(shmoo_axis).Parameter.name.value, .Axes(shmoo_axis).ApplyTo.Pins, .Axes(shmoo_axis).ApplyTo.Timesets, shmoo_val, Port_name
                If .Axes(shmoo_axis).ApplyTo.Pins <> "" Then
                    Shmoo_Spec = .Axes(shmoo_axis).Parameter.name.value & "(" & .Axes(shmoo_axis).ApplyTo.Pins & ")"
                Else
                    Shmoo_Spec = .Axes(shmoo_axis).Parameter.name.value
                End If
                If .Axes(shmoo_axis).ApplyTo.Timesets <> "" Then
                    Shmoo_Spec = .Axes(shmoo_axis).Parameter.name.value & "(" & .Axes(shmoo_axis).ApplyTo.Timesets & ")"
                Else
                    Shmoo_Spec = .Axes(shmoo_axis).Parameter.name.value
                End If
                If Shmoo_setup_str = "" Then
                    Shmoo_setup_str = axis_type & ":" & Shmoo_Spec & "=" & shmoo_val & "; "
                Else
                    Shmoo_setup_str = Shmoo_setup_str & axis_type & ":" & Shmoo_Spec & "=" & shmoo_val & "; "
                End If
                For Each Shmoo_Tracking_Item In .Axes(shmoo_axis).TrackingParameters.list
                    shmoo_pin = .Axes(shmoo_axis).ApplyTo.Pins
                    Shmoo_TimeSet = .Axes(shmoo_axis).ApplyTo.Timesets
                    With .Axes.item(Shmoo_Tracking_Item).Parameter
                        Shmoo_StepSize = (.range.to - .range.from) / shmoo_step
                        shmoo_val = .range.from + Shmoo_Current_Step * Shmoo_StepSize
                        Set_Level_Timing_Spec .type.value, .name.value, shmoo_pin, Shmoo_TimeSet, shmoo_val, Port_name
                    End With
                    If shmoo_pin <> "" Then
                        Shmoo_Spec = .Axes(shmoo_axis).Parameter.name.value & "(" & shmoo_pin & ")"
                    Else
                        Shmoo_Spec = .Axes(shmoo_axis).Parameter.name.value
                    End If
                    If Shmoo_TimeSet <> "" Then
                        Shmoo_Spec = .Axes(shmoo_axis).Parameter.name.value & "(" & Shmoo_TimeSet & ")"
                    End If
                    Shmoo_setup_str = axis_type & ":" & Shmoo_Spec & "=" & shmoo_val & "; "
                Next Shmoo_Tracking_Item
            Next shmoo_axis
        End With
        Flow_Shmoo_X_Last_Value = X_pt
        Flow_Shmoo_Y_Last_Value = Y_pt
        FlowShmooString_GLB = CStr(shmoo_val * 1000)
        theexec.Datalog.WriteComment "*********** Shmoo Point   " & Shmoo_setup_str & "    ***********"
    End If
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Force_Flow_Shmoo_Condition")
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Restart_All_Freerun_Clk()
On Error GoTo errHandler
    Dim site As Variant
    For Each site In theexec.sites
        Exit For
    Next

    Dim nWire_port_ary() As String
    Dim nwp As Variant ', all_ports As String, all_pins As String
    Dim port_pa As String, ac_spec_pa As String, pin_pa As String, global_spec_pa As String
'    If nWire_ports = "" Then nWire_ports = nWire_Ports_GLB
    nWire_port_ary = Split(nWire_Ports_GLB, ",")
    For Each nwp In nWire_port_ary
        Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
        Call FreeRunClk_Disable(port_pa)
    Next nwp
    
    For Each nwp In nWire_port_ary
            Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
            Call VaryFreq(port_pa, theexec.Specs.AC(ac_spec_pa).ContextValue, ac_spec_pa)
    Next nwp
      
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Restart_All_Freerun_Clk")
    If AbortTest Then Exit Function Else Resume Next
End Function





Public Function ReStart_FRC(ports As String)
On Error GoTo errHandler
    Call Enable_FRC(ports)

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "ReStart_FRC")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function FreqMeasDebug(FreqMeasPins As String, V_threshold As Double, t_interval As Double, t_wait As Double)
On Error GoTo errHandler
    Dim MeasFreq As New PinListData, i As Long
    Dim PinMeas As New PinList
    PinMeas = FreqMeasPins
    If theexec.DataManager.PinType(FreqMeasPins) Like "Differential" Then
        thehdw.Digital.Pins(FreqMeasPins).DifferentialLevels.value(chVod) = V_threshold
    Else
        thehdw.Digital.Pins(FreqMeasPins).Levels.value(chVoh) = V_threshold
    End If
    Call Freq_MeasFreqSetup(PinMeas, t_interval, VOH)
    Call HardIP_Freq_MeasFreqStart(PinMeas, t_interval, MeasFreq, CStr(t_wait))
    If theexec.DataManager.PinType(FreqMeasPins) Like "Differential" Then
        For i = 0 To MeasFreq.Pins.Count - 1 Step 2
            theexec.Flow.TestLimit resultVal:=MeasFreq.Pins(i), Tname:="Debug"
        Next i
    Else
        theexec.Flow.TestLimit resultVal:=MeasFreq, Tname:="Debug"
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "FreqMeasDebug")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function PostPoint_test_faillog_2D(argc As Long, argv() As String) 'pclinzg plot faillog for 2D shmoo
On Error GoTo errHandler
    Dim SetupName As String
    Dim StepNamex As String
    Dim StepNamey As String
    Dim value As Double
    Dim Freq As Double
    Dim site As Variant
    
    
    If theexec.enableWord("AI_Fail") = True Then
      For Each site In theexec.sites
        SetupName = theexec.DevChar.Setups.ActiveSetupName
        StepNamex = theexec.DevChar.Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).StepName
        value = theexec.DevChar.Results(SetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_X).value
        theexec.Datalog.WriteComment StepNamex & ":" & Format(value, "0.000")
        
        StepNamey = theexec.DevChar.Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).StepName
        If StepNamey <> "" Then
           Freq = theexec.DevChar.Results(SetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_Y).value
           theexec.Datalog.WriteComment StepNamey & ":" & Format(Freq / 1000000, "0.000") & "Mhz"
       End If
        

        
        HardIP_WriteFuncResult
      Next site
    End If
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "PostPoint_test_faillog_2D")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Getforcecondition_VDD(VDD_Force As String, Interpose_PrePat As String)
On Error GoTo errHandler
Dim force_condition_arr() As String
Dim pin_array() As String
Dim FC As Variant
VDD_Force = vbNullString
force_condition_arr = Split(Interpose_PrePat, ";")

For Each FC In force_condition_arr

    pin_array = Split(FC, ":")
    
    If UBound(pin_array) = 2 Then
        If UCase(pin_array(1)) = "V" Then
            If VDD_Force = "" Then
                VDD_Force = pin_array(0)
            Else
                VDD_Force = VDD_Force & "," & pin_array(0)
            End If
        End If
    End If
Next FC

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Getforcecondition_VDD")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Get_Shmoo_Set_Pin(Shmoo_Apply_Pin As String, VDD_Force As String, pin_count As Long)
On Error GoTo errHandler
          
            Dim tmp_Shmoo_Apply_Pin() As String
            Dim pin_list_arry() As String
            Dim Flag_IO As Boolean, Flag_VDD As Boolean
            Dim i As Long
         
           
            If theexec.DevChar.Setups.IsRunning = True Then
    
                Get_Current_Apply_Pin Shmoo_Apply_Pin
                Call theexec.DataManager.DecomposePinList(Shmoo_Apply_Pin, pin_list_arry, pin_count)
                
                Flag_IO = False
                Flag_VDD = False
                
                For i = 0 To pin_count - 1
                    If UCase(theexec.DataManager.PinType(pin_list_arry(i))) = "I/O" Then Flag_IO = True
                    If UCase(theexec.DataManager.PinType(pin_list_arry(i))) = "POWER" Then Flag_VDD = True
                Next i
                If Flag_IO = True And Flag_VDD = True Then
                    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Get_Shmoo_Set_Pin", "Can not  contain both I/O and Power Pin  for  Shmoo apply pin " & Shmoo_Apply_Pin)
'                    TheExec.ErrorLogMessage "Can not  contain both I/O and Power Pin  for  Shmoo apply pin " & Shmoo_Apply_Pin
                Else
                End If
                
                If Flag_IO = True Then
                   If g_Vbump_function = True Then
                      Shmoo_Apply_Pin = vbNullString
                   Else
                      If VDD_Force = "" Then
                         Shmoo_Apply_Pin = "CorePower"
                      Else
                         Shmoo_Apply_Pin = "CorePower, " & VDD_Force
                      End If
                   End If
                ElseIf Flag_VDD = True Then
                    If g_Vbump_function = True Then
                       Shmoo_Apply_Pin = Shmoo_Apply_Pin
                    Else
                       If VDD_Force = "" Then
                          Shmoo_Apply_Pin = Shmoo_Apply_Pin & ",CorePower"
                       Else
                          Shmoo_Apply_Pin = Shmoo_Apply_Pin & ",CorePower, " & VDD_Force
                       End If
                    End If
                End If
                
            Else
                If g_Vbump_function = True Then
                   Shmoo_Apply_Pin = vbNullString
                Else
                   If VDD_Force = "" Then
                      Shmoo_Apply_Pin = "CorePower"
                   Else
                      Shmoo_Apply_Pin = "CorePower, " & VDD_Force
                   End If
                End If
            End If
            
            Call theexec.DataManager.DecomposePinList(Shmoo_Apply_Pin, pin_list_arry, pin_count)
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Get_Shmoo_Set_Pin")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function GetSrcString_fromEMAArray(Pat As String, TestCase As String, ByRef SrcBitsStr As String, ByRef SrcBitsCount As Double, ByRef SrcBitsAry() As String)
On Error GoTo errHandler
    Dim FindTestCase As Boolean
    Dim FindPattern As Boolean
    Dim Index_SrcStock As Double
    Dim Index_TestCase As Double
    If Dic_SrcStockIndex.Exists(Pat) Then
        FindPattern = True
        Index_SrcStock = Dic_SrcStockIndex.item(Pat)
        
        If SrcStock(Index_SrcStock).TestCase_index.Exists(TestCase) Then
            FindTestCase = True
            Index_TestCase = SrcStock(Index_SrcStock).TestCase_index.item(TestCase)
            SrcBitsStr = SrcStock(Index_SrcStock).TestCase(Index_TestCase).DigSrc_BitStr
            SrcBitsCount = SrcStock(Index_SrcStock).TestCase(Index_TestCase).DigSrc_BitCount
            SrcBitsAry() = SrcStock(Index_SrcStock).TestCase(Index_TestCase).DigSrc_BitStrAry
        Else
            FindTestCase = False
        End If
    Else
        FindPattern = False
    End If
    'error message
    If FindPattern = False Then
        theexec.Datalog.WriteComment "Can NOT find Pattern in Control Table" & vbCrLf
        theexec.Datalog.WriteComment "PatternName:" & Pat & vbCrLf
    Else
        If FindTestCase = False Then
            theexec.Datalog.WriteComment "Can NOT find TestCase in control Table" & vbCrLf
            theexec.Datalog.WriteComment "PatternName:" & Pat & vbCrLf
            theexec.Datalog.WriteComment "TestCase:" & TestCase & vbCrLf
        End If
    End If
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "GetSrcString_fromEMAArray")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Parse_EMA_DigSrcInfo()
On Error GoTo errHandler
    Dim ws As Worksheet
    Dim MaxColumn As Double
    Dim MaxRow As Double
    Dim CurColumn As Double
    Dim CurRow As Double
    Dim tempStr As String
    Dim TempCount As Double
    Dim ExistPatCount As Integer
    Dim ExistTestCount As Integer
    Dim CurPatNum As Integer
    Dim CurTestNum As Integer
    Dim sheetName As String
    
    Dim DicKey_TestCase As String
    Dim DicKey_SrcStock As String
    
    Dim TempAry_SheetContent As Variant
    
    
        If DSSCMappingTableIsRead = False Then
            Dic_SrcStockIndex.compareMode = 1
            Dic_SrcStockIndex.RemoveAll
            Application.ScreenUpdating = False
            sheetName = "CZ_DSSCmappingtable_" & UCase(currentJobName)
            Set ws = Sheets(sheetName)
            MaxColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
            MaxRow = ws.UsedRange.Rows.Count
            TempAry_SheetContent = ws.range(ws.Cells(1, 1), ws.Cells(MaxRow, MaxColumn)).value
            ' define array depth of patterns
            ExistPatCount = 0
            For CurColumn = 1 To MaxColumn
                If TempAry_SheetContent(1, CurColumn) Like "Pattern Name" Then
                    ExistPatCount = ExistPatCount + 1
                    ReDim SrcStock(ExistPatCount - 1) As DynamicSrc
                End If
            Next CurColumn
            
            ExistPatCount = 0
            For CurColumn = 1 To MaxColumn
                If UCase(TempAry_SheetContent(1, CurColumn)) Like UCase("Pattern Name") Then
                    If ExistPatCount > 0 Then
                        ReDim Preserve SrcStock(ExistPatCount - 1).TestCase(CurTestNum - 1) As testCondition
                    End If
                    ExistPatCount = ExistPatCount + 1
                    CurTestNum = 0
            
                ElseIf UCase(TempAry_SheetContent(1, CurColumn)) Like UCase("Test*") Then
                    CurTestNum = CurTestNum + 1
                    If CurColumn = MaxColumn Then
                        ReDim Preserve SrcStock(ExistPatCount - 1).TestCase(CurTestNum - 1) As testCondition
                    End If
                Else
                End If
            Next CurColumn
           
            CurPatNum = 0
            CurTestNum = 0
            For CurColumn = 1 To MaxColumn
                If UCase(TempAry_SheetContent(1, CurColumn)) Like UCase("Pattern Name") Then
                    SrcStock(CurPatNum).PatternName = TempAry_SheetContent(2, CurColumn)
                    
                    'add dictionary for index of pattern in SrcStock array swlinza 20190602
                    DicKey_SrcStock = vbNullString
                    DicKey_SrcStock = SrcStock(CurPatNum).PatternName
                    If Not Dic_SrcStockIndex.Exists(DicKey_SrcStock) Then
                        Dic_SrcStockIndex.Add DicKey_SrcStock, CurPatNum
                    Else
                        theexec.Datalog.WriteComment "There are two same patterns in control table"
                        theexec.Datalog.WriteComment "Pattern Name :" & DicKey_SrcStock
                        theexec.Datalog.WriteComment "Pattern# :" & Dic_SrcStockIndex.item(DicKey_SrcStock) & "," & CurPatNum & vbCrLf
                        GoTo errHandler
                    End If
                    CurPatNum = CurPatNum + 1
                    CurTestNum = 0
                    
                ElseIf UCase(TempAry_SheetContent(1, CurColumn)) Like UCase("Test*") Then
                    SrcStock(CurPatNum - 1).TestCase(CurTestNum).ConditionName = TempAry_SheetContent(1, CurColumn)
                    tempStr = vbNullString
                    TempCount = 0
                    
                    ReDim Preserve SrcStock(CurPatNum - 1).TestCase(CurTestNum).DigSrc_BitStrAry(MaxRow) As String
                    For CurRow = 2 To MaxRow
                        If Not TempAry_SheetContent(CurRow, CurColumn) = "" Then
                            tempStr = tempStr & TempAry_SheetContent(CurRow, CurColumn)
                            SrcStock(CurPatNum - 1).TestCase(CurTestNum).DigSrc_BitStrAry(TempCount) = TempAry_SheetContent(CurRow, CurColumn)
                            TempCount = TempCount + 1
                        End If
                    Next CurRow
                    ReDim Preserve SrcStock(CurPatNum - 1).TestCase(CurTestNum).DigSrc_BitStrAry(TempCount - 1) As String
                    SrcStock(CurPatNum - 1).TestCase(CurTestNum).DigSrc_BitStr = tempStr
                    SrcStock(CurPatNum - 1).TestCase(CurTestNum).DigSrc_BitCount = TempCount
                    
                    'add dictionary for index of all test cases ' swlinza 20190602
                    DicKey_TestCase = vbNullString
                    DicKey_TestCase = SrcStock(CurPatNum - 1).TestCase(CurTestNum).ConditionName
                    SrcStock(CurPatNum - 1).TestCase_index.compareMode = 1
                    If Not SrcStock(CurPatNum - 1).TestCase_index.Exists(DicKey_TestCase) Then
                        SrcStock(CurPatNum - 1).TestCase_index.Add DicKey_TestCase, CurTestNum
                        
                    Else
                        theexec.Datalog.WriteComment "There are two same TestCase in Same Pattern"
                        theexec.Datalog.WriteComment "Pattern Name :" & SrcStock(CurPatNum - 1).PatternName
                        theexec.Datalog.WriteComment "TestCase# :" & SrcStock(CurPatNum - 1).TestCase_index(DicKey_TestCase) + 1 & "," & CurTestNum + 1 & vbCrLf
                        GoTo errHandler
                    End If
                    
                    If TempCount <> Len(tempStr) Then
                        theexec.Datalog.WriteComment "Source Bit is NOT single bit" & vbCrLf
                        theexec.Datalog.WriteComment "PatternName :" & SrcStock(CurPatNum - 1).PatternName & vbCrLf
                        theexec.Datalog.WriteComment "TestCase :" & SrcStock(CurPatNum - 1).TestCase(CurTestNum).ConditionName & vbCrLf
                    End If
    
                    CurTestNum = CurTestNum + 1
                Else
                End If
            Next CurColumn

            DSSCMappingTableIsRead = True
            Application.ScreenUpdating = True
        End If
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Parse_EMA_DigSrcInfo")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Shmoo_Set_Retention_Power(Optional Skip_Sweep_Pin As Boolean = False)
On Error GoTo errHandler
    'Modify for force condition "VRET" 20171213
    Dim i As Long
    Dim rn_ary() As String, rn_ary_fv() As String
    Dim Pin_Ary() As String, p_cnt As Long
    Dim skip_pin As String
    Dim Skip_Pin_Dic  As New Dictionary
                    
    If Skip_Sweep_Pin = True Then
        If theexec.DevChar.Setups.IsRunning = True Then Get_Current_Apply_Pin skip_pin
        If skip_pin <> "" Then Create_Pin_Dic skip_pin, Skip_Pin_Dic
    End If
    
    If g_Retention_VDD <> "" Then
        rn_ary = Split(LCase(g_Retention_VDD), ",")
        rn_ary_fv = Split(g_Retention_ForceV, ",")
        For i = 0 To UBound(rn_ary)
            If Skip_Sweep_Pin = True Then ' Do not set retention power for shmoo pin
                If Skip_Pin_Dic.Exists(rn_ary(i)) = False Then thehdw.DCVS.Pins(rn_ary(i)).Voltage.Main.value = CDbl(rn_ary_fv(i))
            Else
                thehdw.DCVS.Pins(rn_ary(i)).Voltage.Main.value = CDbl(rn_ary_fv(i))
            End If
        Next i
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Set_Retention_Power")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Decide_retetntion_power(Retention_V() As SiteDouble, RetPins As PinList)
On Error GoTo errHandler
    'Modify for force condition "VRET" 20171213
    ' Assume that  shmoo pin must be included in g_Retention_VDD
    Dim i As Long
    Dim rn_ary() As String, rn_ary_fv() As String, rn_cnt As Long
    Dim Pin_Ary() As String, p_cnt As Long
    Dim shmoo_pin As String
    Dim Shmoo_pin_Dic  As New Dictionary
    Dim ShmooSweepPower_Dic  As New Dictionary
    Dim site As Variant 'Carter, 20240304
    If theexec.DevChar.Setups.IsRunning = True Then
        Get_Current_Apply_Pin shmoo_pin
        If g_Retention_VDD <> "" Then
            RetPins.value = g_Retention_VDD
        Else
            RetPins.value = shmoo_pin
        End If
    Else
        If g_Retention_VDD <> "" Then
            RetPins.value = g_Retention_VDD
        Else
            RetPins.value = vbNullString
        End If
    End If
    
    If g_Retention_VDD <> "" Then
        rn_ary = Split(LCase(g_Retention_VDD), ",")
        rn_ary_fv = Split(g_Retention_ForceV, ",")
        If theexec.DevChar.Setups.IsRunning = True Then
            Create_Pin_Dic shmoo_pin, Shmoo_pin_Dic
        End If
        For Each site In theexec.sites
             ShmooSweepPower_Dic.RemoveAll
             theexec.DataManager.DecomposePinList Shmoo_Apply_Pin, Pin_Ary, p_cnt
             For i = 0 To UBound(Pin_Ary)
                 ShmooSweepPower_Dic.Add LCase(Pin_Ary(i)), ShmooSweepPower(i)
             Next i
                          
             For i = 0 To UBound(rn_ary)
                 If ShmooSweepPower_Dic.Exists(rn_ary(i)) = True And Shmoo_pin_Dic.Exists(rn_ary(i)) = True Then
                     Retention_V(i) = ShmooSweepPower_Dic.item(rn_ary(i))
                 Else
                     Retention_V(i)(site) = CDbl(rn_ary_fv(i))
                 End If
             Next i
        Next site
    Else
         If theexec.DevChar.Setups.IsRunning = True Then
            For Each site In theexec.sites
                ShmooSweepPower_Dic.RemoveAll
                theexec.DataManager.DecomposePinList Shmoo_Apply_Pin, Pin_Ary, p_cnt
                For i = 0 To UBound(Pin_Ary)
                    ShmooSweepPower_Dic.Add LCase(Pin_Ary(i)), ShmooSweepPower(i)
                Next i
                
                theexec.DataManager.DecomposePinList shmoo_pin, rn_ary, rn_cnt
                For i = 0 To UBound(rn_ary)
                        Retention_V(i) = ShmooSweepPower_Dic.item(LCase(rn_ary(i)))
                Next i
            Next site
        End If
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Decide_retetntion_power")
    If AbortTest Then Exit Function Else Resume Next
End Function





Public Function Shmoo_Save_core_power_per_site_for_Vbump()
On Error GoTo errHandler
    Dim p_ary() As String, p_cnt As Long, i As Long, InstName As String
    
    Set g_ApplyLevelTimingVmain = Nothing
    Set g_ApplyLevelTimingValt = Nothing
    'PowerGRP MOD 210601
    theexec.DataManager.DecomposePinList "All_Power", p_ary, p_cnt
    For i = 0 To p_cnt - 1
        If theexec.DataManager.ChannelType(p_ary(i)) <> "N/C" Then
           g_ApplyLevelTimingVmain.AddPin UCase((p_ary(i)))
           g_ApplyLevelTimingValt.AddPin UCase((p_ary(i)))
           InstName = GetInstrument(p_ary(i), 0)
           Select Case InstName
                  Case "DC-07"
                        g_ApplyLevelTimingVmain.Pins(UCase(p_ary(i))).value = CDbl(Format(thehdw.DCVI.Pins(p_ary(i)).Voltage, "0.000"))
                        g_ApplyLevelTimingValt.Pins(UCase(p_ary(i))).value = CDbl(Format(thehdw.DCVI.Pins(p_ary(i)).Voltage, "0.000"))
                  Case "VHDVS", "HexVS", "VSM", "VS-5A", "VS-800mA"
                        g_ApplyLevelTimingVmain.Pins(UCase(p_ary(i))).value = CDbl(Format(thehdw.DCVS.Pins(p_ary(i)).Voltage.Main.value, "0.000"))
                        g_ApplyLevelTimingValt.Pins(UCase(p_ary(i))).value = CDbl(Format(thehdw.DCVS.Pins(p_ary(i)).Voltage.Alt.value, "0.000"))
                  Case "HSD-U"
                  Case Else
                        Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Save_core_power_per_site_for_Vbump", "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Save_core_power_per_site_for_Vbump")
'                        TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Save_core_power_per_site"
            End Select
        End If
    Next i
   Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Save_core_power_per_site_for_Vbump")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function dynamic_SELSRM_source_bits_bk(SELSRAM_DSSC As String, BlockType As String, Optional AllPattAry As Variant, Optional DynamicSourceBitAry As Variant, Optional testType As Variant) As String
On Error GoTo errHandler
 Dim BitsDef As String: BitsDef = vbNullString
 Dim inst_portion_Ary() As String
 inst_portion_Ary = Split(theexec.DataManager.instancename, "_")
 'If UCase(TheExec.DataManager.instancename) Like "*_CPU_*" Or UCase(TheExec.DataManager.instancename) = "*_C_*" Or UCase(TheExec.DataManager.instancename) = "*CPUMBIST*" Then
    'BitsDef = "VDD_DISP,VDD_GPU,VDD_DCS_DDR,VDD_SOC,VDD_DISP,VDD_GPU,VDD_DCS_DDR,VDD_SOC"
 'End If
 Dim BitsDefArr() As String
 Dim SELSRAMArr() As String
 Dim i As Long
 Dim j As Long
 Dim k As Long
 Dim L As Long
 Dim BitsOrderInfo As New Dictionary
 Dim BitsNum As Long
 Dim BlockTypeNum As Long
 Dim PattIdx As Long: PattIdx = -1
 Dim logicPin As String
 Dim SELSRM As String
 Dim DSSCSelSrmOpposite As Long
 Dim BitValue() As String
 Dim tmpPatStr() As String
 Dim TestBlock As String
 Dim TestFunc As String
 Dim MatchIdx() As Long
 
 ReDim MatchIdx(UBound(DynamicSourceBitAry)) As Long
 'PCLIN init value
 For i = 0 To UBound(DynamicSourceBitAry)
    MatchIdx(i) = -1
 Next i
' PatStringAry = Split(AllPattStr, ";")
'pp_otc_s_in00_sc_cca0_tdf_com_aut_MS003_DM_DSSC:SELSRM
' Pattern loop for TestType ChrisHsu
If BlockType = "" Then
    For i = 0 To UBound(AllPattAry)
        If AllPattAry(i) <> "" Then
            tmpPatStr = Split(AllPattAry(i), "_")
            If UBound(tmpPatStr) >= 4 And UCase(tmpPatStr(3)) Like "*PL*" Then
                If UCase(tmpPatStr(4)) = "SC" Or UCase(tmpPatStr(4)) = "CH" Then
                    testType(i) = "SCAN"
                ElseIf UCase(tmpPatStr(4)) = "BI" Then 'Or UCase(tmpPatStr(6)) = "BST" Then
                    testType(i) = "BIST"
                Else
                    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "dynamic_SELSRM_source_bits", "Please use correct Pattern for define TestType ")
'                    TheExec.ErrorLogMessage "Please use correct Pattern for define TestType "
                End If
            End If
        End If
    Next i
End If

 ReDim SELSRAMArr(Len(SELSRAM_DSSC) - 1)
 BitsOrderInfo.RemoveAll
 For i = 0 To Len(SELSRAM_DSSC) - 1
    SELSRAMArr(i) = CStr(mid(SELSRAM_DSSC, i + 1, 1))
 Next i
 If BlockType <> "" Then
 'Mask for verify 20191125 ChrisHsu
    For i = 0 To UBound(SelsramMapping)
       If UCase(SelsramMapping(i).BlockName) <> "" Then
         If UCase(BlockType) Like "*" & UCase(SelsramMapping(i).BlockName) & "*" Then
            BlockTypeNum = i
            Exit For
         End If
       End If
    Next i
 Else
    For j = 0 To UBound(AllPattAry) 'Add pattern loop 20191126
        If AllPattAry(j) <> "" Then
            For i = 0 To UBound(SelsramMapping)
                If UCase(AllPattAry(j)) Like SelsramMapping(i).Pattern And InStr(UCase(AllPattAry(j)), ":DIGSRC") = 0 And SelsramMapping(i).Pattern <> "*" Then

                    AllPattAry(j) = AllPattAry(j) & ":SELSRM"
                    MatchIdx(j) = i
                   Exit For
                End If

            Next i
        End If
    Next j
 End If
 
 
 If BlockType <> "" Then
 'Mask for verify 20191125 ChrisHsu
    If BlockTypeNum <> -1 Then
        For L = 0 To UBound(SelsramMapping(BlockTypeNum).logic_Pin)
            If L = 0 Then
                BitsDef = UCase(SelsramMapping(BlockTypeNum).logic_Pin(L))
            Else
                BitsDef = BitsDef & "," & UCase(SelsramMapping(BlockTypeNum).logic_Pin(L))
            End If
        Next L
        BitsDefArr = Split(BitsDef, ",")
        BitsNum = UBound(BitsDefArr)
        If UBound(BitsDefArr) <> UBound(SELSRAMArr) Then
            Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "dynamic_SELSRM_source_bits", "Number of bits not match with SELSRAM Char Info ")
'            TheExec.ErrorLogMessage "Number of bits not match with SELSRAM Char Info "
        Else
            For i = 0 To BitsNum
                If Not BitsOrderInfo.Exists(BitsDefArr(i)) Then
                    BitsOrderInfo.Add (BitsDefArr(i)), SELSRAMArr(i)
                Else
                End If
            Next i
        End If
    
      ReDim BitValue(UBound(SelsramMapping(BlockTypeNum).bitCount))
      For i = 0 To UBound(SelsramMapping(BlockTypeNum).bitCount)
         logicPin = SelsramMapping(BlockTypeNum).logic_Pin(i)
         DSSCSelSrmOpposite = SelsramMapping(BlockTypeNum).SelSrm1(i)

         If BitsOrderInfo.Exists(logicPin) = True Then
            SELSRM = BitsOrderInfo(logicPin)
         Else
            If UCase(logicPin) Like "PRESERVED" Then
                SELSRM = SelsramMapping(BlockTypeNum).SelSrm1(i) 'PCLIN modify for PRESERVED bit
            Else
                Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "dynamic_SELSRM_source_bits", "Wrong Logic Pin Name in SELSRM_Mapping_Table")
'                TheExec.ErrorLogMessage "Wrong Logic Pin Name in SELSRM_Mapping_Table"
            End If
         End If

         If UCase(SELSRM) = "1" Then
            If DSSCSelSrmOpposite = 1 Then
               BitValue(i) = 1
            Else
               BitValue(i) = 0
            End If
         ElseIf UCase(SELSRM) = "0" Then
            If DSSCSelSrmOpposite = 1 Then
               BitValue(i) = 0
            Else
               BitValue(i) = 1
            End If
         ElseIf UCase(SELSRM) = "S" Then
               BitValue(i) = "S"
                 Else
                                theexec.ErrorLogMessage "The filled in SELSRM char can't recognize!"
         End If
      Next i
    End If
    For i = 0 To UBound(DynamicSourceBitAry)
        DynamicSourceBitAry(i) = Join(BitValue, vbNullString)
    Next i
    dynamic_SELSRM_source_bits_bk = Join(BitValue, vbNullString)
    
 Else
    For j = 0 To UBound(MatchIdx)
        If MatchIdx(j) <> -1 Then
            For L = 0 To UBound(SelsramMapping(MatchIdx(j)).logic_Pin)
                If L = 0 Then
                    BitsDef = UCase(SelsramMapping(MatchIdx(j)).logic_Pin(L))
                Else
                    BitsDef = BitsDef & "," & UCase(SelsramMapping(MatchIdx(j)).logic_Pin(L))
                End If
            Next L
            BitsDefArr = Split(BitsDef, ",")
            BitsNum = UBound(BitsDefArr)
            If UBound(BitsDefArr) <> UBound(SELSRAMArr) Then
                Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "dynamic_SELSRM_source_bits", "Number of bits not match with SELSRAM Char Info ")
'                TheExec.ErrorLogMessage "Number of bits not match with SELSRAM Char Info "
            Else
                For i = 0 To BitsNum
                    If Not BitsOrderInfo.Exists(BitsDefArr(i)) Then
                        BitsOrderInfo.Add (BitsDefArr(i)), SELSRAMArr(i)
                    Else
                    End If
                Next i
            End If
        
            ReDim BitValue(UBound(SelsramMapping(MatchIdx(j)).bitCount))
            For i = 0 To UBound(SelsramMapping(MatchIdx(j)).bitCount)
                logicPin = SelsramMapping(MatchIdx(j)).logic_Pin(i)
                DSSCSelSrmOpposite = SelsramMapping(MatchIdx(j)).SelSrm1(i)
                If BitsOrderInfo.Exists(logicPin) = True Then
                    SELSRM = BitsOrderInfo(logicPin)
                Else
                    If UCase(logicPin) Like "PRESERVED" Then
                        SELSRM = SelsramMapping(MatchIdx(j)).SelSrm1(i) 'PCLIN modify for PRESERVED bit
                    Else
                        Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "dynamic_SELSRM_source_bits", "Wrong Logic Pin Name in SELSRM_Mapping_Table")
'                        TheExec.ErrorLogMessage "Wrong Logic Pin Name in SELSRM_Mapping_Table"
                    End If
                End If
                  
                If UCase(SELSRM) = "1" Then
                    If DSSCSelSrmOpposite = 1 Then
                       BitValue(i) = 1
                    Else
                       BitValue(i) = 0
                    End If
                ElseIf UCase(SELSRM) = "0" Then
                    If DSSCSelSrmOpposite = 1 Then
                       BitValue(i) = 0
                    Else
                       BitValue(i) = 1
                    End If
                ElseIf UCase(SELSRM) = "S" Then
                       BitValue(i) = "S"
                End If
            Next i
            DynamicSourceBitAry(j) = Join(BitValue, vbNullString)
        Else
            DynamicSourceBitAry(j) = vbNullString
        End If
    Next j
 End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "dynamic_SELSRM_source_bits")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function DecodingRealSourceBit(Source_Bits As String, BlockType As String, Optional DSSCPatName As String) As String
On Error GoTo errHandler
 Dim BitsDef As String: BitsDef = vbNullString
 Dim inst_portion_Ary() As String
 inst_portion_Ary = Split(theexec.DataManager.instancename, "_")
 Dim BitsDefArr() As String
 Dim RailsDecodingInfo As New Dictionary
 Dim idx As Long
 Dim DSSCSelSrmOpposite As Long
 Dim BitsValue As String
 Dim DcodingRailInfo() As String
 Dim logicPin As String
 Dim i As Long
 Dim j As Long
 
 idx = -1
 
 If BlockType <> "" Then
 'Mask for verify 20191125 ChrisHsu
    For i = 0 To UBound(SelsramMapping)
       If UCase(SelsramMapping(i).BlockName) <> "" Then
         If UCase(BlockType) Like "*" & UCase(SelsramMapping(i).BlockName) & "*" Then
            idx = i
            Exit For
         End If
       End If
    Next i
 Else
    For i = 0 To UBound(SelsramMapping)
        If UCase(DSSCPatName) Like UCase(SelsramMapping(i).Pattern) Then
            idx = i
            Exit For
        End If
    Next i
 End If
 
If idx <> -1 Then ' PC modify for PRESERVED Pin
    For j = 0 To UBound(SelsramMapping(idx).logic_Pin)
        If j = 0 Then
            BitsDef = UCase(SelsramMapping(idx).logic_Pin(j))
        Else
            BitsDef = BitsDef & "," & UCase(SelsramMapping(idx).logic_Pin(j))
        End If
    Next j
    BitsDefArr = Split(BitsDef, ",")
    ReDim Preserve BitsDefArr(UBound(BitsDefArr))
    ReDim DcodingRailInfo(UBound(BitsDefArr))
    For i = 0 To Len(Source_Bits) - 1
        logicPin = SelsramMapping(idx).logic_Pin(i)
        DSSCSelSrmOpposite = SelsramMapping(idx).SelSrm1(i)
        BitsValue = CStr(mid(Source_Bits, i + 1, 1))
        If DSSCSelSrmOpposite = 1 Then
           BitsValue = SelsramMapping(idx).alpha(i) & "=" & BitsValue
        ElseIf DSSCSelSrmOpposite = 0 Then
           BitsValue = SelsramMapping(idx).alpha(i) & "=" & InverStr(BitsValue)
        End If
        If Not UCase(logicPin) Like "PRESERVED" Then
          If Not RailsDecodingInfo.Exists(logicPin) = True Then
              RailsDecodingInfo.Add (logicPin), BitsValue
          End If
        End If
    Next i
End If
 For i = 0 To UBound(BitsDefArr)
  DcodingRailInfo(i) = RailsDecodingInfo(BitsDefArr(i))
 Next i
 
 DecodingRealSourceBit = Join(DcodingRailInfo, ",")
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "DecodingRealSourceBit")
    If AbortTest Then Exit Function Else Resume Next
End Function
    Public Function DecomposePattSet(Init1 As Pattern, Init2 As Pattern, Init3 As Pattern, Init4 As Pattern, Init5 As Pattern, Init6 As Pattern, Init7 As Pattern, Init8 As Pattern, Init9 As Pattern, _
                                      Init10 As Pattern, Payload1 As Pattern, Payload2 As Pattern, Payload3 As Pattern, Payload4 As Pattern, Payload5 As Pattern)
On Error GoTo errHandler
    Dim Pat_init1() As String
    Dim Pats_Num As Long
    Dim PatIdx As Integer
    Dim tmpIN As String: tmpIN = vbNullString
    Dim tmpPL As String: tmpPL = vbNullString
    Dim INIArr() As String
    Dim PLLArr() As String

    Dim PL_Start_Idx As Integer
    PL_Start_Idx = 0
    
    If Init1 <> "" Then
        thehdw.Patterns(Init1).ValidatePatlist
        Pat_init1 = theexec.DataManager.Raw.GetPatternsInSet(CStr(Init1), Pats_Num)
        If UBound(Pat_init1) > 0 Then
           For PatIdx = 0 To Pats_Num - 1
            
            
               If UCase(Pat_init1(PatIdx)) Like "*_IN*" And PL_Start_Idx = 0 Then
                  tmpIN = tmpIN & Pat_init1(PatIdx) & ","
               Else 'If UCase(Pat_init1(PatIdx)) Like "*_PL*" Or UCase(Pat_init1(PatIdx)) Like "*_FULP*" Then
                  PL_Start_Idx = PL_Start_Idx + 1
                  tmpPL = tmpPL & Pat_init1(PatIdx) & ","
               End If
           Next PatIdx
           tmpIN = mid(tmpIN, 1, Len(tmpIN) - 1)
           tmpPL = mid(tmpPL, 1, Len(tmpPL) - 1)
           INIArr() = Split(tmpIN, ",")
           PLLArr() = Split(tmpPL, ",")
           ReDim Preserve INIArr(9)
           ReDim Preserve PLLArr(4)
           Init1.value = INIArr(0)
           Init2.value = INIArr(1)
           Init3.value = INIArr(2)
           Init4.value = INIArr(3)
           Init5.value = INIArr(4)
           Init6.value = INIArr(5)
           Init7.value = INIArr(6)
           Init8.value = INIArr(7)
           Init9.value = INIArr(8)
           Init10.value = INIArr(9)
           Payload1.value = PLLArr(0)
           Payload2.value = PLLArr(1)
           Payload3.value = PLLArr(2)
           Payload4.value = PLLArr(3)
           Payload5.value = PLLArr(4)
        End If
    End If
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "DecomposePattSet")
    If AbortTest Then Exit Function Else Resume Next
    End Function


Public Function Decide_Pmode_ForceVoltage(PerforManceMode As String, Power_pins As String, Pmode_Voltage As String)
On Error GoTo errHandler
    Dim p_ary() As String, p_cnt As Long, i As Long, j As Long, k As Long, L As Long
    Dim SelsrmPinAry() As String, SelsrmPin_cnt As Long
    Dim DC_cat As String, Dc_spec_type As String, dc_sel As String
    Dim SP As Variant, t As String
    Dim PinValue As String
    Dim PerformanceModeArr() As String
    Dim SELSRM_PinDict As New Dictionary
    Dim SELSRM_PinCheck As Boolean
    Dim compare_pinvalue As Double
    Dim SelsrmPin As String
    Dim SelsrmPinArr() As String

    If Power_pins = "" Then Exit Function
    SELSRM_PinDict.RemoveAll
    SELSRM_PinCheck = False
    
'===============================================================    init dc spec use bincut_X_X_X, p_mode use dc spec_BI or SC
    theexec.DataManager.GetInstanceContext DC_cat, dc_sel, t, t, t, t, t, t
    
    PerformanceModeArr = Split(PerforManceMode, ":")
    If UBound(PerformanceModeArr) > 0 Then
        If UCase(PerformanceModeArr(1)) Like "LV" Then
            dc_sel = "MIN"
        ElseIf UCase(PerformanceModeArr(1)) Like "NV" Then
            dc_sel = "TYP"
        ElseIf UCase(PerformanceModeArr(1)) Like "HV" Then
            dc_sel = "MAX"
        Else
            theexec.ErrorLogMessage "DC spec symbol - " & PerformanceModeArr(1) & " can't recognize!"
        End If
    End If
'''''    For Each SP In TheExec.specs.DC.Categories(UCase(DC_cat)).SpecList
        For Each SP In theexec.Specs.DC.Categories(UCase(PerformanceModeArr(0))).SpecList
'===============================================================    init dc spec use bincut_X_X_X, p_mode use dc spec_BI or SC
        SP = LCase(SP)
        If SP Like "*_var_c" Then
            Dc_spec_type = "C"
        ElseIf SP Like "*_var_g" Then
            Dc_spec_type = "G"
        ElseIf SP Like "*_var_s" Then
            Dc_spec_type = "S"
        ElseIf SP Like "*_var_h" Then
            Dc_spec_type = "H"
        ElseIf SP Like "*_var_r" Then
            Dc_spec_type = "R"
        '===================================for dc spec BI/ SC
        ElseIf SP Like "*_var_bi" Then
            Dc_spec_type = "BI"
        ElseIf SP Like "*_var_sc" Then
            Dc_spec_type = "SC"
        '===================================for dc spec BI/ SC
        ElseIf SP Like "*_var" Then ''added case for new DC spec sheets method
            Dc_spec_type = vbNullString
        Else
            Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Decide_Pmode_ForceVoltage", "DC spec " & SP & " is not ended with _VAR_C/S/G/H in " & theexec.DataManager.instancename)
'            TheExec.ErrorLogMessage "DC spec " & SP & " is not ended with _VAR_C/S/G/H in " & TheExec.DataManager.instancename
        End If
        Exit For
    Next SP
    

    If UCase(Dc_spec_type) = "C" Or UCase(Dc_spec_type) = "G" Or UCase(Dc_spec_type) = "S" Or UCase(Dc_spec_type) = "H" Or UCase(Dc_spec_type) = "R" Or UCase(Dc_spec_type) = "BI" Or UCase(Dc_spec_type) = "SC" Then ''added case for new DC spec sheets method
       Dc_spec_type = "_" & Dc_spec_type
    End If
    
    
    If Flag_SelsrmMappingTable_Parsed = True Then
       For j = 0 To UBound(SelsramMapping)
        SelsrmPin = SelsrmPin & Join(SelsramMapping(j).logic_Pin, ",") & ","
        SelsrmPin = SelsrmPin & Join(SelsramMapping(j).sram_Pin, ",") & ","
       Next j
       SelsrmPinArr = Split(Replace(LCase(SelsrmPin), "preserved", ","), ",")
       For k = 0 To UBound(SelsrmPinArr)
        If SelsrmPinArr(k) <> "" Then
            theexec.DataManager.DecomposePinList SelsrmPinArr(k), SelsrmPinAry, SelsrmPin_cnt
            For L = 0 To SelsrmPin_cnt - 1
              If Not SELSRM_PinDict.Exists(LCase(SelsrmPinAry(L))) Then
                SELSRM_PinDict.Add LCase(SelsrmPinAry(L)), True
              End If
            Next L
        End If
       Next k
       SELSRM_PinCheck = True
    End If
    
    theexec.DataManager.DecomposePinList Power_pins, p_ary, p_cnt
    For i = 0 To p_cnt - 1
        p_ary(i) = LCase(p_ary(i))
           
        If SELSRM_PinCheck = True Then
           If SELSRM_PinDict.Exists(p_ary(i)) Then
              If UCase(dc_sel) Like "TYP" Then
                 PinValue = Format(CStr(theexec.Specs.DC.item(p_ary(i) & "_VOP" & "_" & "VAR" & Dc_spec_type).Categories(PerformanceModeArr(0)).Typ.value), "0.000")
              ElseIf UCase(dc_sel) Like "MAX" Then
                 PinValue = Format(CStr(theexec.Specs.DC.item(p_ary(i) & "_VOP" & "_" & "VAR" & Dc_spec_type).Categories(PerformanceModeArr(0)).max.value), "0.000")
              ElseIf UCase(dc_sel) Like "MIN" Then
                 PinValue = Format(CStr(theexec.Specs.DC.item(p_ary(i) & "_VOP" & "_" & "VAR" & Dc_spec_type).Categories(PerformanceModeArr(0)).min.value), "0.000")
              End If
              compare_pinvalue = CDbl(PinValue)
           Else
              If UCase(dc_sel) Like "TYP" Then
                 PinValue = Format(CStr(theexec.Specs.DC.item(p_ary(i) & "_" & "VAR" & Dc_spec_type).Categories(PerformanceModeArr(0)).Typ.value), "0.000")
              ElseIf UCase(dc_sel) Like "MAX" Then
                 PinValue = Format(CStr(theexec.Specs.DC.item(p_ary(i) & "_" & "VAR" & Dc_spec_type).Categories(PerformanceModeArr(0)).max.value), "0.000")
              ElseIf UCase(dc_sel) Like "MIN" Then
                 PinValue = Format(CStr(theexec.Specs.DC.item(p_ary(i) & "_" & "VAR" & Dc_spec_type).Categories(PerformanceModeArr(0)).min.value), "0.000")
              End If
              compare_pinvalue = CDbl(PinValue)
           End If
        Else
    ''\\Hard coding "_VOP"\\
            If UCase(dc_sel) Like "TYP" Then
                PinValue = Format(CStr(theexec.Specs.DC.item(p_ary(i) & "_" & "VAR" & Dc_spec_type).Categories(PerformanceModeArr(0)).Typ.value), "0.000")
            ElseIf UCase(dc_sel) Like "MAX" Then
                PinValue = Format(CStr(theexec.Specs.DC.item(p_ary(i) & "_" & "VAR" & Dc_spec_type).Categories(PerformanceModeArr(0)).max.value), "0.000")
            ElseIf UCase(dc_sel) Like "MIN" Then
                PinValue = Format(CStr(theexec.Specs.DC.item(p_ary(i) & "_" & "VAR" & Dc_spec_type).Categories(PerformanceModeArr(0)).min.value), "0.000")
            End If
            compare_pinvalue = CDbl(PinValue)
        End If
        
        If compare_pinvalue <= 0 Then ' to check if there is any abnormal value in Category
            theexec.Datalog.WriteComment UCase(p_ary(i)) & " is less than " & CStr(compare_pinvalue) & " please check the value of Category: " & PerforManceMode
            If isDebugMode Then theexec.AddOutput UCase(p_ary(i)) & " is less than " & CStr(compare_pinvalue) & " please check the value of Category: " & PerforManceMode, vbRed, True
        End If
        Pmode_Voltage = Pmode_Voltage & ";" & UCase(p_ary(i)) & ":" & "V" & ":" & PinValue
    Next i
    Pmode_Voltage = mid(Pmode_Voltage, 2, Len(Pmode_Voltage) - 1)
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Decide_Pmode_ForceVoltage")
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Char_Process_DigString(DigDSSC_BitSize As String, DigDSSC_Seg As String, DigDSSC_DigPin As String, _
                                       ByRef DigCapName() As String, _
                                       ByRef DigSrcPin As String, _
                                       ByRef DigCapPin As String, ByRef DigSrcSize As String, _
                                       ByRef DigCapSize As String, _
                                       ByRef DigCap_Info_Dict As Dictionary, Optional Repeat As Long) As Long
On Error GoTo errHandler
                                   
                                       
        Dim DigDSSC_Seg_Arr_Split() As String
        Dim DigCapEachSgmt_Info() As String
        Dim DigDSSC_BitSize_Arr() As String
        Dim DigDSSC_DigPin_Arr() As String
        Dim DigDSSC_Seg_Arr() As String
        Dim i As Long
    
        DigDSSC_BitSize_Arr = Split(DigDSSC_BitSize, "|")
        If UBound(DigDSSC_BitSize_Arr) = 1 Then
           DigSrcSize = DigDSSC_BitSize_Arr(0)
           DigCapSize = DigDSSC_BitSize_Arr(1)
        ElseIf UBound(DigDSSC_BitSize_Arr) = 0 Then
           DigSrcSize = DigDSSC_BitSize_Arr(0)
           DigCapSize = vbNullString
        Else
        End If
        DigDSSC_DigPin_Arr = Split(DigDSSC_DigPin, "|")
        If UBound(DigDSSC_DigPin_Arr) = 1 Then
           DigSrcPin = DigDSSC_DigPin_Arr(0)
           DigCapPin = DigDSSC_DigPin_Arr(1)
        ElseIf UBound(DigDSSC_DigPin_Arr) = 0 Then
           DigSrcPin = DigDSSC_DigPin_Arr(0)
           DigCapPin = vbNullString
        Else
        End If
    
        DigDSSC_Seg_Arr = Split(DigDSSC_Seg, "|")
        If UBound(DigDSSC_Seg_Arr) = 1 Then
            DigDSSC_Seg_Arr_Split = Split(DigDSSC_Seg_Arr(1), "+")
            DigCap_Info_Dict.RemoveAll
            ReDim DigCapName(UBound(DigDSSC_Seg_Arr_Split))
            
            For i = 0 To UBound(DigDSSC_Seg_Arr_Split)
                DigCapEachSgmt_Info = Split(DigDSSC_Seg_Arr_Split(i), ":")
                DigCap_Info_Dict.Add DigCapEachSgmt_Info(1), CLng(DigCapEachSgmt_Info(0))
                DigCapName(i) = DigCapEachSgmt_Info(1)
            Next i
         Else
'            DigDSSC_Seg_Arr = Split(DigDSSC_Seg_Arr(0), "+")
'            Repeat = UBound(DigDSSC_Seg_Arr) + 1
         End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Char_Process_DigString")
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Char_Process_DSP_Capture(DigCapName() As String, OutDspWave As DSPWave, DigCap_Info_Dict As Dictionary, DigCap_Pin As String) As Long
On Error GoTo errHandler
       Dim Bin2Dec As Double
       Dim OutPutLen As Long
       Dim DSSC_Capture_Out As String
       Dim CaptureName As Variant
       Dim i As Long, StartNum As Long
       Dim CaptureBits As Long
       Dim DSSC_Capture_Out_Dict As New Dictionary
       Dim DSSC_Sgmt_Name_String As String
       Dim OutPut_Sgmt_Name As String
       Dim site As Variant

              DSSC_Capture_Out_Dict.RemoveAll
              For Each site In theexec.sites
                  StartNum = 0
                  For i = 0 To OutDspWave(site).sampleSize - 1
                      DSSC_Capture_Out = DSSC_Capture_Out & CStr(OutDspWave(site).Element(i))
                  Next i
                  DSSC_Capture_Out_Dict.Add site, DSSC_Capture_Out
                  
                  For Each CaptureName In DigCapName
                   If DigCap_Info_Dict.Exists(CaptureName) Then
                      DSSC_Sgmt_Name_String = vbNullString
                      Bin2Dec = 0
                      CaptureBits = CLng(DigCap_Info_Dict.item(CaptureName))
                      For i = StartNum To (StartNum + CaptureBits - 1)
                          DSSC_Sgmt_Name_String = DSSC_Sgmt_Name_String & CStr(OutDspWave(site).Element(i))
                      Next i
                      OutPutLen = Len(DSSC_Sgmt_Name_String) - 1
                      For i = 0 To OutPutLen
                          Bin2Dec = CDbl(Bin2Dec + mid(DSSC_Sgmt_Name_String, (Len(DSSC_Sgmt_Name_String) - i), 1) * 2 ^ i)
                      Next i
                      OutPut_Sgmt_Name = CaptureName & "(OutPutString:" & DSSC_Sgmt_Name_String & ")"
                      If theexec.DevChar.Setups.IsRunning = True Then
                            If theexec.DevChar.Setups.item(theexec.DevChar.Setups.ActiveSetupName).Output.SuspendDatalog = False Then
                              theexec.Flow.TestLimit resultVal:=Bin2Dec, ForceResults:=tlForceNone, Tname:="DigCap" & ":" & OutPut_Sgmt_Name, PinName:=DigCap_Pin, lowVal:=0, hiVal:=2 ^ (OutPutLen + 1) - 1, tNum:=g_TestNum
                            Else
                              theexec.Flow.TestLimit resultVal:=Bin2Dec, ForceResults:=tlForceNone, Tname:="DigCap" & ":" & OutPut_Sgmt_Name, PinName:=DigCap_Pin, lowVal:=0, hiVal:=2 ^ (OutPutLen + 1) - 1
                            End If
                      Else
                            theexec.Flow.TestLimit resultVal:=Bin2Dec, ForceResults:=tlForceNone, Tname:="DigCap" & ":" & OutPut_Sgmt_Name, PinName:=DigCap_Pin, lowVal:=0, hiVal:=2 ^ (OutPutLen + 1) - 1
                      End If
                      StartNum = StartNum + CaptureBits
                    End If
                  Next CaptureName
              Next site
              g_TestNum = g_TestNum + 1 ' 20193021 update
 
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Char_Process_DSP_Capture")
    If AbortTest Then Exit Function Else Resume Next
End Function




Public Function Decide_DC_Level(DC_Level As PinListData, DC_Level_Alt As PinListData, DC_Level_Vmain As PinListData, BlockType As String, Optional testType As String)
On Error GoTo errHandler

''\\Hard coding\\
If BlockType <> "" Then
  If UCase(BlockType) Like "*CPU*" Then
    If UCase(BlockType) Like "*SCAN*" Or UCase(BlockType) Like "*SA*" Or UCase(BlockType) Like "*TD*" Then
       DC_Level = g_ApplyLevelTimingValt.Copy
       'BlockType = "CPUSCAN"
    ElseIf UCase(BlockType) Like "*BST*" Or UCase(BlockType) Like "*BIR*" Or UCase(BlockType) Like "*BIST*" Or UCase(BlockType) Like "*RET*" Or UCase(BlockType) Like "*MBIST*" Or UCase(BlockType) Like "*BISR*" Then
       DC_Level = g_ApplyLevelTimingVmain.Copy
       'BlockType = "CPUMBIST"
    Else
       DC_Level = g_ApplyLevelTimingValt.Copy
    End If
  ElseIf UCase(BlockType) Like "*SOC*" Then
    If UCase(BlockType) Like "*SCAN*" Or UCase(BlockType) Like "*SA*" Or UCase(BlockType) Like "*TD*" Then
       DC_Level = g_ApplyLevelTimingValt.Copy
       'BlockType = "SOCSCAN"
    ElseIf UCase(BlockType) Like "*BST*" Or UCase(BlockType) Like "*BIR*" Or UCase(BlockType) Like "*BIST*" Or UCase(BlockType) Like "*RET*" Or UCase(BlockType) Like "*MBIST*" Or UCase(BlockType) Like "*BISR*" Then
       DC_Level = g_ApplyLevelTimingVmain.Copy
       'BlockType = "SOCMBIST"
    Else
       DC_Level = g_ApplyLevelTimingValt.Copy
    End If
  ElseIf UCase(BlockType) Like "*GPU*" Or UCase(BlockType) Like "*GFX*" Then
    If UCase(BlockType) Like "*SCAN*" Or UCase(BlockType) Like "*SA*" Or UCase(BlockType) Like "*TD*" Then
       DC_Level = g_ApplyLevelTimingValt.Copy
       'BlockType = "GFXSCAN"
    ElseIf UCase(BlockType) Like "*BST*" Or UCase(BlockType) Like "*BIR*" Or UCase(BlockType) Like "*BIST*" Or UCase(BlockType) Like "*RET*" Or UCase(BlockType) Like "*MBIST*" Or UCase(BlockType) Like "*BISR*" Then
       DC_Level = g_ApplyLevelTimingVmain.Copy
       'BlockType = "GFXMBIST"
    Else
       DC_Level = g_ApplyLevelTimingValt.Copy
    End If
  End If
ElseIf BlockType = "" And testType <> "" Then ' Modify for BlockType is Empty 20191202 ChrisHsu
    If UCase(testType) = "SCAN" Then
    
       DC_Level = g_ApplyLevelTimingValt.Copy
    ElseIf UCase(testType) = "BIST" Then
    
       DC_Level = g_ApplyLevelTimingVmain.Copy
    Else
       DC_Level = g_ApplyLevelTimingValt.Copy
    End If
Else
     DC_Level = g_ApplyLevelTimingValt.Copy
End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Decide_DC_Level")
    If AbortTest Then Exit Function Else Resume Next
End Function




'Public Function Decide_Switching_Bit(digSrc_EQ As String, DSPWaveSwitch As DSPWave, DC_Level As PinListData, BlockType As String, SELSRM_Rails As String, Optional shmoo_pin As String, Optional ShmooPinsVoltage As PinListData, Optional ForcePin As String, Optional SetForceVoltage As Dictionary, Optional DSSCPatName As String, Optional set_init As Boolean, Optional seq As Long, Optional Power_Run_Scenario As String) As String
'On Error GoTo errHandler
'
'  Dim p_ary() As String, p_cnt As Long
'  Dim logicPin As String
'  Dim SramPin As String
'  Dim DSSC_Switching_Voltage As New PinListData
'  Dim Sdomain() As Long
'  Dim DSSCSelSrmOpposite As Long
'  Dim BlockTypeNum As Long
'  Dim PattIdx As Long
'  Dim i As Integer, j As Integer, k As Integer
'  Dim ReturnString() As String
'  Dim LogicValue() As Double
'  Dim SramValue As Double
'  BlockTypeNum = -1
'  PattIdx = -1
'  ReDim ReturnString(Len(digSrc_EQ) - 1)
'  Decide_DSSC_Switching_Voltage DSSC_Switching_Voltage, DC_Level, shmoo_pin, ShmooPinsVoltage, ForcePin, SetForceVoltage, set_init, seq, Power_Run_Scenario
'  If BlockType <> "" Then
'  'Mask for verify 20191125 ChrisHsu
'    For i = 0 To UBound(SelsramMapping)
'       If UCase(SelsramMapping(i).BlockName) <> "" Then
'         If UCase(BlockType) Like "*" & UCase(SelsramMapping(i).BlockName) & "*" Then
'            BlockTypeNum = i
'            Exit For
'         End If
'       End If
'    Next i
'  Else
'    For i = 0 To UBound(SelsramMapping)
'        If DSSCPatName Like SelsramMapping(i).Pattern Then
'            PattIdx = i
'            Exit For
'        End If
'    Next i
'  End If
'
'  If BlockTypeNum <> -1 Or PattIdx <> -1 Then
'    For i = 0 To Len(digSrc_EQ) - 1
'        If UCase(CStr(mid(digSrc_EQ, i + 1, 1))) Like "S" Then
'            If BlockType <> "" Then
'                'Mask for verify 20191125 ChrisHsu
'                logicPin = SelsramMapping(BlockTypeNum).logic_Pin(i)
'                SramPin = SelsramMapping(BlockTypeNum).sram_Pin(i)
'                DSSCSelSrmOpposite = SelsramMapping(BlockTypeNum).SelSrm1(i)
'            Else
'                logicPin = SelsramMapping(PattIdx).logic_Pin(i)
'                SramPin = SelsramMapping(PattIdx).sram_Pin(i)
'                DSSCSelSrmOpposite = SelsramMapping(PattIdx).SelSrm1(i)
'            End If
'            For Each site In TheExec.sites.Active 'PC modify for PRESERVE Pin and Pingroup
'
'                If UCase(logicPin) Like "PRESERVED" Then
'                    DSPWaveSwitch.Element(i) = DSSCSelSrmOpposite
'                    ReturnString(i) = DSSCSelSrmOpposite
'                Else
'                    TheExec.DataManager.DecomposePinList logicPin, p_ary, p_cnt
'                    k = 0
'                    For j = 0 To p_cnt - 1
'                        If TheExec.DataManager.ChannelType(p_ary(j)) <> "N/C" Then
'                            ReDim Preserve LogicValue(k)
'                            LogicValue(k) = CDbl(DSSC_Switching_Voltage.Pins(p_ary(j)).value)
'                            k = k + 1
'                        End If
'                    Next j
'                    SramValue = CDbl(DSSC_Switching_Voltage.Pins(SramPin).value)
'                    If DSSCSelSrmOpposite = 0 Then
'                        ReDim Preserve Sdomain(UBound(LogicValue))
'                        For j = 0 To UBound(LogicValue)
'                            If j = 0 Then
'                                Sdomain(j) = IIf((LogicValue(j) > SramValue), 1, 0)
'                            Else
'                                Sdomain(j) = IIf((LogicValue(j) > SramValue), 1, 0)
'                                If Not Sdomain(j) = Sdomain(j - 1) Then
'                                    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Decide_Switching_Bit", "PinGroup with different SELSRM value")
''                                   TheExec.ErrorLogMessage "PinGroup with different SELSRM value"
'                                End If
'                            End If
'                        Next j
'                        DSPWaveSwitch.Element(i) = Sdomain(LBound(Sdomain))
'                        ReturnString(i) = Sdomain(LBound(Sdomain))
'                    ElseIf DSSCSelSrmOpposite = 1 Then
'                        ReDim Preserve Sdomain(UBound(LogicValue))
'                        For j = 0 To UBound(LogicValue)
'                            If j = 0 Then
'                                Sdomain(j) = IIf((LogicValue(j) > SramValue), 0, 1)
'                            Else
'                                Sdomain(j) = IIf((LogicValue(j) > SramValue), 0, 1)
'                                If Not Sdomain(j) = Sdomain(j - 1) Then
'                                    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Decide_Switching_Bit", "PinGroup with different SELSRM value")
''                                    TheExec.ErrorLogMessage "PinGroup with different SELSRM value"
'                                End If
'                            End If
'                        Next j
'                        DSPWaveSwitch.Element(i) = Sdomain(LBound(Sdomain))
'                        ReturnString(i) = Sdomain(LBound(Sdomain))
'                    End If
'                End If
'            Next site
'        Else
'            For Each site In TheExec.sites.Active
'                DSPWaveSwitch.Element(i) = CDbl(mid(digSrc_EQ, i + 1, 1))
'                ReturnString(i) = CDbl(mid(digSrc_EQ, i + 1, 1))
'            Next site
'        End If
'    Next i
'
'    Set PrintSwitchDspWave = Nothing
'    PrintSwitchDspWave = DSPWaveSwitch
'    g_BlockType = BlockType
'
'    'If theexec.DevChar.Setups.IsRunning = False Then
'       Decide_Switching_Bit = Join(ReturnString, vbNullString)
'       SELSRM_Rails = DecodingRealSourceBit(Decide_Switching_Bit, BlockType, DSSCPatName)
'    'Else
'       'Decide_Switching_Bit = digSrc_EQ
'       'SELSRM_Rails = DecodingRealSourceBit(Decide_Switching_Bit, BlockType, DSSCPatName)
'    'End If
'  End If
'  Exit Function
'errHandler:
'    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Decide_Switching_Bit")
'    If AbortTest Then Exit Function Else Resume Next
'End Function


Public Function Decide_DSSC_Switching_Voltage(DSSC_Switching_Voltage As PinListData, DC_Level As PinListData, Optional Shmoo_Apply_Pin As String, Optional ShmooPinsVoltage As PinListData, Optional ForcePin As String, Optional SetForceVoltage As Dictionary, Optional set_init As Boolean, Optional seq As Long, Optional Power_Run_Scenario As String)
On Error GoTo errHandler

        Dim p_ary() As String, p_cnt As Long, i As Long
        Dim InstName As String
        Dim site As Variant
        Dim CorePower_Pins_Dict As New Dictionary

        'Set PrintDSSCSwitchVoltage = Nothing
        DSSC_Switching_Voltage = DC_Level.Copy
        'PowerGRP MOD 210601
        Create_Pin_Dic "All_Power", CorePower_Pins_Dict

        If ForcePin <> "" Then
          theexec.DataManager.DecomposePinList ForcePin, p_ary, p_cnt
             For i = 0 To p_cnt - 1
               If theexec.DataManager.ChannelType(p_ary(i)) <> "N/C" Then
                 If Not CorePower_Pins_Dict.Exists(LCase(p_ary(i))) = True Then
                    DSSC_Switching_Voltage.AddPin (UCase(p_ary(i)))
                    For Each site In theexec.sites
                      DSSC_Switching_Voltage.Pins(UCase(p_ary(i))).value = SetForceVoltage(UCase(p_ary(i)))
                    Next site
                  Else
                    For Each site In theexec.sites
                      DSSC_Switching_Voltage.Pins(UCase(p_ary(i))).value = SetForceVoltage(UCase(p_ary(i)))
                    Next site
                 End If
               End If
             Next i
         End If

        If Shmoo_Apply_Pin <> "" Then
            theexec.DataManager.DecomposePinList Shmoo_Apply_Pin, p_ary, p_cnt
            For i = 0 To p_cnt - 1
                If theexec.DataManager.ChannelType(p_ary(i)) <> "N/C" Then
                    If Not CorePower_Pins_Dict.Exists(LCase(p_ary(i))) = True Then
                       DSSC_Switching_Voltage.AddPin (UCase(p_ary(i)))
                       For Each site In theexec.sites
                         DSSC_Switching_Voltage.Pins(UCase(p_ary(i))).value = ShmooPinsVoltage.Pins(UCase(p_ary(i))).value
                       Next site
                     Else
                       For Each site In theexec.sites
                         DSSC_Switching_Voltage.Pins(UCase(p_ary(i))).value = ShmooPinsVoltage.Pins(UCase(p_ary(i))).value
                       Next site
                    End If
                End If
            Next i
        End If
 
 Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", Decide_DSSC_Switching_Voltage)
    If AbortTest Then Exit Function Else Resume Next
 End Function
Public Function Decide_Switching_Bit_Mod(digSrc_EQ As String, DSPWaveSwitch As DSPWave, DC_Level As PinListData, MatchIdx As Long, SELSRM_Rails As String, Optional shmoo_pin As String, Optional ShmooPinsVoltage As PinListData, Optional ForcePin As String, Optional SetForceVoltage As Dictionary, Optional DSSCPatName As String, Optional set_init As Boolean, Optional seq As Long, Optional Power_Run_Scenario As String) As String
On Error GoTo errHandler

    Dim p_ary() As String, p_cnt As Long
    Dim logicPin As String
    Dim SramPin As String
    Dim DSSC_Switching_Voltage As New PinListData
    Dim Sdomain() As Long
    Dim DSSCSelSrmOpposite As Long
    Dim i As Integer, j As Integer, k As Integer
    Dim ReturnString() As String
    Dim LogicValue() As Double
    Dim SramValue As Double
    Dim site As Variant
    On Error GoTo errHandler
    ReDim ReturnString(Len(digSrc_EQ) - 1)
    
    Decide_DSSC_Switching_Voltage_Mod DSSC_Switching_Voltage, shmoo_pin, ShmooPinsVoltage, set_init, seq

    If MatchIdx <> -1 Then
        For i = 0 To Len(digSrc_EQ) - 1
            If UCase(CStr(mid(digSrc_EQ, i + 1, 1))) Like "S" Then
    
                logicPin = SelsramMapping(MatchIdx).logic_Pin(i)
                SramPin = SelsramMapping(MatchIdx).sram_Pin(i)
                DSSCSelSrmOpposite = SelsramMapping(MatchIdx).SelSrm1(i)
                For Each site In theexec.sites.Active 'PC modify for PRESERVE Pin and Pingroup
                
                    If UCase(logicPin) Like "PRESERVED" Then
                        DSPWaveSwitch.Element(i) = DSSCSelSrmOpposite
                        ReturnString(i) = DSSCSelSrmOpposite
                    Else
                        theexec.DataManager.DecomposePinList logicPin, p_ary, p_cnt
                        k = 0
                        For j = 0 To p_cnt - 1
                            If theexec.DataManager.ChannelType(p_ary(j)) <> "N/C" Then
                                ReDim Preserve LogicValue(k)
                                LogicValue(k) = CDbl(DSSC_Switching_Voltage.Pins(p_ary(j)).value)
                                k = k + 1
                            End If
                        Next j
                        SramValue = CDbl(DSSC_Switching_Voltage.Pins(SramPin).value)
                        If DSSCSelSrmOpposite = 0 Then
                            ReDim Preserve Sdomain(UBound(LogicValue))
                            For j = 0 To UBound(LogicValue)
                                If j = 0 Then
                                    Sdomain(j) = IIf((LogicValue(j) > SramValue), 1, 0)
                                Else
                                    Sdomain(j) = IIf((LogicValue(j) > SramValue), 1, 0)
                                    If Not Sdomain(j) = Sdomain(j - 1) Then
                                       theexec.ErrorLogMessage "PinGroup with different SELSRM value"
                                    End If
                                End If
                            Next j
                            DSPWaveSwitch.Element(i) = Sdomain(LBound(Sdomain))
                            ReturnString(i) = Sdomain(LBound(Sdomain))
                        ElseIf DSSCSelSrmOpposite = 1 Then
                            ReDim Preserve Sdomain(UBound(LogicValue))
                            For j = 0 To UBound(LogicValue)
                                If j = 0 Then
                                    Sdomain(j) = IIf((LogicValue(j) > SramValue), 0, 1)
                                Else
                                    Sdomain(j) = IIf((LogicValue(j) > SramValue), 0, 1)
                                    If Not Sdomain(j) = Sdomain(j - 1) Then
                                        theexec.ErrorLogMessage "PinGroup with different SELSRM value"
                                    End If
                                End If
                            Next j
                            DSPWaveSwitch.Element(i) = Sdomain(LBound(Sdomain))
                            ReturnString(i) = Sdomain(LBound(Sdomain))
                        End If
                    End If
                Next site
            Else
                For Each site In theexec.sites.Active
                    DSPWaveSwitch.Element(i) = CDbl(mid(digSrc_EQ, i + 1, 1))
                    ReturnString(i) = CDbl(mid(digSrc_EQ, i + 1, 1))
                Next site
            End If
        Next i
    
        Set PrintSwitchDspWave = Nothing
        PrintSwitchDspWave = DSPWaveSwitch
        g_BlockType = vbNullString
    
        Decide_Switching_Bit_Mod = Join(ReturnString, vbNullString)
        SELSRM_Rails = DecodingRealSourceBit_Mod(Decide_Switching_Bit_Mod, MatchIdx)

  End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", Decide_Switching_Bit_Mod)
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Set_Run_Level_Vbump(Power_Run_Scenario As String, powerPin As String, set_init As Boolean, seq As Long)
On Error GoTo errHandler
    Dim VoltageLevel As String, Scenario As String
    Dim i As Long
    Dim init_level As String
    Dim pl_level As String
    Dim Power_Run_Scenario_ary() As String
    Dim inst_name As String
    Dim inst_level As String
    
    SweepGuardBand = False
    
    inst_level = right(theexec.DataManager.instancename, 2)
    init_level = "-99"
    pl_level = "-99"
    

    If set_init = True Then
        init_level = "NV"
        If Not g_FirstSetp = True Then
           If Not (Power_Level_Last Like init_level) Then
               If g_shmoo_ret = True Then
                  Shmoo_Restore_Power_per_site_Vbump_NV True
               End If
               'PowerGRP MOD 210601
               thehdw.DCVS.Pins("All_Power").Voltage.Output = tlDCVSVoltageMain
               thehdw.Wait 0.001
               Shmoo_Restore_Power_per_site_Vbump_NV True, True
           End If
        End If
        g_FirstSetp = False
        Power_Level_Last = init_level
        If init_level Like "-99" Then
            Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Set_Run_Level_Vbump", "Power Run Scenario " & Power_Run_Scenario & " is not supported")
'            TheExec.ErrorLogMessage "Power Run Scenario " & Power_Run_Scenario & " is not supported"
        Else
        End If
    Else
        g_PLSWEEP = False
        If LCase(Power_Run_Scenario) Like LCase("*pl_Sweep*") Then
            pl_level = "Sweep"
            If Not (Power_Level_Last Like pl_level) Then Shmoo_Restore_Power_per_site_Vbump powerPin
            g_PLSWEEP = True
        
        ElseIf LCase(Power_Run_Scenario) Like LCase("*pl_NV*") Then
            pl_level = "NV"
            If Not (Power_Level_Last Like pl_level) Then Shmoo_Restore_Power_per_site_Vbump_NV

        ElseIf LCase(Power_Run_Scenario) Like LCase("*pl" & seq & "_Sweep:*") Then
            SweepGuardBand = True
            pl_level = "SweepGuardBand"
            GuardBandCondition Power_Run_Scenario, seq
            If Not (Power_Level_Last Like pl_level) Then Shmoo_Restore_Power_per_site_Vbump powerPin
            g_PLSWEEP = True
            SweepGuardBand = False
        
        
        ElseIf LCase(Power_Run_Scenario) Like LCase("*pl" & seq & "_Sweep*") Then
            pl_level = "Sweep"
            If Not (Power_Level_Last Like pl_level) Then Shmoo_Restore_Power_per_site_Vbump powerPin
            g_PLSWEEP = True

        ElseIf LCase(Power_Run_Scenario) Like LCase("*pl" & seq & "_NV*") Then
            pl_level = "NV"
            If Not (Power_Level_Last Like pl_level) Then Shmoo_Restore_Power_per_site_Vbump_NV
        ElseIf LCase(Power_Run_Scenario) Like LCase("*pl" & seq & "_VRS*") Then
            pl_level = "VRS"
            If Not (Power_Level_Last Like pl_level) Then Shmoo_Restore_Power_per_site_Vbump_NV , , pl_level
        Else
        End If
           
        Power_Level_Last = pl_level
        If pl_level Like "-99" Then
            Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Set_Run_Level_Vbump", "Power Run Scenario " & Power_Run_Scenario & " is not supported")
'            TheExec.ErrorLogMessage "Power Run Scenario " & Power_Run_Scenario & " is not supported"
        Else
        End If
    End If
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Set_Run_Level_Vbump")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Shmoo_Restore_Power_per_site_Vbump_Retention(Shmoo_Apply_Pin As String, RetentionShmoo As Boolean)
On Error GoTo errHandler
    Dim p_ary() As String, p_cnt As Long, i As Long
    Dim InstName As String
    Dim site As Variant
    Dim Shmoo_Apply_Pin_Arry() As String
    Dim SRAMRampUpFirst As New SiteBoolean
    Dim LogicRampdownFirst As New SiteBoolean
    Dim SramShmooPower As String: SramShmooPower = vbNullString
    Dim RetentionShmoo_Pins_Dict As New Dictionary
    Dim Retention_ForceV_Arr() As String
    Retention_ForceV_Arr = Split(g_Retention_ForceV, ",")
    
    If RetentionShmoo = True Then
       If theexec.Flow.enableWord("Enable_RET_RampDownUp") = False Then
          Retention_RampdownUp Shmoo_Apply_Pin, "DOWN"
       Else
           Create_Pin_Dic Shmoo_Apply_Pin, RetentionShmoo_Pins_Dict
           Shmoo_Apply_Pin_Arry = Split(Shmoo_Apply_Pin, ",")
           For i = 0 To UBound(Shmoo_Apply_Pin_Arry)
              If Not UCase(Shmoo_Apply_Pin_Arry(i)) Like UCase(SramShmooPower) Or SramShmooPower = "" Then
                If theexec.DataManager.ChannelType(Shmoo_Apply_Pin_Arry(i)) <> "N/C" Then
                  InstName = GetInstrument(Shmoo_Apply_Pin_Arry(i), 0)
                  Select Case InstName
                         Case "DC-07"
                            Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Restore_Power_per_site_Vbump_Retention", "Instrument " & InstName & " for pin " & Shmoo_Apply_Pin_Arry(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump")
'                              TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & Shmoo_Apply_Pin_Arry(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump"
                         Case "VHDVS", "HexVS", "VSM", "VS-5A", "VS-800mA"
                              For Each site In theexec.sites
                                  thehdw.DCVS.Pins(Shmoo_Apply_Pin_Arry(i)).Voltage.Main.value = g_Globalpointval.Pins(Shmoo_Apply_Pin_Arry(i)).value
                              Next site
                    Case "HSD-U"
                    Case Else
                        Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Restore_Power_per_site_Vbump_Retention", "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump")
'                              TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump"
                  End Select
                End If
              End If
           Next i
       End If
    End If
     
    If g_Retention_VDD <> "" And theexec.Flow.enableWord("Enable_RET_Ramping") = False Then
       theexec.DataManager.DecomposePinList g_Retention_VDD, p_ary, p_cnt
       For i = 0 To p_cnt - 1
           If theexec.DataManager.ChannelType(p_ary(i)) <> "N/C" Then
              If Not RetentionShmoo_Pins_Dict(p_ary(i)) = True Then
              InstName = GetInstrument(p_ary(i), 0)
                   Select Case InstName
                       Case "DC-07"
                            Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Restore_Power_per_site_Vbump_Retention", "Instrument " & InstName & " for pin " & Shmoo_Apply_Pin_Arry(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump_Retention")
'                             TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump_Retention"
                       Case "VHDVS", "HexVS", "VSM", "VS-5A", "VS-800mA"
                            thehdw.DCVS.Pins(p_ary(i)).Voltage.Main.value = Retention_ForceV_Arr(i)
                       Case "HSD-U"
                       Case Else
                            Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Restore_Power_per_site_Vbump_Retention", "Instrument " & InstName & " for pin " & Shmoo_Apply_Pin_Arry(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump_Retention")
'                            TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump_Retention"
                   End Select
              End If
           End If
       Next i
    End If
    'PowerGRP MOD 210601
    thehdw.DCVS.Pins("All_Power").Voltage.Output = tlDCVSVoltageMain
    thehdw.Wait 0.001
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Restore_Power_per_site_Vbump_Retention")
    If AbortTest Then Exit Function Else Resume Next
End Function



Public Function Decide_DSSC_Switching_Voltage_Mod(DSSC_Switching_Voltage As PinListData, Optional Shmoo_Apply_Pin As String, Optional ShmooPinsVoltage As PinListData, Optional set_init As Boolean, Optional seq As Long)
On Error GoTo errHandler
        Dim p_ary() As String, p_cnt As Long, i As Long
        Dim site As Variant
        Dim ForcePower As New Dictionary

        'Set PrintDSSCSwitchVoltage = Nothing

        ' Set 1st PL force voltage to define SwitchingVoltage
        For i = seq To UBound(g_CharPattInfoAry)
                With g_CharPattInfoAry(i)
                        If .IsInitPattern = False And .Pattern <> "" Then
                                DSSC_Switching_Voltage = .ForceVoltage.Copy
                                Exit For
                        End If
                End With
        Next i

        Create_Pin_Dic g_MergeVDD, ForcePower

         If Shmoo_Apply_Pin <> "" Then
          theexec.DataManager.DecomposePinList Shmoo_Apply_Pin, p_ary, p_cnt
                 For i = 0 To p_cnt - 1
                         If ForcePower.Exists(LCase(p_ary(i))) = False Then DSSC_Switching_Voltage.AddPin (UCase(p_ary(i)))
                                For Each site In theexec.sites
                                        DSSC_Switching_Voltage.Pins(UCase(p_ary(i))).value = ShmooPinsVoltage.Pins(UCase(p_ary(i))).value
                                Next site
                 Next i
         End If
         
         'PrintDSSCSwitchVoltage = DSSC_Switching_Voltage.Copy
 
 Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Decide_DSSC_Switching_Voltage")
    If AbortTest Then Exit Function Else Resume Next
 End Function
Public Function Shmoo_Restore_Power_per_site_Vbump_NV(Optional init As Boolean = False, Optional InitAltRecover As Boolean = False, Optional pmode_sel As String)
On Error GoTo errHandler
    Dim p_ary() As String, p_cnt As Long, i As Long
    Dim InstName As String
    Dim site As Variant
    Dim CorePower_Pins_Dict As New Dictionary
    'PowerGRP MOD 210601
    Dim OtherPower As String: OtherPower = "All_Power"
    Dim Payload_Voltage_Vmain As New PinListData
    Dim Payload_Voltage_Valt As New PinListData
    
    Dim DC_cat As String
    Dim dc_sel As String
    Dim pp_ary As String
    Dim t As String
    theexec.DataManager.GetInstanceContext DC_cat, dc_sel, t, t, t, t, t, t ' get current context for VRS
    
     Payload_Voltage_Vmain = g_ApplyLevelTimingVmain.Copy
     Payload_Voltage_Valt = g_ApplyLevelTimingValt.Copy
     'PowerGRP MOD 210601
     Create_Pin_Dic "All_Power", CorePower_Pins_Dict
  
     If g_ForceCond_VDD <> "" And init = False Then
        theexec.DataManager.DecomposePinList g_ForceCond_VDD, p_ary, p_cnt
           For i = 0 To p_cnt - 1
            If g_CharInputString_Voltage_Dict.Exists((UCase(p_ary(i)))) = True And theexec.DataManager.ChannelType(p_ary(i)) <> "N/C" Then
               If Not CorePower_Pins_Dict.Exists(LCase(p_ary(i))) = True Then
                  Payload_Voltage_Vmain.AddPin (UCase(p_ary(i)))
                  Payload_Voltage_Valt.AddPin (UCase(p_ary(i)))
'                  For Each site In TheExec.sites
                    Payload_Voltage_Vmain.Pins(UCase(p_ary(i))).value = g_CharInputString_Voltage_Dict(UCase(p_ary(i)))
                    Payload_Voltage_Valt.Pins(UCase(p_ary(i))).value = g_CharInputString_Voltage_Dict(UCase(p_ary(i)))
'                  Next site
                  OtherPower = OtherPower & "," & p_ary(i)
                Else
'                  For Each site In TheExec.sites
                    Payload_Voltage_Vmain.Pins(UCase(p_ary(i))).value = g_CharInputString_Voltage_Dict(UCase(p_ary(i)))
                    Payload_Voltage_Valt.Pins(UCase(p_ary(i))).value = g_CharInputString_Voltage_Dict(UCase(p_ary(i)))
'                  Next site
               End If
             End If
           Next i
       End If
     '=========================================== Applied voltage to Valt==================================================
      If init = False Or InitAltRecover = True Then
        theexec.DataManager.DecomposePinList OtherPower, p_ary, p_cnt
            For i = 0 To p_cnt - 1
               If theexec.DataManager.ChannelType(p_ary(i)) <> "N/C" Then
                   InstName = GetInstrument(p_ary(i), 0)
                   Select Case InstName
                      Case "DC-07"
                            'Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Restore_Power_per_site_Vbump_NV", "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump_NV")
                            If pmode_sel Like "*VRS*" Then
'                                p_ary(i) = GetUsedPinNameForDCVI(p_ary(i))
                                If UCase(p_ary(i)) Like "*_FT" Then
                                    pp_ary = p_ary(i)
                                    pp_ary = Trim(Replace(pp_ary, "_FT", " "))
                                    thehdw.DCVI.Pins(p_ary(i)).Voltage = theexec.Specs.DC.item(pp_ary & "_VAR").Categories(DC_cat).Selectors.item(dc_sel).ContextValue
                                Else
                                    thehdw.DCVI.Pins(p_ary(i)).Voltage = theexec.Specs.DC.item(p_ary(i) & "_VAR").Categories(DC_cat).Selectors.item(dc_sel).ContextValue
                                End If
                            Else
                                For Each site In theexec.sites
                                    thehdw.DCVI.Pins(p_ary(i)).Voltage = Payload_Voltage_Valt.Pins(p_ary(i)).value(site)
                                Next site
                            End If
                      Case "VHDVS", "HexVS", "VSM", "VS-5A", "VS-800mA"
                            If glb_TesterType = "Jaguar" Then
'                           For Each site In TheExec.sites
                                If pmode_sel Like "*VRS*" Then
                                    If UCase(p_ary(i)) Like "*_FT" Then
                                        pp_ary = p_ary(i)
                                        pp_ary = Trim(Replace(pp_ary, "_FT", " "))
                                        thehdw.DCVS.Pins(p_ary(i)).Voltage.Alt.value = theexec.Specs.DC.item(pp_ary & "_VAR").Categories(DC_cat).Selectors.item(dc_sel).ContextValue
                                    Else
                                        thehdw.DCVS.Pins(p_ary(i)).Voltage.Alt.value = theexec.Specs.DC.item(p_ary(i) & "_VAR").Categories(DC_cat).Selectors.item(dc_sel).ContextValue
                                    End If
                                Else
                                    thehdw.DCVS.Pins(p_ary(i)).Voltage.Alt.value = Payload_Voltage_Valt.Pins(p_ary(i)).value
                                End If
'                           Next site
                            ElseIf glb_TesterType = "UltraFLEXplus" Then
                                If pmode_sel Like "*VRS*" Then
                                    For Each site In theexec.sites
                                        If UCase(p_ary(i)) Like "*_FT" Then
                                            pp_ary = p_ary(i)
                                            pp_ary = Trim(Replace(pp_ary, "_FT", " "))
                                            thehdw.DCVS.Pins(p_ary(i)).Voltage.Alt.value = theexec.Specs.DC.item(pp_ary & "_VAR").Categories(DC_cat).Selectors.item(dc_sel).ContextValue
                                        Else
                                            thehdw.DCVS.Pins(p_ary(i)).Voltage.Alt.value = theexec.Specs.DC.item(p_ary(i) & "_VAR").Categories(DC_cat).Selectors.item(dc_sel).ContextValue
                                        End If
                                    Next site
                                Else
                                    For Each site In theexec.sites
                                        thehdw.DCVS.Pins(p_ary(i)).Voltage.Alt.value = Payload_Voltage_Valt.Pins(p_ary(i)).value
                                    Next site
                                End If
                            End If
                      Case "HSD-U"
                      Case Else
                            Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Restore_Power_per_site_Vbump", "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump_NV")
'                            TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump"
                   End Select
               End If
            Next i
          If InitAltRecover = False Then
          'PowerGRP MOD 210601
             thehdw.DCVS.Pins("All_Power").Voltage.Output = tlDCVSVoltageAlt
             thehdw.Wait 0.001
          End If
       End If
      '=========================================== Store to Vmain voltage which voltage same as Valt ==================================================
        If (init = True And InitAltRecover = False) Or init = False Then
          theexec.DataManager.DecomposePinList OtherPower, p_ary, p_cnt
            For i = 0 To p_cnt - 1
               If theexec.DataManager.ChannelType(p_ary(i)) <> "N/C" Then
                   InstName = GetInstrument(p_ary(i), 0)
                   Select Case InstName
                      Case "DC-07"
'                            Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Restore_Power_per_site_Vbump_NV", "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump_NV")
                          If UCase(p_ary(i)) Like "*_FT" Then
                              pp_ary = p_ary(i)
                              pp_ary = Trim(Replace(pp_ary, "_FT", " "))
                              thehdw.DCVI.Pins(p_ary(i)).Voltage = Payload_Voltage_Vmain.Pins(pp_ary).value
                          Else
                              thehdw.DCVI.Pins(p_ary(i)).Voltage = Payload_Voltage_Vmain.Pins(p_ary(i)).value
                          End If
'                              TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump"
                      Case "VHDVS", "HexVS", "VSM", "VS-5A", "VS-800mA"
                          thehdw.DCVS.Pins(p_ary(i)).Voltage.Main.value = Payload_Voltage_Vmain.Pins(p_ary(i)).value
                      Case "HSD-U"
                      Case Else
                          Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Restore_Power_per_site_Vbump_NV", "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump_NV")
'                            TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump"
                   End Select
               End If
            Next i
         End If
     Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Restore_Power_per_site_Vbump_NV")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Shmoo_Restore_Power_per_site_Vbump(Shmoo_Apply_Pin As String)
On Error GoTo errHandler
    
    Dim p_ary() As String, p_cnt As Long, i As Long
    Dim InstName As String
    Dim site As Variant
    Dim Shmoo_Apply_Pin_Arry() As String
    Dim SRAMRampUpFirst As New SiteBoolean
    Dim LogicRampdownFirst As New SiteBoolean
    Dim SramShmooPower As String: SramShmooPower = vbNullString
    If g_ForceCond_VDD <> "" Then
        theexec.DataManager.DecomposePinList g_ForceCond_VDD, p_ary, p_cnt
        For i = 0 To p_cnt - 1
           If theexec.DataManager.ChannelType(p_ary(i)) <> "N/C" Then
               InstName = GetInstrument(p_ary(i), 0)
               Select Case InstName
                  Case "DC-07"
                        Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Restore_Power_per_site_Vbump", "Instrument " & InstName & " for pin " & Shmoo_Apply_Pin_Arry(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump")
'                        TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump"
                  Case "VHDVS", "HexVS", "VSM", "VS-5A", "VS-800mA"
                        thehdw.DCVS.Pins(p_ary(i)).Voltage.Alt.value = g_CharInputString_Voltage_Dict(UCase(p_ary(i)))
                  Case "HSD-U"
                  Case Else
                        Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Restore_Power_per_site_Vbump", "Instrument " & InstName & " for pin " & Shmoo_Apply_Pin_Arry(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump")
'                        TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump"
               End Select
           End If
        Next i
     End If

    If Shmoo_Apply_Pin <> "" Then
         Shmoo_Apply_Pin_Arry = Split(Shmoo_Apply_Pin, ",")
         For i = 0 To UBound(Shmoo_Apply_Pin_Arry)
            If Not UCase(Shmoo_Apply_Pin_Arry(i)) Like UCase(SramShmooPower) Or SramShmooPower = "" Then
              If theexec.DataManager.ChannelType(Shmoo_Apply_Pin_Arry(i)) <> "N/C" Then
                InstName = GetInstrument(Shmoo_Apply_Pin_Arry(i), 0)
                   Select Case InstName
                      Case "DC-07"
                            Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Restore_Power_per_site_Vbump", "Instrument " & InstName & " for pin " & Shmoo_Apply_Pin_Arry(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump")
'                            TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & Shmoo_Apply_Pin_Arry(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump"
                      Case "VHDVS", "HexVS", "VSM", "VS-5A", "VS-800mA"
                            For Each site In theexec.sites
                                If SweepGuardBand Then
                                    thehdw.DCVS.Pins(Shmoo_Apply_Pin_Arry(i)).Voltage.Alt.value = CalcPower(g_Globalpointval.Pins(Shmoo_Apply_Pin_Arry(i)).value, CalcSymbol, SweepGuardBandVal)
                                Else
                                    thehdw.DCVS.Pins(Shmoo_Apply_Pin_Arry(i)).Voltage.Alt.value = g_Globalpointval.Pins(Shmoo_Apply_Pin_Arry(i)).value
                                End If
                            Next site
                      Case "HSD-U"
                      Case Else
                            Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Restore_Power_per_site_Vbump", "Instrument " & InstName & " for pin " & Shmoo_Apply_Pin_Arry(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump")
'                             TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump"
                   End Select
              End If
            End If
         Next i
     End If
     'PowerGRP MOD 210601
     thehdw.DCVS.Pins("All_Power").Voltage.Output = tlDCVSVoltageAlt
''wait time for Vmain switch to Valt
     thehdw.Wait 0.001
     Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Restore_Power_per_site_Vbump")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Retention_RampdownUp(Shmoo_Apply_Pin As String, RampDirection As String, Optional FailCycleCollect As Boolean = False, Optional HLVCC_PLD As PinListData)
On Error GoTo errHandler
    Dim StartVoltage As New PinListData
    Dim EndVoltage As New PinListData
    Dim p_ary() As String, p_cnt As Long, i As Long, InstName As String
    Dim step As Integer
    Dim StepNum As Integer
    Dim site As Variant 'Carter, 20240304
    StepNum = 9 ' //hard coding//must be odd number and please modify this number by different project
    
    If UCase(RampDirection) = "DOWN" Then
       StartVoltage = g_ApplyLevelTimingValt.Copy
       If FailCycleCollect = True Then
        EndVoltage = HLVCC_PLD.Copy
       Else
        EndVoltage = g_Globalpointval.Copy
       End If
    ElseIf UCase(RampDirection) = "UP" Then
        If FailCycleCollect = True Then
            StartVoltage = HLVCC_PLD.Copy
        Else
            StartVoltage = g_Globalpointval.Copy
        End If
       EndVoltage = g_ApplyLevelTimingValt.Copy
    End If

    theexec.DataManager.DecomposePinList Shmoo_Apply_Pin, p_ary, p_cnt
    
    If UCase(RampDirection) = "DOWN" Then
        For step = 1 To StepNum
            For i = 0 To p_cnt - 1
                If theexec.DataManager.ChannelType(p_ary(i)) <> "N/C" Then
                   If step Mod 2 = 1 Then
                      For Each site In theexec.sites
                          thehdw.DCVS.Pins(p_ary(i)).Voltage.Main.value = StartVoltage.Pins(p_ary(i)).value - ((StartVoltage.Pins(p_ary(i)).value - EndVoltage.Pins(p_ary(i)).value) / StepNum) * step
                      Next site
                   ElseIf step Mod 2 = 0 Then
                      For Each site In theexec.sites
                          thehdw.DCVS.Pins(p_ary(i)).Voltage.Alt.value = StartVoltage.Pins(p_ary(i)).value - ((StartVoltage.Pins(p_ary(i)).value - EndVoltage.Pins(p_ary(i)).value) / StepNum) * step
                      Next site
                   End If
                End If
            Next i
            If step Mod 2 = 1 Then
               thehdw.DCVS.Pins("All_Power").Voltage.Output = tlDCVSVoltageMain
               thehdw.Wait 20 * 0.000001
            ElseIf step Mod 2 = 0 Then
               thehdw.DCVS.Pins("All_Power").Voltage.Output = tlDCVSVoltageAlt
               thehdw.Wait 20 * 0.000001
            End If
        Next step
    ElseIf UCase(RampDirection) = "UP" Then
        For step = 1 To StepNum
            For i = 0 To p_cnt - 1
                If theexec.DataManager.ChannelType(p_ary(i)) <> "N/C" Then
                   If step Mod 2 = 1 Then
                      For Each site In theexec.sites
                          thehdw.DCVS.Pins(p_ary(i)).Voltage.Alt.value = StartVoltage.Pins(p_ary(i)).value - ((StartVoltage.Pins(p_ary(i)).value - EndVoltage.Pins(p_ary(i)).value) / StepNum) * step
                      Next site
                   ElseIf step Mod 2 = 0 Then
                      For Each site In theexec.sites
                          thehdw.DCVS.Pins(p_ary(i)).Voltage.Main.value = StartVoltage.Pins(p_ary(i)).value - ((StartVoltage.Pins(p_ary(i)).value - EndVoltage.Pins(p_ary(i)).value) / StepNum) * step
                      Next site
                   End If
                End If
            Next i
            If step Mod 2 = 1 Then
               thehdw.DCVS.Pins("All_Power").Voltage.Output = tlDCVSVoltageAlt
               thehdw.Wait 20 * 0.000001
            ElseIf step Mod 2 = 0 Then
               thehdw.DCVS.Pins("All_Power").Voltage.Output = tlDCVSVoltageMain
               thehdw.Wait 20 * 0.000001
            End If
        Next step
    End If
   
   Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Retention_RampdownUp")
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function StoreEachPointResult_3D()
On Error GoTo errHandler

    Dim DevSetupName As String
    Dim AxisOrder As Variant
    Dim Execution_result As New SiteVariant
    Dim site As Variant
    ' Add  start point value and step size of each axis/tracking
    Dim RangeStartX As Double, RangeStartY As Double, RangeStartZ As Double, RangeStepSizeX As Double, RangeStepSizeY As Double, RangeStepSizeZ As Double
    Dim RangeTrackStartX() As Double, RangeTrackStartY() As Double, RangeTrackStartZ() As Double, RangeTrackStepSizeX() As Double, RangeTrackStepSizeY() As Double, RangeTrackStepSizeZ() As Double
    

    DevSetupName = theexec.DevChar.Setups.ActiveSetupName
    AxisOrder = theexec.DevChar.Setups(DevSetupName).Shmoo.AxisOrder
    
    
    If X_Tracking_Point <> 0 Then ReDim RangeTrackStartX(X_Tracking_Point - 1)
    If X_Tracking_Point <> 0 Then ReDim RangeTrackStepSizeX(X_Tracking_Point - 1)
    If Y_Tracking_Point <> 0 Then ReDim RangeTrackStartY(Y_Tracking_Point - 1)
    If Y_Tracking_Point <> 0 Then ReDim RangeTrackStepSizeY(Y_Tracking_Point - 1)
    If Z_Tracking_Point <> 0 Then ReDim RangeTrackStartZ(Z_Tracking_Point - 1)
    If Z_Tracking_Point <> 0 Then ReDim RangeTrackStepSizeZ(Z_Tracking_Point - 1)
    
' Get each axis start points/step size --- 20190711
    Call Get_ShmooInfo_2D3D(RangeStartX, RangeStartY, RangeStartZ, RangeStepSizeX, RangeStepSizeY, RangeStepSizeZ, RangeTrackStartX, RangeTrackStartY, RangeTrackStartZ, RangeTrackStepSizeX, RangeTrackStepSizeY, RangeTrackStepSizeZ)


    ' ZXY
    If AxisOrder = tlDevCharAxisOrder_ZXY Then
        For Each site In theexec.sites
            For MaxArrIndex = 0 To MaxArrIndex - 1
                        If Y_Point <= Yaxis_index Then
                            If X_Point <= Xaxis_index Then
                                If Z_Point <= Zaxis_index Then
                                   Execution_result = theexec.DevChar.Results(DevSetupName).Shmoo.Points(X_Point, Y_Point, Z_Point).ExecutionResult
                                   'Modified for calculate each points value -----20190711
                                    Call StoreEachPointsVoltage(MaxArrIndex, X_Point, Y_Point, Z_Point, RangeStartX, RangeStartY, RangeStartZ, RangeStepSizeX, RangeStepSizeY, RangeStepSizeZ, RangeTrackStartX, RangeTrackStartY, RangeTrackStartZ, RangeTrackStepSizeX, RangeTrackStepSizeY, RangeTrackStepSizeZ)
                                    Select Case Execution_result
                                        Case tlDevCharResult_Pass
                                            g_ShmooResult.Axis_CurrPoint(MaxArrIndex).CurrResult = "P"
                                        Case tlDevCharResult_Fail
                                            g_ShmooResult.Axis_CurrPoint(MaxArrIndex).CurrResult = "F"
                                        Case tlDevCharResult_AssumedPass
                                            g_ShmooResult.Axis_CurrPoint(MaxArrIndex).CurrResult = "*"
                                        Case tlDevCharResult_AssumedFail
                                            g_ShmooResult.Axis_CurrPoint(MaxArrIndex).CurrResult = "~"
                                        Case tlDevCharResult_Alarm
                                            g_ShmooResult.Axis_CurrPoint(MaxArrIndex).CurrResult = "A"
                                        Case tlDevCharResult_Error
                                            g_ShmooResult.Axis_CurrPoint(MaxArrIndex).CurrResult = "E"
                                        Case Else
                                            g_ShmooResult.Axis_CurrPoint(MaxArrIndex).CurrResult = "?"
                                    End Select
                                    Z_Point = Z_Point + 1
                                    If Z_Point = Zaxis_index Then
                                        Z_Point = 0
                                        X_Point = X_Point + 1
                                        If X_Point = Xaxis_index Then
                                            X_Point = 0
                                            Y_Point = Y_Point + 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
            Next MaxArrIndex
            Y_Point = 0
        Next site
        
    ' XYZ
    ElseIf AxisOrder = tlDevCharAxisOrder_XYZ Then
        For Each site In theexec.sites
            For MaxArrIndex = 0 To MaxArrIndex - 1
                If Z_Point <= Zaxis_index Then
                    If Y_Point <= Yaxis_index Then
                        If X_Point <= Xaxis_index Then
                           Execution_result = theexec.DevChar.Results(DevSetupName).Shmoo.Points(X_Point, Y_Point, Z_Point).ExecutionResult
                           Call StoreEachPointsVoltage(MaxArrIndex, X_Point, Y_Point, Z_Point, RangeStartX, RangeStartY, RangeStartZ, RangeStepSizeX, RangeStepSizeY, RangeStepSizeZ, RangeTrackStartX, RangeTrackStartY, RangeTrackStartZ, RangeTrackStepSizeX, RangeTrackStepSizeY, RangeTrackStepSizeZ)
                            Select Case Execution_result
                                Case tlDevCharResult_Pass
                                    g_ShmooResult.Axis_CurrPoint(MaxArrIndex).CurrResult = "P"
                                Case tlDevCharResult_Fail
                                    g_ShmooResult.Axis_CurrPoint(MaxArrIndex).CurrResult = "F"
                                Case tlDevCharResult_AssumedPass
                                    g_ShmooResult.Axis_CurrPoint(MaxArrIndex).CurrResult = "*"
                                Case tlDevCharResult_AssumedFail
                                    g_ShmooResult.Axis_CurrPoint(MaxArrIndex).CurrResult = "~"
                                Case tlDevCharResult_Alarm
                                    g_ShmooResult.Axis_CurrPoint(MaxArrIndex).CurrResult = "A"
                                Case tlDevCharResult_Error
                                    g_ShmooResult.Axis_CurrPoint(MaxArrIndex).CurrResult = "E"
                                Case Else
                                    g_ShmooResult.Axis_CurrPoint(MaxArrIndex).CurrResult = "?"
                            End Select
                            X_Point = X_Point + 1
                            If X_Point = Xaxis_index Then
                                X_Point = 0
                                Y_Point = Y_Point + 1
                                If Y_Point = Yaxis_index Then
                                    Y_Point = 0
                                    Z_Point = Z_Point + 1
                                End If
                            End If
                        End If
                    End If
                End If
            Next MaxArrIndex
            Z_Point = 0
        Next site
    Else
        theexec.Datalog.WriteComment "Please choose Axis Order either ZXY or XYZ  "
    End If
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "StoreEachPointResult_3D")
    If AbortTest Then Exit Function Else Resume Next
End Function


'Public Function StoreEachPointResult(argc As Long, argv() As String)
Public Function StoreEachPointResult_2D()
On Error GoTo errHandler
    Dim Suspend_Flag As Boolean
    Dim DevSetupName As String
'    Dim Execution_result As String
    Dim curr_axis As Variant
    Dim AxisOrder As Variant
    Dim Execution_result As New SiteVariant
    Dim site As Variant
    Dim RangeStartX As Double, RangeStartY As Double, RangeStartZ As Double, RangeStepSizeX As Double, RangeStepSizeY As Double, RangeStepSizeZ As Double
    Dim RangeTrackStartX() As Double, RangeTrackStartY() As Double, RangeTrackStartZ() As Double, RangeTrackStepSizeX() As Double, RangeTrackStepSizeY() As Double, RangeTrackStepSizeZ() As Double
    
'    Exit Function
    DevSetupName = theexec.DevChar.Setups.ActiveSetupName
    Suspend_Flag = theexec.DevChar.Setups.item(DevSetupName).Output.SuspendDatalog
    AxisOrder = theexec.DevChar.Setups(DevSetupName).Shmoo.AxisOrder
    
    If X_Tracking_Point <> 0 Then ReDim RangeTrackStartX(X_Tracking_Point - 1)
    If X_Tracking_Point <> 0 Then ReDim RangeTrackStepSizeX(X_Tracking_Point - 1)
    If Y_Tracking_Point <> 0 Then ReDim RangeTrackStartY(Y_Tracking_Point - 1)
    If Y_Tracking_Point <> 0 Then ReDim RangeTrackStepSizeY(Y_Tracking_Point - 1)
    If Z_Tracking_Point <> 0 Then ReDim RangeTrackStartZ(Z_Tracking_Point - 1)
    If Z_Tracking_Point <> 0 Then ReDim RangeTrackStepSizeZ(Z_Tracking_Point - 1)
    Call Get_ShmooInfo_2D3D(RangeStartX, RangeStartY, RangeStartZ, RangeStepSizeX, RangeStepSizeY, RangeStepSizeZ, RangeTrackStartX, RangeTrackStartY, RangeTrackStartZ, RangeTrackStepSizeX, RangeTrackStepSizeY, RangeTrackStepSizeZ)
    
    For Each site In theexec.sites
        For MaxArrIndex = 0 To MaxArrIndex - 1
'            Select Case AxisOrder
'                Case tlDevCharAxisOrder_ZXY
                    If Y_Point <= Yaxis_index Then
                        If X_Point <= Xaxis_index Then
'                            If Z_Point <= Zaxis_index Then
    '                        Debug.Print CStr(X_Point) & "," & CStr(Y_Point) & "," & CStr(Z_Point)
                               Execution_result = theexec.DevChar.Results(DevSetupName).Shmoo.Points(X_Point, Y_Point).ExecutionResult
                               Call StoreEachPointsVoltage(MaxArrIndex, X_Point, Y_Point, -1, RangeStartX, RangeStartY, RangeStartZ, RangeStepSizeX, RangeStepSizeY, RangeStepSizeZ, RangeTrackStartX, RangeTrackStartY, RangeTrackStartZ, RangeTrackStepSizeX, RangeTrackStepSizeY, RangeTrackStepSizeZ)
    '                            g_ShmooResult.Axis_CurrPoint(Count_Point).CurrResult = Execution_result
                                Select Case Execution_result
'''=======================================================================================================
                                    Case tlDevCharResult_Pass
                                        g_ShmooResult.Axis_CurrPoint(MaxArrIndex).CurrResult = "P"
                                    Case tlDevCharResult_Fail
                                        g_ShmooResult.Axis_CurrPoint(MaxArrIndex).CurrResult = "F"
                                    Case tlDevCharResult_AssumedPass
                                        g_ShmooResult.Axis_CurrPoint(MaxArrIndex).CurrResult = "*"
                                    Case tlDevCharResult_AssumedFail
                                        g_ShmooResult.Axis_CurrPoint(MaxArrIndex).CurrResult = "~"
                                    Case tlDevCharResult_Alarm
                                        g_ShmooResult.Axis_CurrPoint(MaxArrIndex).CurrResult = "A"
                                    Case tlDevCharResult_Error
                                        g_ShmooResult.Axis_CurrPoint(MaxArrIndex).CurrResult = "E"
                                    Case Else
                                        g_ShmooResult.Axis_CurrPoint(MaxArrIndex).CurrResult = "?"
'========================================================================================================
                                End Select
                                X_Point = X_Point + 1
                                If X_Point = Xaxis_index Then
                                    X_Point = 0
                                    Y_Point = Y_Point + 1
'                                    If X_Point = Xaxis_index Then
'                                        X_Point = 0
'                                        Y_Point = Y_Point + 1
'                                    End If
                                End If
'                            End If
                        End If
                    End If
'            End Select
        Next MaxArrIndex
        Y_Point = 0
    Next site
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "StoreEachPointResult_2D")
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Print3DShmooInfo_ZXY(argc As Long, argv() As String)
On Error GoTo errHandler
Dim site As Variant
Dim i, j, k As Long
Dim ShmooResult As String
Dim x_val As Variant
Dim y_val As Variant
Dim z_val As Variant
Dim x_val_tracking() As Variant
Dim y_val_tracking() As Variant
Dim z_val_tracking() As Variant
Dim Lvcc As Double
Dim HVCC As Double
Dim InstName As String
Dim curr_axis As Variant
Dim axis_pin(2) As String
Dim X_Axis_Val() As Double
Dim Y_Axis_Val() As Double
Dim Z_Axis_Val() As Double
Dim DevSetupName As String
'Dim TrackingNum As Integer

Dim tmpstr() As String
Dim tmpStr1() As String
Dim tmpStr2() As String
Dim InstName_H As String
Dim InstName_L As String
Dim TmpVal As Integer
Dim X_Axis_TrackingPara() As String
Dim Y_Axis_TrackingPara() As String
Dim Z_Axis_TrackingPara() As String
Dim X_Axis_TrackingParaFrom() As Variant
Dim Y_Axis_TrackingParaFrom() As Variant
Dim Z_Axis_TrackingParaFrom() As Variant
Dim X_Axis_TrackingParaTo() As Variant
Dim Y_Axis_TrackingParaTo() As Variant
Dim Z_Axis_TrackingParaTo() As Variant
Dim Tracking_Item As Variant
Dim iI, jj, kk As Integer

Dim x_TrackingInfo As String
Dim x_stepsize As Variant
Dim x_num As Integer
Dim y_TrackingInfo As String
Dim y_stepsize As Variant
Dim y_num As Integer
Dim z_TrackingInfo As String
Dim z_stepsize As Variant
Dim z_num As Integer
Dim tmp_Tnum As Long, CurrPointNum As Long
Dim X_result As String
Dim Y_result As String
Dim Z_result As String
Dim X_Tname_result As String
Dim Y_Tname_result As String
Dim Z_Tname_result As String
Dim AxisOrder As Variant
Dim xx As Long
Dim tmp_count1 As Long
Dim tmp_x As Long
Dim yy As Long
Dim tmp_count As Long
Dim tmp_y As Long

iI = 0
jj = 0
kk = 0
''Exit Function
DevSetupName = theexec.DevChar.Setups.ActiveSetupName
For Each curr_axis In theexec.DevChar.Setups(DevSetupName).Shmoo.Axes.list
        Select Case curr_axis
            Case 0
                ReDim X_Axis_Val(Xaxis_index - 1)
                If X_Tracking_Point <> 0 Then 'extract X Axis tracking pin, either ApplyToPins or ParameterName
                    ReDim X_Axis_TrackingPara(X_Tracking_Point - 1)
                    ReDim X_Axis_TrackingParaFrom(X_Tracking_Point - 1)
                    ReDim X_Axis_TrackingParaTo(X_Tracking_Point - 1)
                    With theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters
                        For Each Tracking_Item In .list
                            X_Axis_TrackingPara(iI) = Replace((theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).ApplyTo.Pins), "_", vbNullString)
                            X_Axis_TrackingParaFrom(iI) = theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).range.from
                            X_Axis_TrackingParaTo(iI) = theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).range.to
                            If X_Axis_TrackingPara(iI) = "" Then X_Axis_TrackingPara(iI) = Replace((theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).name), "_", vbNullString)
                            iI = iI + 1
                        Next Tracking_Item
                    End With
                    For x_num = iI - 1 To 0 Step -1
                        x_stepsize = Abs(X_Axis_TrackingParaFrom(x_num) - X_Axis_TrackingParaTo(x_num)) / (Xaxis_index - 1)
                        Call ValueResolution(X_Axis_TrackingParaFrom(x_num), X_Axis_TrackingParaTo(x_num), x_stepsize)
                        x_TrackingInfo = X_Axis_TrackingPara(x_num) & "," & X_Axis_TrackingParaFrom(x_num) & "," & X_Axis_TrackingParaTo(x_num) & "," & x_stepsize & ";" & x_TrackingInfo
                    Next x_num
                    theexec.Datalog.WriteComment " X-Asix TraickingPin Info => " & x_TrackingInfo
                End If
            Case 1
                ReDim Y_Axis_Val(Yaxis_index - 1)
                If Y_Tracking_Point <> 0 Then 'extract Y Axis tracking pin, either ApplyToPins or ParameterName
                    ReDim Y_Axis_TrackingPara(Y_Tracking_Point - 1)
                    ReDim Y_Axis_TrackingParaFrom(Y_Tracking_Point - 1)
                    ReDim Y_Axis_TrackingParaTo(Y_Tracking_Point - 1)
                    With theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters
                        For Each Tracking_Item In .list
                            Y_Axis_TrackingPara(jj) = Replace((theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).ApplyTo.Pins), "_", vbNullString)
                            Y_Axis_TrackingParaFrom(jj) = theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).range.from
                            Y_Axis_TrackingParaTo(jj) = theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).range.to
                            If Y_Axis_TrackingPara(jj) = "" Then Y_Axis_TrackingPara(jj) = Replace((theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).name), "_", vbNullString)
                            jj = jj + 1
                        Next Tracking_Item
                    End With
                    For y_num = jj - 1 To 0 Step -1
                        y_stepsize = Abs(Y_Axis_TrackingParaFrom(y_num) - Y_Axis_TrackingParaTo(y_num)) / (Yaxis_index - 1)
                        Call ValueResolution(Y_Axis_TrackingParaFrom(y_num), Y_Axis_TrackingParaTo(y_num), y_stepsize)
                        y_TrackingInfo = Y_Axis_TrackingPara(y_num) & "," & Y_Axis_TrackingParaFrom(y_num) & "," & Y_Axis_TrackingParaTo(y_num) & "," & y_stepsize & ";" & y_TrackingInfo
                    Next y_num
                    theexec.Datalog.WriteComment " Y-Asix TraickingPin Info => " & y_TrackingInfo
                End If
            Case 2
                ReDim Z_Axis_Val(Zaxis_index - 1)
                If Z_Tracking_Point <> 0 Then 'extract Z Axis tracking pin, either ApplyToPins or ParameterName
                    ReDim Z_Axis_TrackingPara(Z_Tracking_Point - 1)
                    ReDim Z_Axis_TrackingParaFrom(Z_Tracking_Point - 1)
                    ReDim Z_Axis_TrackingParaTo(Z_Tracking_Point - 1)
                    With theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters
                        For Each Tracking_Item In .list
                            Z_Axis_TrackingPara(kk) = Replace((theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).ApplyTo.Pins), "_", vbNullString)
                            Z_Axis_TrackingParaFrom(kk) = theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).range.from
                            Z_Axis_TrackingParaTo(kk) = theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).range.to
                            If Z_Axis_TrackingPara(kk) = "" Then Z_Axis_TrackingPara(kk) = Replace((theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).name), "_", vbNullString)
                            kk = kk + 1
                        Next Tracking_Item
                    End With
                    For z_num = kk - 1 To 0 Step -1
                        z_stepsize = Abs(Z_Axis_TrackingParaFrom(z_num) - Z_Axis_TrackingParaTo(z_num)) / (Zaxis_index - 1)
                        Call ValueResolution(Z_Axis_TrackingParaFrom(z_num), Z_Axis_TrackingParaTo(z_num), z_stepsize)
                        z_TrackingInfo = Z_Axis_TrackingPara(z_num) & "," & Z_Axis_TrackingParaFrom(z_num) & "," & Z_Axis_TrackingParaTo(z_num) & "," & z_stepsize & ";" & z_TrackingInfo
                    Next z_num
                    theexec.Datalog.WriteComment " Z-Asix TraickingPin Info => " & z_TrackingInfo
                End If
        End Select
        axis_pin(curr_axis) = Replace((theexec.DevChar.Setups(DevSetupName).Shmoo.Axes((curr_axis)).ApplyTo.Pins), "_", vbNullString)
        If axis_pin(curr_axis) = "" Then
            axis_pin(curr_axis) = Replace((theexec.DevChar.Setups(DevSetupName).Shmoo.Axes((curr_axis)).Parameter.name), "_", vbNullString)
        End If
        
Next curr_axis


If X_Tracking_Point <> 0 Then ReDim x_val_tracking(X_Tracking_Point - 1)
If Y_Tracking_Point <> 0 Then ReDim y_val_tracking(Y_Tracking_Point - 1)
If Z_Tracking_Point <> 0 Then ReDim z_val_tracking(Z_Tracking_Point - 1)



InstName = theexec.DataManager.instancename
tmpstr = Split(InstName, "_")

tmpStr1 = Split(InstName, "_")
If InStr(tmpStr1(0), "L") > 0 Or InStr(tmpStr1(0), "H") > 0 Then
    TmpVal = InStr(tmpStr1(0), "L")
    tmpStr1(0) = mid(tmpStr1(0), 1, TmpVal) '  "DFTL" H or "MCL" H
    InstName_L = Join(tmpStr1, "_")

    tmpStr2 = Split(InstName, "_")
    tmpStr2(0) = mid(tmpStr2(0), 1, TmpVal - 1) & right(tmpStr2(0), 1) '  H
    InstName_H = Join(tmpStr2, "_")
Else
    InstName_L = InstName
    InstName_H = InstName
End If

k = 0
j = 0
ShmooResult = vbNullString


''  Z-axis
tmp_Tnum = g_TestNum
For Each site In theexec.sites
        If (X_dimemsion = False And Y_dimemsion = False And Z_dimemsion = False) Or Z_dimemsion = True Then
            theexec.Datalog.WriteComment "  [ Z-Axis ] "
            For i = 0 To MaxArrIndex - 1
                Z_Axis_Val(j) = g_ShmooResult.Axis_CurrPoint(i).Z_axis(site)
                j = j + 1
                ShmooResult = ShmooResult & g_ShmooResult.Axis_CurrPoint(i).CurrResult(site)
                x_val = g_ShmooResult.Axis_CurrPoint(i).X_axis(site)
                y_val = g_ShmooResult.Axis_CurrPoint(i).Y_axis(site)
                z_val = g_ShmooResult.Axis_CurrPoint(i).Z_axis(site)
                Call ValueResolution(x_val, y_val, z_val)
                CurrPointNum = i
                Call Multi_Axis_ResultName_3D(CurrPointNum, axis_pin, X_result, X_Tname_result, x_val, x_val_tracking, X_Axis_TrackingPara, Y_result, Y_Tname_result, y_val, y_val_tracking, Y_Axis_TrackingPara, Z_result, Z_Tname_result, z_val, z_val_tracking, Z_Axis_TrackingPara, "Z", site)
                If CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site)) = "*" Then g_ShmooResult.Axis_CurrPoint(i).CurrResult(site) = "AP"
                If CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site)) = "~" Then g_ShmooResult.Axis_CurrPoint(i).CurrResult(site) = "AF"
                
                If i = (Zaxis_index - 1) + k Then
                    Call ShmooResultPF(ShmooResult, Lvcc, HVCC, Z_Axis_Val)
                    Select Case ShmooResult
                        Case "9999"
                            ' still print out the latest point
                            theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                         
                            If UCase(right(tmpstr(0), 2)) = "LH" Then
                            
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-9999, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=9999, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-9999, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=9999, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=9999, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=9999, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-9999, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                     g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-9999, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            Else
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-9999, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=9999, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-9999, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=9999, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                        Case "7777"
                            ' still print out the latest point
                            theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                            If UCase(right(tmpstr(0), 2)) = "LH" Then
                                
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-7777, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=7777, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-7777, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=7777, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=7777, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=7777, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-7777, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-7777, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            Else
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-7777, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=7777, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-7777, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=7777, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                        Case "5555"
                            ' still print out the latest point
                            theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                            If UCase(right(tmpstr(0), 2)) = "LH" Then
                                
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-5555, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=5555, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-5555, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=5555, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=5555, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=5555, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-5555, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                     g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-5555, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            Else
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-5555, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=5555, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-5555, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=5555, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                        Case Else ' search LVCC/HVCC point
                            ' still print out the latest point
                            theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                            If UCase(right(tmpstr(0), 2)) = "LH" Then
                            
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                     g_TestNum = g_TestNum + 1
                                   theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                     g_TestNum = g_TestNum + 1
                               End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                     g_TestNum = g_TestNum + 1
                               End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            Else
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                     g_TestNum = g_TestNum + 1
                                   theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                     g_TestNum = g_TestNum + 1
                               End If
                            End If
                    End Select
                    ShmooResult = vbNullString
                    k = k + Zaxis_index
                    j = 0
                Else
                    theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                End If
                
            Next i
            k = 0 ' initial it fot multi site
        End If
    ''  =============================================================================================================================================
    ''  X-axis
        If X_dimemsion = True Then
            theexec.Datalog.WriteComment "  [ X-Axis ] "
            'xx = 2
            tmp_count1 = 0
            For xx = 1 To Zaxis_index
                For i = tmp_count1 To MaxArrIndex * Zaxis_index - 1 Step Zaxis_index
                    If i > xx * MaxArrIndex - 1 Then
                        tmp_count1 = i
                        Exit For
                    Else
                        tmp_x = i - (xx - 1) * (MaxArrIndex - 1)
                        ShmooResult = ShmooResult & g_ShmooResult.Axis_CurrPoint(tmp_x).CurrResult(site)
                        X_Axis_Val(j) = g_ShmooResult.Axis_CurrPoint(tmp_x).X_axis(site)
                        j = j + 1
                        
                        x_val = g_ShmooResult.Axis_CurrPoint(tmp_x).X_axis(site)
                        y_val = g_ShmooResult.Axis_CurrPoint(tmp_x).Y_axis(site)
                        z_val = g_ShmooResult.Axis_CurrPoint(tmp_x).Z_axis(site)
                        Call ValueResolution(x_val, y_val, z_val)
                        CurrPointNum = tmp_x
                        Call Multi_Axis_ResultName_3D(CurrPointNum, axis_pin, X_result, X_Tname_result, x_val, x_val_tracking, X_Axis_TrackingPara, Y_result, Y_Tname_result, y_val, y_val_tracking, Y_Axis_TrackingPara, Z_result, Z_Tname_result, z_val, z_val_tracking, Z_Axis_TrackingPara, "X", site)
                        If CStr(g_ShmooResult.Axis_CurrPoint(tmp_x).CurrResult(site)) = "*" Then g_ShmooResult.Axis_CurrPoint(tmp_x).CurrResult(site) = "AP"
                        If CStr(g_ShmooResult.Axis_CurrPoint(tmp_x).CurrResult(site)) = "~" Then g_ShmooResult.Axis_CurrPoint(tmp_x).CurrResult(site) = "AF"
                    End If
    
            'Debug.Print tmp_x
                    If j = Xaxis_index Then
                        
                        Call ShmooResultPF(ShmooResult, Lvcc, HVCC, X_Axis_Val)
                        Select Case ShmooResult
                            Case "9999"
                                ' still print out the latest point
                                If i > MaxArrIndex - 1 Then
                                    theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_x).CurrResult(site))
                                Else
                                    theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                                End If
            
                                If UCase(right(tmpstr(0), 2)) = "LH" Then
            
                                    If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=-9999, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=9999, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=-9999, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=9999, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    End If
                                ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                                    If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=9999, hiVal:=Z_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=9999, hiVal:=Z_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    End If
                                ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                                    If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=-9999, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=-9999, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    End If
                                Else
                                    If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=-9999, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=9999, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=-9999, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=9999, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    End If
                                End If
                            Case "7777"
                                ' still print out the latest point
                                If i > MaxArrIndex - 1 Then
                                    theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_x).CurrResult(site))
                                Else
                                    theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                                End If
                                If UCase(right(tmpstr(0), 2)) = "LH" Then
            
                                    If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=-7777, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=7777, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=-7777, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=7777, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    End If
                                ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                                    If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=7777, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=7777, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    End If
                                ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                                    If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=-7777, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=-7777, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    End If
                                Else
                                    If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=-7777, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=7777, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=-7777, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=7777, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    End If
                                End If
                            Case "5555"
                                ' still print out the latest point
                                If i > MaxArrIndex - 1 Then
                                    theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_x).CurrResult(site))
                                Else
                                    theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                                End If
                                If UCase(right(tmpstr(0), 2)) = "LH" Then
            
                                    If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=-5555, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=5555, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=-5555, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=5555, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    End If
                                ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                                    If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=5555, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=5555, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    End If
                                ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                                    If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=-5555, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=-5555, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    End If
                                Else
                                    If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=-5555, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=5555, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=-5555, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=5555, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    End If
                                End If
                            Case Else ' search LVCC/HVCC point
                                ' still print out the latest point
                                If i > MaxArrIndex - 1 Then
                                    theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_x).CurrResult(site))
                                Else
                                    theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                                End If
                                If UCase(right(tmpstr(0), 2)) = "LH" Then
            
                                    If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    End If
                                ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                                    If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                   End If
                                ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                                    If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                   End If
                                Else
                                   If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    End If
                                End If
                        End Select
                        ShmooResult = vbNullString
                        j = 0
                    Else
                        If i > MaxArrIndex - 1 Then
                            theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_x).CurrResult(site))
                        Else
                            theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                        End If
                    End If
            
                Next i
            '    i = tmp_count1
            Next xx
            k = 0 ' initial it fot multi site
        End If
    '''''  =============================================================================================================================================
    '''''  Y-axis
        If Y_dimemsion = True Then
            theexec.Datalog.WriteComment "  [ Y-Axis ] "
            'yy = 2
            tmp_count = 0
            For yy = 1 To Zaxis_index * Xaxis_index
                For i = tmp_count To (MaxArrIndex * Zaxis_index * Xaxis_index) - 1 Step (Zaxis_index * Xaxis_index)
            
                    If i > yy * MaxArrIndex - 1 Then
                        tmp_count = i
                        Exit For
                    Else
                        tmp_y = i - (yy - 1) * (MaxArrIndex - 1)
                        ShmooResult = ShmooResult & g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site)
                        Y_Axis_Val(j) = g_ShmooResult.Axis_CurrPoint(tmp_y).Y_axis(site)
                        j = j + 1
                        x_val = g_ShmooResult.Axis_CurrPoint(tmp_y).X_axis(site)
                        y_val = g_ShmooResult.Axis_CurrPoint(tmp_y).Y_axis(site)
                        z_val = g_ShmooResult.Axis_CurrPoint(tmp_y).Z_axis(site)
                        Call ValueResolution(x_val, y_val, z_val)
                        CurrPointNum = tmp_y
                        Call Multi_Axis_ResultName_3D(CurrPointNum, axis_pin, X_result, X_Tname_result, x_val, x_val_tracking, X_Axis_TrackingPara, Y_result, Y_Tname_result, y_val, y_val_tracking, Y_Axis_TrackingPara, Z_result, Z_Tname_result, z_val, z_val_tracking, Z_Axis_TrackingPara, "Y", site)
                        If CStr(g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site)) = "*" Then g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site) = "AP"
                        If CStr(g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site)) = "~" Then g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site) = "AF"
                    End If
    
            'Debug.Print tmp_y
                    If j = Yaxis_index Then
                        Call ShmooResultPF(ShmooResult, Lvcc, HVCC, Y_Axis_Val)
                        Select Case ShmooResult
                            Case "9999"
                                ' still print out the latest point
                                If i > MaxArrIndex - 1 Then
                                    theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site))
                                Else
                                    theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                                End If
            
                                If UCase(right(tmpstr(0), 2)) = "LH" Then
            
                                    If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=-9999, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=9999, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=-9999, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=9999, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    End If
                                ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                                    If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=9999, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=9999, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                   End If
                                ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                                    If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=-9999, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=-9999, hiVal:=X_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    End If
                                Else
                                    If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=-9999, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=9999, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=-9999, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=9999, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    End If
                                End If
                            Case "7777"
                                ' still print out the latest point
                                If i > MaxArrIndex - 1 Then
                                    theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site))
                                Else
                                    theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                                End If
                                If UCase(right(tmpstr(0), 2)) = "LH" Then
            
                                    If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=-7777, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=7777, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=-7777, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=7777, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    End If
                                ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                                    If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=7777, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=7777, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    End If
                                ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                                    If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=-7777, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=-7777, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    End If
                                Else
                                    If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=-7777, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=7777, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=-7777, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=7777, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    End If
                                End If
                            Case "5555"
                                ' still print out the latest point
                                If i > MaxArrIndex - 1 Then
                                    theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site))
                                Else
                                    theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                                End If
                                If UCase(right(tmpstr(0), 2)) = "LH" Then
            
                                    If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=-5555, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=5555, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=-5555, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=5555, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    End If
                                ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                                    If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=5555, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=5555, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    End If
                                ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                                    If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=-5555, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=-5555, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    End If
                                Else
                                    If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=-5555, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=5555, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=-5555, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=5555, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    End If
                                End If
                            Case Else ' search LVCC/HVCC point
                                ' still print out the latest point
                                If i > MaxArrIndex - 1 Then
                                    theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site))
                                Else
                                    theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                                End If
                                If UCase(right(tmpstr(0), 2)) = "LH" Then
            
                                    If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    End If
                                ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                                    If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    End If
                                ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                                    If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    End If
                                Else
                                    If RangeSeq(0) = True Then 'X-axis small--->large
                                        theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    Else
                                        theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                        theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                        g_TestNum = g_TestNum + 1
                                    End If
                                End If
                        End Select
                        ShmooResult = vbNullString
                        j = 0
                    Else
                        If i > MaxArrIndex - 1 Then
                            theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site))
                        Else
                            theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                        End If
                    End If
            
                Next i
            Next yy
            k = 0 ' initial it fot multi site
            '''
        End If
        g_TestNum = tmp_Tnum
Next site

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Print3DShmooInfo_ZXY")
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Print2DShmooInfo(argc As Long, argv() As String)
On Error GoTo errHandler
Dim site As Variant
Dim i, j, k As Long
Dim ShmooResult As String
Dim x_val As Variant
Dim y_val As Variant
Dim z_val As Variant
Dim Lvcc As Double
Dim HVCC As Double
Dim InstName As String
Dim curr_axis As Variant
Dim axis_pin(2) As String
Dim X_Axis_Val() As Double
Dim Y_Axis_Val() As Double
Dim Z_Axis_Val() As Double
Dim x_val_tracking() As Variant
Dim y_val_tracking() As Variant
Dim z_val_tracking() As Variant
Dim DevSetupName As String

Dim tmpstr() As String
Dim tmpStr1() As String
Dim tmpStr2() As String
Dim InstName_H As String
Dim InstName_L As String
Dim TmpVal As Integer
Dim X_Axis_TrackingPara() As String
Dim Y_Axis_TrackingPara() As String
Dim Z_Axis_TrackingPara() As String
Dim X_Axis_TrackingParaFrom() As Variant
Dim Y_Axis_TrackingParaFrom() As Variant
Dim Z_Axis_TrackingParaFrom() As Variant
Dim X_Axis_TrackingParaTo() As Variant
Dim Y_Axis_TrackingParaTo() As Variant
Dim Z_Axis_TrackingParaTo() As Variant
Dim Tracking_Item As Variant
Dim iI, jj, kk As Integer

Dim x_TrackingInfo As String
Dim x_stepsize As Variant
Dim x_num As Integer
Dim y_TrackingInfo As String
Dim y_stepsize As Variant
Dim y_num As Integer
Dim z_TrackingInfo As String
Dim z_stepsize As Variant
Dim z_num As Integer
Dim tmp_Tnum As Long, CurrPointNum As Long
Dim X_result As String
Dim Y_result As String
Dim Z_result As String
Dim X_Tname_result As String
Dim Y_Tname_result As String
Dim Z_Tname_result As String

iI = 0
jj = 0
kk = 0
''Exit Function
DevSetupName = theexec.DevChar.Setups.ActiveSetupName
For Each curr_axis In theexec.DevChar.Setups(DevSetupName).Shmoo.Axes.list
        Select Case curr_axis
            Case 0
                ReDim X_Axis_Val(Xaxis_index - 1)
                If X_Tracking_Point <> 0 Then 'extract X Axis tracking pin, either ApplyToPins or ParameterName
                    ReDim X_Axis_TrackingPara(X_Tracking_Point - 1)
                    ReDim X_Axis_TrackingParaFrom(X_Tracking_Point - 1)
                    ReDim X_Axis_TrackingParaTo(X_Tracking_Point - 1)
                    With theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters
                        For Each Tracking_Item In .list
                            X_Axis_TrackingPara(iI) = Replace((theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).ApplyTo.Pins), "_", vbNullString)
                            X_Axis_TrackingParaFrom(iI) = theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).range.from
                            X_Axis_TrackingParaTo(iI) = theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).range.to
                            If X_Axis_TrackingPara(iI) = "" Then X_Axis_TrackingPara(iI) = Replace((theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).name), "_", vbNullString)
                            iI = iI + 1
                        Next Tracking_Item
                    End With
                    For x_num = iI - 1 To 0 Step -1
                        x_stepsize = Abs(X_Axis_TrackingParaFrom(x_num) - X_Axis_TrackingParaTo(x_num)) / (Xaxis_index - 1)
                        Call ValueResolution(X_Axis_TrackingParaFrom(x_num), X_Axis_TrackingParaTo(x_num), x_stepsize)
                        x_TrackingInfo = X_Axis_TrackingPara(x_num) & "," & X_Axis_TrackingParaFrom(x_num) & "," & X_Axis_TrackingParaTo(x_num) & "," & x_stepsize & ";" & x_TrackingInfo
                    Next x_num
                    theexec.Datalog.WriteComment " X-Asix TraickingPin Info => " & x_TrackingInfo
                End If
            Case 1
                ReDim Y_Axis_Val(Yaxis_index - 1)
                If Y_Tracking_Point <> 0 Then 'extract Y Axis tracking pin, either ApplyToPins or ParameterName
                    ReDim Y_Axis_TrackingPara(Y_Tracking_Point - 1)
                    ReDim Y_Axis_TrackingParaFrom(Y_Tracking_Point - 1)
                    ReDim Y_Axis_TrackingParaTo(Y_Tracking_Point - 1)
                    With theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters
                        For Each Tracking_Item In .list
                            Y_Axis_TrackingPara(jj) = Replace((theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).ApplyTo.Pins), "_", vbNullString)
                            Y_Axis_TrackingParaFrom(jj) = theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).range.from
                            Y_Axis_TrackingParaTo(jj) = theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).range.to
                            If Y_Axis_TrackingPara(jj) = "" Then Y_Axis_TrackingPara(jj) = Replace((theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).name), "_", vbNullString)
                            jj = jj + 1
                        Next Tracking_Item
                    End With
                    For y_num = jj - 1 To 0 Step -1
                        y_stepsize = Abs(Y_Axis_TrackingParaFrom(y_num) - Y_Axis_TrackingParaTo(y_num)) / (Yaxis_index - 1)
                        Call ValueResolution(Y_Axis_TrackingParaFrom(y_num), Y_Axis_TrackingParaTo(y_num), y_stepsize)
                        y_TrackingInfo = Y_Axis_TrackingPara(y_num) & "," & Y_Axis_TrackingParaFrom(y_num) & "," & Y_Axis_TrackingParaTo(y_num) & "," & y_stepsize & ";" & y_TrackingInfo
                    Next y_num
                    theexec.Datalog.WriteComment " Y-Asix TraickingPin Info => " & y_TrackingInfo
                End If
            Case 2
                ReDim Z_Axis_Val(Zaxis_index - 1)
                If Z_Tracking_Point <> 0 Then 'extract Z Axis tracking pin, either ApplyToPins or ParameterName
                    ReDim Z_Axis_TrackingPara(Z_Tracking_Point - 1)
                    ReDim Z_Axis_TrackingParaFrom(Z_Tracking_Point - 1)
                    ReDim Z_Axis_TrackingParaTo(Z_Tracking_Point - 1)
                    With theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters
                        For Each Tracking_Item In .list
                            Z_Axis_TrackingPara(kk) = Replace((theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).ApplyTo.Pins), "_", vbNullString)
                            Z_Axis_TrackingParaFrom(kk) = theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).range.from
                            Z_Axis_TrackingParaTo(kk) = theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).range.to
                            If Z_Axis_TrackingPara(kk) = "" Then Z_Axis_TrackingPara(kk) = Replace((theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).name), "_", vbNullString)
                            kk = kk + 1
                        Next Tracking_Item
                    End With
                    For z_num = kk - 1 To 0 Step -1
                        z_stepsize = Abs(Z_Axis_TrackingParaFrom(z_num) - Z_Axis_TrackingParaTo(z_num)) / (Zaxis_index - 1)
                        Call ValueResolution(Z_Axis_TrackingParaFrom(z_num), Z_Axis_TrackingParaTo(z_num), z_stepsize)
                        z_TrackingInfo = Z_Axis_TrackingPara(z_num) & "," & Z_Axis_TrackingParaFrom(z_num) & "," & Z_Axis_TrackingParaTo(z_num) & "," & z_stepsize & ";" & z_TrackingInfo
                    Next z_num
                    theexec.Datalog.WriteComment " Z-Asix TraickingPin Info => " & z_TrackingInfo
                End If
        End Select
        axis_pin(curr_axis) = Replace((theexec.DevChar.Setups(DevSetupName).Shmoo.Axes((curr_axis)).ApplyTo.Pins), "_", vbNullString)
        If axis_pin(curr_axis) = "" Then
            axis_pin(curr_axis) = Replace((theexec.DevChar.Setups(DevSetupName).Shmoo.Axes((curr_axis)).Parameter.name), "_", vbNullString)
        End If
        
Next curr_axis

If X_Tracking_Point <> 0 Then ReDim x_val_tracking(X_Tracking_Point - 1)
If Y_Tracking_Point <> 0 Then ReDim y_val_tracking(Y_Tracking_Point - 1)
If Z_Tracking_Point <> 0 Then ReDim z_val_tracking(Z_Tracking_Point - 1)


InstName = theexec.DataManager.instancename
tmpstr = Split(InstName, "_")

tmpStr1 = Split(InstName, "_")
If InStr(tmpStr1(0), "L") > 0 Or InStr(tmpStr1(0), "H") > 0 Then
    TmpVal = InStr(tmpStr1(0), "L")
    tmpStr1(0) = mid(tmpStr1(0), 1, TmpVal) '  "DFTL" H or "MCL" H
    InstName_L = Join(tmpStr1, "_")

    tmpStr2 = Split(InstName, "_")
    tmpStr2(0) = mid(tmpStr2(0), 1, TmpVal - 1) & right(tmpStr2(0), 1) '  H
    InstName_H = Join(tmpStr2, "_")
Else
    InstName_L = InstName
    InstName_H = InstName
End If
 
k = 0
j = 0
ShmooResult = vbNullString


''  =============================================================================================================================================
''  X-axis
tmp_Tnum = g_TestNum
For Each site In theexec.sites
    If X_dimemsion = True Then
        theexec.Datalog.WriteComment "  [ X-Axis ] "
        Dim xx As Long
        Dim tmp_count1 As Long
        Dim tmp_x As Long
        'xx = 2
        tmp_count1 = 0
'        For xx = 1 To Zaxis_index
            For i = tmp_count1 To MaxArrIndex - 1
'                If i > xx * MaxArrIndex - 1 Then
'                    tmp_count1 = i
'                    Exit For
'                Else
'                    tmp_x = i - (xx - 1) * (MaxArrIndex - 1)
                    ShmooResult = ShmooResult & g_ShmooResult.Axis_CurrPoint(i).CurrResult(site)
                    X_Axis_Val(j) = g_ShmooResult.Axis_CurrPoint(i).X_axis(site)
                    j = j + 1
                    x_val = g_ShmooResult.Axis_CurrPoint(i).X_axis(site)
                    y_val = g_ShmooResult.Axis_CurrPoint(i).Y_axis(site)
                    Call ValueResolution(x_val, y_val)
                    CurrPointNum = i
                    Call Multi_Axis_ResultName_3D(CurrPointNum, axis_pin, X_result, X_Tname_result, x_val, x_val_tracking, X_Axis_TrackingPara, Y_result, Y_Tname_result, y_val, y_val_tracking, Y_Axis_TrackingPara, Z_result, Z_Tname_result, z_val, z_val_tracking, Z_Axis_TrackingPara, "X", site)
                    If CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site)) = "*" Then g_ShmooResult.Axis_CurrPoint(i).CurrResult(site) = "AP"
                    If CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site)) = "~" Then g_ShmooResult.Axis_CurrPoint(i).CurrResult(site) = "AF"
'                End If
        'Debug.Print tmp_x
                If j = Xaxis_index Then
                    
                    Call ShmooResultPF(ShmooResult, Lvcc, HVCC, X_Axis_Val)
                    Select Case ShmooResult
                        Case "9999"
                            ' still print out the latest point
'                            If i > MaxArrIndex - 1 Then
'                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_x).CurrResult(site))
'                            Else
                                theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
'                            End If
        
                            If UCase(right(tmpstr(0), 2)) = "LH" Then
        
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-9999, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=9999, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-9999, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=9999, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=9999, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=9999, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-9999, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-9999, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            Else
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-9999, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=9999, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-9999, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=9999, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                        Case "7777"
                            ' still print out the latest point
'                            If i > MaxArrIndex - 1 Then
'                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_x).CurrResult(site))
'                            Else
                                theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
'                            End If
                            If UCase(right(tmpstr(0), 2)) = "LH" Then
        
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-7777, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=7777, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-7777, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=7777, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=7777, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=7777, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-7777, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-7777, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            Else
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-7777, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=7777, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-7777, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=7777, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                        Case "5555"
                            ' still print out the latest point
                                theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                            If UCase(right(tmpstr(0), 2)) = "LH" Then
        
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-5555, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=5555, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-5555, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=5555, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=5555, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=5555, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-5555, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-5555, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            Else
                               If RangeSeq(0) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-5555, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=5555, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-5555, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=5555, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                        Case Else ' search LVCC/HVCC point
                                theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                            If UCase(right(tmpstr(0), 2)) = "LH" Then
        
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                               End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                               End If
                            Else
                               If RangeSeq(0) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                    End Select
        '''''            theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & " LVCC or HVCC or shmoo hole "
                    ShmooResult = vbNullString
                    j = 0
                Else
                        theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                End If
        
            Next i
        k = 0 ' initial it fot multi site
    End If
'''''  =============================================================================================================================================
'''''  Y-axis
    If Y_dimemsion = True Then
        theexec.Datalog.WriteComment "  [ Y-Axis ] "
        Dim yy As Long
        Dim tmp_count As Long
        Dim tmp_y As Long
        'yy = 2
        tmp_count = 0
        For yy = 1 To Xaxis_index
            For i = tmp_count To (MaxArrIndex * Xaxis_index) - 1 Step Xaxis_index
        
                If i > yy * MaxArrIndex - 1 Then
                    tmp_count = i
                    Exit For
                Else
                    tmp_y = i - (yy - 1) * (MaxArrIndex - 1)
                    ShmooResult = ShmooResult & g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site)
                    Y_Axis_Val(j) = g_ShmooResult.Axis_CurrPoint(tmp_y).Y_axis(site)
                    j = j + 1
                    x_val = g_ShmooResult.Axis_CurrPoint(tmp_y).X_axis(site)
                    y_val = g_ShmooResult.Axis_CurrPoint(tmp_y).Y_axis(site)
                    z_val = g_ShmooResult.Axis_CurrPoint(tmp_y).Z_axis(site)
                    Call ValueResolution(x_val, y_val)
                    CurrPointNum = tmp_y
                    Call Multi_Axis_ResultName_3D(CurrPointNum, axis_pin, X_result, X_Tname_result, x_val, x_val_tracking, X_Axis_TrackingPara, Y_result, Y_Tname_result, y_val, y_val_tracking, Y_Axis_TrackingPara, Z_result, Z_Tname_result, z_val, z_val_tracking, Z_Axis_TrackingPara, "Y", site)
                    If CStr(g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site)) = "*" Then g_ShmooResult.Axis_CurrPoint(i).CurrResult(site) = "AP"
                    If CStr(g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site)) = "~" Then g_ShmooResult.Axis_CurrPoint(i).CurrResult(site) = "AF"
                End If
        'Debug.Print tmp_y
                If j = Yaxis_index Then
        '            theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                    Call ShmooResultPF(ShmooResult, Lvcc, HVCC, Y_Axis_Val)
                    Select Case ShmooResult
                        Case "9999"
                            ' still print out the latest point
                            If i > MaxArrIndex - 1 Then
                                theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site))
                            Else
                                theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                            End If
        
                            If UCase(right(tmpstr(0), 2)) = "LH" Then
        
                                If RangeSeq(1) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-9999, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=9999, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-9999, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=9999, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                                If RangeSeq(1) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=9999, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=9999, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                               End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                                If RangeSeq(1) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-9999, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-9999, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            Else
                                If RangeSeq(1) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-9999, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=9999, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-9999, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=9999, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                        Case "7777"
                            ' still print out the latest point
                            If i > MaxArrIndex - 1 Then
                                theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site))
                            Else
                                theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                            End If
                            If UCase(right(tmpstr(0), 2)) = "LH" Then
        
                                If RangeSeq(1) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-7777, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=7777, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-7777, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=7777, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                                If RangeSeq(1) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=7777, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=7777, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                                If RangeSeq(1) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-7777, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-7777, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            Else
                                If RangeSeq(1) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-7777, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=7777, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-7777, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=7777, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                        Case "5555"
                            ' still print out the latest point
                            If i > MaxArrIndex - 1 Then
                                theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site))
                            Else
                                theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                            End If
                            If UCase(right(tmpstr(0), 2)) = "LH" Then
        
                                If RangeSeq(1) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-5555, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=5555, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-5555, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=5555, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                                If RangeSeq(1) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=5555, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=5555, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                                If RangeSeq(1) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-5555, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-5555, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            Else
                                If RangeSeq(1) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-5555, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=5555, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-5555, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=5555, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                        Case Else ' search LVCC/HVCC point
                            ' still print out the latest point
                            If i > MaxArrIndex - 1 Then
                                theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site))
                            Else
                                theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                            End If
                            If UCase(right(tmpstr(0), 2)) = "LH" Then
        
                                If RangeSeq(1) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                                If RangeSeq(1) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                                If RangeSeq(1) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            Else
                                If RangeSeq(1) = True Then 'X-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                    End Select
        ''''            theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & " LVCC or HVCC or shmoo hole "
                    ShmooResult = vbNullString
                    j = 0
                Else
                    If i > MaxArrIndex - 1 Then
                        theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site))
                    Else
                        theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                    End If
        '            TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                End If
        
            Next i
        Next yy
        k = 0 ' initial it fot multi site

    End If
    g_TestNum = tmp_Tnum
Next site

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Print2DShmooInfo")
    If AbortTest Then Exit Function Else Resume Next
End Function



Public Function ShmooResultPF(ShmRes As String, Lvcc As Double, HVCC As Double, ShmooEachPoint() As Double)
On Error GoTo errHandler
'Public Function ShmooResultPF(ShmRes As String, LVCC As Double, HVCC As Double)

Dim strLen As Long
Dim i, j, k As Long
Dim Count As Long
Dim LocResult As String
Dim PreviousResult As String
Dim PreviousResult_tmp As String
Dim FPPoint(1000) As Integer
Dim FPCount As Integer
Dim PFPoint(1000) As Integer
Dim PFCount As Integer
Dim FirstPassLoc As Integer
Dim FirstFailLoc As Integer
Dim FPTransit As Integer
Dim PFTransit As Integer
'Dim LVCC As Double
Dim FPTransit_1 As Integer
Dim PFTransit_1 As Integer
Dim Entry_Flag As Boolean
Dim FailToPassCounr As Integer


    PFCount = 0
    FPCount = 0
    FirstPassLoc = -1
    FirstFailLoc = -1
    FPTransit = -1
    PFTransit = -1
    FPTransit_1 = -1
    PFTransit_1 = -1
    Count = 0
    Entry_Flag = True
    FailToPassCounr = 0
    k = InStr(ShmRes, "P") ' for FFPPFF case

ShmRes = Replace(Replace(Replace(Replace(ShmRes, "*", "P"), "~", "F"), "AP", "P"), "AF", "F")

''Dim ShmooEachPoint(5) As Double
''ShmooEachPoint(0) = 1
''ShmooEachPoint(1) = 1.2
''ShmooEachPoint(2) = 1.4
''ShmooEachPoint(3) = 1.6
''ShmooEachPoint(4) = 1.8
''ShmooEachPoint(5) = 2

''ShmooEachPoint(0) = 2
''ShmooEachPoint(1) = 1.8
''ShmooEachPoint(2) = 1.6
''ShmooEachPoint(3) = 1.4
''ShmooEachPoint(4) = 1.2
''ShmooEachPoint(5) = 1

If InStr(ShmRes, "A") Then 'Alarm
    ShmRes = "7777"
ElseIf InStr(ShmRes, "E") Then 'Error
    ShmRes = "7777"
Else
''    If LVCC_flag = True Then
        strLen = Len(ShmRes)
        For i = 1 To strLen
            LocResult = mid(ShmRes, i, 1)
            If (i = 1) Then
                PreviousResult = LocResult
                PreviousResult_tmp = LocResult
            Else
                If (PreviousResult <> LocResult) Then
                    If (PreviousResult = "P") Then
                        PFPoint(PFCount) = i
                        PFCount = PFCount + 1
                    Else
                        FPPoint(FPCount) = i
                        FPCount = FPCount + 1
                    End If
                    PreviousResult = LocResult
                End If
            End If
            
            If (LocResult = "P") Then
                If (FirstPassLoc = -1) Then
                    FirstPassLoc = i - 1
                End If
                If (FirstFailLoc <> -1 And PFTransit = -1) Then
                    PFTransit = i - 1
                End If

            End If
            
            If (LocResult = "F") Then
                If (FirstFailLoc = -1) Then
                    FirstFailLoc = i - 1
                End If
                If (FirstPassLoc <> -1 And FPTransit = -1) Then
                    FPTransit = i - 1
                End If
                
           End If
           
            If left(ShmRes, 1) = "F" Then ' only allow first point is "F"
                If mid(ShmRes, k + 1, 1) = "P" Then ' skip  FFFFF 'P' FFFFF case
                    If FirstPassLoc <> -1 And (PreviousResult_tmp <> LocResult) Then
                        PreviousResult_tmp = LocResult
                        FailToPassCounr = FailToPassCounr + 1 '
                    End If
                End If
            End If
        Next i
    
        If (PFTransit = -1 And FPTransit = -1 And FirstPassLoc <> -1 And FirstFailLoc = -1) Then ' All PASS
            If ShmooEachPoint(UBound(ShmooEachPoint)) > ShmooEachPoint(LBound(ShmooEachPoint)) Then
                Lvcc = ShmooEachPoint(LBound(ShmooEachPoint))
                HVCC = ShmooEachPoint(UBound(ShmooEachPoint))
            Else
                Lvcc = ShmooEachPoint(UBound(ShmooEachPoint))
                HVCC = ShmooEachPoint(LBound(ShmooEachPoint))
            End If
        ElseIf (PFTransit = -1 And FPTransit = -1 And FirstPassLoc = -1 And FirstFailLoc <> -1) Then ' All FAIL
                ShmRes = "9999"

        ElseIf (PFTransit <> -1 And FPTransit = -1 And FirstPassLoc <> -1 And FirstFailLoc <> -1) Then ' Fail-Pass transition point
            If ShmooEachPoint(UBound(ShmooEachPoint)) > ShmooEachPoint(LBound(ShmooEachPoint)) Then
                Lvcc = ShmooEachPoint(FPPoint(0) - 1)
                HVCC = ShmooEachPoint(UBound(ShmooEachPoint))
            Else
                Lvcc = ShmooEachPoint(UBound(ShmooEachPoint))
                HVCC = ShmooEachPoint(FPPoint(0) - 1)
            End If
        ElseIf (PFTransit = -1 And FPTransit <> -1 And FirstPassLoc <> -1 And FirstFailLoc <> -1) Then ' Pass-Fail transition point
            If ShmooEachPoint(UBound(ShmooEachPoint)) > ShmooEachPoint(LBound(ShmooEachPoint)) Then
                Lvcc = ShmooEachPoint(0)
                HVCC = ShmooEachPoint(PFPoint(0) - 2)
            Else
                Lvcc = ShmooEachPoint(PFPoint(0) - 2)
                HVCC = ShmooEachPoint(0)
            End If
        ElseIf (PFTransit <> -1 And FPTransit <> -1 And FirstPassLoc <> -1 And FirstFailLoc <> -1 And FailToPassCounr <> 2) Then ' Shmoo hole, ex: PPPFFFPPP
'            If ShmooEachPoint(UBound(ShmooEachPoint)) > ShmooEachPoint(LBound(ShmooEachPoint)) Then
                ShmRes = "5555"
'            Else
'                ShmRes = "-5555"
'            End If
        ElseIf (FailToPassCounr = 2) Then '  ex: FFPPPFFF
            If ShmooEachPoint(UBound(ShmooEachPoint)) > ShmooEachPoint(LBound(ShmooEachPoint)) Then
                Lvcc = ShmooEachPoint(FPPoint(0) - 1)
                HVCC = ShmooEachPoint(PFPoint(0) - 2)
            Else
                Lvcc = ShmooEachPoint(PFPoint(0) - 2)
                HVCC = ShmooEachPoint(FPPoint(0) - 1)
            End If
        End If
    
End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "ShmooResultPF")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Decide_Switching_Bit_Debug_LVCC(digSrc_EQ As String, DSPWaveSwitch As DSPWave, DC_Level As PinListData, BlockType As String, SELSRM_Rails As String, Optional shmoo_pin As String, Optional ShmooPinsVoltage As PinListData, Optional ForcePin As String, Optional SetForceVoltage As Dictionary, Optional DSSCPatName As String, _
                                                Optional SEL_print As SiteVariant, Optional enable_print As Boolean = False) As String
On Error GoTo errHandler
  Dim p_ary() As String, p_cnt As Long
  Dim logicPin As String
  Dim SramPin As String
  Dim DSSC_Switching_Voltage As New PinListData
  Dim Sdomain As Long
  Dim DSSCSelSrmOpposite As Long
  Dim BlockTypeNum As Long
  Dim PattIdx As Long
  Dim i As Integer, j As Integer
  Dim ReturnString() As String
  Dim LogicValue As Double
  Dim tmpLogicVal As Double
  Dim SramValue As Double
  Dim sl_ReturnString() As New SiteLong
  Dim s_Statement As String
  BlockTypeNum = -1
  PattIdx = -1
  ReDim ReturnString(Len(digSrc_EQ) - 1)
  ReDim sl_ReturnString(Len(digSrc_EQ) - 1)
    '''=========================================================='''
    ''' this is to solve pin group not decompose in LVCC Boundary'''
    Dim pin As Variant
    Dim cnt As Long
    Dim NewShmooPinsVoltage As New PinListData
    For Each pin In ShmooPinsVoltage.Pins
        theexec.DataManager.DecomposePinList pin, p_ary, p_cnt
        For cnt = 0 To p_cnt - 1
            NewShmooPinsVoltage.AddPin p_ary(cnt)
            NewShmooPinsVoltage.Pins(p_ary(cnt)).value = ShmooPinsVoltage.Pins(pin).value
        Next cnt
    Next pin
    '''=========================================================='''
  Decide_DSSC_Switching_Voltage DSSC_Switching_Voltage, DC_Level, shmoo_pin, NewShmooPinsVoltage, ForcePin, SetForceVoltage
    
        Dim l_Selsram_index As Long
    l_Selsram_index = SelSRAM_Index_Select(SelsramMapping, BlockType, DSSCPatName)
   
    If l_Selsram_index <> -1 Then
        Call SelSRAM_DigSrc_Bit(l_Selsram_index, digSrc_EQ, DSSC_Switching_Voltage, DSPWaveSwitch, ReturnString, sl_ReturnString, , SEL_print, True)
        
        'If theexec.DevChar.Setups.IsRunning = False Then
           Decide_Switching_Bit_Debug_LVCC = Join(ReturnString, "")
           SELSRM_Rails = DecodingRealSourceBit(Decide_Switching_Bit_Debug_LVCC, BlockType, DSSCPatName)
'        Else
'           Decide_Switching_Bit_Debug_LVCC = digSrc_EQ
'           SELSRM_Rails = DecodingRealSourceBit(Decide_Switching_Bit_Debug_LVCC, BlockType, DSSCPatName)
'        End If
        
    Else
        s_Statement = "We could not find from the SelSRAM Mapping Table properly."
        ''Standardize error message
    End If

  Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Decide_Switching_Bit_Debug_LVCC")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ShmooPostStep3D(argc As Long, argv() As String)
On Error GoTo errHandler
    Dim SetupName As String
    Dim i As Long
    Dim OutputString As String
    Dim instancename As String
    Dim TestNum As Long
    Dim lvccf As Integer
    Dim Lvcc As Double
    Dim site As Variant
    Dim v_Xi0 As Double
    Dim TestVoltage As String
    Dim StartVoltage As Double, EndVoltage As Double, StepSize As Double
    Dim Patt_String As String, Shmoo_Result As String
    Dim Pat As Variant
    Dim PinName As Variant
    Dim StepName As Variant
    Dim RangeFrom As Double, RangeTo As Double, RangeStepSize As Double, RangeSteps As Long
    Dim allPowerPins As String
    Dim PowerPinCnt As Long, PowerPinAry() As String
    Dim FlagFirstPass As Boolean
    Dim last_point_result As tlDevCharResult, current_point_result As tlDevCharResult
    Dim min_point As Long, max_point As Long, current_point As Long
    Dim Vcc_min As String, Vcc_max As String
    Dim patt_ary() As String, pat_count As Long, p As Variant
    Dim Pin_Ary() As String, Pin_Cnt As Long
    Dim shmoo_pin_string As String
    Dim tmp As String
    Dim Search_String As String, Search_String_X As String, Search_String_Y As String, Search_String_Z As String
    Dim Group As Boolean
    Dim Label As String
    Dim Shmoo_Pattset As New Pattern
    Dim VIL_Flag As Boolean
    Dim step_Start As Long
    Dim step_Stop As Long
    Dim Step_x As Long
    Dim RangeLow As Double, RangeStart As Double
    Dim Shmoo_hole As String
    Dim RangeCalcType As tlDevCharRangeField
    Dim xio_spec As String
    Dim Range_temp As Double
    Dim range_plus As Long
    
    Dim HIO_PinName_Updated As Boolean      '20180515 TER
    
    Dim index As Long
    Dim vbump_value As String
    
    
    instancename = theexec.DataManager.instancename     '20180616 TER add
'''    Call Get_Tname_FromFlowSheet(InstanceName, HIO_PinName_Updated)      '20180515 TER add
    
    For Each site In theexec.sites
        OutputString = vbNullString
        lvccf = 0
        
        Dim nWire_port_ary() As String
        Dim nwp As Variant
        Dim port_pa As String, ac_spec_pa As String, pin_pa As String, global_spec_pa As String
        Dim FRC_Name As String, FRC_Value As Double, All_FRC_Status As String
        All_FRC_Status = vbNullString
        nWire_port_ary = Split(nWire_Ports_GLB, ",")
        For Each nwp In nWire_port_ary
            Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
            If thehdw.Protocol.ports(port_pa).Enabled = True Then
                FRC_Name = Replace(UCase(ac_spec_pa), "_FREQ_VAR", vbNullString)
                FRC_Value = theexec.Specs.AC(ac_spec_pa).CurrentValue
                If All_FRC_Status = vbNullString Then
                    All_FRC_Status = FRC_Name & "=" & FRC_Value
                Else
                    All_FRC_Status = All_FRC_Status & ";" & FRC_Name & "=" & FRC_Value
                End If
            End If
        Next nwp
        If FRC_Name = "" Then ' Default use XI0, if no input of FRC info
            FRC_Name = XI0
            FRC_Value = TheExec.Specs.AC(XI0_Diff & "_Freq_VAR").CurrentValue      '"XI0_Freq_VAR"
            All_FRC_Status = FRC_Name & "=" & FRC_Value
        End If
    
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        
        ''Read X axis setup information
        SetupName = theexec.DevChar.Setups.ActiveSetupName
        
        With theexec.DevChar
            StepName = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).StepName
            RangeFrom = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.from
            RangeTo = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.to
            RangeSteps = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.Steps + 1
            RangeStepSize = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.StepSize
            RangeCalcType = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.CalculatedField

            If RangeFrom < RangeTo Then ' always start from lower Value
                step_Start = 0
                step_Stop = RangeSteps - 1
                Step_x = 1
                RangeLow = RangeFrom
                range_plus = -1
            Else
                step_Start = RangeSteps - 1
                step_Stop = 0
                Step_x = -1
                If RangeCalcType = tlDevCharRangeField_Steps Then 'calculate step
                    RangeTo = RangeFrom - (RangeSteps - 1) * RangeStepSize
                End If
                RangeLow = RangeTo
                range_plus = 1
            End If
        End With
        
        Patt_String = vbNullString
        Dim patset As Variant, j As Long
        Shmoo_Pattset.value = Shmoo_Pattern
        Patt_String = PatSetToPat(Shmoo_Pattset)
        gen_search_string SetupName, Search_String_X, tlDevCharShmooAxis_X, RangeFrom, RangeTo, RangeStepSize, RangeSteps
        
        ''Read Y axis setup information
        With theexec.DevChar
            StepName = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).StepName
            RangeFrom = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.range.from
            RangeTo = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.range.to
            RangeSteps = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.range.Steps + 1
            RangeStepSize = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.range.StepSize
            RangeCalcType = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.range.CalculatedField

            If RangeFrom < RangeTo Then ' always start from lower Value
                step_Start = 0
                step_Stop = RangeSteps - 1
                Step_x = 1
                RangeLow = RangeFrom
                range_plus = -1
            Else
                step_Start = RangeSteps - 1
                step_Stop = 0
                Step_x = -1
                If RangeCalcType = tlDevCharRangeField_Steps Then 'calculate step
                    RangeTo = RangeFrom - (RangeSteps - 1) * RangeStepSize
                End If
                RangeLow = RangeTo
                range_plus = 1
            End If
        End With
        
        gen_search_string SetupName, Search_String_Y, tlDevCharShmooAxis_Y, RangeFrom, RangeTo, RangeStepSize, RangeSteps
        
        With theexec.DevChar
            StepName = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_Z).StepName
            RangeFrom = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_Z).Parameter.range.from
            RangeTo = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_Z).Parameter.range.to
            RangeSteps = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_Z).Parameter.range.Steps + 1
            RangeStepSize = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_Z).Parameter.range.StepSize
            RangeCalcType = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_Z).Parameter.range.CalculatedField

            If RangeFrom < RangeTo Then ' always start from lower Value
                step_Start = 0
                step_Stop = RangeSteps - 1
                Step_x = 1
                RangeLow = RangeFrom
                range_plus = -1
            Else
                step_Start = RangeSteps - 1
                step_Stop = 0
                Step_x = -1
                If RangeCalcType = tlDevCharRangeField_Steps Then 'calculate step
                    RangeTo = RangeFrom - (RangeSteps - 1) * RangeStepSize
                End If
                RangeLow = RangeTo
                range_plus = 1
            End If
        End With
        
        gen_search_string SetupName, Search_String_Z, tlDevCharShmooAxis_Z, RangeFrom, RangeTo, RangeStepSize, RangeSteps
        Search_String_Z = "Z@" & Search_String_Z
        Search_String = Search_String_X & Search_String_Y & Search_String_Z
        
        If InStr(theexec.DataManager.instancename, "_NV") Then TestVoltage = "NV"
        If InStr(theexec.DataManager.instancename, "_HV") Then TestVoltage = "HV"
        If InStr(theexec.DataManager.instancename, "_LV") Then TestVoltage = "LV"


        OutputString = OutputString & "[V," & site & "," & All_FRC_Status & "," & HramLotId(site) & "-" & CStr(HramWaferId(site)) & "," & CStr(XCoord(site)) & "," & CStr(YCoord(site)) & ","  '20180716 Auto parsing FRC info

        OutputString = OutputString & theexec.DataManager.instancename & "," & SetupName & "," & CStr(g_TestNum) & ","

        OutputString = OutputString & Patt_String & ","
        OutputString = OutputString & TestVoltage & ","
         PinName = argv(0) 'setup voltage
        If argv(0) <> Empty Then
            theexec.DataManager.DecomposePinList argv(0), Pin_Ary, Pin_Cnt
            PinName = argv(0) 'setup voltage
        End If
        
         
        If Vbump_for_Interpose = True Then
                Dim PL_DC_conditions_str As String
                PL_DC_conditions_str = Replace(PL_DC_conditions_GLB, ":V:", "=")
                PL_DC_conditions_str = Replace(PL_DC_conditions_str, ";", ",")
                OutputString = OutputString & PL_DC_conditions_str
            
        Else
            For j = 0 To Pin_Cnt - 1
                PinName = Pin_Ary(j)
                If theexec.DataManager.ChannelType(PinName) <> "N/C" Then
                    If j = 0 Then
                        OutputString = OutputString & PinName & "=" & Format(thehdw.DCVS.Pins(PinName).Voltage.Main.value, "0.000")
                    Else
                         OutputString = OutputString & "," & PinName & "=" & Format(thehdw.DCVS.Pins(PinName).Voltage.Main.value, "0.000")
                    End If
                End If
            Next j
        End If
        
        For i = 1 To argc - 1
          If UCase(argv(i)) = "VIL" Or UCase(argv(i)) = "VOL" Then
            VIL_Flag = True
          Else
            theexec.DataManager.DecomposePinList argv(i), Pin_Ary, Pin_Cnt
            For j = 0 To Pin_Cnt - 1
                PinName = Pin_Ary(j)
                If theexec.DataManager.ChannelType(PinName) <> "N/C" Then
                    If Vbump_for_Interpose = True Then
                        index = InStr(LCase(PL_DC_conditions_str), PinName & "=")
                        vbump_value = mid(LCase(PL_DC_conditions_str), index + Len(PinName) + 1, 5)
                        OutputString = OutputString & "," & PinName & "=" & vbump_value
                        
                    Else
                        OutputString = OutputString & "," & PinName & "=" & Format(thehdw.DCVS.Pins(PinName).Voltage.Main.value, "0.000")
                    End If
                End If
            Next j
          End If
        Next
        PL_DC_conditions_str = vbNullString
        OutputString = OutputString & ","
       Search_String = mid(Search_String, 1, Len(Search_String) - 1) 'take out last ","
        OutputString = OutputString & Search_String
        OutputString = OutputString & "]"
        theexec.Datalog.WriteComment OutputString
    Next site

          ''clear forcecondition before exit function
          Charz_Force_Power_condition = vbNullString
          g_TestNum = g_TestNum + 1
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "ShmooPostStep3D")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function StoreEachPointsVoltage(Curr_Point As Long, X_Point As Long, Y_Point As Long, Z_Point As Long, RangeStartX As Double, RangeStartY As Double, RangeStartZ As Double, RangeStepSizeX As Double, RangeStepSizeY As Double, RangeStepSizeZ As Double, RangeTrackStartX() As Double, RangeTrackStartY() As Double, RangeTrackStartZ() As Double, RangeTrackStepSizeX() As Double, RangeTrackStepSizeY() As Double, RangeTrackStepSizeZ() As Double)
On Error GoTo errHandler

    Dim curr_num As Long

    If X_Point <> -1 Then g_ShmooResult.Axis_CurrPoint(Curr_Point).X_axis = RangeStartX + X_Point * RangeStepSizeX
    If Y_Point <> -1 Then g_ShmooResult.Axis_CurrPoint(Curr_Point).Y_axis = RangeStartY + Y_Point * RangeStepSizeY
    If Z_Point <> -1 Then g_ShmooResult.Axis_CurrPoint(Curr_Point).Z_axis = RangeStartZ + Z_Point * RangeStepSizeZ
    
    If X_Tracking_Point > 0 Then
       For curr_num = 0 To UBound(RangeTrackStartX)
           g_ShmooResult.Axis_CurrPoint(Curr_Point).X_axis_Tracking(curr_num) = RangeTrackStartX(curr_num) + X_Point * RangeTrackStepSizeX(curr_num)
       Next curr_num
    ElseIf Y_Tracking_Point > 0 Then
       For curr_num = 0 To UBound(RangeTrackStartY)
           g_ShmooResult.Axis_CurrPoint(Curr_Point).Y_axis_Tracking(curr_num) = RangeTrackStartY(curr_num) + Y_Point * RangeTrackStepSizeY(curr_num)
       Next curr_num
    ElseIf Z_Tracking_Point > 0 Then
       For curr_num = 0 To UBound(RangeTrackStartZ)
           g_ShmooResult.Axis_CurrPoint(Curr_Point).Z_axis_Tracking(curr_num) = RangeTrackStartZ(curr_num) + Z_Point * RangeTrackStepSizeZ(curr_num)
       Next curr_num
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "StoreEachPointsVoltage")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Get_ShmooInfo_2D3D(RangeStartX As Double, RangeStartY As Double, RangeStartZ As Double, RangeStepSizeX As Double, RangeStepSizeY As Double, RangeStepSizeZ As Double, RangeTrackStartX() As Double, RangeTrackStartY() As Double, RangeTrackStartZ() As Double, RangeTrackStepSizeX() As Double, RangeTrackStepSizeY() As Double, RangeTrackStepSizeZ() As Double)
On Error GoTo errHandler

    Dim curr_axis As Variant, curr_track As Variant, curr_num As Long
    Dim RangeFrom As Double, RangeTo As Double, RangeStepSize As Double, RangeSteps As Long
    Dim DevSetupName As String
    
    DevSetupName = theexec.DevChar.Setups.ActiveSetupName
    
    For Each curr_axis In theexec.DevChar.Setups(DevSetupName).Shmoo.Axes.list
            With theexec.DevChar
                    RangeFrom = .Setups(DevSetupName).Shmoo.Axes(curr_axis).Parameter.range.from
                    RangeTo = .Setups(DevSetupName).Shmoo.Axes(curr_axis).Parameter.range.to
                    RangeStepSize = .Setups(DevSetupName).Shmoo.Axes(curr_axis).Parameter.range.StepSize
                    RangeSteps = .Setups(DevSetupName).Shmoo.Axes(curr_axis).Parameter.range.Steps
                    Select Case curr_axis
                        Case 0
                             If RangeFrom < RangeTo Then
                                RangeStepSizeX = RangeStepSize
                             Else
                                RangeStepSizeX = -RangeStepSize
                             End If
                             RangeStartX = RangeFrom
                             If X_Tracking_Point > 0 Then
                                 curr_num = 0
                                 For Each curr_track In .Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.list
                                     RangeTrackStartX(curr_num) = .Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(curr_track).range.from
                                     RangeTrackStepSizeX(curr_num) = ((.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(curr_track).range.to) - (.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(curr_track).range.from)) / RangeSteps
                                     curr_num = curr_num + 1
                                 Next curr_track
                             End If
                        Case 1
                             If RangeFrom < RangeTo Then
                                RangeStepSizeY = RangeStepSize
                             Else
                                RangeStepSizeY = -RangeStepSize
                             End If
                             RangeStartY = RangeFrom
                             If Y_Tracking_Point > 0 Then
                                 curr_num = 0
                                 For Each curr_track In .Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.list
                                     RangeTrackStartY(curr_num) = .Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(curr_track).range.from
                                     RangeTrackStepSizeY(curr_num) = ((.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(curr_track).range.to) - (.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(curr_track).range.from)) / RangeSteps
                                     curr_num = curr_num + 1
                                 Next curr_track
                             End If
                        Case 2
                             If RangeFrom < RangeTo Then
                                RangeStepSizeZ = RangeStepSize
                             Else
                                RangeStepSizeZ = -RangeStepSize
                             End If
                             RangeStartZ = RangeFrom
                             If Z_Tracking_Point > 0 Then
                                 curr_num = 0
                                 For Each curr_track In .Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.list
                                     RangeTrackStartZ(curr_num) = .Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(curr_track).range.from
                                     RangeTrackStepSizeZ(curr_num) = ((.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(curr_track).range.to) - (.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(curr_track).range.from)) / RangeSteps
                                     curr_num = curr_num + 1
                                 Next curr_track
                             End If
                    End Select
            End With
        Next curr_axis
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Get_ShmooInfo_2D3D")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Multi_Axis_ResultName_3D(CurrPointNum As Long, axis_pin() As String, X_result As String, X_Tname_result As String, x_val As Variant, x_val_tracking() As Variant, X_Axis_TrackingPara() As String, Y_result As String, Y_Tname_result As String, y_val As Variant, y_val_tracking() As Variant, Y_Axis_TrackingPara() As String, Z_result As String, Z_Tname_result As String, z_val As Variant, z_val_tracking() As Variant, Z_Axis_TrackingPara() As String, AxisResult As String, site As Variant)
On Error GoTo errHandler

            Dim TrackingNum As Integer
            Dim Xaxis_pin_Arr() As String, Yaxis_pin_Arr() As String, Zaxis_pin_Arr() As String
            Dim PinsNum As Integer
            
      Select Case AxisResult
         Case "Z"
            If X_Tracking_Point <> 0 Then
               X_result = vbNullString
               X_Tname_result = vbNullString

               If InStr(CStr(axis_pin(0)), ",") > 0 Then
                  Xaxis_pin_Arr = Split(CStr(axis_pin(0)), ",")
                  For PinsNum = 0 To UBound(Xaxis_pin_Arr)
                      X_result = "_" & Xaxis_pin_Arr(PinsNum) & x_val & X_result
                  Next PinsNum
               Else
                  X_result = "_" & CStr(axis_pin(0)) & x_val & X_result
               End If
               X_Tname_result = X_result
            Else
               X_result = vbNullString
               X_Tname_result = vbNullString
               If InStr(CStr(axis_pin(0)), ",") > 0 Then
                  Xaxis_pin_Arr = Split(CStr(axis_pin(0)), ",")
                  For PinsNum = 0 To UBound(Xaxis_pin_Arr)
                      X_result = "_" & Xaxis_pin_Arr(PinsNum) & x_val
                  Next PinsNum
               Else
                  X_result = "_" & CStr(axis_pin(0)) & x_val
               End If
               X_Tname_result = X_result
            End If
            
            If Y_Tracking_Point <> 0 Then
               Y_result = vbNullString
               Y_Tname_result = vbNullString
               
               If InStr(CStr(axis_pin(1)), ",") > 0 Then
                  Yaxis_pin_Arr = Split(CStr(axis_pin(1)), ",")
                  For PinsNum = 0 To UBound(Yaxis_pin_Arr)
                      Y_result = "_" & Yaxis_pin_Arr(PinsNum) & y_val & Y_result
                  Next PinsNum
               Else
                  Y_result = "_" & CStr(axis_pin(1)) & y_val & Y_result
               End If
               Y_Tname_result = Y_result
            Else
               Y_result = vbNullString
               Y_Tname_result = vbNullString
               If InStr(CStr(axis_pin(1)), ",") > 0 Then
                  Yaxis_pin_Arr = Split(CStr(axis_pin(1)), ",")
                  For PinsNum = 0 To UBound(Yaxis_pin_Arr)
                      Y_result = "_" & Yaxis_pin_Arr(PinsNum) & y_val
                  Next PinsNum
               Else
                  Y_result = "_" & CStr(axis_pin(1)) & y_val
               End If
               Y_Tname_result = Y_result
            End If
            
            If Z_Tracking_Point <> 0 Then
               Z_result = vbNullString
               Z_Tname_result = vbNullString

               If InStr(CStr(axis_pin(2)), ",") > 0 Then
                  Zaxis_pin_Arr = Split(CStr(axis_pin(2)), ",")
                  For PinsNum = 0 To UBound(Zaxis_pin_Arr)
                      Z_result = "_" & Zaxis_pin_Arr(PinsNum) & z_val & Z_result
                      Z_Tname_result = "_" & Zaxis_pin_Arr(PinsNum) & "result" & Z_Tname_result
                  Next PinsNum
               Else
                  Z_result = "_" & CStr(axis_pin(2)) & z_val & Z_result
                  Z_Tname_result = "_" & CStr(axis_pin(2)) & "result" & Z_Tname_result
               End If
            Else
               Z_result = vbNullString
               Z_Tname_result = vbNullString
               If InStr(CStr(axis_pin(2)), ",") > 0 Then
                  Zaxis_pin_Arr = Split(CStr(axis_pin(2)), ",")
                  For PinsNum = 0 To UBound(Zaxis_pin_Arr)
                      Z_result = "_" & Zaxis_pin_Arr(PinsNum) & z_val & Z_result
                      Z_Tname_result = "_" & Zaxis_pin_Arr(PinsNum) & "result" & Z_Tname_result
                  Next PinsNum
               Else
                  Z_result = "_" & CStr(axis_pin(2)) & z_val
                  Z_Tname_result = "_" & CStr(axis_pin(2)) & "result"
               End If
            End If
            
         Case "X"
            If X_Tracking_Point <> 0 Then
               X_result = vbNullString
               X_Tname_result = vbNullString

               If InStr(CStr(axis_pin(0)), ",") > 0 Then
                  Xaxis_pin_Arr = Split(CStr(axis_pin(0)), ",")
                  For PinsNum = 0 To UBound(Xaxis_pin_Arr)
                      X_result = "_" & Xaxis_pin_Arr(PinsNum) & x_val & X_result
                      X_Tname_result = "_" & Xaxis_pin_Arr(PinsNum) & "result" & X_Tname_result
                  Next PinsNum
               Else
                  X_result = "_" & CStr(axis_pin(0)) & x_val & X_result
                  X_Tname_result = "_" & CStr(axis_pin(0)) & "result"
               End If
            Else
               X_result = vbNullString
               X_Tname_result = vbNullString
               If InStr(CStr(axis_pin(0)), ",") > 0 Then
                  Xaxis_pin_Arr = Split(CStr(axis_pin(0)), ",")
                  For PinsNum = 0 To UBound(Xaxis_pin_Arr)
                      X_result = "_" & Xaxis_pin_Arr(PinsNum) & x_val & X_result
                      X_Tname_result = "_" & Xaxis_pin_Arr(PinsNum) & "result" & X_Tname_result
                  Next PinsNum
               Else
                  X_result = "_" & CStr(axis_pin(0)) & x_val
                  X_Tname_result = "_" & CStr(axis_pin(0)) & "result"
               End If
            End If
            
            If Y_Tracking_Point <> 0 Then
               Y_result = vbNullString
               Y_Tname_result = vbNullString

               If InStr(CStr(axis_pin(1)), ",") > 0 Then
                  Yaxis_pin_Arr = Split(CStr(axis_pin(1)), ",")
                  For PinsNum = 0 To UBound(Yaxis_pin_Arr)
                      Y_result = "_" & Yaxis_pin_Arr(PinsNum) & y_val & Y_result
                  Next PinsNum
               Else
                  Y_result = "_" & CStr(axis_pin(1)) & y_val & Y_result
               End If
               Y_Tname_result = Y_result
            Else
               Y_result = vbNullString
               Y_Tname_result = vbNullString
               If InStr(CStr(axis_pin(1)), ",") > 0 Then
                  Yaxis_pin_Arr = Split(CStr(axis_pin(1)), ",")
                  For PinsNum = 0 To UBound(Yaxis_pin_Arr)
                      Y_result = "_" & Yaxis_pin_Arr(PinsNum) & y_val
                  Next PinsNum
               Else
                  Y_result = "_" & CStr(axis_pin(1)) & y_val
               End If
               Y_Tname_result = Y_result
            End If
            
            If Z_Tracking_Point <> 0 Then
               Z_result = vbNullString
               Z_Tname_result = vbNullString

               If InStr(CStr(axis_pin(2)), ",") > 0 Then
                  Zaxis_pin_Arr = Split(CStr(axis_pin(2)), ",")
                  For PinsNum = 0 To UBound(Zaxis_pin_Arr)
                      Z_result = "_" & Zaxis_pin_Arr(PinsNum) & z_val & Z_result
                  Next PinsNum
               Else
                  Z_result = "_" & CStr(axis_pin(2)) & z_val & Z_result
               End If
               Z_Tname_result = Z_result
            Else
               Z_result = vbNullString
               Z_Tname_result = vbNullString
               If InStr(CStr(axis_pin(2)), ",") > 0 Then
                  Zaxis_pin_Arr = Split(CStr(axis_pin(2)), ",")
                  For PinsNum = 0 To UBound(Zaxis_pin_Arr)
                      Z_result = "_" & Zaxis_pin_Arr(PinsNum) & z_val & Z_result
                  Next PinsNum
               Else
                  Z_result = "_" & CStr(axis_pin(2)) & z_val
               End If
               Z_Tname_result = Z_result
            End If
               
         Case "Y"
               If X_Tracking_Point <> 0 Then
                  X_result = vbNullString
                  X_Tname_result = vbNullString

                  If InStr(CStr(axis_pin(0)), ",") > 0 Then
                     Xaxis_pin_Arr = Split(CStr(axis_pin(0)), ",")
                     For PinsNum = 0 To UBound(Xaxis_pin_Arr)
                        X_result = "_" & Xaxis_pin_Arr(PinsNum) & x_val & X_result
                     Next PinsNum
                  Else
                     X_result = "_" & CStr(axis_pin(0)) & x_val & X_result
                  End If
                  X_Tname_result = X_result
               Else
                  X_result = vbNullString
                  X_Tname_result = vbNullString
                  If InStr(CStr(axis_pin(0)), ",") > 0 Then
                     Xaxis_pin_Arr = Split(CStr(axis_pin(0)), ",")
                     For PinsNum = 0 To UBound(Xaxis_pin_Arr)
                         X_result = "_" & Xaxis_pin_Arr(PinsNum) & x_val
                     Next PinsNum
                  Else
                     X_result = "_" & CStr(axis_pin(0)) & x_val
                  End If
                  X_Tname_result = X_result
               End If
                
               If Y_Tracking_Point <> 0 Then
                  Y_result = vbNullString
                  Y_Tname_result = vbNullString

                  If InStr(CStr(axis_pin(0)), ",") > 0 Then
                     Xaxis_pin_Arr = Split(CStr(axis_pin(0)), ",")
                     For PinsNum = 0 To UBound(Xaxis_pin_Arr)
                        Y_result = "_" & Yaxis_pin_Arr(PinsNum) & y_val & Y_result
                        Y_Tname_result = "_" & Yaxis_pin_Arr(PinsNum) & "result" & Y_Tname_result
                     Next PinsNum
                  Else
                     Y_result = "_" & CStr(axis_pin(1)) & y_val & Y_result
                     Y_Tname_result = "_" & CStr(axis_pin(1)) & "result"
                  End If
               Else
                  Y_result = vbNullString
                  Y_Tname_result = vbNullString
                  If InStr(CStr(axis_pin(1)), ",") > 0 Then
                     Yaxis_pin_Arr = Split(CStr(axis_pin(1)), ",")
                     For PinsNum = 0 To UBound(Xaxis_pin_Arr)
                         Y_result = "_" & Yaxis_pin_Arr(PinsNum) & y_val & Y_result
                         Y_Tname_result = "_" & Yaxis_pin_Arr(PinsNum) & "result" & Y_Tname_result
                     Next PinsNum
                  Else
                     Y_result = "_" & CStr(axis_pin(1)) & y_val
                     Y_Tname_result = "_" & CStr(axis_pin(1)) & "result"
                  End If
               End If
                
               If Z_Tracking_Point <> 0 Then
                  Z_result = vbNullString
                  Z_Tname_result = vbNullString

                  If InStr(CStr(axis_pin(2)), ",") > 0 Then
                     Zaxis_pin_Arr = Split(CStr(axis_pin(2)), ",")
                     For PinsNum = 0 To UBound(Zaxis_pin_Arr)
                         Z_result = "_" & Zaxis_pin_Arr(PinsNum) & z_val & Z_result
                     Next PinsNum
                  Else
                     Z_result = "_" & CStr(axis_pin(2)) & z_val & Z_result
                  End If
                  Z_Tname_result = Z_result
               Else
                  Z_result = vbNullString
                  Z_Tname_result = vbNullString
                  If InStr(CStr(axis_pin(2)), ",") > 0 Then
                     Zaxis_pin_Arr = Split(CStr(axis_pin(2)), ",")
                     For PinsNum = 0 To UBound(Zaxis_pin_Arr)
                         Z_result = "_" & Zaxis_pin_Arr(PinsNum) & z_val & Z_result
                     Next PinsNum
                  Else
                     Z_result = "_" & CStr(axis_pin(2)) & z_val
                  End If
                  Z_Tname_result = Z_result
               End If
     End Select
          
     If X_result <> "" Then X_result = mid(X_result, 2, Len(X_result))
     If X_Tname_result <> "" Then X_Tname_result = mid(X_Tname_result, 2, Len(X_Tname_result))
     If Y_result <> "" Then Y_result = mid(Y_result, 2, Len(Y_result))
     If Y_Tname_result <> "" Then Y_Tname_result = mid(Y_Tname_result, 2, Len(Y_Tname_result))
     If Z_result <> "" Then Z_result = mid(Z_result, 2, Len(Z_result))
     If Z_Tname_result <> "" Then Z_Tname_result = mid(Z_Tname_result, 2, Len(Z_Tname_result))

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Multi_Axis_ResultName_3D")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Print3DShmooInfo_XYZ(argc As Long, argv() As String)
On Error GoTo errHandler
Dim site As Variant
Dim i, j, k As Long
Dim ShmooResult As String
Dim x_val As Variant
Dim y_val As Variant
Dim z_val As Variant
Dim x_val_tracking() As Variant
Dim y_val_tracking() As Variant
Dim z_val_tracking() As Variant
Dim Lvcc As Double
Dim HVCC As Double
Dim InstName As String
Dim curr_axis As Variant
Dim axis_pin(2) As String
Dim X_Axis_Val() As Double
Dim Y_Axis_Val() As Double
Dim Z_Axis_Val() As Double
Dim DevSetupName As String
'Dim TrackingNum As Integer

Dim tmpstr() As String
Dim tmpStr1() As String
Dim tmpStr2() As String
Dim InstName_H As String
Dim InstName_L As String
Dim TmpVal As Integer
Dim X_Axis_TrackingPara() As String
Dim Y_Axis_TrackingPara() As String
Dim Z_Axis_TrackingPara() As String
Dim X_Axis_TrackingParaFrom() As Variant
Dim Y_Axis_TrackingParaFrom() As Variant
Dim Z_Axis_TrackingParaFrom() As Variant
Dim X_Axis_TrackingParaTo() As Variant
Dim Y_Axis_TrackingParaTo() As Variant
Dim Z_Axis_TrackingParaTo() As Variant
Dim Tracking_Item As Variant
Dim iI, jj, kk As Integer

Dim x_TrackingInfo As String
Dim x_stepsize As Variant
Dim x_num As Integer
Dim y_TrackingInfo As String
Dim y_stepsize As Variant
Dim y_num As Integer
Dim z_TrackingInfo As String
Dim z_stepsize As Variant
Dim z_num As Integer
Dim tmp_Tnum As Long, CurrPointNum As Long
Dim X_result As String
Dim Y_result As String
Dim Z_result As String
Dim X_Tname_result As String
Dim Y_Tname_result As String
Dim Z_Tname_result As String
Dim AxisOrder As Variant
Dim xx As Long
Dim tmp_count1 As Long
Dim tmp_x As Long
Dim yy As Long
Dim tmp_count As Long
Dim tmp_y As Long

iI = 0
jj = 0
kk = 0
''Exit Function

DevSetupName = theexec.DevChar.Setups.ActiveSetupName

For Each curr_axis In theexec.DevChar.Setups(DevSetupName).Shmoo.Axes.list
        Select Case curr_axis
            Case 0
                ReDim X_Axis_Val(Xaxis_index - 1)
                If X_Tracking_Point <> 0 Then 'extract X Axis tracking pin, either ApplyToPins or ParameterName
                    ReDim X_Axis_TrackingPara(X_Tracking_Point - 1)
                    ReDim X_Axis_TrackingParaFrom(X_Tracking_Point - 1)
                    ReDim X_Axis_TrackingParaTo(X_Tracking_Point - 1)
                    With theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters
                        For Each Tracking_Item In .list
                            X_Axis_TrackingPara(iI) = Replace((theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).ApplyTo.Pins), "_", vbNullString)
                            X_Axis_TrackingParaFrom(iI) = theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).range.from
                            X_Axis_TrackingParaTo(iI) = theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).range.to
                            If X_Axis_TrackingPara(iI) = vbNullString Then X_Axis_TrackingPara(iI) = Replace((theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).name), "_", vbNullString)
                            iI = iI + 1
                        Next Tracking_Item
                    End With
                    For x_num = iI - 1 To 0 Step -1
                        x_stepsize = Abs(X_Axis_TrackingParaFrom(x_num) - X_Axis_TrackingParaTo(x_num)) / (Xaxis_index - 1)
                        Call ValueResolution(X_Axis_TrackingParaFrom(x_num), X_Axis_TrackingParaTo(x_num), x_stepsize)
                        x_TrackingInfo = X_Axis_TrackingPara(x_num) & "," & X_Axis_TrackingParaFrom(x_num) & "," & X_Axis_TrackingParaTo(x_num) & "," & x_stepsize & ";" & x_TrackingInfo
                    Next x_num
                    x_TrackingInfo = mid(x_TrackingInfo, 1, Len(x_TrackingInfo) - 1)
                    theexec.Datalog.WriteComment " X-Asix TraickingPin Info => " & x_TrackingInfo
                End If
            Case 1
                ReDim Y_Axis_Val(Yaxis_index - 1)
                If Y_Tracking_Point <> 0 Then 'extract Y Axis tracking pin, either ApplyToPins or ParameterName
                    ReDim Y_Axis_TrackingPara(Y_Tracking_Point - 1)
                    ReDim Y_Axis_TrackingParaFrom(Y_Tracking_Point - 1)
                    ReDim Y_Axis_TrackingParaTo(Y_Tracking_Point - 1)
                    With theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters
                        For Each Tracking_Item In .list
                            Y_Axis_TrackingPara(jj) = Replace((theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).ApplyTo.Pins), "_", vbNullString)
                            Y_Axis_TrackingParaFrom(jj) = theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).range.from
                            Y_Axis_TrackingParaTo(jj) = theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).range.to
                            If Y_Axis_TrackingPara(jj) = vbNullString Then Y_Axis_TrackingPara(jj) = Replace((theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).name), "_", vbNullString)
                            jj = jj + 1
                        Next Tracking_Item
                    End With
                    For y_num = jj - 1 To 0 Step -1
                        y_stepsize = Abs(Y_Axis_TrackingParaFrom(y_num) - Y_Axis_TrackingParaTo(y_num)) / (Yaxis_index - 1)
                        Call ValueResolution(Y_Axis_TrackingParaFrom(y_num), Y_Axis_TrackingParaTo(y_num), y_stepsize)
                        y_TrackingInfo = Y_Axis_TrackingPara(y_num) & "," & Y_Axis_TrackingParaFrom(y_num) & "," & Y_Axis_TrackingParaTo(y_num) & "," & y_stepsize & ";" & y_TrackingInfo
                    Next y_num
                    y_TrackingInfo = mid(y_TrackingInfo, 1, Len(y_TrackingInfo) - 1)
                    theexec.Datalog.WriteComment " Y-Asix TraickingPin Info => " & y_TrackingInfo
                End If
            Case 2
                ReDim Z_Axis_Val(Zaxis_index - 1)
                If Z_Tracking_Point <> 0 Then 'extract Z Axis tracking pin, either ApplyToPins or ParameterName
                    ReDim Z_Axis_TrackingPara(Z_Tracking_Point - 1)
                    ReDim Z_Axis_TrackingParaFrom(Z_Tracking_Point - 1)
                    ReDim Z_Axis_TrackingParaTo(Z_Tracking_Point - 1)
                    With theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters
                        For Each Tracking_Item In .list
                            Z_Axis_TrackingPara(kk) = Replace((theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).ApplyTo.Pins), "_", vbNullString)
                            Z_Axis_TrackingParaFrom(kk) = theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).range.from
                            Z_Axis_TrackingParaTo(kk) = theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).range.to
                            If Z_Axis_TrackingPara(kk) = vbNullString Then Z_Axis_TrackingPara(kk) = Replace((theexec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.item(Tracking_Item).name), "_", vbNullString)
                            kk = kk + 1
                        Next Tracking_Item
                    End With
                    For z_num = kk - 1 To 0 Step -1
                        z_stepsize = Abs(Z_Axis_TrackingParaFrom(z_num) - Z_Axis_TrackingParaTo(z_num)) / (Zaxis_index - 1)
                        Call ValueResolution(Z_Axis_TrackingParaFrom(z_num), Z_Axis_TrackingParaTo(z_num), z_stepsize)
                        z_TrackingInfo = Z_Axis_TrackingPara(z_num) & "," & Z_Axis_TrackingParaFrom(z_num) & "," & Z_Axis_TrackingParaTo(z_num) & "," & z_stepsize & ";" & z_TrackingInfo
                    Next z_num
                    z_TrackingInfo = mid(z_TrackingInfo, 1, Len(z_TrackingInfo) - 1)
                    theexec.Datalog.WriteComment " Z-Asix TraickingPin Info => " & z_TrackingInfo
                End If
        End Select
        axis_pin(curr_axis) = Replace((theexec.DevChar.Setups(DevSetupName).Shmoo.Axes((curr_axis)).ApplyTo.Pins), "_", vbNullString)
        If axis_pin(curr_axis) = "" Then
            axis_pin(curr_axis) = Replace((theexec.DevChar.Setups(DevSetupName).Shmoo.Axes((curr_axis)).Parameter.name), "_", vbNullString)
        End If
        
Next curr_axis


If X_Tracking_Point <> 0 Then ReDim x_val_tracking(X_Tracking_Point - 1)
If Y_Tracking_Point <> 0 Then ReDim y_val_tracking(Y_Tracking_Point - 1)
If Z_Tracking_Point <> 0 Then ReDim z_val_tracking(Z_Tracking_Point - 1)



InstName = theexec.DataManager.instancename
tmpstr = Split(InstName, "_")

tmpStr1 = Split(InstName, "_")

If InStr(tmpStr1(0), "L") > 0 Or InStr(tmpStr1(0), "H") > 0 Then
    TmpVal = InStr(tmpStr1(0), "L")
    tmpStr1(0) = mid(tmpStr1(0), 1, TmpVal) '  "DFTL" H or "MCL" H
    InstName_L = Join(tmpStr1, "_")
    tmpStr2 = Split(InstName, "_")
    tmpStr2(0) = mid(tmpStr2(0), 1, TmpVal - 1) & right(tmpStr2(0), 1) '  H
    InstName_H = Join(tmpStr2, "_")
Else
    InstName_L = InstName
    InstName_H = InstName
End If

k = 0
j = 0
ShmooResult = vbNullString

tmp_Tnum = g_TestNum
For Each site In theexec.sites
  ''  =============================================================================================================================================
  ''  X-axis
    If (X_dimemsion = False And Y_dimemsion = False And Z_dimemsion = False) Or X_dimemsion = True Then
        theexec.Datalog.WriteComment "  [ X-Axis ] "
        For i = 0 To MaxArrIndex - 1
            X_Axis_Val(j) = g_ShmooResult.Axis_CurrPoint(i).X_axis(site)
            j = j + 1
            ShmooResult = ShmooResult & g_ShmooResult.Axis_CurrPoint(i).CurrResult(site)
            x_val = g_ShmooResult.Axis_CurrPoint(i).X_axis(site)
            y_val = g_ShmooResult.Axis_CurrPoint(i).Y_axis(site)
            z_val = g_ShmooResult.Axis_CurrPoint(i).Z_axis(site)
            Call ValueResolution(x_val, y_val, z_val)
            CurrPointNum = i
            Call Multi_Axis_ResultName_3D(CurrPointNum, axis_pin, X_result, X_Tname_result, x_val, x_val_tracking, X_Axis_TrackingPara, Y_result, Y_Tname_result, y_val, y_val_tracking, Y_Axis_TrackingPara, Z_result, Z_Tname_result, z_val, z_val_tracking, Z_Axis_TrackingPara, "X", site)
            If CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site)) = "*" Then g_ShmooResult.Axis_CurrPoint(i).CurrResult(site) = "AP"
            If CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site)) = "~" Then g_ShmooResult.Axis_CurrPoint(i).CurrResult(site) = "AF"
    
            If i = (Xaxis_index - 1) + k Then
                Call ShmooResultPF(ShmooResult, Lvcc, HVCC, X_Axis_Val)
                Select Case ShmooResult
                    Case "9999"
                        ' still print out the latest point
                        theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                        If UCase(right(tmpstr(0), 2)) = "LH" Then
                            If RangeSeq(0) = True Then  'X-axis small--->large
                                theexec.Flow.TestLimit resultVal:=-9999, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                                theexec.Flow.TestLimit resultVal:=9999, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            Else
                                theexec.Flow.TestLimit resultVal:=-9999, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                                theexec.Flow.TestLimit resultVal:=9999, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            End If
                        ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                            If RangeSeq(2) = True Then  'X-axis small--->large
                                theexec.Flow.TestLimit resultVal:=9999, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            Else
                                theexec.Flow.TestLimit resultVal:=9999, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            End If
                        ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                            If RangeSeq(0) = True Then  'X-axis small--->large
                                theexec.Flow.TestLimit resultVal:=-9999, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                 g_TestNum = g_TestNum + 1
                           Else
                                theexec.Flow.TestLimit resultVal:=-9999, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            End If
                        Else
                            If RangeSeq(0) = True Then  'X-axis small--->large
                                theexec.Flow.TestLimit resultVal:=-9999, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                                theexec.Flow.TestLimit resultVal:=9999, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            Else
                                theexec.Flow.TestLimit resultVal:=-9999, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                                theexec.Flow.TestLimit resultVal:=9999, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            End If
                        End If
                    Case "7777"
                        ' still print out the latest point
                        theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                        If UCase(right(tmpstr(0), 2)) = "LH" Then
                            If RangeSeq(0) = True Then  'X-axis small--->large
                                theexec.Flow.TestLimit resultVal:=-7777, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                                theexec.Flow.TestLimit resultVal:=7777, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            Else
                                theexec.Flow.TestLimit resultVal:=-7777, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                                theexec.Flow.TestLimit resultVal:=7777, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            End If
                        ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                            If RangeSeq(0) = True Then  'X-axis small--->large
                                theexec.Flow.TestLimit resultVal:=7777, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            Else
                                theexec.Flow.TestLimit resultVal:=7777, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            End If
                        ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                            If RangeSeq(0) = True Then  'X-axis small--->large
                                theexec.Flow.TestLimit resultVal:=-7777, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            Else
                                theexec.Flow.TestLimit resultVal:=-7777, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            End If
                        Else
                            If RangeSeq(0) = True Then  'X-axis small--->large
                                theexec.Flow.TestLimit resultVal:=-7777, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                                theexec.Flow.TestLimit resultVal:=7777, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            Else
                                theexec.Flow.TestLimit resultVal:=-7777, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                                theexec.Flow.TestLimit resultVal:=7777, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            End If
                        End If
                    Case "5555"
                        ' still print out the latest point
                        theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                        If UCase(right(tmpstr(0), 2)) = "LH" Then
                            If RangeSeq(0) = True Then  'X-axis small--->large
                                theexec.Flow.TestLimit resultVal:=-5555, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                                theexec.Flow.TestLimit resultVal:=5555, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            Else
                                theexec.Flow.TestLimit resultVal:=-5555, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                                theexec.Flow.TestLimit resultVal:=5555, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            End If
                        ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                            If RangeSeq(0) = True Then  'X-axis small--->large
                                theexec.Flow.TestLimit resultVal:=5555, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            Else
                                theexec.Flow.TestLimit resultVal:=5555, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            End If
                        ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                            If RangeSeq(0) = True Then  'X-axis small--->large
                                theexec.Flow.TestLimit resultVal:=-5555, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                 g_TestNum = g_TestNum + 1
                            Else
                                theexec.Flow.TestLimit resultVal:=-5555, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            End If
                        Else
                            If RangeSeq(0) = True Then  'X-axis small--->large
                                theexec.Flow.TestLimit resultVal:=-5555, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                                theexec.Flow.TestLimit resultVal:=5555, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            Else
                                theexec.Flow.TestLimit resultVal:=-5555, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                                theexec.Flow.TestLimit resultVal:=5555, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            End If
                        End If
                    Case Else ' search LVCC/HVCC point
                        ' still print out the latest point
                        theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                        If UCase(right(tmpstr(0), 2)) = "LH" Then
                            If RangeSeq(0) = True Then  'X-axis small--->large
                                theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                                theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            Else
                                theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                 g_TestNum = g_TestNum + 1
                               theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                 g_TestNum = g_TestNum + 1
                            End If
                        ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                            If RangeSeq(0) = True Then  'X-axis small--->large
                                theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            Else
                                theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                 g_TestNum = g_TestNum + 1
                            End If
                        ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                            If RangeSeq(0) = True Then 'X-axis small--->large
                                theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            Else
                                theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            End If
                        Else
                            If RangeSeq(0) = True Then  'X-axis small--->large
                                theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                                theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            Else
                                theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                 g_TestNum = g_TestNum + 1
                               theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                 g_TestNum = g_TestNum + 1
                            End If
                        End If
                End Select
                ShmooResult = vbNullString
                k = k + Xaxis_index
                j = 0
            Else
                theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
            End If
    
        Next i
        k = 0 ' initial it fot multi site
    End If
    ''  =============================================================================================================================================
    ''  Y-axis
    If Y_dimemsion = True Then
        theexec.Datalog.WriteComment "  [ Y-Axis ] "
        'xx = 2
        tmp_count1 = 0
        For xx = 1 To Xaxis_index
            For i = tmp_count1 To MaxArrIndex * Xaxis_index - 1 Step Xaxis_index
                If i > xx * MaxArrIndex - 1 Then
                    tmp_count1 = i
                    Exit For
                Else
                    tmp_x = i - (xx - 1) * (MaxArrIndex - 1)
                    ShmooResult = ShmooResult & g_ShmooResult.Axis_CurrPoint(tmp_x).CurrResult(site)
                    Y_Axis_Val(j) = g_ShmooResult.Axis_CurrPoint(tmp_x).Y_axis(site)
                    j = j + 1
                    x_val = g_ShmooResult.Axis_CurrPoint(tmp_x).X_axis(site)
                    y_val = g_ShmooResult.Axis_CurrPoint(tmp_x).Y_axis(site)
                    z_val = g_ShmooResult.Axis_CurrPoint(tmp_x).Z_axis(site)
                    Call ValueResolution(x_val, y_val, z_val)
                    CurrPointNum = tmp_x
                    Call Multi_Axis_ResultName_3D(CurrPointNum, axis_pin, X_result, X_Tname_result, x_val, x_val_tracking, X_Axis_TrackingPara, Y_result, Y_Tname_result, y_val, y_val_tracking, Y_Axis_TrackingPara, Z_result, Z_Tname_result, z_val, z_val_tracking, Z_Axis_TrackingPara, "Y", site)
                    If CStr(g_ShmooResult.Axis_CurrPoint(tmp_x).CurrResult(site)) = "*" Then g_ShmooResult.Axis_CurrPoint(i).CurrResult(site) = "AP"
                    If CStr(g_ShmooResult.Axis_CurrPoint(tmp_x).CurrResult(site)) = "~" Then g_ShmooResult.Axis_CurrPoint(i).CurrResult(site) = "AF"
                End If
        'Debug.Print tmp_x
                If j = Yaxis_index Then
    
                    Call ShmooResultPF(ShmooResult, Lvcc, HVCC, Y_Axis_Val)
                    Select Case ShmooResult
                        Case "9999"
                            ' still print out the latest point
                            If i > MaxArrIndex - 1 Then
                                theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_x).CurrResult(site))
                            Else
                                theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                            End If
    
                            If UCase(right(tmpstr(0), 2)) = "LH" Then
                                If RangeSeq(1) = True Then 'Y-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-9999, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=9999, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-9999, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=9999, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                               End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                                If RangeSeq(1) = True Then 'Y-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=9999, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=9999, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                                If RangeSeq(1) = True Then 'Y-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-9999, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-9999, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            Else
                                If RangeSeq(1) = True Then 'Y-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-9999, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=9999, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-9999, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=9999, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                        Case "7777"
                            ' still print out the latest point
                            If i > MaxArrIndex - 1 Then
                                theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_x).CurrResult(site))
                            Else
                                theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                            End If
                            If UCase(right(tmpstr(0), 2)) = "LH" Then
                                If RangeSeq(1) = True Then 'Y-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-7777, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=7777, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-7777, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=7777, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                                If RangeSeq(1) = True Then 'Y-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=7777, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=7777, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                                If RangeSeq(1) = True Then 'Y-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-7777, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-7777, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            Else
                                If RangeSeq(1) = True Then 'Y-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-7777, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=7777, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-7777, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=7777, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                        Case "5555"
                            ' still print out the latest point
                            If i > MaxArrIndex - 1 Then
                                theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_x).CurrResult(site))
                            Else
                                theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                            End If
                            If UCase(right(tmpstr(0), 2)) = "LH" Then
                                If RangeSeq(1) = True Then 'Y-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-5555, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=5555, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-5555, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=5555, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                                If RangeSeq(1) = True Then 'Y-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=5555, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=5555, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                                If RangeSeq(1) = True Then 'Y-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-5555, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-5555, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            Else
                                If RangeSeq(1) = True Then 'Y-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-5555, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=5555, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-5555, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=5555, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                        Case Else ' search LVCC/HVCC point
                            ' still print out the latest point
                            If i > MaxArrIndex - 1 Then
                                theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_x).CurrResult(site))
                            Else
                                theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                            End If
                            If UCase(right(tmpstr(0), 2)) = "LH" Then
                                If RangeSeq(1) = True Then 'Y-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                                If RangeSeq(1) = True Then 'Y-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                               End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                                If RangeSeq(1) = True Then 'Y-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                               End If
                            Else
                                If RangeSeq(1) = True Then 'Y-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                    End Select
                    ShmooResult = vbNullString
                    j = 0
                Else
                    If i > MaxArrIndex - 1 Then
                        theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_x).CurrResult(site))
                    Else
                        theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                    End If
                End If
    
            Next i
        '    i = tmp_count1
        Next xx
        k = 0 ' initial it fot multi site
    End If
    '''''  =============================================================================================================================================
    '''''  Z-axis
    If Z_dimemsion = True Then
        theexec.Datalog.WriteComment "  [ Z-Axis ] "
        'yy = 2
        tmp_count = 0
        For yy = 1 To Xaxis_index * Yaxis_index
            For i = tmp_count To (MaxArrIndex * Xaxis_index * Yaxis_index) - 1 Step (Xaxis_index * Yaxis_index)
    
                If i > yy * MaxArrIndex - 1 Then
                    tmp_count = i
                    Exit For
                Else
                    tmp_y = i - (yy - 1) * (MaxArrIndex - 1)
                    ShmooResult = ShmooResult & g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site)
                    Z_Axis_Val(j) = g_ShmooResult.Axis_CurrPoint(tmp_y).Z_axis(site)
                    j = j + 1
                    x_val = g_ShmooResult.Axis_CurrPoint(tmp_y).X_axis(site)
                    y_val = g_ShmooResult.Axis_CurrPoint(tmp_y).Y_axis(site)
                    z_val = g_ShmooResult.Axis_CurrPoint(tmp_y).Z_axis(site)
                    Call ValueResolution(x_val, y_val, z_val)
                    CurrPointNum = tmp_y
                    Call Multi_Axis_ResultName_3D(CurrPointNum, axis_pin, X_result, X_Tname_result, x_val, x_val_tracking, X_Axis_TrackingPara, Y_result, Y_Tname_result, y_val, y_val_tracking, Y_Axis_TrackingPara, Z_result, Z_Tname_result, z_val, z_val_tracking, Z_Axis_TrackingPara, "Z", site)
                    If CStr(g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site)) = "*" Then g_ShmooResult.Axis_CurrPoint(i).CurrResult(site) = "AP"
                    If CStr(g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site)) = "~" Then g_ShmooResult.Axis_CurrPoint(i).CurrResult(site) = "AF"
                End If
        'Debug.Print tmp_y
                If j = Zaxis_index Then
                    Call ShmooResultPF(ShmooResult, Lvcc, HVCC, Z_Axis_Val)
                    Select Case ShmooResult
                        Case "9999"
                            ' still print out the latest point
                            If i > MaxArrIndex - 1 Then
                                theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site))
                            Else
                                theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                            End If
    
                            If UCase(right(tmpstr(0), 2)) = "LH" Then
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-9999, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=9999, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-9999, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=9999, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=9999, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=9999, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                               End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-9999, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-9999, hiVal:=X_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            Else
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-9999, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=9999, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-9999, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=9999, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                        Case "7777"
                            ' still print out the latest point
                            If i > MaxArrIndex - 1 Then
                                theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site))
                            Else
                                theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                            End If
                            If UCase(right(tmpstr(0), 2)) = "LH" Then
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-7777, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=7777, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-7777, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=7777, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=7777, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=7777, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-7777, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-7777, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            Else
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-7777, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=7777, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-7777, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=7777, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                        Case "5555"
                            ' still print out the latest point
                            If i > MaxArrIndex - 1 Then
                                theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site))
                            Else
                                theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                            End If
                            If UCase(right(tmpstr(0), 2)) = "LH" Then
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-5555, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=5555, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-5555, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=5555, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=5555, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=5555, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-5555, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-5555, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            Else
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=-5555, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=5555, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=-5555, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=5555, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                        Case Else ' search LVCC/HVCC point
                            ' still print out the latest point
                            If i > MaxArrIndex - 1 Then
                                theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site))
                            Else
                                theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                            End If
                            If UCase(right(tmpstr(0), 2)) = "LH" Then
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "H" Then
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(right(tmpstr(0), 1)) = "L" Then
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            Else
                                If RangeSeq(2) = True Then 'Z-axis small--->large
                                    theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    theexec.Flow.TestLimit resultVal:=Lvcc, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_L & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    theexec.Flow.TestLimit resultVal:=HVCC, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_H & "_" & X_Tname_result & "_" & Y_Tname_result & "_" & Z_Tname_result, tNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                    End Select
                    ShmooResult = vbNullString
                    j = 0
                Else
                    If i > MaxArrIndex - 1 Then
                        theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site))
                    Else
                        theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & X_result & "_" & Y_result & "_" & Z_result & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                    End If
                End If
    
            Next i
        Next yy
        k = 0 ' initial it fot multi site
        '''
        '''
    End If
    g_TestNum = tmp_Tnum
Next site

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Print3DShmooInfo_XYZ")
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Feature_Check_in_Functional_T_char()
On Error GoTo errHandler

    Dim patt_maxrow As Integer
    Dim patt_maxcol As Integer
    Dim index_CharPat As Integer
    Dim PasetsString As String
    Dim split_Pasets() As String
    Dim Split_PasetsCheckString() As String
    Dim PasetsCheckString As String
    Dim dic_Patsets_All As New Dictionary
    Dim dic_ErrorPatsets As New Dictionary
    Dim PatSets_in_char As New Dictionary
    Dim ErrorPatsetsCheckString As String
    Dim i, j, x, y As Integer
    Dim m As Integer:: m = 0
    Dim n As Integer:: n = 0
    Dim Inst_SheetName() As String
    Dim List_Error_Inst() As Error_Inst
    Dim split_Pmode() As String
    Dim sheettypeSelected As String:: sheettypeSelected = DMGR_SHEET_TYPE_TESTINSTANCESSHEET
    Dim Function_name As String:: Function_name = "Functional_T_char"
    Dim MaxRow As Integer
    Dim maxcol As Integer
    Dim Error_InstName() As String
    Dim F_wait As Boolean:: F_wait = False
    Dim patt As String
    ''Check if pattern exists in PatSets_all sheet or not
    Dim Pat_Path As String
    Dim File_path As String
    Dim split_FilePath_in_PatSetAll() As String
    Dim split_FilePath() As String
    Dim Pat_Latest As String
    Dim Flag_old_version As Boolean
    
'''    Dim char_funcTime As Double:: char_funcTime = TheExec.Timer(0)   'check test time of this func
    

    '' User dedined project folder path, MUST VERIFIED IT BEFORE USEDDDDDDD.
    If theexec.TesterMode = testModeOffline Then
        Pat_Path = "O:\TP-to-C651\Staten\SVN_user\Share_trunk\Pattern"
    Else
        Pat_Path = "U:\TP-to-C651\Staten\SVN_user\Share_trunk\Pattern"
    End If
    
    Application.ScreenUpdating = False
    Inst_SheetName = theexec.job.GetSheetNamesOfType(sheettypeSelected)

    ReDim Preserve List_Error_Inst(n) As Error_Inst
    For i = 0 To UBound(Inst_SheetName)
        Worksheets(Inst_SheetName(i)).Activate
        MaxRow = Worksheets(Inst_SheetName(i)).UsedRange.Rows.Count
        maxcol = Worksheets(Inst_SheetName(i)).UsedRange.Columns.Count
        For x = 5 To MaxRow
            If Cells(x, 2) = "" Then
                MaxRow = x  ''Update maxrow to avoid no activite instance
                Exit For
            End If
        Next x
        '' Search for functional_T_char module
        If Cells(4, 4) = "Name" Then
            For x = 5 To MaxRow
                If Cells(x, 4) = Function_name Then
                    If Cells(x, 67) <> "" Then
                        split_Pmode = Split(Cells(x, 67), ":")
                        If theexec.Specs.DC.Categories.Contains(split_Pmode(0)) = False Then
                        '' Check Pmode if it exists in DC spec catagory
                            n = n + 1
                            ReDim Preserve List_Error_Inst(n) As Error_Inst
                            List_Error_Inst(n).Error_status.Pmode_Naming_error = True
                            List_Error_Inst(n).instance_name = Cells(x, 2)
                            List_Error_Inst(n).ErrorPmode = split_Pmode(0)
                        End If
                        If Cells(x, 67) Like "*ERT*" Or Cells(x, 67) Like "*NRT*" Or Cells(x, 67) Like "*SRT*" Then
                        '' Check retension item has right format in Power_Run_Scenario and Wait column
                            If UCase(Cells(x, 59)) <> "INIT_NV_PL_NV" Then
                                n = n + 1
                                ReDim Preserve List_Error_Inst(n) As Error_Inst
                                List_Error_Inst(n).Error_status.Ret_Scenario_error = True
                                List_Error_Inst(n).instance_name = Cells(x, 2)
                            End If  '' End If UCase(Cells(x, 59)) <> "INIT_NV_PL_NV"
    
                            Dim temp_split_wait() As String
                            temp_split_wait() = Split(Cells(x, 60), ",")
                            For j = 0 To UBound(temp_split_wait)
                                If temp_split_wait(j) <> "" Then
                                    F_wait = True
                                    Exit For
                                End If
                            Next j
                            If F_wait = False Then
                                n = n + 1
                                ReDim Preserve List_Error_Inst(n) As Error_Inst
                                List_Error_Inst(n).Error_status.Ret_WaitTime_error = True
                                List_Error_Inst(n).instance_name = Cells(x, 2)
                            End If  '' End If F_wait = False
                            F_wait = False
                        End If  '' End If Cells(x, 67) Like "*ERT*"
                    End If  '' End If Cells(x, 67) <> ""
                    
                    ''Get char pattern
                    For y = 44 To 58    ''Init_Patt1~10,PayLoad_Patt1~5
                        If Cells(x, y) <> "" Then
                            Split_PasetsCheckString = Split(UCase(Cells(x, y)), ",")
                            For m = 0 To UBound(Split_PasetsCheckString)
                                PasetsCheckString = Split(UCase(Cells(x, y)), ":")(0)
                                If PatSets_in_char.Exists(PasetsCheckString) = False Then
                                    index_CharPat = index_CharPat + 1
                                    PatSets_in_char.Add PasetsCheckString, index_CharPat
                                End If
                            Next m
                        End If '' End If Cells(x, y) <> ""
                    Next y
                End If  ''End If Cells(x, 4) = Function_name
            Next x
        End If  '' End If Cells(4, 4) = "Name"
    Next i
    
    ''Check if pattern exists in PatSets_all sheet or not
    ''Get PatSets_all Sheets as dictionary
    Worksheets("PatSets_All").Activate
    patt_maxrow = Worksheets("PatSets_All").UsedRange.Rows.Count
    patt_maxcol = Worksheets("PatSets_All").UsedRange.Columns.Count
    For x = 4 To patt_maxrow
        If Cells(x, 2) = "" Then
            ''Update maxrow to avoid no activite Patsets
            patt_maxrow = x - 1
            Exit For
        End If
    Next x
    For x = 4 To patt_maxrow
        dic_Patsets_All.Add UCase(Cells(x, 2)), UCase(Cells(x, 5))
    Next x
    
    For index_CharPat = 0 To PatSets_in_char.Count - 1
        If dic_Patsets_All.Exists(PatSets_in_char.Keys(index_CharPat)) = True Then
            ''Check pattern file version to update the latest version in PatSets_all sheet
            Dim Ver As Integer
            Dim split_FileName() As String
            split_FileName = Split(Split(dic_Patsets_All(PatSets_in_char.Keys(index_CharPat)), ":")(1), "_")
            Ver = split_FileName(UBound(split_FileName) - 2)
            split_FilePath_in_PatSetAll = Split(dic_Patsets_All(PatSets_in_char.Keys(index_CharPat)), "\")
            File_path = Pat_Path & "\" & split_FilePath_in_PatSetAll(UBound(split_FilePath_in_PatSetAll()) - 1) & "\"
            Do
                Pat_Latest = Split(Dir(File_path & PatSets_in_char.Keys(index_CharPat) & "_" & Ver & "*"), ".")(0)
                Ver = Ver + 1
            Loop Until Dir(File_path & PatSets_in_char.Keys(index_CharPat) & "_" & Ver & "*") = ""
            If Split(dic_Patsets_All(PatSets_in_char.Keys(index_CharPat)), ":")(1) <> Pat_Latest Then
                Flag_old_version = True
                If isDebugMode = True Then theexec.AddOutput PatSets_in_char.Keys(index_CharPat) & " has the newer version " & Pat_Latest & ", please check the PatSets_all sheet."
            End If
        Else
            n = n + 1
            ReDim Preserve List_Error_Inst(n) As Error_Inst
            List_Error_Inst(n).Error_status.Patsets_error = True
            'List_Error_Inst(n).Instance_name = Cells(x, 2)
            List_Error_Inst(n).ErrorPatsets = PatSets_in_char.Keys(index_CharPat)
            'ErrorPatsetsCheckString = List_Error_Inst(n).ErrorPatsets
            'Debug.Print List_Error_Inst(n).ErrorPatsets
'            If dic_ErrorPatsets.Exists(ErrorPatsetsCheckString) = True Then
'            Else
'                List_Error_Inst(n).Error_status.Patsets_error_print = True
'                dic_ErrorPatsets.Add ErrorPatsetsCheckString, n
'                List_Error_Inst(n).ErrorPatsetsPrint = ErrorPatsetsCheckString
'                'Debug.Print List_Error_Inst(n).ErrorPatsetsPrint
'            End If ''End  If ErrorPatsetsCheck Exists
        End If '' End If PatsetsArray Check
    Next index_CharPat

    '' Output the Feature check result
    If UBound(List_Error_Inst) = 0 Then
        If Flag_old_version = False Then
            If isDebugMode = True Then theexec.AddOutput "All Patterns in Functional_T_char have defined to the latest version!"
            If isDebugMode = True Then theexec.AddOutput "Feature_Check_in_Functional_T_char success !"
        Else
            If isDebugMode = True Then theexec.AddOutput "Feature_Check_in_Functional_T_char have issue !"
        End If
    Else
        For n = 1 To UBound(List_Error_Inst)
            With List_Error_Inst(n).Error_status
                If .Pmode_Naming_error = True Then
                    If isDebugMode = True Then theexec.AddOutput "Error Pmode " & List_Error_Inst(n).ErrorPmode & " in " & List_Error_Inst(n).instance_name & ", please check it out."
                ElseIf .Ret_Scenario_error = True Then
                    If isDebugMode = True Then theexec.AddOutput "In Retension test " & List_Error_Inst(n).instance_name & ", the Power_Run_Scenario should be 'INIT_NV_PL_NV'. Please check if it is for the special case or not."
                ElseIf .Ret_WaitTime_error = True Then
                    If isDebugMode = True Then theexec.AddOutput "No wait time in Retension test " & List_Error_Inst(n).instance_name & ", please check the wait time setting."
                 ElseIf .Patsets_error = True Then
                    If isDebugMode = True Then theexec.AddOutput "Unknown pattern " & List_Error_Inst(n).ErrorPatsets & ", please check the pattern exists ot not."
                End If
            End With
        Next n
        If isDebugMode = True Then theexec.AddOutput "Feature_Check_in_Functional_T_char have issue !"
    End If
        
    Application.ScreenUpdating = True
    
'''    char_funcTime = TheExec.Timer(char_funcTime)   'check test time of this func
'''    TheExec.AddOutput " Test time : " & char_funcTime   'check test time of this func
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Feature_Check_in_Functional_T_char")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function RetentionMeasureParser(MeasPins As String)
On Error GoTo errHandler
    Dim i As Long
    Dim j As Long
    Dim PinsInfoAry() As String
    Dim PinAry() As String
    Dim PinCnt As Long
    Dim tmpPinCnt As Long
    Dim InstName As String
    Dim DCVIPins As String
    Dim DCVSPins As String
    Dim DigitalPins As String
    Dim MeasureCond() As String
    Dim DCVIInstName As String
    Dim DCVSInstName As String
    Dim DigitalPinsInstName As String
    Dim DCVIForceValue As String
    Dim DCVSForceValue As String
    Dim DigitalForceValue As String
    Dim DCVIMeasureRange As String
    Dim DCVSMeasureRange As String
    Dim DigitalMeasureRange As String
    Dim PinFoece As String
    Dim PinMeasure As String
    Dim PinName As String
    Dim ForceValueArr() As String
    Dim MeasureRangeArr() As String
    Dim ForceValueFlag As Boolean
    
    PinsInfoAry = Split(MeasPins, "|")
    ReDim glb_RET_MeasAry(UBound(PinsInfoAry))  'Ubound is 2, MeasureCase V/I/F
    
    For i = 0 To UBound(PinsInfoAry)
        ForceValueFlag = False
        If PinsInfoAry(i) <> "" Then
            If InStr(PinsInfoAry(i), "@") <> 0 Then 'E.g. "VDD_SOC@0@0.2||"
                MeasureCond = Split(PinsInfoAry(i), "@")
                If UBound(MeasureCond) <> 2 Then
                    theexec.Datalog.WriteComment "<Warning> fill in the incorrect @ size."
                End If
                PinName = MeasureCond(0)
                ForceValueArr = Split(MeasureCond(1), ",")
                MeasureRangeArr = Split(MeasureCond(2), ",")
                ForceValueFlag = True
            Else
                PinName = PinsInfoAry(i)
                PinFoece = vbNullString
                PinMeasure = vbNullString
            End If
            
            theexec.DataManager.DecomposePinList PinName, PinAry, PinCnt
            If ForceValueFlag Then
                If PinCnt <> UBound(ForceValueArr) + 1 And PinCnt <> UBound(MeasureRangeArr) + 1 Then
                    theexec.Datalog.WriteComment "<Warning> size of pins, force values or measure range aren't match."
                End If
            End If
            For tmpPinCnt = 0 To UBound(PinAry)
                If ForceValueFlag = True Then
                    PinFoece = ForceValueArr(tmpPinCnt)
                    PinMeasure = MeasureRangeArr(tmpPinCnt)
                End If
                If theexec.DataManager.ChannelType(PinAry(tmpPinCnt)) <> "N/C" Then
                    InstName = GetInstrument(PinAry(tmpPinCnt), 0)
                    Select Case InstName
                        Case "DC-07", "DC-30"
                            If DCVIPins = "" Then
                                DCVIPins = PinAry(tmpPinCnt)
                                DCVIInstName = InstName
                                DCVIForceValue = PinFoece
                                DCVIMeasureRange = PinMeasure
                            Else
                                DCVIPins = DCVIPins & "," & PinAry(tmpPinCnt)
                                DCVIInstName = DCVIInstName & "," & InstName
                                DCVIForceValue = DCVIForceValue & "," & PinFoece
                                DCVIMeasureRange = DCVIMeasureRange & "," & PinMeasure
                            End If
                        Case "VHDVS", "HexVS", "VSM", "VS-5A", "VS-800mA"
                            If DCVSPins = "" Then
                                DCVSPins = PinAry(tmpPinCnt)
                                DCVSInstName = InstName
                                DCVSForceValue = PinFoece
                                DCVSMeasureRange = PinMeasure
                            Else
                                DCVSPins = DCVSPins & "," & PinAry(tmpPinCnt)
                                DCVSInstName = DCVSInstName & "," & InstName
                                DCVSForceValue = DCVSForceValue & "," & PinFoece
                                DCVSMeasureRange = DCVSMeasureRange & "," & PinMeasure
                            End If
                        Case "HSD-U", "HSDP"
                            If DigitalPins = "" Then
                                DigitalPins = PinAry(tmpPinCnt)
                                DigitalPinsInstName = InstName
                                DigitalForceValue = PinFoece
                                DigitalMeasureRange = PinMeasure
                            Else
                                DigitalPins = DigitalPins & "," & PinAry(tmpPinCnt)
                                DigitalPinsInstName = DigitalPinsInstName & "," & InstName
                                DigitalForceValue = DigitalForceValue & "," & PinFoece
                                DigitalMeasureRange = DigitalMeasureRange & "," & PinMeasure
                            End If
                        Case Else
                            theexec.Datalog.WriteComment " Measurement Pin Type not Support !!! "
                    End Select
                End If
            Next tmpPinCnt
            With glb_RET_MeasAry(i)
                .DCVI_Pins.PinNameAry = Split(DCVIPins, ",")
                .DCVI_Pins.InstrumentType = Split(DCVIInstName, ",")
                If DCVIForceValue = "" And DCVIPins <> "" Then
                    ReDim .DCVI_Pins.ForceValueAry(UBound(.DCVI_Pins.PinNameAry))
                Else
                    .DCVI_Pins.ForceValueAry = Split(DCVIForceValue, ",")
                End If
                If DCVIMeasureRange = "" And DCVIPins <> "" Then
                    ReDim .DCVI_Pins.MeasureRangeAry(UBound(.DCVI_Pins.PinNameAry))
                Else
                    .DCVI_Pins.MeasureRangeAry = Split(DCVIMeasureRange, ",")
                End If
                If DCVIPins <> "" Then ReDim .DCVI_Pins.RestoreGateOffAry(UBound(.DCVI_Pins.PinNameAry))
                
                .DCVS_Pins.PinNameAry = Split(DCVSPins, ",")
                .DCVS_Pins.InstrumentType = Split(DCVSInstName, ",")
                If DCVSForceValue = "" And DCVSPins <> "" Then
                    ReDim .DCVS_Pins.ForceValueAry(UBound(.DCVS_Pins.PinNameAry))
                Else
                    .DCVS_Pins.ForceValueAry = Split(DCVSForceValue, ",")
                End If
                If DCVSMeasureRange = "" And DCVSPins <> "" Then
                    ReDim .DCVS_Pins.MeasureRangeAry(UBound(.DCVS_Pins.PinNameAry))
                Else
                    .DCVS_Pins.MeasureRangeAry = Split(DCVSMeasureRange, ",")
                End If
                If DCVSPins <> "" Then ReDim .DCVS_Pins.RestoreGateOffAry(UBound(.DCVS_Pins.PinNameAry))
                
                .digital_pins.PinNameAry = Split(DigitalPins, ",")
                .digital_pins.InstrumentType = Split(DigitalPinsInstName, ",")
                If DigitalForceValue = "" And DigitalPins <> "" Then
                    ReDim .digital_pins.ForceValueAry(UBound(.digital_pins.PinNameAry))
                Else
                    .digital_pins.ForceValueAry = Split(DigitalForceValue, ",")
                End If
                If DCVSMeasureRange = "" And DigitalPins <> "" Then
                    ReDim .digital_pins.MeasureRangeAry(UBound(.digital_pins.PinNameAry))
                Else
                    .digital_pins.MeasureRangeAry = Split(DigitalMeasureRange, ",")
                End If
                If DigitalPins <> "" Then ReDim .digital_pins.RestoreGateOffAry(UBound(.digital_pins.PinNameAry))
                
            End With
            glb_RET_MeasAry(i).DCVI_Pins.PinNameAry = Split(DCVIPins, ",")
            glb_RET_MeasAry(i).DCVS_Pins.PinNameAry = Split(DCVSPins, ",")
            glb_RET_MeasAry(i).digital_pins.PinNameAry = Split(DigitalPins, ",")
        Else
            glb_RET_MeasAry(i).DCVI_Pins.PinNameAry = Split(vbNullString, ",")
            glb_RET_MeasAry(i).DCVS_Pins.PinNameAry = Split(vbNullString, ",")
            glb_RET_MeasAry(i).digital_pins.PinNameAry = Split(vbNullString, ",")
        End If
        DCVIPins = vbNullString
        DCVSPins = vbNullString
        DigitalPins = vbNullString
        DCVIInstName = vbNullString
        DCVSInstName = vbNullString
        DigitalPinsInstName = vbNullString
        DCVIForceValue = vbNullString
        DCVSForceValue = vbNullString
        DigitalForceValue = vbNullString
        DCVIMeasureRange = vbNullString
        DCVSMeasureRange = vbNullString
        DigitalMeasureRange = vbNullString
    Next i
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "RetentionMeasureParser")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Pins_Measure_Case(ConditionString As String)
On Error GoTo errHandler
    Dim measureCase As Long 'V/I/F
    Dim PinType As Long
    For measureCase = 0 To UBound(glb_RET_MeasAry)
        With glb_RET_MeasAry(measureCase)
            If UBound(.DCVI_Pins.PinNameAry) <> -1 Then Retention_MeasDCVI measureCase, ConditionString
            If UBound(.DCVS_Pins.PinNameAry) <> -1 Then Retention_MeasDCVS measureCase, ConditionString
            If UBound(.digital_pins.PinNameAry) <> -1 Then Retention_MeasDigital measureCase, ConditionString
        End With
    Next measureCase
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Pins_Measure_Case")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Retention_MeasDCVI(MeasCase As Long, condition As String)
On Error GoTo errHandler
    Dim i As Long
    Dim PinInfoArr As MeasureCond
    Dim TestInstName As String
    Dim MeasResultV As New PinListData
    Dim MeasResultI As New PinListData
    Select Case MeasCase
        Case RET_MeasV
            TestInstName = "MeasureVoltage"
            PinInfoArr = glb_RET_MeasAry(RET_MeasV).DCVI_Pins
            For i = 0 To UBound(PinInfoArr.PinNameAry)
                With thehdw.DCVI.Pins(PinInfoArr.PinNameAry(i))
                    If PinInfoArr.ForceValueAry(i) = "" And thehdw.DCVI.Pins(PinInfoArr.PinNameAry(i)).Gate <> True Then
                        .mode = tlDCVIModeHighImpedance
                        .Voltage = pc_Def_VFI_UVI80_VoltCalmp
                        .BleederResistor = tlDCVIBleederResistorOff
                        .Current = 0
                        .Connect tlDCVIConnectHighSense
                        PinInfoArr.RestoreGateOffAry(i) = True
                    Else
                        .mode = tlDCVIModeCurrent
                        .Voltage = pc_Def_VFI_UVI80_VoltCalmp
                        If PinInfoArr.ForceValueAry(i) <> "" Then
                            .Current = CDbl(PinInfoArr.ForceValueAry(i))
                        Else
                            .Current = 0
                        End If
                        .Connect tlDCVIConnectDefault
                    End If
                    .Gate = True
                    .Meter.mode = tlDCVIMeterVoltage
                End With
                thehdw.Wait 0.001
            Next i
            MeasResultV = thehdw.DCVI.Pins(Join(PinInfoArr.PinNameAry, ",")).Meter.Read(tlStrobe, pc_Def_VFI_UVI80_ReadPoint)
            theexec.Flow.TestLimit MeasResultV, , , , , , unitVolt, , TestInstName & condition
            
            For i = 0 To UBound(PinInfoArr.PinNameAry)
                If PinInfoArr.RestoreGateOffAry(i) Then
                    thehdw.DCVI.Pins(PinInfoArr.PinNameAry(i)).Gate = False
                End If
            Next i
            
        Case RET_MeasI
            TestInstName = "MeasureCurrent"
            PinInfoArr = glb_RET_MeasAry(RET_MeasI).DCVI_Pins
            For i = 0 To UBound(PinInfoArr.PinNameAry)
                With thehdw.DCVI.Pins(PinInfoArr.PinNameAry(i))
                    If UCase(theexec.DataManager.PinType(PinInfoArr.PinNameAry(i))) <> UCase("Power") Then
                        .Gate = False
                        .mode = tlDCVIModeVoltage
                        .Voltage = 0
                    End If
                    If PinInfoArr.ForceValueAry(i) <> "" Then
                        .Voltage = CDbl(PinInfoArr.ForceValueAry(i))
                    End If
                    .VoltageRange.Autorange = True
                    .CurrentRange.Autorange = True
                    .Current = pc_Def_UVI80_Init_MeasCurrRange
                    .Connect tlDCVIConnectDefault
                    .Gate = True
                End With
                thehdw.Wait 0.001
                With thehdw.DCVI.Pins(PinInfoArr.PinNameAry(i))
                    .Meter.mode = tlDCVIMeterCurrent
                    If PinInfoArr.MeasureRangeAry(i) <> "" Then
                        .SetCurrentAndRange CDbl(PinInfoArr.MeasureRangeAry(i)), CDbl(PinInfoArr.MeasureRangeAry(i))
                    Else
                        .SetCurrentAndRange 0.2, 0.2
                    End If
                    .CurrentRange.Autorange = True
                End With
            Next i
            MeasResultI = thehdw.DCVI.Pins(Join(PinInfoArr.PinNameAry, ",")).Meter.Read(tlStrobe, 10)
            theexec.Flow.TestLimit MeasResultI, , , , , , unitAmp, , TestInstName & condition
        Case Else
    End Select
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Retention_MeasDCVI")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Retention_MeasDCVS(MeasCase As Long, condition As String)
On Error GoTo errHandler
    Dim i As Long
    Dim PinInfoArr As MeasureCond
    Dim TestInstName As String
    Dim MeasResultV As New PinListData
    Dim MeasResultI As New PinListData
    Dim ForceValue1 As Double
    Dim AnalogMeas As Boolean
    Dim MeasurePinStr As String
    Dim ReadPoint As Long
    Select Case MeasCase
        Case RET_MeasV
            TestInstName = "MeasureVoltage"
            PinInfoArr = glb_RET_MeasAry(RET_MeasV).DCVS_Pins
            For i = 0 To UBound(PinInfoArr.PinNameAry)
                With thehdw.DCVS.Pins(PinInfoArr.PinNameAry(i))
                    If PinInfoArr.InstrumentType(i) = "VS-5A" Or PinInfoArr.InstrumentType(i) = "VS-800mA" Then
                        If PinInfoArr.ForceValueAry(i) = "" Then
                            If .Gate = False Then
                                .Disconnect
                                .mode = tlDCVSModeHighImpedance
                                .Meter.mode = tlDCVSMeterVoltage
                                .Connect
                                .Gate = True
                                PinInfoArr.RestoreGateOffAry(i) = True
                            Else
                                .Meter.mode = tlDCVSMeterVoltage
                                .Connect tlDCVSConnectDefault
                            End If
                        Else
                            If CDbl(PinInfoArr.ForceValueAry(i)) > 0 Then
                                .Voltage.value = 1 + 0.5
                            Else
                                .VoltageRange.value = 5.5
                                .Voltage.value = -1 - 0.5
                            End If
                            ForceValue1 = CDbl(PinInfoArr.ForceValueAry(i))
                            .CurrentRange.value = Abs(ForceValue1)
                            .CurrentLimit.Source.FoldLimit.level.value = Abs(ForceValue1)
                            .CurrentLimit.Sink.FoldLimit.level.value = Abs(ForceValue1)
                            .Meter.mode = tlDCVSMeterVoltage
                            .Connect
                            If .Gate = False Then
                                PinInfoArr.RestoreGateOffAry(i) = True
                                .Gate = True
                            End If
                            .mode = tlDCVSModeCurrent
                        End If
                    Else
                        If PinInfoArr.InstrumentType(i) = "VHDVS" Then
                            ReadPoint = 1
                        Else
                            ReadPoint = 10
                        End If
                        .Meter.mode = tlDCVSMeterVoltage
                        .Connect
                        .Gate = True
                    End If
                End With
            Next i
            MeasResultV = thehdw.DCVS.Pins(Join(PinInfoArr.PinNameAry, ",")).Meter.Read(tlStrobe, ReadPoint)    'pc_Def_UVS256HP_ReadPoint)
            theexec.Flow.TestLimit MeasResultV, , , , , , unitVolt, , TestInstName & condition
            
            For i = 0 To UBound(PinInfoArr.PinNameAry)
                If PinInfoArr.RestoreGateOffAry(i) = True Then
                    thehdw.DCVS.Pins(PinInfoArr.PinNameAry(i)).Gate = False
                    thehdw.DCVS.Pins(PinInfoArr.PinNameAry(i)).Disconnect tlDCVSConnectDefault
                End If
            Next i
        Case RET_MeasI
            TestInstName = "MeasureCurrent"
            PinInfoArr = glb_RET_MeasAry(RET_MeasI).DCVS_Pins
            For i = 0 To UBound(PinInfoArr.PinNameAry)
                With thehdw.DCVS.Pins(PinInfoArr.PinNameAry(i))
                
                    If .Gate = False And (PinInfoArr.InstrumentType(i) = "VS-5A" Or PinInfoArr.InstrumentType(i) = "VS-800mA") Then
                        .Disconnect tlDCVSConnectDefault
                        .Meter.mode = tlDCVSMeterCurrent
                        .mode = tlDCVSModeVoltage
                        .SetCurrentRanges pc_Def_UVS256HP_Init_MeasCurrRange, pc_Def_UVS256HP_Init_MeasCurrRange
                        If PinInfoArr.MeasureRangeAry(i) <> "" Then .SetCurrentRanges CDbl(PinInfoArr.MeasureRangeAry(i)), CDbl(PinInfoArr.MeasureRangeAry(i))
                        .Voltage.value = 0#
                        .Connect tlDCVSConnectDefault
                        .Gate = True
                        If PinInfoArr.ForceValueAry(i) <> "" Then
                            .Voltage.value = CDbl(PinInfoArr.ForceValueAry(i))
                        End If
                        PinInfoArr.RestoreGateOffAry(i) = True
                    Else
                        If PinInfoArr.ForceValueAry(i) <> "" Then
                            .Voltage.value = CDbl(PinInfoArr.ForceValueAry(i))
                        End If
                        .Meter.mode = tlDCVSMeterCurrent
                        If PinInfoArr.MeasureRangeAry(i) <> "" Then
                            .SetCurrentRanges CDbl(PinInfoArr.MeasureRangeAry(i)), PinInfoArr.MeasureRangeAry(i)
                        End If
                        .Gate = True
                    End If
                End With
                thehdw.Wait 0.001
            Next i
            MeasurePinStr = Join(PinInfoArr.PinNameAry, ",")
            If Enable_AutoRange Then
                Call DCVS_MeasureCurrent_AutoRange(MeasurePinStr, Enable_AutoRange, MeasurePinStr, , MeasurePinStr)
            End If
            MeasResultI = thehdw.DCVS.Pins(MeasurePinStr).Meter.Read(tlStrobe, 10, 10000, tlDCVSMeterReadingFormatAverage)   'pc_Def_HexVS_ReadPoint
            theexec.Flow.TestLimit MeasResultI, , , , , , unitAmp, , TestInstName & condition
            For i = 0 To UBound(PinInfoArr.PinNameAry)
                If PinInfoArr.RestoreGateOffAry(i) = True Then
                    thehdw.DCVS.Pins(PinInfoArr.PinNameAry(i)).Gate = False
                    thehdw.DCVS.Pins(PinInfoArr.PinNameAry(i)).Disconnect tlDCVSConnectDefault
                End If
            Next i
        Case Else
    End Select
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Retention_MeasDCVS")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Retention_MeasDigital(MeasCase As Long, condition As String)
On Error GoTo errHandler
    Dim i As Long
    Dim PinInfoArr As MeasureCond
    Dim TestInstName As String
    Dim MeasResultV As New PinListData
    Dim MeasResultI As New PinListData
    Dim MeasResultF As New PinListData
    Dim CounterValue As New PinListData
    Dim FreqInterval As Double
    FreqInterval = 0.01
    Select Case MeasCase
        Case RET_MeasV
            TestInstName = "MeasureVoltage"
            PinInfoArr = glb_RET_MeasAry(RET_MeasV).digital_pins
            thehdw.Digital.Pins(Join(PinInfoArr.PinNameAry, ",")).Disconnect
            With thehdw.PPMU.Pins(Join(PinInfoArr.PinNameAry, ","))
                .Gate = tlOff
                If PinInfoArr.ForceValueAry(i) <> "" Then
                    If glb_TesterType = "UltraFLEXplus" Then
                        Dim ForceIPerSite As New SiteDouble
                        ForceIPerSite = CDbl(PinInfoArr.ForceValueAry(i))
                        .ForceIPerSite ForceIPerSite, ForceIPerSite
                    Else
                        .ForceI CDbl(PinInfoArr.ForceValueAry(i)), CDbl(PinInfoArr.ForceValueAry(i))
                    End If
                End If
                .Connect
                .Gate = tlOn
            End With
            thehdw.Wait 0.001
            MeasResultV = thehdw.PPMU.Pins(Join(PinInfoArr.PinNameAry, ",")).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
            With thehdw.PPMU.Pins(Join(PinInfoArr.PinNameAry, ","))
                .ForceV pc_Def_PPMU_InitialValue_FV, pc_Def_PPMU_Max_InitialValue_FI_Range
                .Disconnect
                .Gate = tlOff
            End With
            theexec.Flow.TestLimit MeasResultV, , , , , , unitVolt, , TestInstName & condition
        Case RET_MeasI
            TestInstName = "MeasureCurrent"
            PinInfoArr = glb_RET_MeasAry(RET_MeasI).digital_pins
            thehdw.Digital.Pins(Join(PinInfoArr.PinNameAry, ",")).Disconnect
            With thehdw.PPMU.Pins(Join(PinInfoArr.PinNameAry, ","))
                .Gate = tlOff
                .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_Max_InitialValue_FI_Range
                If PinInfoArr.ForceValueAry(i) <> "" Then
                    If glb_TesterType = "UltraFLEXplus" Then
                        Dim ForceVPerSite As New SiteDouble
                        ForceVPerSite = CDbl(PinInfoArr.ForceValueAry(i))
                        .ForceVPerSite ForceVPerSite, ForceVPerSite
                    Else
                        .ForceV CDbl(PinInfoArr.ForceValueAry(i)), CDbl(PinInfoArr.ForceValueAry(i))
                    End If
                End If
                .Connect
                .Gate = tlOn
            End With
            thehdw.Wait 0.1
            MeasResultI = thehdw.PPMU.Pins(Join(PinInfoArr.PinNameAry, ",")).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
            theexec.Flow.TestLimit MeasResultI, , , , , , unitAmp, , TestInstName & condition
        Case RET_MeasF
            TestInstName = "MeasureFrequency"
            PinInfoArr = glb_RET_MeasAry(RET_MeasF).digital_pins
'            thehdw.Digital.Pins(Join(PinInfoArr.PinNameAry, ",")).Levels.value(chVoh) = 0.5
'            thehdw.Digital.Pins(Join(PinInfoArr.PinNameAry, ",")).Levels.value(chVol) = 0.6
            With thehdw.Digital.Pins(Join(PinInfoArr.PinNameAry, ",")).FreqCtr
                .EventSource = VOH
                .EventSlope = Positive
                .Interval = FreqInterval
                .Enable = IntervalEnable
                .Clear
                thehdw.Wait 0.001
                .start
                CounterValue = .Read()
            End With
            MeasResultF = CounterValue.Math.divide(FreqInterval)
            theexec.Flow.TestLimit MeasResultF, , , , , , unitHz, , TestInstName & condition
        Case Else
    End Select
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Retention_MeasDigital")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Shmoo_Save_Power(Optional CZ_item As Boolean = True)
On Error GoTo errHandler
    
    Dim p_ary() As String, p_cnt As Long, i As Long, InstName As String
    Dim ContextVmainValue As New PinListData
    Dim ContextVAltValue As New PinListData
        Dim PwrPin As String
    
    Set g_ApplyLevelTimingVmain = Nothing
    Set g_ApplyLevelTimingValt = Nothing
    Set g_ContextVmainValue = Nothing
    Set g_ContextVAltValue = Nothing
    Set ContextVmainValue = Nothing
    Set ContextVAltValue = Nothing
    
    PwrPin = ALL_Power_DCVS_pins + IIf(ALL_Power_DCVS_pins <> "" And ALL_Power_DCVI_pins <> "", ",", "") + ALL_Power_DCVI_pins
    theexec.DataManager.DecomposePinList PwrPin, p_ary, p_cnt

    '#If IGXL_VER_1030 = True Then
'    If ALL_Power_DCVS_pins <> "" Then
'        ContextVmainValue = TheHdw.DCVS.Pins(ALL_Power_DCVS_pins).Voltage.Main.PinListData
'        ContextVAltValue = TheHdw.DCVS.Pins(ALL_Power_DCVS_pins).Voltage.Alt.PinListData
'    End If
    '#Else
        For i = 0 To p_cnt - 1
            If gl_GetInstrument_Dic.Exists(LCase(p_ary(i))) Then
            'If Not NC_PIN.Exists(p_ary(i)) Then
            ContextVmainValue.AddPin UCase((p_ary(i)))
            ContextVAltValue.AddPin UCase((p_ary(i)))
            InstName = gl_GetInstrument_Dic.item(LCase(p_ary(i)))
            Select Case InstName
                 Case "VHDVS", "HexVS", "VSM", "VS-5A", "VS-800mA"
                       ContextVmainValue.Pins(UCase(p_ary(i))).value = CDbl(Format(thehdw.DCVS.Pins(p_ary(i)).Voltage.Main.value, "0.000"))
                       ContextVAltValue.Pins(UCase(p_ary(i))).value = CDbl(Format(thehdw.DCVS.Pins(p_ary(i)).Voltage.Alt.value, "0.000"))
                 Case "DC-07", "DC-30"
                       ContextVmainValue.Pins(UCase(p_ary(i))).value = CDbl(Format(thehdw.DCVI.Pins(p_ary(i)).Voltage, "0.000"))
                       ContextVAltValue.Pins(UCase(p_ary(i))).value = CDbl(Format(thehdw.DCVI.Pins(p_ary(i)).Voltage, "0.000"))
                 Case "HSD-U"
                 Case Else
                       theexec.ErrorLogMessage "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Save_core_power_per_site"
             End Select
            End If
        Next i
    '#End If
    
	If CZ_item then
		If DICT_DCVS_PIN_INDEX.Count = 0 Then
			For i = 0 To p_cnt - 1
				If gl_GetInstrument_Dic.Exists(LCase(ContextVmainValue.Pins(i).name)) Then
					DICT_DCVS_PIN_INDEX.Add ContextVmainValue.Pins.item(i).name, i
				End If
			Next i
		End If
	End If
    g_ApplyLevelTimingVmain = ContextVmainValue.Copy
    g_ApplyLevelTimingValt = ContextVAltValue.Copy
    g_ContextVmainValue = ContextVmainValue.Copy
    g_ContextVAltValue = ContextVAltValue.Copy
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Save_Power")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function RemoveSpecialKeyWord(PinName As String) As String
On Error GoTo errHandler

    Dim NameTmp As String
    Dim NameTmpArr() As String
    
    NameTmp = PinName
    NameTmpArr = Split(NameTmp, "_")
    If UCase(NameTmpArr(UBound(NameTmpArr))) = "CP" Or UCase(NameTmpArr(UBound(NameTmpArr))) = "FT" Then
        ReDim Preserve NameTmpArr(UBound(NameTmpArr) - 1)
        NameTmp = Join(NameTmpArr, "_")
    End If
    RemoveSpecialKeyWord = NameTmp
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "RemoveSpecialKeyWord")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function unitMapping(MappingValue As Variant) As Variant
On Error GoTo errHandler

    Dim val As Variant
    
    val = MappingValue
    If val >= 1000000 Then
       val = (val / 1000000) & "M"
    ElseIf val >= 1000 And val < 1000000 Then
       val = (val / 1000) & "K"
    ElseIf val >= 1 And val < 1000 Then
       val = (val * 1000) & "m"
    ElseIf val >= 0.001 And val < 1 Then
       val = (val * 1000) & "m"
    ElseIf val >= 0.000001 And val < 0.001 Then
       val = (val * 1000000) & "u"
    End If
    
    unitMapping = val
    Exit Function
errHandler:
        Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "unitMapping")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GetCharSetUp(SetupName As String, axis As tlDevCharShmooAxis, Optional StepName As Variant, Optional RangeFrom As Double, Optional RangeTo As Double, _
                            Optional RangeSteps As Long, Optional RangeStepSize As Double, Optional RangeCalcType As tlDevCharRangeField, Optional PinName As String, _
                            Optional ParameterName As String)
On Error GoTo errHandler
 
        With theexec.DevChar.Setups(SetupName).Shmoo.Axes(axis)
            StepName = .StepName
            RangeFrom = .Parameter.range.from
            RangeTo = .Parameter.range.to
            RangeSteps = .Parameter.range.Steps + 1
            RangeStepSize = .Parameter.range.StepSize
            RangeCalcType = .Parameter.range.CalculatedField
            PinName = .ApplyTo.Pins
            ParameterName = .Parameter.name.value
        End With
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "GetCharSetUp")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function AddPinToPinListData(ApplyPin As String, ApplyPinArr() As String, ApplyPinListData As PinListData)
On Error GoTo errHandler
    
    Dim i As Long
    
    If ApplyPin <> "" Then
        ApplyPinArr = Split(ApplyPin, ",")
        For i = 0 To UBound(ApplyPinArr)
            If gl_GetInstrument_Dic.Exists(LCase(ApplyPinArr(i))) Then
            'If Not NC_PIN.Exists(ApplyPinArr(i)) Then
            'If TheExec.DataManager.ChannelType(ApplyPinArr(i)) <> "N/C" Then
                ApplyPinListData.AddPin ApplyPinArr(i)
            End If
        Next i
    Else
        theexec.Datalog.WriteComment "The apply pin is empty!"
    End If
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "AddPinToPinListData")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function SetTargetVol(PinName As String, OutputSel As Long, CurrentPinListData As PinListData, TargetVol As PinListData, _
                            set_init As Boolean, RestoreMainOrAlt As PinListData, seq As Long)
On Error GoTo errHandler

    Dim InstrumentType As String
    Dim site As Variant
    Const OutputIsVmain = 1
    Const OutputIsValt = 2
    
    If gl_GetInstrument_Dic.Exists(LCase(PinName)) Then
        InstrumentType = gl_GetInstrument_Dic.item(LCase(PinName))
        Select Case InstrumentType
            Case "VHDVS", "HexVS", "VSM", "VS-5A", "VS-800mA"
               For Each site In theexec.sites
                    If OutputSel = OutputIsValt Then
                        CurrentPinListData.Pins(PinName).value = thehdw.DCVS.Pins(PinName).Voltage.Alt.value
                        thehdw.DCVS.Pins(PinName).Voltage.Main.value = TargetVol.Pins(PinName).value
                    Else
                        CurrentPinListData.Pins(PinName).value = thehdw.DCVS.Pins(PinName).Voltage.Main.value
                        thehdw.DCVS.Pins(PinName).Voltage.Alt.value = TargetVol.Pins(PinName).value
                    End If
               Next site
            Case "DC-07", "DC-30"
                CurrentPinListData.Pins(PinName).value = thehdw.DCVI.Pins(PinName).Voltage
                thehdw.DCVI.Pins(PinName).Voltage = TargetVol.Pins(PinName).value
            Case "HSD-U"
            Case Else
                'theexec.ErrorLogMessage "Instrument " & InstrumentType & " for pin " & PinName & " is not supported in Shmoo_Restore_Power_per_site_Vbump"
                Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "SetTargetVol", _
                "Instrument " & InstrumentType & " for pin " & PinName & " is not supported in retention")
        End Select
    End If
                               
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "SetTargetVol")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function SetRecoverVol(PinName As String, OutputSel As Long, RecoverVol As PinListData)
On Error GoTo errHandler

    Dim InstrumentType As String
    Dim site As Variant
    Const OutputIsVmain = 1
    Const OutputIsValt = 2
    
        If gl_GetInstrument_Dic.Exists(LCase(PinName)) Then
                InstrumentType = gl_GetInstrument_Dic.item(LCase(PinName))
                Select Case InstrumentType
                        Case "VHDVS", "HexVS", "VSM", "VS-5A", "VS-800mA"
                                For Each site In theexec.sites
                                        If OutputSel = OutputIsValt Then
                                                thehdw.DCVS.Pins(PinName).Voltage.Main.value = RecoverVol.Pins(PinName).value
                                        Else
                                                thehdw.DCVS.Pins(PinName).Voltage.Alt.value = RecoverVol.Pins(PinName).value
                                        End If
                                Next
                        Case "DC-07", "DC-30"
                                For Each site In theexec.sites
                                        thehdw.DCVI.Pins(PinName).Voltage.value = RecoverVol.Pins(PinName).value
                                Next
                        Case "HSD-U"
                        Case Else
                                 'theexec.ErrorLogMessage "Instrument " & InstrumentType & " for pin " & PinName & " is not supported in Shmoo_Restore_Power_per_site_Vbump_Retention"
                                Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "SetRecoverVol", _
                                "Instrument " & InstrumentType & " for pin " & PinName & " is not supported in retention")
                End Select
        End If
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "SetRecoverVol")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub RecoverVmainValtAfterRampUp(PinName As String, RecoverVoltage As PinListData, IsInit As Boolean)
On Error GoTo errHandler
    
    Dim InstrumentType As String
    Dim site As Variant
    
    If gl_GetInstrument_Dic.Exists(LCase(PinName)) Then
        InstrumentType = gl_GetInstrument_Dic.item(LCase(PinName))
        Select Case InstrumentType
            Case "VHDVS", "HexVS", "VSM", "VS-5A", "VS-800mA"
                For Each site In theexec.sites
                    If IsInit = True Then
                        thehdw.DCVS.Pins(PinName).Voltage.Alt.value = RecoverVoltage.Pins(PinName).value
                    Else
                        thehdw.DCVS.Pins(PinName).Voltage.Main.value = RecoverVoltage.Pins(PinName).value
                    End If
                Next site
            Case "DC-07", "DC-30"
                For Each site In theexec.sites
                    thehdw.DCVI.Pins(PinName).Voltage.value = RecoverVoltage.Pins(PinName).value
                Next
            Case "HSD-U"
            Case Else
                'theexec.ErrorLogMessage "Instrument " & InstrumentType & " for pin " & PinName & " is not supported in Shmoo_Restore_Power_per_site_Vbump_Retention"
                Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "SetRecoverVol", _
                "Instrument " & InstrumentType & " for pin " & PinName & " is not supported in retention")
        End Select
    End If
    
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "RecoverVmainValtAfterRampUp")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub RecoverVmainValtAfterVoltageUpDirectly(PinName As String, OutputSel As Long, IsInit As Boolean, Optional RestoreValue As PinListData)
On Error GoTo errHandler

    Dim InstrumentType As String
    Dim site As Variant
    Const OutputIsVmain = 1
    Const OutputIsValt = 2
    
    If gl_GetInstrument_Dic.Exists(LCase(PinName)) Then
        InstrumentType = gl_GetInstrument_Dic.item(LCase(PinName))
        Select Case InstrumentType
            Case "VHDVS", "HexVS", "VSM", "VS-5A", "VS-800mA"
                For Each site In theexec.sites
                    If IsInit Then
                        If OutputSel = OutputIsValt Then
                            thehdw.DCVS.Pins(PinName).Voltage.Main.value = g_ContextVmainValue.Pins(PinName).value
                        Else
                            thehdw.DCVS.Pins(PinName).Voltage.Alt.value = g_ContextVAltValue.Pins(PinName).value
                        End If
                    Else
                        If OutputSel = OutputIsValt Then
                            thehdw.DCVS.Pins(PinName).Voltage.Main.value = g_ContextVmainValue.Pins(PinName).value
                        Else
                            thehdw.DCVS.Pins(PinName).Voltage.Alt.value = RestoreValue.Pins(PinName).value
                        End If
                    End If
                Next site
            Case "DC-07", "DC-30"
                For Each site In theexec.sites
                    thehdw.DCVI.Pins(PinName).Voltage.value = g_ContextVmainValue.Pins(PinName).value
                Next
                'theexec.ErrorLogMessage "Instrument " & InstrumentType & " for pin " & PinName & " is not supported in Shmoo_Restore_Power_per_site_Vbump_Retention"
'                            Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "SetRecoverVol", _
'                            "Instrument " & InstrumentType & " for pin " & PinName & " is not supported in retention")
            Case "HSD-U"
            Case Else
                Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "SetRecoverVol", _
                "Instrument " & InstrumentType & " for pin " & PinName & " is not supported in retention")
                'theexec.ErrorLogMessage "Instrument " & InstrumentType & " for pin " & PinName & " is not supported in Shmoo_Restore_Power_per_site_Vbump_Retention"
        End Select
    End If
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "RecoverVmainValtAfterVoltageUpDirectly")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function ProcessPattInfo(PattAry() As Pattern, SELSRAM_DSSC As String _
                , DigSrc_BitSize As String, DigSrc_Seg As String, DigSrc_pin As String, digSrc_EQ As String, ForceCondition As String _
                , WaitTime As String, Ramp_Setting As String, BlockType As String, HarvestHeader As String, PowerRunScenario As String, pmode As String)
                ', WaitTime As String, BlockType As String, HarvestHeader As String, PowerRunScenario As String, pmode As String)
On Error GoTo errHandler
 
    Erase g_CharPattInfoAry
    ReDim g_CharPattInfoAry(InitialPatCnt + PayloadPatCnt - 1)
    Dim tmpScenarioAry() As String
    Dim DigSrc_BitSizeAry() As String
    Dim DigSrc_SegAry() As String
    Dim DigSrc_PinAry() As String
    Dim DigSrc_EQAry() As String
    Dim MergedCond As New SiteVariant
    Dim MergedVDDPins As String
    Dim Init_NV As New PinListData
    Dim Init_HV As New PinListData
    Dim Init_LV As New PinListData
    Dim Payload_NV As New PinListData
    Dim Payload_HV As New PinListData
    Dim Payload_LV As New PinListData
    Dim InitSel As String
    Dim PmodeSel As String
    
    Dim PatternCnt As Long
    Dim SelSramIdx As Long
    Dim WaitTimeAry() As String
    Dim i As Long, j As Long
    
    Dim tmpGBAry() As String
    Dim tmpOffsetVal As String
    Dim voltageType As String
    Dim PwrPin As String
    
    Dim tmpPatAry() As String
    Dim SplitInfo() As String
    Dim subSplitinfo() As String
    Dim SelectType As String
    Dim EMAKeyWord As String
    Dim PatternLoopCnt As Long
    
    Dim EQarray() As String
    Dim SegmentVar As Variant
    
    g_MergeVDD = vbNullString
    g_MergeCond = vbNullString
    PwrPin = ALL_Power_DCVS_pins + IIf(ALL_Power_DCVS_pins <> "" And ALL_Power_DCVI_pins <> "", ",", "") + ALL_Power_DCVI_pins
    Decide_Shmoo_Voltage pmode, PwrPin, ForceCondition, PowerRunScenario, Init_NV, Init_HV, Init_LV _
                                            , Payload_NV, Payload_HV, Payload_LV, InitSel, PmodeSel, MergedVDDPins, MergedCond
                                            
    g_MergeVDD = MergedVDDPins
    g_MergeCond = MergedCond
    Select Case UCase(InitSel)
    Case "TYP"
        Power_Level_Last = "INIT_NV"    '"First_init_NV"
        voltageType = "NV"
    Case "MAX"
        Power_Level_Last = "INIT_HV"    '"First_init_HV"
        voltageType = "HV"
    Case "MIN"
        Power_Level_Last = "INIT_LV"    '"First_init_LV"
        voltageType = "LV"
    End Select
    
    If PowerRunScenario <> "" And InStr(PowerRunScenario, "_") <> 0 Then tmpScenarioAry = Split(PowerRunScenario, "_")
    'If DigSrc_Seg <> "" And DigSrc_BitSize <> "" Then Process_DigSrc_Setting DigSrc_BitSize, DigSrc_Seg, DigSrc_pin, digSrc_EQ
    If digSrc_EQ <> "" And DigSrc_BitSize <> "" Then Process_DigSrc_Setting DigSrc_BitSize, digSrc_EQ, DigSrc_pin, DigSrc_Seg
    
    
    If SELSRAM_DSSC <> "" Then 'And BlockType <> "" Then
        If InStr(SELSRAM_DSSC, "'") > 0 Then SELSRAM_DSSC = Replace(SELSRAM_DSSC, "'", vbNullString)
        If UCase(SELSRAM_DSSC) Like "SELSRM*" Or UCase(SELSRAM_DSSC) Like "SELSRAM*" Then
            SELSRAM_DSSC = Replace(UCase(SELSRAM_DSSC), "SELSRAM", vbNullString)
            SELSRAM_DSSC = Replace(UCase(SELSRAM_DSSC), "SELSRM", vbNullString)
        ElseIf UCase(SELSRAM_DSSC) Like "DSELSRM*" Or UCase(SELSRAM_DSSC) Like "DSELSRAM*" Then
            SELSRAM_DSSC = Replace(UCase(SELSRAM_DSSC), "DSELSRAM", vbNullString)
            SELSRAM_DSSC = Replace(UCase(SELSRAM_DSSC), "DSELSRM", vbNullString)
            Call InverStr(SELSRAM_DSSC)
        End If
    End If
    
    'pattern case : SELSRAM, EMA, DIGSRC, LOOP
    'argument ex  : PP_PALA0_H_IN00_SC_XXXX_XXX_XXX_AUT_ALLFRV_SI:SELSRM
    '             : PP_PALA0_H_IN00_SC_XXXX_XXX_XXX_AUT_ALLFRV_SI:DIGSRC:EMA
    'Loop case    : PP_PALA0_H_IN00_SC_XXXX_XXX_XXX_AUT_ALLFRV_SI:3
    If WaitTime <> "" Then Process_Ret_Setting WaitTime, Ramp_Setting

    If UBound(g_CharPattInfoAry) = UBound(PattAry) Then
        For PatternCnt = 0 To UBound(g_CharPattInfoAry)
            If PattAry(PatternCnt) <> "" Then
                With g_CharPattInfoAry(PatternCnt)
                    If PatternCnt < InitialPatCnt Then
                        'Judge the Selsram Info
                        .IsInitPattern = True
                        .PowerRunCond = voltageType
                    Else
                        .Pattern.value = PattAry(PatternCnt).value
                        .IsInitPattern = False
                        .PowerRunCond = "SWEEP"
                    End If
                    .SelSramMatchIdx = -1
                    .PatternLoopCnt = 1
                    .Sequence = PatternCnt + 1
                    
                    If .digSrc_EQ <> "" Then
                        EQarray = Split(.digSrc_EQ, "+")
                        For i = 0 To UBound(EQarray)
                            If UCase(.SegDict(EQarray(i))(0)) = "SELSRAM" And SELSRAM_DSSC <> "" Then
                                For SelSramIdx = 0 To UBound(SelsramMapping)
                                    If UCase(PattAry(PatternCnt).value) Like UCase(SelsramMapping(SelSramIdx).Pattern) And SelsramMapping(SelSramIdx).Pattern <> "*" Then
                                        .SelSramMatchIdx = SelSramIdx
                                        .DynamicSourceBit = dynamic_SELSRM_source_bits(SELSRAM_DSSC, BlockType, SelSramIdx)
                                        .DigSrcType = "SELSRAM"
                                        Exit For
                                    End If
                                Next SelSramIdx
                            End If
                        Next
                    End If
                    .Pattern.value = PattAry(PatternCnt).value
                    .testType = PatternTypeChecker(PattAry(PatternCnt).value)
                End With
                    

            End If
        Next PatternCnt
    Else
        theexec.Datalog.WriteComment " Pattern Counts Abnormal, Please Check Input Patterns ! "
    End If
    
    For i = 0 To UBound(tmpScenarioAry) Step 2
        If UCase(tmpScenarioAry(i)) Like "INIT*" Then
            If UCase(tmpScenarioAry(i)) Like "INIT#" Or UCase(tmpScenarioAry(i)) Like "INIT##" Then
                tmpScenarioAry(i) = CLng(Replace(UCase(tmpScenarioAry(i)), "INIT", ""))
                g_CharPattInfoAry(tmpScenarioAry(i) - 1).PowerRunCond = UCase(tmpScenarioAry(i + 1))
            Else
                For j = 0 To InitialPatCnt - 1
                    g_CharPattInfoAry(j).PowerRunCond = UCase(tmpScenarioAry(i + 1))
                Next j
            End If
        ElseIf UCase(tmpScenarioAry(i)) Like "PL*" Then
            If UCase(tmpScenarioAry(i)) Like "PL#" Then
                tmpScenarioAry(i) = CLng(Replace(UCase(tmpScenarioAry(i)), "PL", ""))
                g_CharPattInfoAry(InitialPatCnt + tmpScenarioAry(i) - 1).PowerRunCond = UCase(tmpScenarioAry(i + 1))
            Else
                For j = 0 To PayloadPatCnt - 1
                    g_CharPattInfoAry(InitialPatCnt + j).PowerRunCond = UCase(tmpScenarioAry(i + 1))
                Next j
            End If
        End If
    Next i
    
    For i = 0 To UBound(g_CharPattInfoAry)
        If g_CharPattInfoAry(i).Pattern.value <> "" Then
            If g_CharPattInfoAry(i).DigSrcType = "SELSRAM" Then
                Select Case UCase(PmodeSel)
                Case "TYP"
                    Set g_CharPattInfoAry(i).DictApplyVol = Dict_PL_NV
                Case "MAX"
                    Set g_CharPattInfoAry(i).DictApplyVol = Dict_PL_HV
                Case "MIN"
                    Set g_CharPattInfoAry(i).DictApplyVol = Dict_PL_LV
                End Select
            End If
            If i < InitialPatCnt Then
                With g_CharPattInfoAry(i)
                    Select Case UCase(.PowerRunCond)
                    Case "NV"
                        .ForceVoltage = Init_NV.Copy
                    Case "HV"
                        .ForceVoltage = Init_HV.Copy
                    Case "LV"
                        .ForceVoltage = Init_LV.Copy
                    Case Else
                        If UCase(.PowerRunCond) Like "*SWEEP*" Then
                            If InStr(.PowerRunCond, ":") <> 0 Then
                                tmpOffsetVal = vbNullString
                                tmpGBAry = Split(.PowerRunCond, ":")
                                'Only Suppory one calculation symbol
                                .GuardBandSymbol = mid(tmpGBAry(1), 1, 1)
                                tmpOffsetVal = LCase(mid(tmpGBAry(1), 2, Len(tmpGBAry(1))))
                                If InStr(tmpOffsetVal, "mv") <> 0 Then
                                    .GuardBandVal = Format((CDbl(Replace(tmpOffsetVal, "mv", "")) / 1000), "0.000")
                                ElseIf InStr(tmpOffsetVal, "v") <> 0 Then
                                    .GuardBandVal = Format(CDbl(Replace(tmpOffsetVal, "v", "")), "0.000")
                                Else
                                    theexec.Datalog.WriteComment "Please Check the Scenario Formate: " & tmpOffsetVal
                                End If
                                .PowerRunCond = "SweepGuardBand"
                            End If
                        
                            Select Case UCase(InitSel)
                            Case "TYP"
                                .ForceVoltage = Init_NV.Copy
                            Case "MAX"
                                .ForceVoltage = Init_HV.Copy
                            Case "MIN"
                                .ForceVoltage = Init_LV.Copy
                            End Select
                        End If
                    End Select
                End With
            Else
                With g_CharPattInfoAry(i)
                    Select Case UCase(g_CharPattInfoAry(i).PowerRunCond)
                    Case "NV"
                        .ForceVoltage = Payload_NV.Copy
                    Case "HV"
                        .ForceVoltage = Payload_HV.Copy
                    Case "LV"
                        .ForceVoltage = Payload_LV.Copy
                    Case "VRS"
                        Select Case UCase(InitSel)
                        Case "TYP"
                            .ForceVoltage = Init_NV.Copy
                        Case "MAX"
                            .ForceVoltage = Init_HV.Copy
                        Case "MIN"
                            .ForceVoltage = Init_LV.Copy
                        End Select
                    Case Else
                        If UCase(.PowerRunCond) Like "*SWEEP*" Then
                            If InStr(.PowerRunCond, ":") <> 0 Then
                                tmpOffsetVal = vbNullString
                                tmpGBAry = Split(.PowerRunCond, ":")
                                'Only Suppory one calculation symbol
                                .GuardBandSymbol = mid(tmpGBAry(1), 1, 1)
                                tmpOffsetVal = LCase(mid(tmpGBAry(1), 2, Len(tmpGBAry(1))))
                                If InStr(tmpOffsetVal, "mv") <> 0 Then
                                    .GuardBandVal = Format((CDbl(Replace(tmpOffsetVal, "mv", "")) / 1000), "0.000")
                                ElseIf InStr(tmpOffsetVal, "v") <> 0 Then
                                    .GuardBandVal = Format(CDbl(Replace(tmpOffsetVal, "v", "")), "0.000")
                                Else
                                    theexec.Datalog.WriteComment "Please Check the Scenario Formate: " & tmpOffsetVal
                                End If
                                .PowerRunCond = "SweepGuardBand"
                            End If
                            Select Case UCase(PmodeSel)
                            Case "TYP"
                                .ForceVoltage = Payload_NV.Copy
                            Case "MAX"
                                .ForceVoltage = Payload_HV.Copy
                            Case "MIN"
                                .ForceVoltage = Payload_LV.Copy
                            End Select
                        End If
                    End Select
                End With
            End If
        End If
    Next i
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "ProcessPattInfo")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Decide_Shmoo_Voltage(PerforManceMode As String, Power_pins As String, ForceCond As String, RunScenario As String _
                                                                    , Init_NV As PinListData, Init_HV As PinListData, Init_LV As PinListData _
                                                                    , Payload_NV As PinListData, Payload_HV As PinListData, Payload_LV As PinListData _
                                                                    , InitSel As String, PMSel As String, MergeFoeceVDD As String, MergeForceCond As SiteVariant)
On Error GoTo errHandler

    Dim PerformanceModeArr() As String
    Dim Dc_spec_type As String
    Dim p_ary() As String, p_cnt As Long, i As Long
    Dim PinValue As New SiteDouble
    Dim ContextCat As String
    Dim ContextSel As String
    Dim PmodeCat As String
    Dim PmodeSel As String
    Dim DummyStr As String
    Dim Specs As Variant
    Dim Init_NV_Dict As New Dictionary
    Dim Init_HV_Dict As New Dictionary
    Dim Init_LV_Dict As New Dictionary
    Dim Payload_NV_Dict As New Dictionary
    Dim Payload_HV_Dict As New Dictionary
    Dim Payload_LV_Dict As New Dictionary
    Dim ForcePin_Dict As New Dictionary
    
    Dim ContentNV As Boolean: ContentNV = False
    Dim ContentHV As Boolean: ContentHV = False
    Dim ContentLV As Boolean: ContentLV = False
    Dim Read_Init_NV As Boolean: Read_Init_NV = False
    Dim Read_Init_HV As Boolean: Read_Init_HV = False
    Dim Read_Init_LV As Boolean: Read_Init_LV = False
    Dim Read_PL_NV As Boolean: Read_PL_NV = False
    Dim Read_PL_HV As Boolean: Read_PL_HV = False
    Dim Read_PL_LV As Boolean: Read_PL_LV = False
    
    Dim FC As Variant
    Dim ForceCondAry() As String
    Dim pininfo() As String
    Dim ForcePinAry() As String
    Dim FroceSpecAry() As String
    Dim ForceValueAry() As New SiteDouble
    Dim ForceVDD_Cnt As Long
    Dim tmpPinAry() As String
    Dim tmpSpecAry() As String
    Dim tmpPinCnt As Long
    
    Dim Var_name As String
    Dim VOP_VAR_Name As String
    Dim PL_VAR_Name As String
    Dim tmpDict_Value As String
    Dim t_SpecPin As String
    Dim PLD_Vmain As New PinListData
    Dim DCVS_pin As String
    Dim DCVI_pin As String
    Dim InstName As String
    
    Dim BinCutPinAry() As String        '20240327 applyBV
    Dim BinCutDict As New Dictionary
    Dim BinCutDomain As String
    Dim site As Variant
    Dim BinCutValue As New SiteDouble
    Dim tmpName As String
    
    If Power_pins = "" Then Exit Function
    g_CharInputString_Voltage_Dict.RemoveAll
    Dict_INI_NV.RemoveAll
    Dict_INI_HV.RemoveAll
    Dict_INI_LV.RemoveAll
    Dict_PL_NV.RemoveAll
    Dict_PL_HV.RemoveAll
    Dict_PL_LV.RemoveAll
    MergeFoeceVDD = vbNullString
    MergeForceCond = vbNullString
    g_Retention_VDD = vbNullString
    g_Retention_ForceV = vbNullString
    Set g_RetntionVal = Nothing
    ForceVDD_Cnt = 0
    PLD_Vmain = g_ApplyLevelTimingVmain.Copy
    '20240327 applyBV
    If gb_ApplyBV = True Then
        For Each FC In pinGroup_BinCut
            Set BinCutValue = New SiteDouble
            BinCutValue = BinCut_Payload_Voltage(VddBinStr2Enum(CStr(FC))).divide(1000) 'unit = V
            theexec.DataManager.DecomposePinList FC, tmpPinAry, tmpPinCnt
            For i = 0 To tmpPinCnt - 1
                If gl_GetInstrument_Dic.Exists(LCase(tmpPinAry(i))) Then
                    If Not BinCutDict.Exists(tmpPinAry(i)) Then
                    BinCutDict.Add tmpPinAry(i), BinCutValue
                    End If
                End If
            Next i
        Next FC
    End If
    'process force condition
    If ForceCond <> "" Then
        theexec.Datalog.WriteComment "Force Condtion:" & ForceCond
        ForceCond = UCase(ForceCond)
        ForceCondAry = Split(ForceCond, ";")
        For Each FC In ForceCondAry
            pininfo = Split(FC, ":")
            If UBound(pininfo) = 2 Then
                theexec.DataManager.DecomposePinList pininfo(0), tmpPinAry, tmpPinCnt
                If pininfo(1) = "BV" Or pininfo(1) = "V" Then   '20240327 applyBV
                    ReDim tmpSpecAry(tmpPinCnt - 1) As String
                    For i = 0 To tmpPinCnt - 1
                        If gl_GetInstrument_Dic.Exists(LCase(tmpPinAry(i))) Then
                            tmpSpecAry(i) = RemoveSpecialKeyWord(tmpPinAry(i))
                            ReDim Preserve ForcePinAry(ForceVDD_Cnt)
                            ReDim Preserve FroceSpecAry(ForceVDD_Cnt)
                            ReDim Preserve ForceValueAry(ForceVDD_Cnt)
                        
                            ForcePinAry(ForceVDD_Cnt) = tmpPinAry(i)
                            FroceSpecAry(ForceVDD_Cnt) = tmpSpecAry(i)
                            If pininfo(1) = "BV" And BinCutDict.Exists(UCase(tmpPinAry(i))) Then
                                ForceValueAry(ForceVDD_Cnt) = BinCutDict.item(UCase(tmpPinAry(i)))
                            Else
                            ForceValueAry(ForceVDD_Cnt) = CDbl(pininfo(2))
                            End If
                            ForceVDD_Cnt = ForceVDD_Cnt + 1
                            If Not ForcePin_Dict.Exists(tmpPinAry(i)) Then ForcePin_Dict.Add tmpPinAry(i), tmpPinAry(i)
                        End If
                    Next i
                ElseIf pininfo(1) = "VRET" Then
                    For i = 0 To tmpPinCnt - 1
                        If gl_GetInstrument_Dic.Exists(LCase(tmpPinAry(i))) Then
                            g_RetntionVal.AddPin tmpPinAry(i)
                            g_RetntionVal.Pins(tmpPinAry(i)).value = pininfo(2)
                            If g_Retention_VDD = "" Then
                                g_Retention_VDD = tmpPinAry(i)
                            Else
                                g_Retention_VDD = g_Retention_VDD & "," & tmpPinAry(i)
                            End If
                        End If
                    Next i
                End If
            ElseIf UBound(pininfo) = 1 And UCase(pininfo(0)) = "USL" Then
                CHAR_USL_HVCC = FormatNumber(CDbl(pininfo(1)), 3)
				CHAR_USL_LVCC = FormatNumber(CDbl(pininfo(1)), 3)
            ElseIf UBound(pininfo) = 1 And UCase(pininfo(0)) = "LSL" Then
                CHAR_LSL_HVCC  = FormatNumber(CDbl(pininfo(1)), 3)
				CHAR_LSL_LVCC = FormatNumber(CDbl(pininfo(1)), 3)
            End If
        Next FC
    End If
    For Each FC In BinCutDict.Keys
        If BinCutDict.Exists(FC) And Not ForcePin_Dict.Exists(FC) Then
            tmpName = RemoveSpecialKeyWord(CStr(FC))
            ReDim Preserve ForcePinAry(ForceVDD_Cnt)
            ReDim Preserve FroceSpecAry(ForceVDD_Cnt)
            ReDim Preserve ForceValueAry(ForceVDD_Cnt)
            ForcePinAry(ForceVDD_Cnt) = FC
            FroceSpecAry(ForceVDD_Cnt) = tmpName
            ForceValueAry(ForceVDD_Cnt) = BinCutDict.item(UCase(FC))
            ForceVDD_Cnt = ForceVDD_Cnt + 1
            ForcePin_Dict.Add FC, FC
        End If
    Next
    'process context
    theexec.DataManager.GetInstanceContext ContextCat, ContextSel, DummyStr, DummyStr, DummyStr, DummyStr, DummyStr, DummyStr
    InitSel = ContextSel
    Select Case UCase(ContextSel)
    Case "TYP"
        'Read_Init_NV = True
        ContentNV = True
        If UCase(RunScenario) Like "*INIT#_HV*" Or UCase(RunScenario) Like "*INIT##_HV*" Or UCase(RunScenario) Like "INIT_HV*" Then Read_Init_HV = True
        If UCase(RunScenario) Like "*INIT#_LV*" Or UCase(RunScenario) Like "*INIT##_LV*" Or UCase(RunScenario) Like "INIT_LV*" Then Read_Init_LV = True
    Case "MAX"
        'Read_Init_HV = True
        ContentHV = True
        If UCase(RunScenario) Like "*INIT#_NV*" Or UCase(RunScenario) Like "*INIT##_NV*" Or UCase(RunScenario) Like "INIT_NV*" Then Read_Init_NV = True
        If UCase(RunScenario) Like "*INIT#_LV*" Or UCase(RunScenario) Like "*INIT##_LV*" Or UCase(RunScenario) Like "INIT_LV*" Then Read_Init_LV = True
    Case "MIN"
        'Read_Init_LV = True
        ContentLV = True
        If UCase(RunScenario) Like "*INIT#_NV*" Or UCase(RunScenario) Like "*INIT##_NV*" Or UCase(RunScenario) Like "INIT_NV*" Then Read_Init_NV = True
        If UCase(RunScenario) Like "*INIT#_HV*" Or UCase(RunScenario) Like "*INIT##_HV*" Or UCase(RunScenario) Like "INIT_HV*" Then Read_Init_HV = True
    End Select

    'process pmode
    PerformanceModeArr = Split(PerforManceMode, ":")
    If UBound(PerformanceModeArr) = -1 Then 'if argument is empty
        PmodeCat = UCase(ContextCat)
        PmodeSel = UCase(ContextSel)
    Else
        PmodeCat = UCase(PerformanceModeArr(0))
        If UBound(PerformanceModeArr) > 0 Then
            If UCase(PerformanceModeArr(1)) Like "NV" Then PmodeSel = "TYP"
            If UCase(PerformanceModeArr(1)) Like "HV" Then PmodeSel = "MAX"
            If UCase(PerformanceModeArr(1)) Like "LV" Then PmodeSel = "MIN"
        Else    'Using NV as default setting
            PmodeSel = "TYP"
        End If
    End If
    PMSel = PmodeSel
    Select Case UCase(PmodeSel)
    Case "TYP"
        Read_PL_NV = True
        If UCase(RunScenario) Like "*PL#_HV*" Or UCase(RunScenario) Like "*PL##_HV*" Then Read_PL_HV = True
        If UCase(RunScenario) Like "*PL#_LV*" Or UCase(RunScenario) Like "*PL##_LV*" Then Read_PL_LV = True
    Case "MAX"
        Read_PL_HV = True
        If UCase(RunScenario) Like "*PL#_NV*" Or UCase(RunScenario) Like "*PL##_NV*" Then Read_PL_NV = True
        If UCase(RunScenario) Like "*PL#_LV*" Or UCase(RunScenario) Like "*PL##_LV*" Then Read_PL_LV = True
    Case "MIN"
        Read_PL_LV = True
        If UCase(RunScenario) Like "*PL#_NV*" Or UCase(RunScenario) Like "*PL##_NV*" Then Read_PL_NV = True
        If UCase(RunScenario) Like "*PL#_HV*" Or UCase(RunScenario) Like "*PL##_HV*" Then Read_PL_HV = True

    End Select
    
    
    For Each Specs In theexec.Specs.DC.Categories(PmodeCat).SpecList
        Specs = LCase(Specs)  'Define DC Spec Type
        If Specs Like "*_var_c" Then
            Dc_spec_type = "_C"
        ElseIf Specs Like "*_var_g" Then
            Dc_spec_type = "_G"
        ElseIf Specs Like "*_var_s" Then
            Dc_spec_type = "_S"
        ElseIf Specs Like "*_var_h" Then
            Dc_spec_type = "_H"
        ElseIf Specs Like "*_var_r" Then
            Dc_spec_type = "_R"
        ElseIf Specs Like "*_var_bi" Then
            Dc_spec_type = "_BI"
        ElseIf Specs Like "*_var_sc" Then
            Dc_spec_type = "_SC"
        ElseIf Specs Like "*_var" Then
            Dc_spec_type = ""
        Else
            theexec.ErrorLogMessage "DC spec " & Specs & " is not ended with _VAR_C/S/G/H in " & theexec.DataManager.instancename
        End If
        Exit For
    Next Specs
    
    theexec.DataManager.DecomposePinList Power_pins, p_ary, p_cnt
    
    For i = 0 To p_cnt - 1
        t_SpecPin = vbNullString
        p_ary(i) = UCase(p_ary(i))
        If gl_GetInstrument_Dic.Exists(LCase(p_ary(i))) Then
            t_SpecPin = RemoveSpecialKeyWord(p_ary(i))
            Var_name = t_SpecPin & "_" & "VAR" & Dc_spec_type
            VOP_VAR_Name = t_SpecPin & "_VOP" & "_" & "VAR" & Dc_spec_type
            
            'Store Init voltage for reference
            If ContentNV = True Then
                Init_NV.AddPin p_ary(i)
                Init_NV.Pins(p_ary(i)).value = PLD_Vmain.Pins(DICT_DCVS_PIN_INDEX.item(p_ary(i))).value
                If Not Dict_INI_NV.Exists(UCase(p_ary(i))) Then
                    Dict_INI_NV.Add UCase(p_ary(i)), PLD_Vmain.Pins(DICT_DCVS_PIN_INDEX.item(p_ary(i))).value
                End If
            ElseIf Read_Init_NV = True Then
                tmpDict_Value = Format(theexec.Specs.DC.item(Var_name).Categories(ContextCat).Typ.value, "0.000")
                Init_NV.AddPin p_ary(i)
                Init_NV.Pins(p_ary(i)).value = tmpDict_Value
                If Not Dict_INI_NV.Exists(UCase(p_ary(i))) Then
                    Dict_INI_NV.Add UCase(p_ary(i)), CDbl(tmpDict_Value)
                End If
            End If
            If ContentHV = True Then
                Init_HV.AddPin p_ary(i)
                Init_HV.Pins(p_ary(i)).value = PLD_Vmain.Pins(DICT_DCVS_PIN_INDEX.item(p_ary(i))).value
                If Not Dict_INI_HV.Exists(UCase(p_ary(i))) Then
                    Dict_INI_HV.Add UCase(p_ary(i)), PLD_Vmain.Pins(DICT_DCVS_PIN_INDEX.item(p_ary(i))).value
                End If
            ElseIf Read_Init_HV = True Then
                tmpDict_Value = Format(theexec.Specs.DC.item(Var_name).Categories(ContextCat).max.value, "0.000")
                Init_HV.AddPin p_ary(i)
                Init_HV.Pins(p_ary(i)).value = tmpDict_Value
                If Not Dict_INI_HV.Exists(UCase(p_ary(i))) Then
                    Dict_INI_HV.Add UCase(p_ary(i)), CDbl(tmpDict_Value)
                End If
            End If
            If ContentLV = True Then
                Init_LV.AddPin p_ary(i)
                Init_LV.Pins(p_ary(i)).value = PLD_Vmain.Pins(DICT_DCVS_PIN_INDEX.item(p_ary(i))).value
                If Not Dict_INI_LV.Exists(UCase(p_ary(i))) Then
                    Dict_INI_LV.Add UCase(p_ary(i)), PLD_Vmain.Pins(DICT_DCVS_PIN_INDEX.item(p_ary(i))).value
                End If
            ElseIf Read_Init_LV = True Then
                tmpDict_Value = Format(theexec.Specs.DC.item(Var_name).Categories(ContextCat).min.value, "0.000")
                Init_LV.AddPin p_ary(i)
                Init_LV.Pins(p_ary(i)).value = tmpDict_Value
                If Not Dict_INI_LV.Exists(UCase(p_ary(i))) Then
                    Dict_INI_LV.Add UCase(p_ary(i)), CDbl(tmpDict_Value)
                End If
            End If

            'Define DC spec Name
            If theexec.Specs.DC.Contains(VOP_VAR_Name) Then
                PL_VAR_Name = VOP_VAR_Name
            ElseIf theexec.Specs.DC.Contains(Var_name) Then
                PL_VAR_Name = Var_name
            Else
                theexec.Datalog.WriteComment " Specs Name: " & VOP_VAR_Name & " Or " & Var_name & " Not Foind ! "
            End If
            
            'Store Payload Voltage for Reference
            'Payload voltage = pmode + interpose_prepat(force voltage)
            If Read_PL_NV = True Then
                If i = 0 And ForceVDD_Cnt <> 0 Then
                    For ForceVDD_Cnt = 0 To UBound(ForcePinAry)
                        'add force voltage first
                        If ForcePinAry(ForceVDD_Cnt) <> "" Then
                            Payload_NV.AddPin ForcePinAry(ForceVDD_Cnt)
                            For Each site In theexec.sites
                                Payload_NV.Pins(ForcePinAry(ForceVDD_Cnt)).value = CDbl(ForceValueAry(ForceVDD_Cnt))
                                ForceValueAry(ForceVDD_Cnt) = CDbl(ForceValueAry(ForceVDD_Cnt))
                            Next site
                            If Not Dict_PL_NV.Exists(ForcePinAry(ForceVDD_Cnt)) Then
                                Dict_PL_NV.Add ForcePinAry(ForceVDD_Cnt), ForceValueAry(ForceVDD_Cnt)
                            End If
                        End If
                    Next ForceVDD_Cnt
                End If
                If ForcePin_Dict.Exists(p_ary(i)) = False Then
                    'add voltage from pmode category if the pin isn't exist in interpose_prepat argument
                    tmpDict_Value = Format(theexec.Specs.DC.item(PL_VAR_Name).Categories(PmodeCat).Typ.value, "0.000")
                    Payload_NV.AddPin p_ary(i)
                    Payload_NV.Pins(p_ary(i)).value = tmpDict_Value
                    If Not Dict_PL_NV.Exists(p_ary(i)) Then
                        Dict_PL_NV.Add p_ary(i), CDbl(tmpDict_Value)
                    End If
                End If
                For Each site In theexec.sites
                    PinValue = Payload_NV.Pins(p_ary(i)).value
                Next site
            End If
            If Read_PL_HV = True Then
                If i = 0 And ForceVDD_Cnt <> 0 Then
                    For ForceVDD_Cnt = 0 To UBound(ForcePinAry)
                        If ForcePinAry(ForceVDD_Cnt) <> "" Then
                            Payload_HV.AddPin ForcePinAry(ForceVDD_Cnt)
                            For Each site In theexec.sites
                                Payload_HV.Pins(ForcePinAry(ForceVDD_Cnt)).value = CDbl(ForceValueAry(ForceVDD_Cnt))
                                ForceValueAry(ForceVDD_Cnt) = CDbl(ForceValueAry(ForceVDD_Cnt))
                            Next site
                            If Not Dict_PL_HV.Exists(ForcePinAry(ForceVDD_Cnt)) Then
                                Dict_PL_HV.Add ForcePinAry(ForceVDD_Cnt), ForceValueAry(ForceVDD_Cnt)
                            End If
                        End If
                    Next ForceVDD_Cnt
                End If
                If ForcePin_Dict.Exists(p_ary(i)) = False Then
                    tmpDict_Value = Format(theexec.Specs.DC.item(PL_VAR_Name).Categories(PmodeCat).max.value, "0.000")
                    Payload_HV.AddPin p_ary(i)
                    Payload_HV.Pins(p_ary(i)).value = tmpDict_Value
                    If Not Dict_PL_HV.Exists(p_ary(i)) Then
                        Dict_PL_HV.Add p_ary(i), CDbl(tmpDict_Value)
                    End If
                End If
                For Each site In theexec.sites
                    PinValue = Payload_HV.Pins(p_ary(i)).value
                Next site
            End If
            If Read_PL_LV = True Then
                If i = 0 And ForceVDD_Cnt <> 0 Then
                    For ForceVDD_Cnt = 0 To UBound(ForcePinAry)
                        If ForcePinAry(ForceVDD_Cnt) <> "" Then
                            Payload_LV.AddPin ForcePinAry(ForceVDD_Cnt)
                            For Each site In theexec.sites
                                Payload_LV.Pins(ForcePinAry(ForceVDD_Cnt)).value = CDbl(ForceValueAry(ForceVDD_Cnt))
                                ForceValueAry(ForceVDD_Cnt) = CDbl(ForceValueAry(ForceVDD_Cnt))
                            Next site
                            If Not Dict_PL_LV.Exists(ForcePinAry(ForceVDD_Cnt)) Then
                                Dict_PL_LV.Add ForcePinAry(ForceVDD_Cnt), ForceValueAry(ForceVDD_Cnt)
                            End If
                        End If
                    Next ForceVDD_Cnt
                End If
                If ForcePin_Dict.Exists(p_ary(i)) = False Then
                    tmpDict_Value = Format(theexec.Specs.DC.item(PL_VAR_Name).Categories(PmodeCat).min.value, "0.000")
                    Payload_LV.AddPin p_ary(i)
                    Payload_LV.Pins(p_ary(i)).value = tmpDict_Value
                    If Not Dict_PL_LV.Exists(p_ary(i)) Then
                        Dict_PL_LV.Add p_ary(i), CDbl(tmpDict_Value)
                    End If
                End If
                For Each site In theexec.sites
                    PinValue = Payload_LV.Pins(p_ary(i)).value
                Next site
            End If
            
            MergeFoeceVDD = MergeFoeceVDD & IIf(i <> 0, ",", "") & p_ary(i)
            For Each site In theexec.sites
                MergeForceCond = MergeForceCond & IIf(i <> 0, ";", "") & p_ary(i) & ":V:" & Format(PinValue, "0.000")
            Next site
            SortAllPinInstrumentType LCase(p_ary(i)), DCVS_pin, DCVI_pin
            
'            If PinValue <= 0 Then ' to check if there is any abnormal value in Category
'                TheExec.Datalog.WriteComment UCase(p_ary(i)) & " is less than " & CStr(PinValue) & " please check the value of Category: " & PerforManceMode
'                TheExec.AddOutput UCase(p_ary(i)) & " is less than " & CStr(PinValue) & " please check the value of Category: " & PerforManceMode, vbRed, True
'            End If
            If PinValue.compare(LessThanOrEqualTo, 0).Any(True) Then
                theexec.Datalog.WriteComment UCase(p_ary(i)) & " is less than 0" & " please check the value of Category: " & PerforManceMode
                theexec.AddOutput UCase(p_ary(i)) & " is less than 0" & " please check the value of Category: " & PerforManceMode, vbRed, True
            End If
        End If
    Next i
    g_ForceDCVS = DCVS_pin
    g_ForceDCVI = DCVI_pin
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Decide_Shmoo_Voltage")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function dynamic_SELSRM_source_bits(SELSRAM_DSSC As String, BlockType As String, MatchedIdx As Long, Optional AllPattAry As Variant, Optional DynamicSourceBitAry As Variant, Optional testType As Variant) As String
On Error GoTo errHandler

'Mask for verify 20191125 ChrisHsu
''DSSC pin seq need modified for each project\\Hard coding\\
    Dim BitsDef As String: BitsDef = "" 'VDD_DISP,VDD_AVE,VDD_GPU,VDD_DCS_DDR,VDD_SOC
    Dim BitsDefArr() As String
    Dim SELSRAMArr() As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim L As Long
    Dim BitsOrderInfo As New Dictionary
    Dim BitsNum As Long
    Dim BlockTypeNum As Long
    Dim PattIdx As Long: PattIdx = -1
    Dim logicPin As String
    Dim SELSRM As String
    Dim DSSCSelSrmOpposite As Long
    Dim BitValue() As String
    Dim tmpPatStr() As String
    Dim TestBlock As String
    Dim TestFunc As String

    ReDim SELSRAMArr(Len(SELSRAM_DSSC) - 1)
    BitsOrderInfo.RemoveAll
    BitsDefArr = Split(BitsDef, ",")
    For i = 0 To Len(SELSRAM_DSSC) - 1
       SELSRAMArr(i) = CStr(mid(SELSRAM_DSSC, i + 1, 1))
    Next i
    
    With SelsramMapping(MatchedIdx)
        For L = 0 To UBound(.bitCount)
            If Not BitsOrderInfo.Exists(.logic_Pin(L)) Then
                BitsOrderInfo.Add (.logic_Pin(L)), SELSRAMArr(L)
            End If
        Next L
    End With
    
    ReDim BitValue(UBound(SelsramMapping(MatchedIdx).bitCount))
    For i = 0 To UBound(SelsramMapping(MatchedIdx).bitCount)
        logicPin = SelsramMapping(MatchedIdx).logic_Pin(i)
        DSSCSelSrmOpposite = SelsramMapping(MatchedIdx).SelSrm1(i)
        
        If BitsOrderInfo.Exists(logicPin) = True Then
            SELSRM = BitsOrderInfo(logicPin)
        Else
            If UCase(logicPin) Like "PRESERVED" Then
                SELSRM = SelsramMapping(MatchedIdx).SelSrm1(i) 'PCLIN modify for PRESERVED bit
            Else
                'theexec.ErrorLogMessage "Wrong Logic Pin Name in SELSRM_Mapping_Table"
                Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "dynamic_SELSRM_source_bits", "Wrong Logic Pin Name in SELSRM_Mapping_Table")
            End If
        End If
            
        If UCase(SELSRM) = "1" Then
            If DSSCSelSrmOpposite = 1 Then
                BitValue(i) = 1
            Else
                BitValue(i) = 0
            End If
        ElseIf UCase(SELSRM) = "0" Then
            If DSSCSelSrmOpposite = 1 Then
                BitValue(i) = 0
            Else
                BitValue(i) = 1
            End If
        ElseIf UCase(SELSRM) = "S" Then
            BitValue(i) = "S"
        End If
    Next i
    dynamic_SELSRM_source_bits = Join(BitValue, "")
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "dynamic_SELSRM_source_bits")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function GetSegmentData(SubSrcData As String, SrcBitsCnt As Long) As DSPWave
On Error GoTo errHandler
    Dim DataType As String
    Dim sub_wave As New DSPWave
    Dim NOT_Numeric As Boolean: NOT_Numeric = False
    Dim j As Long

    DataType = mid(SubSrcData, 1, 2)
    Select Case UCase(DataType)
        Case "0X"
            SubSrcData = GlbUtility.xHex2BinStr(SubSrcData)
        Case "0B"
            SubSrcData = Replace(UCase(SubSrcData), "0B", "")
        Case "0D"
            SubSrcData = GlbUtility.Dec2Bin(Replace(UCase(SubSrcData), "0D", ""), SrcBitsCnt)
        Case Else
            If IsNumeric(SubSrcData) Then
                'TheExec.Datalog.WriteComment "The DigSrc data format didn't support. Please check it!"
            Else
                SrcBitFromArgument = False
                sub_wave = GetStoredCaptureData(SubSrcData)
                SrcBitsCnt = sub_wave.sampleSize
                NOT_Numeric = True
            End If
    End Select
    
    If NOT_Numeric = False Then
        sub_wave.CreateConstant 0, Len(SubSrcData)
        For j = 0 To Len(SubSrcData) - 1
            sub_wave.Element(j) = CDbl(mid(SubSrcData, j + 1, 1))
        Next j
    End If
    Set GetSegmentData = sub_wave

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "GetSegmentData")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function DSPWaveReverse(SrcBitsCnt As Long, SplitSize As Long, NumberBit As Long, SubSrcData As String, InputWave As DSPWave) As DSPWave ''''for LSB case
'Public Function DSPWaveReverse(SrcBitsCnt As Long, SplitSize As Long, NumberBit As Long, SrcDataType As String, SubSrcData As String, InputWave As DSPWave, Optional s_SELSRAM As SiteVariant) As DSPWave

On Error GoTo errHandler

    Dim InvWave As New DSPWave
    Dim ResultWave As New DSPWave
    Dim Tmp_SelSram() As String
    Dim SplitWave As New DSPWave
    Dim j As Long
    Dim site As Variant
    
    ReDim Tmp_SelSram(theexec.sites.Existing.Count - 1)
    InvWave.CreateConstant 0, 1
    
    If SrcBitsCnt < SplitSize Then
        For Each site In theexec.sites.Active
            InvWave = InputWave.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspParallel, SrcBitsCnt, 0, Bit0IsMsb)
            ResultWave = InvWave.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspSerial, SrcBitsCnt, 0)
        Next
    Else
        For Each site In theexec.sites.Active
            For j = 0 To Ceiling(SrcBitsCnt / SplitSize) - 1
                If j = Ceiling(SrcBitsCnt / SplitSize) - 1 Then NumberBit = SrcBitsCnt Mod SplitSize
                SplitWave = InputWave.Select(j * SplitSize, 1, NumberBit).Copy
                InvWave = SplitWave.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspParallel, NumberBit, 0, Bit0IsMsb)
                ResultWave = InvWave.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspSerial, NumberBit, 0).Concatenate(ResultWave)
            Next
            
        Next
    End If
    
'    If UCase(SrcDataType) = "SELSRAM" Then
'        For Each site In TheExec.sites.Active
'            'Tmp_SelSram(site) = Selsram_print(site)
'            s_SELSRAM(site) = StrReverse(s_SELSRAM(site))
'        Next
'    Else
        SubSrcData = StrReverse(SubSrcData)
'    End If
    
    For Each site In theexec.sites.Active
        InputWave = ResultWave.Copy
    Next

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "DSPWaveReverse")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Shmoo_Test_Pattern(ByVal patt As Pattern, ReportResult As PFType, ResultMode As tlResultMode, ConcurrentMode As tlPatConcurrentMode, _
                                    Power_Run_Scenario As String, powerPin As String, set_init As Boolean, seq As Long, PowerCond As PinListData, wait_time As Double, _
                                    instSSNinfo As Inst_SSN, inst_info As Instance_Info, _
                                    Optional DigSrc_BitSize As String, Optional DigSrc_Seg As String, Optional DigSrc_DigSrcPin As String, Optional digSrc_EQ As String, _
                                    Optional DynamicSelSrmBits As String, Optional DigSrcType As String, Optional SelSramMatchIdx As Long, Optional PatternLoop As Long, Optional testType As String, _
                                    Optional SegmentDict As Dictionary, Optional LSB_first As Boolean, Optional DictApplyVol As Dictionary)
On Error GoTo errHandler
                                                                        
    Dim site As Variant
    Dim lPatternCount As Long
    Dim TestCase As String
    Dim DigSrc_Size As Double
    'Dim DigSrc_flag As Boolean
    Dim digcap_flag As Boolean 'add for DigCap function
    Dim DigSrc_wav As New DSPWave
    Dim DigSrc_pin As New PinList
    Dim PattArray() As String
    Dim PatCount As Long

    Dim i, j As Integer
    '========================== 'add for Multi Pat function ==========================
    Dim MultiPatAry() As String
    Dim MultiPat As Boolean
    Dim MultiPatCount As Long
    Dim CountMultiPat As Long
    '========================== 'add for Multi Pat function ==========================
    
    '========================== 'add for DigCap function ============================
    Dim DigCapName() As String
    Dim DigSrcPin As String, DigCapPin As String, DigSrcSize As String, DigCapSize As String
    Dim DigCap_Info_Dict As New Dictionary
    Dim DigCap_Pin As New PinList
    Dim OutDspWave As New DSPWave
    '========================== 'add for DigCap function ============================
    
    Dim DC_Spec_Level As New PinListData
    Dim DecodeingString As String
    Dim DigSrcTypeAry() As String
    Dim SrcBitAry() As String
    Dim Count As Long
    Dim TestNumber As Long
    Dim Suspend_Flag As Boolean
    Dim TnameCombShmooInfo As String
    Dim pattstr As String
    
    Dim RET_Shmoo As Boolean
    Dim ConditionString As String
    Dim DataType As String
    Dim sub_wave As New DSPWave
    Dim tmpWave1 As New DSPWave
    'Dim tmpWave2 As New DSPWave
    Dim DigSegAry() As String
    Dim SubSourceData As String
    Dim tmpSrcBitsCnt As Long
    'Dim tmpSrcBit As String
    Dim tmpSrcBits As Long
    Dim NOT_Numeric As Boolean
    Dim SourceBitCount As Long
    Dim sub_wave_size As Long
    
    Dim SplitSize As Long: SplitSize = 32
    Dim SplitWave As New DSPWave
    Dim InvWave As New DSPWave
    Dim NumberBit As Long: NumberBit = 32
    Dim sub_result_wave As New DSPWave
    Dim DigSrcPrintOut() As String
    Dim DictPrintOut As New Dictionary
    Dim Tmp_Output As String
    Dim Tmp_Argument As String
    Dim Tmp_SelSram() As String
    Dim SEL_print As New SiteVariant
    ReDim Tmp_SelSram(TheExec.sites.Existing.Count - 1)
    ReDim DigSrcPrintOut(TheExec.sites.Existing.Count - 1)
    
    Dim DigEQAry() As String
    
    pattstr = patt.value
    '' Add for Pattern loop ,20160607, KS
    MultiPat = False 'add for Multi Pat function
    digcap_flag = False 'add for DigCap function
    g_Retention_Shmoo = False 'add for SelSram function
    lPatternCount = 0 'initial PatternCount for pat loops
    MultiPatCount = 0 'initial Multi patterns count

    TheExec.Datalog.WriteComment ""
    If wait_time <> 0 Then g_Retention_Shmoo = True ' use for non SELSRM Function
        '========================================================================Process SELSRM format=====================================================================

    If PatternLoop > 1 Then TheExec.Datalog.WriteComment "Loop Pattern :" & patt & "_" & "Repeat count :" & PatternLoop
    If LCase(patt.value) Like "*.pat*" Then
        ReDim PattArray(0)
        PattArray(0) = patt.value
        PatCount = 1
    Else
        Call PATT_GetPatListFromPatternSet(patt.value, PattArray, PatCount)
    End If    
    If digSrc_EQ <> "" Then
        Set DigSrc_wav = Nothing
        DigEQAry = Split(digSrc_EQ, "+")
        DigSrcPin = IIf(DigSrc_DigSrcPin <> "", DigSrc_DigSrcPin, "JTAG_TDI")
        
        For i = 0 To UBound(DigEQAry)
            Set sub_wave = Nothing
            Set InvWave = Nothing
            sub_wave_size = 0
            NOT_Numeric = False
            SubSourceData = vbNullString
            tmpSrcBits = 0
            
            'SELSRAM case
'            If UCase(DigEQAry(i)) = "SELSRAM" Then
'                If InStr(UCase(DynamicSelSrmBits), "S") > 0 Then
'                    sub_wave.CreateConstant 0, Len(DynamicSelSrmBits)
'                    Decide_DC_Level_Mod DC_Spec_Level, testType
'                    DynamicSelSrmBits = Decide_Switching_Bit(DynamicSelSrmBits, sub_wave, DC_Spec_Level, "", DecodeingString, powerPin, g_Globalpointval, g_ForceCond_VDD, g_CharInputString_Voltage_Dict, patt.value, , , , SEL_print)
'                End If
'                tmpSrcBitsCnt = CLng(Len(DynamicSelSrmBits))
'                For Each site In TheExec.sites.Active
'                    Tmp_SelSram(site) = SEL_print(site)
'                Next
'            'EMA case
'            ElseIf UCase(DigEQAry(i)) = "EMA" Then
'                TestCase = DynamicSelSrmBits
'                Call GetSrcString_fromEMAArray(patt.value, TestCase, SubSourceData, DigSrc_Size, SrcBitAry)
'                tmpSrcBitsCnt = CLng(DigSrc_Size)
'                'sub_Wave
'            'MULTIFSTP
'            ElseIf UCase(DigEQAry(i)) = "MULTIFSTP" Then
'                'sub_Wave
'            'SEGMENT & DICTIONARY
'            Else
                If SegmentDict.Exists(DigEQAry(i)) Then
                    SubSourceData = SegmentDict.item(DigEQAry(i))(0)
                    tmpSrcBitsCnt = SegmentDict.item(DigEQAry(i))(1)
                    If UCase(SubSourceData) = "SELSRAM" Then
                            sub_wave.CreateConstant 0, Len(DynamicSelSrmBits)
                            Decide_DC_Level_Mod DC_Spec_Level, testType
                        DynamicSelSrmBits = Decide_Switching_Bit(DynamicSelSrmBits, sub_wave, DC_Spec_Level, "", DecodeingString, powerPin, g_Globalpointval, g_ForceCond_VDD, DictApplyVol, patt.value, , , , SEL_print, True)
                        digSrc_EQ_GB = DynamicSelSrmBits

                        'tmpSrcBitsCnt = CLng(Len(DynamicSelSrmBits))
                        For Each site In theexec.sites.Active
                            Tmp_SelSram(site) = SEL_print(site)
                        Next
                    Else
                        SrcBitFromArgument = True
                        sub_wave = GetSegmentData(SubSourceData, tmpSrcBitsCnt)
                    End If
                Else
                    sub_wave.CreateConstant 0, tmpSrcBitsCnt
                End If
            'End If
            
            For Each site In theexec.sites.Active
                sub_wave_size = sub_wave(site).sampleSize
                Exit For
            Next site
            
            'Final source bit size > data bit size - add dummy bit
            If tmpSrcBitsCnt > sub_wave_size Then
                TheExec.Datalog.WriteComment "<warning> The source bit size:" & tmpSrcBitsCnt & "> data bit size:" & sub_wave_size & "! Please check!"
                tmpWave1.CreateConstant 0, tmpSrcBitsCnt - sub_wave_size
                For Each site In theexec.sites.Active
                    sub_wave = tmpWave1.Concatenate(sub_wave.Copy)
                Next
                'tmpSrcBitsCnt = sub_wave_size
                If UCase(SubSourceData) = "SELSRAM" Then
                    For Each site In TheExec.sites.Active
                        SEL_print(site) = String(tmpSrcBitsCnt - sub_wave_size, "0") + SEL_print(site)
                    Next site
                Else
                SubSourceData = String(tmpSrcBitsCnt - sub_wave_size, "0") & SubSourceData
                End If
            'Final source bit size < data bit size - remove dummy bit
            ElseIf tmpSrcBitsCnt < sub_wave_size Then
                TheExec.Datalog.WriteComment "<warning> The source bit size:" & tmpSrcBitsCnt & "< data bit size:" & sub_wave_size & "! Please check!"
                Dim chk_wave As New DSPWave
                For Each site In TheExec.sites.Active
                chk_wave = sub_wave.Select(1, 1, sub_wave_size - tmpSrcBitsCnt).Copy
                If chk_wave.CalcSum <> 0 Then theexec.Datalog.WriteComment "<warning> The dummy bit is included non-zero value! Please check!"
                sub_wave = sub_wave.Select(sub_wave_size - tmpSrcBitsCnt, 1, tmpSrcBitsCnt).Copy
                Next site
                    
                    If UCase(SubSourceData) = "SELSRAM" Then
                        For Each site In TheExec.sites.Active
                            SEL_print(site) = mid(SEL_print(site), sub_wave_size - tmpSrcBitsCnt + 1, tmpSrcBitsCnt)
                        Next site
                    Else
                SubSourceData = mid(SubSourceData, sub_wave_size - tmpSrcBitsCnt + 1, tmpSrcBitsCnt)
            End If

            End If

            If LSB_first And SrcBitFromArgument Then
'                If UCase(DigSegAry(i)) = "SELSRAM" Then
'                    For Each site In TheExec.sites.Active
'                        Tmp_SelSram(site) = SEL_print(site)
'                    Next
'                End If
                sub_wave = DSPWaveReverse(tmpSrcBitsCnt, SplitSize, NumberBit, SubSourceData, sub_wave)
                SrcBitFromArgument = False
            End If
            
            If UCase(DigEQAry(i)) <> "SELSRAM" And SegmentDict.Exists(DigEQAry(i)) Then _
                theexec.Datalog.WriteComment "DigSrc data from argument : " & DigEQAry(i) & " = " & SegmentDict.item(DigEQAry(i))(0)
            
            For Each site In theexec.sites.Active
                DigSrc_wav = DigSrc_wav.Concatenate(sub_wave.Copy)
                If UCase(SubSourceData) = "SELSRAM" Then
                    DigSrcPrintOut(site) = DigSrcPrintOut(site) + SEL_print(site)
                    TheExec.Datalog.WriteComment "site " & site & " DigSrc data from SelSram: " & DigEQAry(i) & " = " & SEL_print(site)
                Else
                    DigSrcPrintOut(site) = DigSrcPrintOut(site) + SubSourceData
                End If
            Next
            SourceBitCount = SourceBitCount + tmpSrcBitsCnt
        Next i
        'source to hardware
        DigSrc_pin.value = DigSrcPin
        dssc_pat_init_GB = PattArray(0)
        Dim sSrcSigName As String
        Dim tempVarArray As Variant
        tempVarArray = thehdw.DSSC.Pins(DigSrc_pin).Pattern(PattArray(0)).Source.Labels.list
        sSrcSigName = tempVarArray(0)
        If sSrcSigName = "" Then
            sSrcSigName = "FUNC_SRC"
        End If
        Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, sSrcSigName, SourceBitCount, DigSrc_wav)
            
'        If TheExec.Flow.EnableWord("Debug_LVCC_VminBoundary") = True Or TheExec.Flow.EnableWord("Debug_HVCC_VminBoundary") = True Then
'            digSrc_EQ_GB = digSrc_EQ:: BlockType_GB = "GFXSCAN":: DigSrcSize_GB = DigSrcSize:: dssc_pat_init_GB = PattArray(0):: DigSrc_pin_GB = DigSrc_pin:: TestType_GB = testType
        
        DigSrc_pin_GB = DigSrc_pin
        Tmp_Output = "DigSrc pattern = " & IIf(set_init = True, "Init", "Payload") & seq & ": " & patt & "," & "Src Bits = " & SourceBitCount & "," & "Output String [ LSB(L) ==> MSB(R) ]:" & "SourceCode:"
        
        For Each site In theexec.sites.Active
            theexec.Datalog.WriteComment "Site " & CStr(site) & " " & Tmp_Output & DigSrcPrintOut(site)
        Next
        theexec.Datalog.WriteComment ""
    End If
    ''20240521 MFSTP create Digital source Wave
    If CZ_inst_info.MultiFSTP_Enable Then
        If InStr(PattArray(0), inst_info.digSrcPatterns(inst_info.idxDigSrcPattern)) Then
            inst_info.Harvest_Core_DigSrc_Pin = IIf(DigSrc_DigSrcPin <> "", DigSrc_DigSrcPin, "JTAG_TDI")
            sSrcSigName = CheckPatLabel(PattArray(0), inst_info.Harvest_Core_DigSrc_Pin.value)
            Call Calculate_Harvest_Core_DSSC_Source_For_UserFunction(inst_info.digSrcPatterns(inst_info.idxDigSrcPattern), CStr(inst_info.digSrcLabel(inst_info.idxDigSrcPattern)), inst_info.Harvest_Core_DigSrc_Pin, sSrcSigName)
            'inst_info.idxDigSrcPattern = inst_info.idxDigSrcPattern + 1
        End If
    End If
    
    If InStr(patt, ",") > 0 Then 'Multi Pattern function
        MultiPatAry = Split(patt, ",")
        MultiPatCount = UBound(MultiPatAry)
        MultiPat = True
    End If
''===========================================================SET RUN LEVEL=========================================================
    Set_Run_Level Power_Run_Scenario, powerPin, set_init, seq
''===========================================================SET RUN LEVEL=========================================================

Dim pattmp() As String

    For CountMultiPat = 0 To MultiPatCount  'Multi pat function
    
        If MultiPat = True Then
           Call TheHdw.Patterns(MultiPatAry(CountMultiPat)).Load
        Else
           Call TheHdw.Patterns(patt).Load
        End If
        
        ' Setup HRAM
        With TheHdw.Digital
            .Patgen.HaltMode = tlHaltOnOpcode
            .hram.size = gl_HRAMmaxDepth
            .hram.CaptureType = captFail
            .hram.SetTrigger trigFirst, True, 0, True
        End With
     
        'For Count = 0 To lPatternCount
     '********************************************************** SSN VBT **************************************
        Dim shortPatName As String
        Dim SCAN_Site_Blooean As New SiteBoolean
        Dim patpath As Variant
        Dim PatCnt As Long
        	If LCase(patt.value) Like "*.pat*" Then
                ReDim patpath(0)
                patpath(0) = patt.value
            Else
                patpath = TheExec.DataManager.Raw.GetPatternsInSet(patt.value, PatCnt)
            End If
            If glb_TesterType = "UltraFLEXplus" Then
                If instSSNinfo.bSSNTest = True And LCase(patt) Like "*_pl*" Then     '20231120: Added pl only
                
                    shortPatName = SPN(CStr(patpath(0)))       ''Split pattern name with path
                    If (TheExec.Flow.IsCharacterizing = False) Then
                        If instSSNinfo.bSSNCoreHarvest = True Then
                            Dim SSN_ScanPins As String: SSN_ScanPins = ""
                            Dim maxFailsPerPin As Long: maxFailsPerPin = glb_SSN_CaptureLimit
                            
                            SSN_PreBody True, SSN_ScanPins, maxFailsPerPin, tlDatalogScanResultMode_Module, True, shortPatName
                        End If
                    Else
                        ' Running Char flow
                    End If
                End If
            End If
    '********************************************************** SSN VBT **************************************
        
            ''    Dim sSrcSigName As String
             ''   Dim tempVarArray As Variant
        
              ''  tempVarArray = TheHdw.DSSC.Pins("JTAG_TDI").Pattern(patt).Source.Labels.list
               '' sSrcSigName = tempVarArray(0)
    
        
        
            If MultiPat = True Then
              Call thehdw.Patterns(MultiPatAry(CountMultiPat)).start
            Else
                'Call thehdw.Patterns(patt).start
                'Exit Function
                pattmp = Split(patpath(0), ":")
                Call thehdw.Patterns(pattmp(0)).start
            End If
            While thehdw.Digital.Patgen.IsRunning = True
                thehdw.Digital.Patgen.Continue 0, cpuA
            Wend
            thehdw.Digital.Patgen.HaltWait
            
             '20231219: Modified to remove site loop
            Dim sitePatResult As New SiteBoolean
            sitePatResult = thehdw.Digital.Patgen.PatternBurstPassedPerSite
            
            SCAN_Site_Blooean = SCAN_Site_Blooean.LogicalAnd(sitePatResult)
            
    
            
        'Next Count
        
        ' If  Suspend_Flag=False the Tname need to include all information from X,Y,Z current point
        If theexec.DevChar.Setups.IsRunning = True Then Suspend_Flag = theexec.DevChar.Setups.item(theexec.DevChar.Setups.ActiveSetupName).Output.SuspendDatalog
        
        If TPModeAsCharz_GLB = True Then
            For Each site In theexec.sites
                TnameCombShmooInfo = vbNullString
                TestNumber = theexec.sites.item(site).TestNumber
                If thehdw.Digital.Patgen.PatternBurstPassed(site) Then
                    theexec.sites.item(site).testResult = sitePass
                    If theexec.DevChar.Setups.IsRunning = True Then
                        If Suspend_Flag = False Then
                            Call PrintEachPoint_TestName(TnameCombShmooInfo)
                            ConditionString = TnameCombShmooInfo
                            TnameCombShmooInfo = theexec.DataManager.instancename & TnameCombShmooInfo
                            If g_TestNum <> -1 Then
                                Call theexec.Datalog.WriteFunctionalResult(site, g_TestNum, logTestPass, , TnameCombShmooInfo)
                            Else
                                Call theexec.Datalog.WriteFunctionalResult(site, TestNumber, logTestPass, , TnameCombShmooInfo)
                            End If
                        Else
                            Call theexec.Datalog.WriteFunctionalResult(site, TestNumber, logTestPass)
                        End If
                    Else
                        Call theexec.Datalog.WriteFunctionalResult(site, TestNumber, logTestPass)
                    End If
                Else
                    theexec.sites.item(site).testResult = siteFail
                    If theexec.DevChar.Setups.IsRunning = True Then
                        If Suspend_Flag = False Then
                            Call PrintEachPoint_TestName(TnameCombShmooInfo)
                            TnameCombShmooInfo = theexec.DataManager.instancename & TnameCombShmooInfo
                            ConditionString = TnameCombShmooInfo
                            If g_TestNum <> -1 Then
                               Call theexec.Datalog.WriteFunctionalResult(site, g_TestNum, logTestFail, , TnameCombShmooInfo)
                            Else
                               Call theexec.Datalog.WriteFunctionalResult(site, TestNumber, logTestFail, , TnameCombShmooInfo)
                            End If
                        Else
                            Call theexec.Datalog.WriteFunctionalResult(site, TestNumber, logTestFail)
                        End If
                    Else
                        Call theexec.Datalog.WriteFunctionalResult(site, TestNumber, logTestFail)
                    End If
                End If
                If theexec.DevChar.Setups.IsRunning = False Then theexec.sites.item(site).TestNumber = TestNumber + 1
                If theexec.DevChar.Setups.IsRunning = True And Suspend_Flag = False And g_TestNum = -1 Then theexec.sites.item(site).TestNumber = TestNumber + 1
            Next site
        Else
            HardIP_WriteFuncResult , , , , pattstr
        End If
        
        If Suspend_Flag = False And g_TestNum <> -1 Then g_TestNum = g_TestNum + 1
        
                
     '********************************************************** SSN VBT **************************************
        If (glb_TesterType = "UltraFLEXplus") Then
            If instSSNinfo.bSSNTest = True And instSSNinfo.bSSNCoreHarvest = True Then
                If LCase(patt) Like "*_pl*" Then   '20231120: Added pl only
                    If (theexec.Flow.IsCharacterizing = False) Then
                        Dim SSN_InterposeFunc As InterposeName ': SSN_InterposeFunc = "LogPinCoreFailsBySite"
                        SSN_Body instSSNinfo, True, True, shortPatName, , , SSN_InterposeFunc  ', SSN_ScanPins
                        SSN_PostBody False
                    End If
                Else    ''Init SSN fail
'                        For Each site In TheExec.sites
'                            If SCAN_Site_Blooean(site) = False Then
'                                TheExec.sites(site).FlagState(glb_SSN_Failflag) = logicTrue
'                                TheExec.Flow.TestLimit resultVal:=1, lowVal:=0, hiVal:=0, Tname:=CStr(glb_SSN_Failflag), ForceResults:=tlForceNone
'                            End If
'                        Next site
                End If
            End If
        End If
    '********************************************************** SSN VBT **************************************
            
    Next CountMultiPat
    
    If DigCapPin <> "" Then ' Process DSP Capture Data
       Call CreateSimulateDataDSPWave(OutDspWave, CLng(DigCapSize), CLng(DigCapSize))
       Call Char_Process_DSP_Capture(DigCapName, OutDspWave, DigCap_Info_Dict, CStr(DigCap_Pin))
    End If
            
    For Each site In theexec.sites
        ' Prevent over write shmoo pattern
        DebugPrintFunc patt.value
    Next site
    
    If wait_time <> 0 Then
        If g_PLSWEEP = False Then 'non-SWEEP case
            If theexec.DevChar.Setups.IsRunning = True Then
                RET_Shmoo = True
                g_Retention_Info = g_Retention_Info & IIf(g_Retention_Info <> "", ",", "") & patt.value & "|" & wait_time
            Else
                RET_Shmoo = False
            End If
            Shmoo_Restore_Retention_Power powerPin, wait_time, RET_Shmoo, set_init, seq
        Else 'SWEEP case
           TheHdw.Wait wait_time
        End If
         If set_init = True Then
           TheExec.Datalog.WriteComment "wait " & wait_time & " after init pattern " & seq
        Else
           TheExec.Datalog.WriteComment "wait " & wait_time & " after payload pattern " & seq - InitialPatCnt
        End If
    End If
    
    patt.value = pattstr
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Test_Pattern")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Set_Run_Level(Power_Run_Cond As String, powerPin As String, set_init As Boolean, seq As Long)
On Error GoTo errHandler

    Dim current_level As String
    Dim EnableSweep As Boolean:: EnableSweep = False
    
    SweepGuardBand = False
    current_level = "-99"
    
    If set_init = True Then 'Init Condition
        Select Case UCase(Power_Run_Cond)
        Case "NV"
            current_level = "INIT_NV"
        Case "HV"
            current_level = "INIT_HV"
        Case "LV"
            current_level = "INIT_LV"
        Case "SWEEP"
            current_level = "INIT_Sweep"
            EnableSweep = True
        Case "SweepGuardBand"
            current_level = "INIT_SweepGuardBand"
            EnableSweep = True
        Case Else
            theexec.Datalog.WriteComment " The Scenario " & Power_Run_Cond & " ,Not Support ! "
        End Select
    Else    'Playload Condition
        Select Case UCase(Power_Run_Cond)
        Case "NV"
            current_level = "PL_NV"
        Case "HV"
            current_level = "PL_HV"
        Case "LV"
            current_level = "PL_LV"
        Case "VRS"
            'thehdw.DCVS.pins(g_AllDCVSPin).Voltage.Output = tlDCVSVoltageMain
            current_level = "PL_VRS"
        Case "SWEEP"
            current_level = "PL_Sweep"
            EnableSweep = True
        Case "SWEEPGUARDBAND"
            current_level = "PL_SweepGuardBand"
            EnableSweep = True
        Case Else
            theexec.Datalog.WriteComment " The Scenario " & Power_Run_Cond & " ,Not Support ! "
        End Select
        If g_INIT_PAT_DONE = True Then current_level = current_level & "_" & CStr(Sweep_cnt)
    End If
    If current_level <> "-99" And Not (Power_Level_Last Like current_level) Then
        If current_level = "PL_VRS" Then thehdw.DCVS.Pins(ALL_Power_DCVS_pins).Voltage.Output = tlDCVSVoltageMain
        Shmoo_Restore_Power powerPin, set_init, seq, EnableSweep
    End If
    Power_Level_Last = current_level
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Set_Run_Level")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Shmoo_Restore_Power(Shmoo_Apply_Pin As String, SetInit As Boolean, seq As Long, isSweep As Boolean)
On Error GoTo errHandler
    
    Dim p_ary() As String, p_cnt As Long, i As Long, j As Long
    Dim InstName As String
    Dim site As Variant
    Dim Shmoo_Apply_Pin_Arry() As String
    Dim SramShmooPower As String: SramShmooPower = ""
    Dim Shmoo_Dict As New Dictionary
    Dim DCVS_flag As Boolean: DCVS_flag = False
    On Error GoTo errHandler
    
    If isSweep = True Then
        Shmoo_Apply_Pin_Arry = Split(Shmoo_Apply_Pin, ",")
        For i = 0 To UBound(Shmoo_Apply_Pin_Arry)
            Shmoo_Dict.Add LCase(Shmoo_Apply_Pin_Arry(i)), True
        Next i
    End If
    
    If g_ForceCond_VDD <> "" Then
        theexec.DataManager.DecomposePinList g_ForceCond_VDD, p_ary, p_cnt
        For i = 0 To p_cnt - 1
            If gl_GetInstrument_Dic.Exists(LCase(p_ary(i))) Then
                If isSweep = True And Shmoo_Dict.Exists(LCase(p_ary(i))) Then
                    'Bypass Apply Shmoo Voltage
                Else
                    InstName = gl_GetInstrument_Dic.item(LCase(p_ary(i)))
                    Select Case InstName
                    Case "VHDVS", "HexVS", "VSM", "VS-5A", "VS-800mA"
                        DCVS_flag = True
                        If SetInit = True Then
                            thehdw.DCVS.Pins(p_ary(i)).Voltage.Main.value = g_CharPattInfoAry(seq - 1).ForceVoltage.Pins(UCase(p_ary(i))).value
                        Else
                            For Each site In theexec.sites
                                thehdw.DCVS.Pins(p_ary(i)).Voltage.Alt.value = g_CharPattInfoAry(seq - 1).ForceVoltage.Pins(UCase(p_ary(i))).value
                            Next site
                        End If
                    Case "DC-07", "DC-30"
                        'If SetInit = True Then
                        For Each site In theexec.sites
                            thehdw.DCVI.Pins(p_ary(i)).Voltage = g_CharPattInfoAry(seq - 1).ForceVoltage.Pins(UCase(p_ary(i))).value
                        Next site
                        'Else
                        '    thehdw.DCVI.Pins(p_ary(i)).Voltage = g_CharPattInfoAry(seq - 1).ForceVoltage.Pins(UCase(p_ary(i))).value
                        'End If
                    Case "HSD-U"
                    Case Else
                        Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Restore_Power", _
                        "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported")
                    End Select
                End If
            End If
        Next i
     End If
    
    If isSweep = True And Shmoo_Apply_Pin <> "" Then
        Shmoo_Apply_Pin_Arry = Split(Shmoo_Apply_Pin, ",")
        For i = 0 To UBound(Shmoo_Apply_Pin_Arry)
            If Not UCase(Shmoo_Apply_Pin_Arry(i)) Like UCase(SramShmooPower) Or SramShmooPower = "" Then
                If gl_GetInstrument_Dic.Exists(LCase(Shmoo_Apply_Pin_Arry(i))) Then
                InstName = gl_GetInstrument_Dic.item(LCase(Shmoo_Apply_Pin_Arry(i)))
                    Select Case InstName
                        Case "VHDVS", "HexVS", "VSM", "VS-5A", "VS-800mA"
                            DCVS_flag = True
                            If SetInit = True Then
                                For Each site In theexec.sites
                                    With g_CharPattInfoAry(seq - 1)
                                        If .GuardBandVal <> 0 Then
                                            thehdw.DCVS.Pins(Shmoo_Apply_Pin_Arry(i)).Voltage.Main.value = CalcPower(g_Globalpointval.Pins(Shmoo_Apply_Pin_Arry(i)).value, .GuardBandSymbol, .GuardBandVal)
                                        Else
                                            thehdw.DCVS.Pins(Shmoo_Apply_Pin_Arry(i)).Voltage.Main.value = g_Globalpointval.Pins(Shmoo_Apply_Pin_Arry(i)).value
                                        End If
                                    End With
                                Next site
                            Else
                                For Each site In theexec.sites
                                    With g_CharPattInfoAry(seq - 1)
                                        If .GuardBandVal <> 0 Then
                                            thehdw.DCVS.Pins(Shmoo_Apply_Pin_Arry(i)).Voltage.Alt.value = CalcPower(g_Globalpointval.Pins(Shmoo_Apply_Pin_Arry(i)).value, .GuardBandSymbol, .GuardBandVal)
                                        Else
                                            thehdw.DCVS.Pins(Shmoo_Apply_Pin_Arry(i)).Voltage.Alt.value = g_Globalpointval.Pins(Shmoo_Apply_Pin_Arry(i)).value
                                        End If
                                    End With
                                Next site
                                
                                'TheExec.Flow.TestLimit thehdw.DCVS.Pins(p_ary(i)).Voltage.Alt.value, , , , , , , , p_ary(i) & "_shmoo_Valt"
                             End If
                        Case "DC-07", "DC-30"
                            If SetInit = True Then
                                For Each site In theexec.sites
                                    With g_CharPattInfoAry(seq - 1)
                                        If .GuardBandVal <> 0 Then
                                            thehdw.DCVI.Pins(Shmoo_Apply_Pin_Arry(i)).Voltage = CalcPower(g_Globalpointval.Pins(Shmoo_Apply_Pin_Arry(i)).value, .GuardBandSymbol, .GuardBandVal)
                                        Else
                                            thehdw.DCVI.Pins(Shmoo_Apply_Pin_Arry(i)).Voltage = g_Globalpointval.Pins(Shmoo_Apply_Pin_Arry(i)).value
                                        End If
                                    End With
                                Next site
                             Else
                                For Each site In theexec.sites
                                    With g_CharPattInfoAry(seq + InitialPatCnt - 1)
                                        If .GuardBandVal <> 0 Then
                                            thehdw.DCVI.Pins(Shmoo_Apply_Pin_Arry(i)).Voltage = CalcPower(g_Globalpointval.Pins(Shmoo_Apply_Pin_Arry(i)).value, .GuardBandSymbol, .GuardBandVal)
                                        Else
                                            thehdw.DCVI.Pins(Shmoo_Apply_Pin_Arry(i)).Voltage = g_Globalpointval.Pins(Shmoo_Apply_Pin_Arry(i)).value
                                        End If
                                    End With
                                Next site
                             End If
                        Case "HSD-U"
                        Case Else
                             Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Restore_Power", _
                            "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported")
                    End Select
             End If
           End If
        Next i
        'End If
    End If
    If DCVS_flag Then
        If SetInit = True Then
            thehdw.DCVS.Pins(g_ForceDCVS).Voltage.Output = tlDCVSVoltageMain
        Else
            thehdw.DCVS.Pins(g_ForceDCVS).Voltage.Output = tlDCVSVoltageAlt
        End If
    End If
    
''wait time for Vmain switch to Valt
     thehdw.Wait 0.001
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Restore_Power")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Shmoo_Restore_Retention_Power(Shmoo_Apply_Pin As String, WaitTime As Double, RetentionShmoo As Boolean, set_init As Boolean, seq As Long, _
                                            Optional ConditionString As String = vbNullString)
On Error GoTo errHandler
                                                                                        
    Dim p_ary() As String, p_cnt As Long, i As Long
    Dim InstName As String
    Dim site As Variant
    Dim Shmoo_Apply_Pin_Arry() As String
    Dim SRAMRampUpFirst As New SiteBoolean
    Dim LogicRampdownFirst As New SiteBoolean
    Dim SramShmooPower As String: SramShmooPower = vbNullString
    Dim RetentionShmoo_Pins_Dict As New Dictionary
    Dim Ret_Dict As New Dictionary
    Dim Retention_ForceV_Arr() As String
    Dim RetentionPinsAry() As String
    Dim RetentionRestorePLD As New PinListData
    Dim RestoreMainOrAlt As New PinListData
    Dim DCVSOutPutSelector As Long
    Dim dDCVSOutPutSelector As Long
    On Error GoTo errHandler

    '----------------------------------------------------------------------------------------------------
    'Retention flow
    '----------------------------------------------------------------------------------------------------
    '- Drop Voltage
    '   a. Shmoo case (Shmoo Pin)
    '       1.[Enable Word "Disable_RET_RampDownUp"] = TRUE : Drop voltage to shmoo voltage directly
    '       2.[Enable Word "Disable_RET_RampDownUp"] = FALSE: Drop voltage to shmoo voltage by rampping
    '   b. Retention pin & voltage from argument case (Retention Pin)
    '       1.[Enable Word "Disable_RET_RampDownUp"] = TRUE : Drop voltage to retention voltage directly
    '
    '- Wait time
    '
    '- Rise Voltage
    '   a. Shmoo case (Shmoo Pin)
    '       1.[Enable Word "Disable_RET_RampDownUp"] = TRUE : Rise voltage to restore voltage directly
    '       2.[Enable Word "Disable_RET_RampDownUp"] = FALSE: Rise voltage to restore voltage by rampping
    '   b. Retention pin & voltage from argument case (Retention Pin)
    '       1.[Enable Word "Disable_RET_RampDownUp"] = TRUE : Rise voltage to restore voltage directly
    '
    '- Need to restore Vmain or Valt ("not Output one")
    '-----------------------------------------------------------------------------------------------------
    
    Dim PinName As String
    Dim ApplyPin As String
    Dim Disable_RET_Ramp As Boolean: Disable_RET_Ramp = theexec.Flow.enableWord("Disable_RET_RampDownUp")
    Dim RegetOutput As Long
    Dim RetentionPin As String
    Dim ApplyPinArr() As String
    Dim TargetVal As New PinListData
    Dim ExistPosition As Long
    Dim ReplaceStr As String
    Dim T0, T1 As Double
    Const OutputIsVmain = 1
    Const OutputIsValt = 2
    Const VmainFirst = 1
    Const ValtFirst = 2
    Dim RampStep As Integer: RampStep = 9
    Dim RampTime As Double: RampTime = 20 * 0.000001
    Dim RampOffset As Double: RampOffset = 0
    Dim RampOffsetSymbol As String
    Dim DCVS_RET As String
    Dim DCVI_RET As String
    Dim p_ret_ary() As String
    Dim RetPin_dic As New Dictionary
    Dim ShmooPin_dic As New Dictionary


    With g_CharPattInfoAry(seq - 1)
        If .RampStep <> 0 Then RampStep = .RampStep
        If .RamppingTime <> 0 Then RampTime = .RamppingTime
        RampOffset = .RampOffset
        RampOffsetSymbol = .RampOffsetSymbol
    End With
    
    If Shmoo_Apply_Pin = "" And g_Retention_VDD = "" Then Exit Function
    
    If ALL_Power_DCVS_pins <> "" Then
        DCVSOutPutSelector = thehdw.DCVS.Pins(ALL_Power_DCVS_pins).Voltage.Output
    Else
        DCVSOutPutSelector = -1
        dDCVSOutPutSelector = -1
    End If
    
    If DCVSOutPutSelector = OutputIsVmain Then dDCVSOutPutSelector = ValtFirst
    If DCVSOutPutSelector = OutputIsValt Then dDCVSOutPutSelector = VmainFirst
    Set g_RestoreMainOrAlt = Nothing
    
    If Shmoo_Apply_Pin = "" And g_Retention_VDD = "" Then Exit Function
    
    If g_Retention_VDD <> "" Then
        RetentionPin = g_Retention_VDD
        p_ret_ary = Split(g_Retention_VDD, ",")
        
        For i = 0 To UBound(p_ret_ary)
            'Check_DCVSorDCVI LCase(p_ret_ary(i)), DCVS_RET, DCVI_RET
            SortAllPinInstrumentType LCase(p_ret_ary(i)), DCVS_RET, DCVI_RET
        Next
    End If
    
    If Shmoo_Apply_Pin <> "" Then
        p_ary = Split(Shmoo_Apply_Pin, ",")
        For i = 0 To UBound(p_ary)
            'Check_DCVSorDCVI LCase(p_ary(i)), DCVS_RET, DCVI_RET
            SortAllPinInstrumentType LCase(p_ary(i)), DCVS_RET, DCVI_RET
        Next
    End If
   
    '''20240704 remove duplicate pins in RetentionPin''''''
    
     p_ary = Split(RetentionPin, ",")
     RetentionPin = vbNullString
     For i = 0 To UBound(p_ary)
        If RetPin_dic.Exists(UCase(p_ary(i))) = False Then
            RetPin_dic.Add UCase(p_ary(i)), i
            RetentionPin = RetentionPin & IIf(RetentionPin <> "", ",", "") & p_ary(i)
        Else
            Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Restore_Retention_Power", "Duplicate pin, please check")
            For Each site In TheExec.sites
                TheExec.sites.item(site).result = tlResultFail
            Next site
        End If
     Next i
     p_ary = Split(Shmoo_Apply_Pin, ",")
     For i = 0 To UBound(p_ary)
        If ShmooPin_dic.Exists(UCase(p_ary(i))) = False Then
            If RetPin_dic.Exists(UCase(p_ary(i))) = False Then
                ShmooPin_dic.Add UCase(p_ary(i)), i
                RetentionPin = RetentionPin & IIf(RetentionPin <> "", ",", "") & p_ary(i)
            Else
				ShmooPin_dic.Add UCase(p_ary(i)), i
            End If
        Else
            Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Restore_Retention_Power", "Duplicate pin, please check")
            For Each site In TheExec.sites
                TheExec.sites.item(site).result = tlResultFail
            Next site
        End If
     Next i

    
    Call AddPinToPinListData(RetentionPin, RetentionPinsAry, RestoreMainOrAlt)
    RetentionRestorePLD = RestoreMainOrAlt.Copy
    TargetVal = RestoreMainOrAlt.Copy
    
    For i = 0 To RestoreMainOrAlt.Pins.Count - 1
        PinName = RestoreMainOrAlt.Pins(i).name
        For Each site In TheExec.sites
            'If set_init = True Then
                RestoreMainOrAlt.Pins(PinName).value = g_CharPattInfoAry(seq - 1).ForceVoltage.Pins(PinName).value
            'Else
            '    RestoreMainOrAlt.Pins(PinName).value = g_CharPattInfoAry(seq - 1).ForceVoltage.Pins(PinName).value
            'End If
			If RetPin_dic.Exists(UCase(PinName)) = True Then TargetVal.Pins(PinName).value = g_RetntionVal.Pins(PinName).value
            If ShmooPin_dic.Exists(UCase(PinName)) = True Then TargetVal.Pins(PinName).value = g_Globalpointval.Pins(PinName).value
            If RampOffsetSymbol <> "" Then
                Select Case RampOffsetSymbol
                    Case "+":
                        TargetVal.Pins(PinName).value = TargetVal.Pins(PinName).value + RampOffset
                    Case "-":
                        TargetVal.Pins(PinName).value = TargetVal.Pins(PinName).value - RampOffset
                    Case Else
                End Select
            End If
        Next
    Next
    
    'Drop to target vol directly
    If Disable_RET_Ramp Then
        For i = 0 To RestoreMainOrAlt.Pins.Count - 1
            PinName = RestoreMainOrAlt.Pins(i).name
            SetTargetVol PinName, DCVSOutPutSelector, RetentionRestorePLD, TargetVal, set_init, RestoreMainOrAlt, seq
        Next i
    'Drop to target vol by rampping
    Else
        Retention_Ramping RetentionPin, g_CharPattInfoAry(seq - 1).ForceVoltage, TargetVal, DCVSOutPutSelector, set_init, RampStep, RampTime, DCVS_RET
        If set_init = True Then
            'Retention_Ramping RetentionPin, g_CharPattInfoAry(seq - 1).ForceVoltage, TargetVal, DCVSOutPutSelector, set_init, RampStep, RampTime
        Else
            'Retention_Ramping RetentionPin, g_CharPattInfoAry(seq - 1).ForceVoltage, TargetVal, DCVSOutPutSelector, set_init, RampStep, RampTime
            If RampOffset > 0 Then
                For i = 0 To RestoreMainOrAlt.Pins.Count - 1
                    PinName = RestoreMainOrAlt.Pins(i).name
                    For Each site In theexec.sites
                        theexec.Datalog.WriteComment "Site: " & site & ";" & PinName & ": " & TargetVal.Pins(PinName).value & " V(Guard Band)"
                    Next
                Next
            End If
        End If
    End If
    
    If Disable_RET_Ramp Then
        If DCVSOutPutSelector = OutputIsValt Then
            thehdw.DCVS.Pins(DCVS_RET).Voltage.Output = tlDCVSVoltageMain
            RegetOutput = thehdw.DCVS.Pins(DCVS_RET).Voltage.Output
        ElseIf DCVSOutPutSelector = OutputIsVmain Then
            thehdw.DCVS.Pins(DCVS_RET).Voltage.Output = tlDCVSVoltageAlt
            RegetOutput = thehdw.DCVS.Pins(DCVS_RET).Voltage.Output
        End If
    End If
    
    thehdw.Wait 0.001
    
    thehdw.Wait CDbl(WaitTime) / 2
    T0 = theexec.Timer(0)
    If Glb_RetentionMeasurement = True Then
        Pins_Measure_Case ConditionString
    End If
    T1 = theexec.Timer(T0)
    If T1 >= (CDbl(WaitTime) / 2) Then
    Else
        thehdw.Wait (CDbl(WaitTime) / 2) - T1
    End If
    
    If Disable_RET_Ramp Then
        For i = 0 To RestoreMainOrAlt.Pins.Count - 1
            PinName = RestoreMainOrAlt.Pins(i).name
            SetRecoverVol PinName, RegetOutput, RetentionRestorePLD
        Next i
    Else
        'If set_init = True Then
            Retention_Ramping RetentionPin, TargetVal, g_CharPattInfoAry(seq - 1).ForceVoltage, dDCVSOutPutSelector, set_init, RampStep, RampTime, DCVS_RET
        'Else
        '    Retention_Ramping RetentionPin, TargetVal, g_CharPattInfoAry(seq - 1).ForceVoltage, dDCVSOutPutSelector, set_init, RampStep  ' Scott 20230925
        'End If
        For i = 0 To RestoreMainOrAlt.Pins.Count - 1
            PinName = RestoreMainOrAlt.Pins(i).name
            RecoverVmainValtAfterRampUp PinName, RestoreMainOrAlt, set_init
        Next i
    End If
   
    If Disable_RET_Ramp Then
        If RegetOutput = OutputIsValt Then
            thehdw.DCVS.Pins(DCVS_RET).Voltage.Output = tlDCVSVoltageMain
            RegetOutput = thehdw.DCVS.Pins(DCVS_RET).Voltage.Output
        ElseIf DCVSOutPutSelector = OutputIsVmain Then
            thehdw.DCVS.Pins(DCVS_RET).Voltage.Output = tlDCVSVoltageAlt
            RegetOutput = thehdw.DCVS.Pins(DCVS_RET).Voltage.Output
        End If
    End If
    
    For i = 0 To RestoreMainOrAlt.Pins.Count - 1
        PinName = RestoreMainOrAlt.Pins(i).name
        RecoverVmainValtAfterVoltageUpDirectly PinName, RegetOutput, set_init, RestoreMainOrAlt
    Next i
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Shmoo_Restore_Retention_Power")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function PatternTypeChecker(PatStr As String) As String
On Error GoTo errHandler

    Dim i As Long
    Dim tmpPatStrAry() As String
    Dim tmpPatStr() As String
    Dim PatType As String
    
	If LCase(PatStr) Like "*.pat*" Then
        tmpPatStrAry = Split(PatStr, ":")
        tmpPatStrAry(0) = tmpPatStrAry(1)
        ReDim Preserve tmpPatStrAry(0)
    Else
        tmpPatStrAry = Split(PatStr, ",")
    End If
    PatType = vbNullString
    For i = 0 To UBound(tmpPatStrAry)
        tmpPatStr = Split(tmpPatStrAry(i), "_")
        If UBound(tmpPatStr) >= 4 And UCase(tmpPatStr(3)) Like "*PL*" Then
            If UCase(tmpPatStr(4)) = "SC" Or UCase(tmpPatStr(4)) = "CH" Or UCase(tmpPatStr(4)) = "FU" Then '240227
                If PatType = "" Then
                    PatType = "SCAN"
                Else
                    If PatType <> "SCAN" Then theexec.ErrorLogMessage "Please use correct Patterns for TestType define. "
                End If
            ElseIf UCase(tmpPatStr(4)) = "BI" Then 'Or UCase(tmpPatStr(6)) = "BST" Then
                If PatType = "" Then
                    PatType = "BIST"
                Else
                    If PatType <> "BIST" Then theexec.ErrorLogMessage "Please use correct Patterns for TestType define. "
                End If
            Else
                theexec.ErrorLogMessage "Please use correct Patterns for TestType define. "
            End If
        End If
    Next i
    PatternTypeChecker = PatType
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "PatternTypeChecker")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Decide_DC_Level_Mod(DC_Level As PinListData, testType As String)
On Error GoTo errHandler
    
    If UCase(testType) = "BIST" Then
        DC_Level = g_ApplyLevelTimingVmain.Copy
    Else
        DC_Level = g_ApplyLevelTimingValt.Copy
    End If
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Decide_DC_Level_Mod")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function DecodingRealSourceBit_Mod(Source_Bits As String, idx As Long) As String
On Error GoTo errHandler

    Dim BitsDef As String: BitsDef = g_BitsDef
    Dim BitsDefArr() As String
    Dim RailsDecodingInfo As New Dictionary
    Dim DSSCSelSrmOpposite As Long
    Dim BitsValue As String
    Dim DcodingRailInfo() As String
    Dim logicPin As String
    Dim i As Long
 
    BitsDefArr = Split(BitsDef, ",")
    ReDim Preserve BitsDefArr(UBound(BitsDefArr))
    ReDim DcodingRailInfo(UBound(BitsDefArr))

    ReDim BitsDefArr(UBound(SelsramMapping(idx).logic_Pin))
    For i = 0 To UBound(BitsDefArr)
        BitsDefArr(i) = SelsramMapping(idx).logic_Pin(i)
    Next i
    
    ReDim DcodingRailInfo(UBound(BitsDefArr))
    For i = 0 To Len(Source_Bits) - 1
        logicPin = SelsramMapping(idx).logic_Pin(i)
        DSSCSelSrmOpposite = SelsramMapping(idx).SelSrm1(i)
        BitsValue = CStr(mid(Source_Bits, i + 1, 1))
        If Not RailsDecodingInfo.Exists(logicPin) = True Then
            If DSSCSelSrmOpposite = 1 Then
                RailsDecodingInfo.Add (logicPin), BitsValue
            Else
                If BitsValue = "1" Then
                    RailsDecodingInfo.Add (logicPin), 0
                ElseIf BitsValue = "0" Then
                    RailsDecodingInfo.Add (logicPin), 1
                ElseIf UCase(BitsValue) = "S" Then
                    RailsDecodingInfo.Add (logicPin), "S"
                End If
            End If
        End If
    Next i
    
    For i = 0 To UBound(BitsDefArr)
        DcodingRailInfo(i) = RailsDecodingInfo(BitsDefArr(i))
    Next i
 
    DecodingRealSourceBit_Mod = Join(DcodingRailInfo, vbNullString)
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "DecodingRealSourceBit_Mod")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function DecomposePattSet_Mod(Init1 As Pattern, DecompPatAry() As Pattern)
On Error GoTo errHandler

    Dim Pat_init1() As String
    Dim Pats_Num As Long
    Dim PatIdx As Integer
    On Error GoTo errHandler
    
    If Init1 <> "" Then
        thehdw.Patterns(Init1).ValidatePatlist
        Pat_init1 = theexec.DataManager.Raw.GetPatternsInSet(CStr(Init1), Pats_Num)
        ReDim DecompPatAry(Pats_Num) As Pattern
        If UBound(Pat_init1) > 0 Then
           For PatIdx = 0 To Pats_Num - 1
               If UCase(Pat_init1(PatIdx)) Like "*_IN*" And PayloadPatCnt = 0 Then
                   DecompPatAry(PatIdx).value = Pat_init1(PatIdx)
                   InitialPatCnt = InitialPatCnt + 1
               Else
                  DecompPatAry(PatIdx).value = Pat_init1(PatIdx)
                  PayloadPatCnt = PayloadPatCnt + 1
               End If
           Next PatIdx
        End If
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "DecomposePattSet_Mod")
    If AbortTest Then Exit Function Else Resume Next
End Function
    
Public Function Retention_Ramping(Shmoo_Apply_Pin As String, StartVoltage As PinListData, TargetVoltage As PinListData, OutputSel As Long, set_init As Boolean, _
                                    Optional StepCnt As Integer = 9, Optional RampTime As Double = 0.00002, Optional DCVSpin As String)
On Error GoTo errHandler

    Dim p_ary() As String, p_cnt As Long, i As Long, InstName As String
    Dim step As Integer
    Dim RestoreMainOrAlt As New PinListData
    Dim RestoreMain As Boolean
    Dim RestoreAlt As Boolean

    On Error GoTo errHandler

    Const IsVmain = 1
    Const IsValt = 0
    Dim CalValue As Double
    Dim MainAltStatus As Integer
    Dim PinType() As String
    Dim InstrumentType As String
    Dim site As Variant 'Carter, 20240304
    If OutputSel = 2 Then
        MainAltStatus = IsValt
    ElseIf OutputSel = 1 Then
        MainAltStatus = IsVmain
    Else
        MainAltStatus = -1
    End If
    
    theexec.DataManager.DecomposePinList Shmoo_Apply_Pin, p_ary, p_cnt
    
    ReDim PinType(p_cnt)
    For i = 0 To p_cnt - 1
        If gl_GetInstrument_Dic.Exists(LCase(p_ary(i))) Then
            InstrumentType = gl_GetInstrument_Dic.item(LCase(p_ary(i)))
            Select Case InstrumentType
                Case "DC-07", "DC-30"
                    PinType(i) = "DCVI"
                Case Else
                    PinType(i) = "DCVS"
            End Select
        End If
    Next
    
    For step = 1 To StepCnt
        For i = 0 To p_cnt - 1
            If gl_GetInstrument_Dic.Exists(LCase(p_ary(i))) Then
                If PinType(i) = "DCVS" Then
                    For Each site In theexec.sites
                        CalValue = StartVoltage.Pins(p_ary(i)).value - ((StartVoltage.Pins(p_ary(i)).value - TargetVoltage.Pins(p_ary(i)).value) / StepCnt) * step
                        If MainAltStatus = IsValt Then   'Output is Valt now -> ready to switch to Vmain
                            thehdw.DCVS.Pins(p_ary(i)).Voltage.Main.value = CalValue
                        Else
                            thehdw.DCVS.Pins(p_ary(i)).Voltage.Alt.value = CalValue
                        End If
                    Next
                ElseIf PinType(i) = "DCVI" Then
                    For Each site In theexec.sites
                        CalValue = StartVoltage.Pins(p_ary(i)).value - ((StartVoltage.Pins(p_ary(i)).value - TargetVoltage.Pins(p_ary(i)).value) / StepCnt) * step
                        thehdw.DCVI.Pins(p_ary(i)).Voltage = CalValue
                    Next
                End If
            End If
        Next i
        If MainAltStatus = IsValt Then
            thehdw.DCVS.Pins(DCVSpin).Voltage.Output = tlDCVSVoltageMain
        ElseIf MainAltStatus = IsVmain Then
            thehdw.DCVS.Pins(DCVSpin).Voltage.Output = tlDCVSVoltageAlt
        End If
        thehdw.Wait RampTime
        If MainAltStatus <> -1 Then MainAltStatus = (MainAltStatus + 1) Mod 2
    Next step
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Retention_Ramping")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function PatternIndexChk(PatHeader As Variant) As Long

    If UCase(PatHeader) Like "INIT#" Or UCase(PatHeader) Like "INIT##" Then
        PatternIndexChk = CLng(Replace(UCase(PatHeader), "INIT", ""))
    ElseIf UCase(PatHeader) Like "PL#" Or UCase(PatHeader) Like "PL##" Then
        PatternIndexChk = CLng(Replace(UCase(PatHeader), "PL", "")) + InitialPatCnt
    Else
        theexec.Datalog.WriteComment "The argument of pattern index header name isn't INIT or PL. Please check it."
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "PatternIndexChk")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Process_Ret_Setting(Wait As String, Ret_Ramp_Setting As String)
    
    '-----------------------------------------------------------------------
    'wait - ex - INIT1:0.1:+0.25
    'tmp3(0) = PatIndexHeader, tmp3(1) = WaitTime, tmp3(2) = OffsetVoltage
    '
    'ret_ramp_setting - ex - INIT1:9:0.001
    'tmp3(0) = PatIndexHeader, tmp3(1) = RampStep, tmp3(2) = RamppingTime
    '-----------------------------------------------------------------------

    Dim tmp1() As String
    Dim tmp2() As String
    Dim tmp3() As String
    Dim i, j As Integer
    Dim symbol As String
    Dim patindex As Long
    
    Const PatHeader = 0
    Const WaitTimeIdx = 1
    Const OffsetIdx = 2
    Const StepIdx = 1
    Const RampTimeIdx = 2
    
    tmp1 = Split(Wait, ",")
    tmp2 = Split(Ret_Ramp_Setting, ",")
    
    For i = 0 To UBound(tmp1)
        'wait time
        tmp3 = Split(tmp1(i), ":")
        patindex = PatternIndexChk(tmp3(PatHeader))
        With g_CharPattInfoAry(patindex - 1)
            If tmp3(WaitTimeIdx) <> "" Then .WaitTime = CDbl(tmp3(WaitTimeIdx))
            'ret offset
            If tmp3(OffsetIdx) <> "" Then
                symbol = mid(tmp3(OffsetIdx), 1, 1)
                Select Case symbol
                    Case "+"
                        .RampOffsetSymbol = "+"
                    Case "-"
                        .RampOffsetSymbol = "-"
                    Case Else
                        theexec.Datalog.WriteComment "The offset of retention needs to add the symbol!"
                End Select
                .RampOffset = CDbl(mid(tmp3(OffsetIdx), 2))
            End If
        End With
    Next
    
    'ret_ramp_setting
    For i = 0 To UBound(tmp2)
        tmp3 = Split(tmp2(i), ":")
        patindex = PatternIndexChk(tmp3(PatHeader))
        With g_CharPattInfoAry(patindex - 1)
            If .WaitTime <> 0 Then
                If tmp3(StepIdx) <> "" Then .RampStep = CInt(tmp3(StepIdx))
                If tmp3(RampTimeIdx) <> "" Then .RamppingTime = CDbl(tmp3(RampTimeIdx))
            End If
        End With
    Next

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Process_Ret_Setting")
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Process_DigSrc_Setting(BitSize As String, EQ As String, PinName As String, Seg As String)
'Public Function Process_DigSrc_Setting(BitSize As String, Seg As String, PinName As String, EQ As String)

    '------------------------------------------------------------------------------------------
    'DigSrc_BitSize - ex - INIT1:size20,PL1:size5
    '
    'DigSrc_Seg - ex - INIT1:sgm0_10=0000011101+sgm1_5=11111+sgm2_5=00010,PL1:sgm0_5=10101
    '           - ex - PL2=sgmt0_8+sgmt1_4+SELSRAM+sgmt2_2
    '
    'DigSrc_Pin - ex - "INIT1:JTAG_TDI,PL1:JTAD_TDI2"
    '
    'DigSrc_Eq - ex - "INIT1=smg0+sgm2,PL1=smg0"
    '          - ex 0 PL2:sgmt0=0B00001110+sgmt1=0B1111+sgmt2=0B01
    '------------------------------------------------------------------------------------------

    Dim tmp_size() As String
    Dim tmp_seg() As String
    Dim tmp_name() As String
    Dim tmp_eq() As String
    Dim tmp_info() As String
    Dim tmp_info2() As String
    Dim eq_info() As String
    Dim seg_info() As String
    Dim i, j As Integer
    Dim patindex As Long
    Dim PatIndex2 As Long
    Dim SegString As String
    Dim Dict_tmp() As New Dictionary
    ReDim Dict_tmp(InitialPatCnt + PayloadPatCnt - 1)
    
    Dim EQString As String
    
    
    Const PatHeader = 0
    
    tmp_size = Split(BitSize, ",")
    tmp_seg = Split(Seg, ",")
    tmp_name = Split(PinName, ",")
    tmp_eq = Split(EQ, ",")
    
    If UBound(tmp_eq) <> UBound(tmp_size) Then _
    theexec.Datalog.WriteComment "<Warning>  DigSrc_Eq and DigSrc_Size amount didn't match. Please check it!"
    
    For i = 0 To UBound(tmp_size)
        'Digsrc_size
        tmp_info = Split(tmp_size(i), ":")
        patindex = PatternIndexChk(tmp_info(PatHeader))
        g_CharPattInfoAry(patindex - 1).DigSrc_BitSize = tmp_info(1)
        
        'Digsrc_EQ
        tmp_info = Split(tmp_eq(i), ":")
        PatIndex2 = PatternIndexChk(tmp_info(PatHeader))
        If PatIndex2 = patindex Then
            EQString = ""
            tmp_info2 = Split(tmp_info(1), "+")
            For j = 0 To UBound(tmp_info2)
                eq_info = Split(tmp_info2(j), "_")
                EQString = EQString & IIf(EQString = "", "", "+") & eq_info(0)
                If UBound(eq_info) = 1 Then Dict_tmp(PatIndex2 - 1).Add eq_info(0), eq_info(1)
            Next
            g_CharPattInfoAry(patindex - 1).digSrc_EQ = EQString
        End If

''        'Digsrc_seg
''        tmp_info = Split(tmp_seg(i), ":")
''        PatIndex2 = PatternIndexChk(tmp_info(PatHeader))
''        If PatIndex2 = patindex Then
''            SegString = ""
''            tmp_info2 = Split(tmp_info(1), "+")
''            For j = 0 To UBound(tmp_info2)
''                seg_info = Split(tmp_info2(j), "_")
''                SegString = SegString & IIf(SegString = "", "", "+") & seg_info(0)
''                If UBound(seg_info) = 1 Then Dict_tmp(PatIndex2 - 1).Add seg_info(0), seg_info(1)
''            Next
''            g_CharPattInfoAry(patindex - 1).DigSrc_Seg = SegString
''        End If
    Next
    
    'Digsrc_seg
    If UBound(tmp_seg) <> -1 Then
        For i = 0 To UBound(tmp_seg)
            tmp_info = Split(tmp_seg(i), ":")
            patindex = PatternIndexChk(tmp_info(PatHeader))
            tmp_info2 = Split(tmp_info(1), ";")
            For j = 0 To UBound(tmp_info2)
                seg_info = Split(tmp_info2(j), "=")
                g_CharPattInfoAry(patindex - 1).SegDict.Add seg_info(0), Array(seg_info(1), CLng(Dict_tmp(patindex - 1).item(seg_info(0))))
                'g_CharPattInfoAry(patindex - 1).EqDict.Add eq_info(0), Array(eq_info(1), CLng(Dict_tmp(patindex - 1).item(eq_info(0))))
            Next j
        Next i
    End If
    
''    'digSrc_EQ
''    If UBound(tmp_eq) <> -1 Then
''        For i = 0 To UBound(tmp_eq)
''            tmp_info = Split(tmp_eq(i), ":")
''            patindex = PatternIndexChk(tmp_info(PatHeader))
''            tmp_info2 = Split(tmp_info(1), "+")
''            For j = 0 To UBound(tmp_info2)
''                eq_info = Split(tmp_info2(j), "=")
''                g_CharPattInfoAry(patindex - 1).EqDict.Add eq_info(0), Array(eq_info(1), CLng(Dict_tmp(patindex - 1).item(eq_info(0))))
''            Next
''        Next
''    End If
''
    'Digsrc_pin
    For i = 0 To UBound(tmp_name)
        tmp_info = Split(tmp_name(i), ":")
        patindex = PatternIndexChk(tmp_info(PatHeader))
        g_CharPattInfoAry(patindex - 1).DigSrc_pin = tmp_info(1)
    Next

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "Process_DigSrc_Setting")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Shmoo_Save_core_power_per_site_for_Vbump_Site_Loop(Optional PinGPName As String)
    Dim p_ary() As String, p_cnt As Long, i As Long, InstName As String
    On Error GoTo errHandler
    Dim site As Variant 'Carter, 20240304
    Set g_ApplyLevelTimingVmain = Nothing
    Set g_ApplyLevelTimingValt = Nothing

    If PinGPName <> "" Then
        theexec.DataManager.DecomposePinList PinGPName, p_ary, p_cnt
    Else
        theexec.DataManager.DecomposePinList "All_power", p_ary, p_cnt
    End If

    For i = 0 To p_cnt - 1
        If theexec.DataManager.ChannelType(p_ary(i)) <> "N/C" Then
            g_ApplyLevelTimingVmain.AddPin UCase((p_ary(i)))
            g_ApplyLevelTimingValt.AddPin UCase((p_ary(i)))
            InstName = GetInstrument(p_ary(i), 0)
            For Each site In theexec.sites
                Select Case InstName
                    Case "DC-07"
                        g_ApplyLevelTimingVmain.Pins(UCase(p_ary(i))).value = CDbl(Format(thehdw.DCVI.Pins(p_ary(i)).Voltage, "0.000"))
                        g_ApplyLevelTimingValt.Pins(UCase(p_ary(i))).value = CDbl(Format(thehdw.DCVI.Pins(p_ary(i)).Voltage, "0.000"))
                    Case "VHDVS"
                        g_ApplyLevelTimingVmain.Pins(UCase(p_ary(i))).value = CDbl(Format(thehdw.DCVS.Pins(p_ary(i)).Voltage.Main.value, "0.000"))
                        g_ApplyLevelTimingValt.Pins(UCase(p_ary(i))).value = CDbl(Format(thehdw.DCVS.Pins(p_ary(i)).Voltage.Alt.value, "0.000"))
                    Case "HexVS"
                        g_ApplyLevelTimingVmain.Pins(UCase(p_ary(i))).value = CDbl(Format(thehdw.DCVS.Pins(p_ary(i)).Voltage.Main.value, "0.000"))
                        g_ApplyLevelTimingValt.Pins(UCase(p_ary(i))).value = CDbl(Format(thehdw.DCVS.Pins(p_ary(i)).Voltage.Alt.value, "0.000"))
                    Case "VS-5A"
                        g_ApplyLevelTimingVmain.Pins(UCase(p_ary(i))).value = CDbl(Format(thehdw.DCVS.Pins(p_ary(i)).Voltage.Main.value, "0.000"))
                        g_ApplyLevelTimingValt.Pins(UCase(p_ary(i))).value = CDbl(Format(thehdw.DCVS.Pins(p_ary(i)).Voltage.Alt.value, "0.000"))
                    Case "VS-800mA"
                        g_ApplyLevelTimingVmain.Pins(UCase(p_ary(i))).value = CDbl(Format(thehdw.DCVS.Pins(p_ary(i)).Voltage.Main.value, "0.000"))
                        g_ApplyLevelTimingValt.Pins(UCase(p_ary(i))).value = CDbl(Format(thehdw.DCVS.Pins(p_ary(i)).Voltage.Alt.value, "0.000"))
                    Case "HSD-U"
                    Case Else
                          theexec.ErrorLogMessage "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Save_core_power_per_site"
                 End Select
             Next site
        End If
    Next i
    Exit Function

errHandler:
    theexec.Datalog.WriteComment "<Error> Shmoo_Save_core_power_per_site_for_Vbump_Site_Loop:: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function CheckPatLabel(Pat As String, DigSrcPin As String) As String
On Error GoTo errHandler

    Dim sSrcSigName As String
    Dim tempVarArray As Variant
    tempVarArray = thehdw.DSSC.Pins(DigSrcPin).Pattern(Pat).Source.Labels.list
    sSrcSigName = tempVarArray(0)
    If sSrcSigName = "" Then
        sSrcSigName = "FUNC_SRC"
    End If
    
    CheckPatLabel = sSrcSigName

    
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error>Shmoo_Save_core_power_per_site_for_Vbump_Site_Loop:: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Sub CharSetUpCheck()
On Error GoTo errHandler

    ReviseSetup = False

    If TheExec.Flow.enableWord("CZsetup_Chk") = False And Enable_CZsetup_Chk = True Then
        
        If Enable_CZsetup_Chk = True Then 
			CheckSetupRange
        else
		end if
        If TheExec.RunMode = runModeProduction Then
            Enable_CZsetup_Chk = False
        Else
            Enable_CZsetup_Chk = True
        End If
	else
    End If
    
    
    
    
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "CharSetUpCheck")
    If AbortTest Then Exit Sub Else Resume Next
End Sub
Public Sub CheckSetupRange()
On Error GoTo errHandler
    
    'Dim SheetName() As String
    'Dim tmpSheet As Variant
    Dim tmpSetup As Variant
    Dim SetupName() As String
    Dim tmpAxis As Variant
    Dim ParameterType As String
    Dim site As Variant

    'SheetName = theexec.Job.GetSheetNamesOfType(DMGR_SHEET_TYPE_CHARACTERIZATIONSHEET)
    With theexec.DevChar
        SetupName = .Setups.list
        For Each tmpSetup In SetupName
            For Each tmpAxis In .Setups(tmpSetup).Shmoo.Axes.list
                With .Setups(tmpSetup).Shmoo.Axes(tmpAxis).Parameter
                    ParameterType = .type
                    Select Case UCase(ParameterType)
                        Case "GLOBAL SPEC", " LEVEL":
                            VolHigher1p3 .range.from, .range.to, CStr(tmpSetup)
                            FromLowToHigh .range.from, .range.to, CStr(tmpSetup)
                        Case Else
                    End Select
                End With
            Next tmpAxis
        Next tmpSetup
    End With

    If ReviseSetup Then
		TheExec.Flow.TestLimit resultVal:=1, lowVal:=0, hiVal:=0, ForceResults:=tlForceNone, Tname:="CZ setup check"
	else
	end if
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "CheckSetupRange")
    If AbortTest Then Exit Sub Else Resume Next
End Sub
Public Sub VolHigher1p3(FromValue As Double, ToValue As Double, SetupName As String)

    If FromValue >= 1.3 Or ToValue >= 1.3 Then
        TheExec.Datalog.WriteComment "<Warning> char_set_up : " & SetupName & ": The setup of range is higher then 1.3V. Please check!"
        ReviseSetup = True
	else
    End If
    
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "VolHigher1p3")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub FromLowToHigh(FromValue As Double, ToValue As Double, SetupName As String)

    If FromValue < ToValue Then
        TheExec.Datalog.WriteComment "<Warning> char_set_up : " & SetupName & " : The from value of setup is lower then the to value . Please revise!"
        ReviseSetup = True
	else
    End If
    
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Shmoo", "FromValue < ToValue")
    If AbortTest Then Exit Sub Else Resume Next
End Sub
