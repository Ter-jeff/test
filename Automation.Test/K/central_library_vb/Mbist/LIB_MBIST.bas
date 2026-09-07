Attribute VB_Name = "LIB_MBIST"
Option Explicit
Public gS_currPayload_pattSetName As String
Public MatchFlag As Boolean
''''-----------------------------------------------------------------
Public gB_findCpuMbist_flag As Boolean
Public gB_findGpuMbist_flag As Boolean
Public gB_findSocMbist_flag As Boolean
Public gB_findPwrPin_flag As Boolean
Public gS_SocMbist_sheetName As String
Public gS_CpuMbist_sheetName As String
Public gS_GpuMbist_sheetName As String
Public gl_burst_pat As New Dictionary
Public Flag_BurstPat_INIT As Boolean ''carter 20191118

'=======================20160301=======================================

Public Function MbistRampApplyLevel_AutoReadingContext(Optional ByVal ApplyPins As String = "CorePower", Optional RampingStep As Double = 10, Optional RampWaitTime As Double = 0, Optional instancename As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    ''SWLINZA20171120, for ramping voltage for each ATPG and Mbist instance

    Dim Apply_Pins_Ary() As String
    Dim Apply_Pins_count As Long
    Dim Extra_RampingTime As Double: Extra_RampingTime = RampWaitTime 'RampDown_Time = 0
    'Dim RampingStep As Double
    Dim Original_voltage() As Double
    Dim Apply_TargetVoltage() As Double
    Dim DiffVoltage() As Double
    Dim RampingVoltage() As Double
    Dim Voltage_from_HW As String
    Dim i, j As Integer
    Dim Current_DCCategory As String
    Dim Current_DCSelector As String
    Dim TestBlock As String
    Dim SepcSymbolic As String
    Dim ApplyPins_String As String
    Dim ApplyPins_Boolean() As Boolean
    'Dim AllPins_needApply As Boolean
    Dim Dummy_tempStr As String
    
    If theexec.enableWord("Ramping_MbistATPG") = False Then Exit Function
    
    theexec.DataManager.DecomposePinList ApplyPins, Apply_Pins_Ary(), Apply_Pins_count
    ReDim Original_voltage(Apply_Pins_count - 1) As Double
    ReDim DiffVoltage(Apply_Pins_count - 1) As Double
    ReDim RampingVoltage(Apply_Pins_count - 1) As Double
    ReDim Apply_TargetVoltage(Apply_Pins_count - 1) As Double
    ReDim ApplyPins_Boolean(Apply_Pins_count - 1) As Boolean
    
    '----- to get target voltage from DC spec for each instance -----
    'Apply_TargetVoltage
    
    'Swlinza 20180126, to save test time in IGXL9.0, use this command instead of following two
    theexec.DataManager.GetInstanceContext Current_DCCategory, Current_DCSelector, Dummy_tempStr, Dummy_tempStr, Dummy_tempStr, Dummy_tempStr, Dummy_tempStr, Dummy_tempStr
    'Current_DCCategory = TheExec.TestInstances.Item(InstanceName).TimingAndLevels.DCCategory
    'Current_DCSelector = TheExec.TestInstances.Item(InstanceName).TimingAndLevels.DCSelector
    
    If Current_DCCategory = Previous_DCCategory And Current_DCSelector = Previous_DCSelector Then
        Exit Function
    Else
        Previous_DCCategory = Current_DCCategory
        Previous_DCSelector = Current_DCSelector
    End If
    
    TestBlock = mid(Current_DCCategory, 1, 3)
    
    Select Case UCase(TestBlock)
        Case UCase("Soc")
            SepcSymbolic = "_VAR_S"
        Case UCase("Cpu")
            SepcSymbolic = "_VAR_C"
        Case UCase("Gfx")
            SepcSymbolic = "_VAR_G"
        Case UCase("RTO")
            SepcSymbolic = "_VAR_R"
        Case Else
            SepcSymbolic = "_VAR_H"
    End Select
    
    '------ to calculate ramping voltage for each pins ------
    'AllPins_needApply = False
    For i = 0 To Apply_Pins_count - 1
        Original_voltage(i) = FormatNumber(TheHdw.DCVS.Pins(Apply_Pins_Ary(i)).Voltage.Main, 3)
        Apply_TargetVoltage(i) = theexec.Specs.DC.item(Apply_Pins_Ary(i) & SepcSymbolic).Categories.item(Current_DCCategory).Selectors.item(Current_DCSelector).ContextValue
        DiffVoltage(i) = Original_voltage(i) - Apply_TargetVoltage(i)
        RampingVoltage(i) = FormatNumber((DiffVoltage(i) / RampingStep), 3)
        If Apply_TargetVoltage(i) = Original_voltage(i) Or Abs(DiffVoltage(i)) < 0.001 * RampingStep Then
            ApplyPins_Boolean(i) = False
        Else
            ApplyPins_Boolean(i) = True
            'AllPins_needApply = True
        End If
    Next i
    'If AllPins_needApply = False Then Exit Function
    '--------- Ramp down for retention voltage ------'
    For i = 0 To RampingStep - 1
        For j = 0 To Apply_Pins_count - 1
            If ApplyPins_Boolean(j) = True Then
                If i = RampingStep - 1 Then
                    TheHdw.DCVS.Pins(Apply_Pins_Ary(j)).Voltage.Main = Apply_TargetVoltage(j)
                Else
                    TheHdw.DCVS.Pins(Apply_Pins_Ary(j)).Voltage.Main = Original_voltage(j) - RampingVoltage(j) * i
                End If
            End If
        Next j
        TheHdw.Wait Extra_RampingTime / RampingStep
    Next i
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_MBIST", "MbistRampApplyLevel_AutoReadingContext") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' [20230425][All][Carter] Add Finger print new syntex
Public Function Finger_print(pattern_load As String, RunFailCycle As Boolean, Optional Flag_Name As String, Optional mbist_loop As Boolean = False)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
 
    If glb_TesterType = "UltraFLEXplus" Then
        Call Finger_Print_NewSyntex(pattern_load, RunFailCycle, Flag_Name, mbist_loop)      'PLUS as version 10.30.90 use function
    Else
        Call Finger_print_Org(pattern_load, RunFailCycle, Flag_Name, mbist_loop)
    End If
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_MBIST", "Finger_print") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function ReduceBlkLen(InpStr As String, OutpStr As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    'Dim InputStr As String: InputStr = "Proc48_Mem121_Mem120_Mem123"   '_Mem119_Mem118"
    Dim SplitStr() As String
    Dim BlkFirstStr As String
    Dim BlkSecondStr As String
    Dim ReplaceStr As String
    Dim ReduceStr As String
    Dim BlkIdx As Long
    Dim DupIdx As Long
    Dim DupLen As Long
    Dim DupStr As String
    Dim IdxStr As String, i As Integer
    
    SplitStr = Split(InpStr, "_")
    If UBound(SplitStr) > 1 Then
        BlkFirstStr = SplitStr(0)
        BlkSecondStr = SplitStr(1)
        ReplaceStr = right(InpStr, Len(InpStr) - Len(BlkFirstStr) - Len(BlkSecondStr) - 2)
            For DupIdx = 0 To Len(BlkSecondStr) - 1
            IdxStr = mid(BlkSecondStr, DupIdx + 1, 1)
                If IsNumeric(IdxStr) Then
                    DupLen = DupIdx
                    DupStr = left(BlkSecondStr, DupLen)
                    Exit For
                End If
        
            Next
        ReduceStr = Replace(ReplaceStr, DupStr, "")
        OutpStr = BlkFirstStr & "_" & BlkSecondStr & "_" & ReduceStr
    Else
        OutpStr = InpStr
    End If
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_MBIST", "ReduceBlkLen") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' [20230907][All][Oliver] modify for only DCVS pin use tlDCVSVoltageMain
Public Function ATPG_offline(pattern_load As String, ResultMode As tlResultMode)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    ''Carter, 20191120
    
    Dim Pins As Variant
    Dim patt As Variant
    
    Dim PinName() As String
    Dim m_testName As String
    Dim rtnPatternNames() As String
    
    Dim offline_patallpass As Boolean
    Dim offline_pat_status As New SiteBoolean
    
    Dim NumberPins As Long
    Dim rtnPatternCount As Long
    
    Dim Core_Vmain As Double

    m_testName = theexec.DataManager.instancename
    offline_patallpass = True
    offline_pat_status = False
    TheHdw.DCVS.Pins(Core_Power_DCVS_pins).Voltage.Output = tlDCVSVoltageMain
    Call theexec.DataManager.DecomposePinList("CorePower", PinName(), NumberPins)
    If InStr(pattern_load, "\") = 0 Then ''Burst Pattern, exclude walkingZ pattern
        Call PATT_GetPatListFromPatternSet(pattern_load, rtnPatternNames, rtnPatternCount)
        For Each patt In rtnPatternNames
            Call ATPG_offline_Simulation(patt, ResultMode, offline_patallpass, offline_pat_status)
        Next patt
    Else ''Single Pattern
        Call ATPG_offline_Simulation(pattern_load, ResultMode, offline_patallpass, offline_pat_status)
    End If
    
Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    If isDebugMode Then theexec.AddOutput "Error in the VBT ATPG_offline"
    Call Print_Error_Message(Error_Info, "LIB_MBIST", "ATPG_offline") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function ATPG_offline_Simulation(pattern_load As Variant, ResultMode As tlResultMode, offline_patallpass As Boolean, offline_pat_status As SiteBoolean)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim site As Variant 'Carter, 20240304
    If LCase(pattern_load) Like "*_in*" Then
        Call TheHdw.Patterns(pattern_load).test(pfAlways, 0, ResultMode)
    Else
        offline_patallpass = True
              
        If theexec.enableWord("Golden_Default") = False Then
            For Each site In theexec.sites
                If offline_pat_status(site) = False Then
                    offline_pat_status(site) = IIf(Round(WorksheetFunction.min(1, Rnd * 30), 0) = 1, True, False)
                End If
            Next site
        Else
            offline_pat_status = True
        End If
        
        For Each site In theexec.sites
            offline_patallpass = offline_patallpass And offline_pat_status(site)
        Next site
        
        If offline_patallpass = True Then
            Call TheHdw.Patterns(pattern_load).test(pfAlways, 0, ResultMode)
    
        Else
            Call TheHdw.Patterns(pattern_load).test(pfNever, 0, ResultMode)
            For Each site In theexec.sites
                If offline_pat_status(site) = False Then
                    Call theexec.Datalog.WriteFunctionalResult(site, theexec.sites.item(site).TestNumber, logTestFail)
    
                Else
                    Call theexec.Datalog.WriteFunctionalResult(site, theexec.sites.item(site).TestNumber, logTestPass)
    
                End If
            Next
        End If
    End If
    
Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    If isDebugMode Then theexec.AddOutput "Error in the VBT ATPG_offline_pat"
    Call Print_Error_Message(Error_Info, "LIB_MBIST", "ATPG_offline_Simulation") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Sub Sub_SetVol_toAllPins(PinAry() As String, VolAry() As Double, RampStep As Double)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim i As Long
    Dim PinName As String
    Dim DropVoltage As Double
    Dim State_index As Integer
    
    State_index = MbistERT_GroupCurrentVol.Count
    Sub_GroupRamp PinAry, VolAry, State_index
    
    If UBound(VolAry) = 0 Then
        ReDim Preserve VolAry(UBound(PinAry())) As Double
        For i = 1 To UBound(VolAry)
            VolAry(i) = VolAry(0)
        Next i
    ElseIf UBound(VolAry) = UBound(PinAry) Then
        ' same count then do nothing
    Else
        theexec.ErrorLogMessage "MBist Retention Vol Setting count is not match with Pin Count"
    End If
    
    For i = State_index To MbistERT_GroupCurrentVol.Count - 1
        PinName = Split(MbistERT_GroupPinNmame(Group_Name & i), ",")(0)
        MbistERT_OriginVol.Add Group_Name & i, lclPinListData_PreviousCorePower.Pins(PinName).value
        MbistERT_TargetVol.Add Group_Name & i, MbistERT_GroupCurrentVol(Group_Name & i)
        
        DropVoltage = MbistERT_OriginVol(Group_Name & i) - MbistERT_TargetVol(Group_Name & i)
        MbistERT_DropVol.Add Group_Name & i, DropVoltage
        MbistERT_DropVol_PerSite.Add Group_Name & i, FormatNumber((DropVoltage / RampStep), 3)
    Next i
       
Exit Sub 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_MBIST", "Sub_SetVol_toAllPins") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/08/18
End Sub


Sub Sub_GroupRamp(PinAry() As String, VolAry() As Double, State_index As Integer)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim Pin_index As Integer
    Dim Value_index As Integer
    Dim Group_index As Integer
    Dim PinName As String
    Dim F_diff As Boolean
    
    F_diff = False
    Group_index = State_index
    PinName = vbNullString
    
    MbistERT_GroupCurrentVol.Add Group_Name & Group_index, VolAry(0)
    MbistERT_GroupPreviousVol.Add Group_Name & Group_index, lclPinListData_PreviousCorePower.Pins(PinAry(0)).value
    Group_index = Group_index + 1
    
    For Pin_index = 1 To UBound(PinAry)
        For Value_index = State_index To Group_index - 1
            If MbistERT_GroupCurrentVol(Group_Name & Value_index) = VolAry(Pin_index) And MbistERT_GroupPreviousVol(Group_Name & Value_index) = lclPinListData_PreviousCorePower.Pins(PinAry(Pin_index)).value Then
                F_diff = True
                Exit For
            End If
        Next Value_index
        If Not F_diff Then
            MbistERT_GroupCurrentVol.Add Group_Name & Group_index, VolAry(Pin_index)
            MbistERT_GroupPreviousVol.Add Group_Name & Group_index, lclPinListData_PreviousCorePower.Pins(PinAry(Pin_index)).value
            Group_index = Group_index + 1
        End If
        F_diff = False
    Next Pin_index
    
    For Value_index = State_index To MbistERT_GroupCurrentVol.Count - 1
        For Pin_index = 0 To UBound(PinAry)
            If MbistERT_GroupCurrentVol(Group_Name & Value_index) = VolAry(Pin_index) And MbistERT_GroupPreviousVol(Group_Name & Value_index) = lclPinListData_PreviousCorePower.Pins(PinAry(Pin_index)).value Then
                PinName = IIf(PinName = "", PinAry(Pin_index), PinName & "," & PinAry(Pin_index))
            End If
        Next Pin_index
        MbistERT_GroupPinNmame.Add Group_Name & Value_index, PinName
        PinName = vbNullString
    Next Value_index
Exit Sub 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_MBIST", "Sub_GroupRamp") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/08/18
End Sub


Sub Sub_VoltageRamping(RampStep As Double, ExtraWaitTiom_forRamp As Double, RampDir As RET_RampingDir)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim Ramp_index, Group_index As Integer
    Dim RampingDir As Double
    Dim Vol_from As Double
    Dim Vol_to As Double
    Dim Start_index, End_index As Integer
    
    Select Case RampDir
        Case RampDown:
            RampingDir = 1
            Start_index = MbistERT_GroupCurrentVol.Count - 1
            End_index = 0
        Case RampUp:
            RampingDir = -1
            Start_index = 0
            End_index = MbistERT_GroupCurrentVol.Count - 1
        Case Else:
            Call Print_Error_Message(Error_Warning_Info.Warning_Info, "LIB_MBIST", "Sub_VoltageRamping", "Give wrong ramp enum !!") 'Add ErrHandler 2023/08/18
    End Select
    
    For Ramp_index = 1 To RampStep
        For Group_index = Start_index To End_index Step (RampingDir * -1)
'            If Ramp_index = RampStep - 1 Then
'                Vol_from = IIf(RampDir = RampDown, MbistERT_TargetVol(Group_Name & Group_index), MbistERT_OriginVol(Group_Name & Group_index))
'                TheHdw.DCVS.Pins(MbistERT_GroupPinNmame(Group_Name & Group_index)).Voltage.value = Vol_from
'            Else
            Vol_to = IIf(RampDir = RampDown, MbistERT_OriginVol(Group_Name & Group_index), MbistERT_TargetVol(Group_Name & Group_index))
            TheHdw.DCVS.Pins(MbistERT_GroupPinNmame(Group_Name & Group_index)).Voltage.value = Vol_to - MbistERT_DropVol_PerSite(Group_Name & Group_index) * Ramp_index * RampingDir
'            End If
        Next Group_index
        TheHdw.Wait ExtraWaitTiom_forRamp / RampStep
    Next Ramp_index
    
Exit Sub 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_MBIST", "Sub_VoltageRamping") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/08/18
End Sub


Sub Sub_CStrToDblAry(str As String, Ary2() As Double, Delimiter As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    
    Dim i As Long
    Dim StrAry() As String
    StrAry() = Split(str, Delimiter)
    ReDim Ary2(UBound(StrAry)) As Double
    For i = 0 To UBound(StrAry)
        Ary2(i) = CDbl(StrAry(i))
    Next i

Exit Sub 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_MBIST", "Sub_CStrToDblAry") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/08/18
End Sub

'[20231106][T-ALL][Oliver] add for UltraFlexPlus will not change alarmfail flag state
Public Function Finger_print_Org(pattern_load As String, RunFailCycle As Boolean, Optional Flag_Name As String, Optional mbist_loop As Boolean = False)
 
On Error GoTo errHandler
    Dim maxDepth As Integer
    Dim patt As Variant
    Dim site As Variant
 
    Dim rtnPatternNames() As String, rtnPatternCount As Long
    Dim astrPattPathSplit() As String
    Dim astrPattPathSplit_01() As String
    Dim blPatPass As New SiteBoolean
    Dim numcap As Long
    Dim PinData_d As New PinListData
    Dim Mbist_repair_cycle As Long
    Dim Pins As New PinData
    Dim Cdata As Variant
    Dim TestNumber As New SiteLong
    Dim ins_new_name As String
    Dim tested As New SiteBoolean
    Dim strPattName As String
    Dim inst_match As Boolean
    Dim temp As Long
    Dim AllPins As String
    Dim PinData As New PinListData

    Dim LogLen As Long
    Dim LogLimited As Long
    Dim PrintTimes As Integer
    Dim PrintIdx As Integer
    Dim DecomposeLog() As String
    Dim ReviseStr As String

    Dim blMbistFP_Binout As Boolean
    Dim MBISTFailBlockFlag As Boolean
    Dim PassOrFail As New SiteLong
    Dim lGetFlagIdx As Long
    Dim blJump As Boolean
    Dim m_testName As String
    Dim k As Long, p As Long, g As Long, j As Long, i As Long:: k = 0:: p = 0:: g = 0:: j = 0:: i = 0

    Dim m_tn As Long
    Dim m_tn_restore As Long
    Dim m_tn_BurstIndex As Long
    Dim Mbist_repair_vector As Long
    
    '----------------------------------------------------------------
    ' SWLINZA 20181128, for MemFP DTR, C651/Si requests 2018/08
    '----------------------------------------------------------------
    Dim Pattern_Desc As String
    Dim Pattern_Server As String
    Dim CurCount_FailAry_Element As New SiteLong
    Dim NewFmt_Printing_Header As String
    Dim CharacterNumbers As Long
    Dim Pattern_GenericName() As String
    Dim MFP_pattern_idx As Long
    Dim MFP_flow_idx As Variant
    Dim Pattern_Failure_Cycles() As Variant
 
    m_tn_BurstIndex = 2
    AllPins = "JTAG_TDO"
    LogLimited = 255
    m_testName = theexec.DataManager.instancename
    Call PATT_GetPatListFromPatternSet(pattern_load, rtnPatternNames, rtnPatternCount)
        
    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
    maxDepth = gl_HRAMmaxDepth            'maxDepth = TheHdw.Digital.HRAM.maxDepth    'Flex UP1600 max depth 512    'Plus org max depth 16k, but use 512 to let speed faster
    TheHdw.Digital.hram.size = maxDepth
    TheHdw.Digital.hram.CaptureType = captFail
    
    TheHdw.Digital.hram.SetTrigger trigFail, False, 0, True
    TheHdw.Digital.Patgen.ClearFail
        
    For Each patt In rtnPatternNames
        numcap = 0
        '==================================================================='''''''''''finger print_block_01_begin
        If theexec.enableWord("Mbist_FingerPrint") = True Then
            astrPattPathSplit = Split(CStr(patt), "\")
            strPattName = UCase(astrPattPathSplit(UBound(astrPattPathSplit)))
'        If strPattName Like "*:*" Then
'            astrPattPathSplit_01 = Split(strPattName, ":")
'            strPattName = astrPattPathSplit_01(0)
'        End If
            If strPattName Like "*.GZ" Then strPattName = Replace(strPattName, ".GZ", vbNullString)
            
            '---------------------------------------------
            'SWLINZA 20181128 for MFP DTR, to split patset
            '---------------------------------------------
            Pattern_GenericName() = Split(strPattName, ":")
            Pattern_GenericName(0) = UCase(Pattern_GenericName(0))
            
            If glb_TesterType = "Jaguar" Then
                If Pattern_GenericName(0) Like "*.PAT" Then Pattern_GenericName(0) = Replace(Pattern_GenericName(0), ".PAT", vbNullString)
            Else
                If Pattern_GenericName(0) Like "*.PATX" Then Pattern_GenericName(0) = Replace(Pattern_GenericName(0), ".PATX", vbNullString)
            End If
            
            MatchFlag = False
            For temp = 0 To UBound(tpCycleBlockInfor)
                If UCase(tpCycleBlockInfor(temp).strPattName) = strPattName Then
                    tpEvaPattCycleBlockInfor = tpCycleBlockInfor(temp).tpMbistCycleBlock
                    MatchFlag = True
                    MFP_pattern_idx = temp 'SWLINZa 20180907 for MFP DTR, to indentify pattern#
                    Exit For
                End If
            Next temp
        End If
        '==================================================================='''''''''''finger print_block_01_end
        Call TheHdw.Patterns(patt).test(pfAlways, 0, tlResultModeDomain)
        '//////////////////////////////////////////////////////////////////////////////////////////////////'''''''''''finger print_block_03_begin
        If RunFailCycle = True And theexec.enableWord("Mbist_FingerPrint") = True Then

            If MatchFlag = False Then
                theexec.Datalog.WriteComment ("Warning!! Pattern Name not match ")
                'Exit Function
            'End If
            Else
                If MatchFlag And blMbistFP_Binout Then
                    For k = 0 To UBound(tpEvaPattCycleBlockInfor)
                        If tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                            lGetFlagIdx = GetFlagInfoArrIndex(tpEvaPattCycleBlockInfor(k).strFlagName)
                            If lGetFlagIdx >= 0 Then
                                tyFlagInfoArr(lGetFlagIdx).CheckInfo = True
                            End If
                        End If
                    Next k
                End If
                
                blPatPass = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
                
                '========================================================================================
                For Each site In theexec.sites
                    m_tn = theexec.sites.item(site).TestNumber
                    m_tn_restore = m_tn
'                    TheExec.sites.item(site).TestNumber = m_tn * 1000 + 1
                    '############################################################################################

                    If blPatPass(site) = False Then     ''  patt fail
                        '--------------------------------------------------------------------------
                        'SWLINZA 20181128 for MFP DTR, to get "flow-condition"(form flow) and compose DTR header
                        '--------------------------------------------------------------------------
                        MFP_flow_idx = theexec.sites.item(site).SiteVariableValue("MFP_Flow_Idx")
                        Pattern_Desc = tpCycleBlockInfor(MFP_pattern_idx).strDecsName(MFP_flow_idx)
                        Pattern_Server = tpCycleBlockInfor(MFP_pattern_idx).strServerName(MFP_flow_idx)
                        NewFmt_Printing_Header = "MemFP,1" & "," & site & "," & Pattern_Server & "," & Pattern_Desc & "," & UCase(m_testName) & "," & Pattern_GenericName(0) & ","
                        CharacterNumbers = LogLimited - Len(NewFmt_Printing_Header)
                        numcap = TheHdw.Digital.hram.CapturedCycles
                        If numcap > 0 Then
                            ReDim Pattern_Failure_Cycles(numcap - 1) As Variant
                        End If
                        If CharacterNumbers <= 0 Then
                            theexec.Datalog.WriteComment "The length of header is over than" & LogLimited & "."
                            theexec.Datalog.WriteComment "Please  Check the length of header which consist of instance name and pattern name."
                        End If

                        For i = 0 To UBound(Pattern_Failure_Cycles())
                            Pattern_Failure_Cycles(i) = vbNullString
                        Next i
                        CurCount_FailAry_Element = 0

                        For i = 0 To numcap - 1
                            Set PinData = TheHdw.Digital.Pins(AllPins).hram.PinData(i)
                            If Mbist_Repair_CompareType = "Cycle" Then
                                Mbist_repair_cycle = TheHdw.Digital.hram.PatGenInfo(i, pgCycle)
                                For j = 0 To UBound(tpEvaPattCycleBlockInfor)
                                    If Mbist_repair_cycle = tpEvaPattCycleBlockInfor(j).lCycle Then
                                        MBISTFailBlockFlag = True
                                        Exit For
                                    End If
                                Next j
                            ElseIf Mbist_Repair_CompareType = "Vector" Then
                                Mbist_repair_vector = TheHdw.Digital.hram.PatGenInfo(i, pgVector)
                                For j = 0 To UBound(tpEvaPattCycleBlockInfor)
                                    If Mbist_repair_vector = tpEvaPattCycleBlockInfor(j).lVector Then
                                        MBISTFailBlockFlag = True
                                        Exit For
                                    End If
                                Next j
                            End If
                            
                            For Each Pins In PinData.Pins
                                Cdata = Pins.value(site)
                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                If LCase(theexec.DataManager.instancename) Like "*bist*" Then

                                    If MBISTFailBlockFlag Then
                                        MBISTFailBlockFlag = False
                                        '=================================================================================
                                        If Mbist_Repair_CompareType = "Cycle" Then
                                        'If Mbist_repair_cycle = tpEvaPattCycleBlockInfor(k).lCycle Then
                                            For k = Count To UBound(tpEvaPattCycleBlockInfor)
                                                If Mbist_repair_cycle = tpEvaPattCycleBlockInfor(k).lCycle Then
                                                    If tpEvaPattCycleBlockInfor(k).strCompare <> Cdata Then
                                                        PassOrFail(site) = 0
                                                        If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                                             theexec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicTrue
                                                        End If
                                                                                                    
                                                        '---------------------------------------------------------------------------------------------
                                                        ' SWLINZA 20181128, for MemFP DTR, C651/Si requests 2018/08/M,
                                                        ' to count Fail Cycles and store in ary for print later
                                                        ' Array element + Header must lower than 255, if it's over than 255 then print warning message
                                                        '----------------------------------------------------------------------------------------------
                                                        If Len("," & CStr(Mbist_repair_cycle)) < CharacterNumbers Then
                                                            If Len(Pattern_Failure_Cycles(CurCount_FailAry_Element) & "," & CStr(Mbist_repair_cycle)) > CharacterNumbers Then
                                                                CurCount_FailAry_Element = CurCount_FailAry_Element + 1
                                                            Else
                                                                CurCount_FailAry_Element = CurCount_FailAry_Element
                                                            End If
                                                            If CurCount_FailAry_Element > 50 Then
                                                                ReDim Preserve Pattern_Failure_Cycles(CurCount_FailAry_Element)
                                                            End If
                                                            If Pattern_Failure_Cycles(CurCount_FailAry_Element) = "" Then
                                                                Pattern_Failure_Cycles(CurCount_FailAry_Element) = CStr(Mbist_repair_cycle)
                                                            Else
                                                                Pattern_Failure_Cycles(CurCount_FailAry_Element) = Pattern_Failure_Cycles(CurCount_FailAry_Element) & "," & CStr(Mbist_repair_cycle)
                                                            End If
                                                        Else
                                                            theexec.Datalog.WriteComment "The Remianing Character Numbers is not enough to output failing cycles."
                                                            theexec.Datalog.WriteComment "Please check the length of header which consist of instance name and pattern name."
                                                        End If
                                                        
                                                    Else
                                                        PassOrFail(site) = 1
                                                        If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                                            If theexec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                                                                 theexec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                                                            End If
                                                        End If
                                                    End If
                                                    blJump = True
                                                    Count = k + 1
                                                Else
                                                    PassOrFail(site) = 1
                                                    If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                                        If theexec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                                                             theexec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                                                        End If
                                                    End If
                                                End If
                                                
                                                ReduceBlkLen tpEvaPattCycleBlockInfor(k).strBlaclName, ReviseStr
                                                LogLen = Len(ReviseStr)
                                                If LogLen Mod LogLimited <> 0 Then PrintTimes = (LogLen \ LogLimited) + 1
                                                For PrintIdx = 0 To PrintTimes - 1
                                                    ReDim Preserve DecomposeLog(PrintTimes - 1)
                                                    DecomposeLog(PrintIdx) = mid(ReviseStr, 1 + PrintIdx * LogLimited, LogLimited)
                                                    theexec.Flow.TestLimit PassOrFail(site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & m_testName, , _
                                                            DecomposeLog(PrintIdx) & " " & tpEvaPattCycleBlockInfor(k).lCycle, , , , , tlForceNone
                                                Next PrintIdx
    '''                                                    TheExec.Flow.TestLimit PassOrFail(site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & m_testName, , _
    '''                                                            tpEvaPattCycleBlockInfor(k).strBlaclName & " " & tpEvaPattCycleBlockInfor(k).lCycle, , , , , tlForceNone
                                                If blJump = True Then
                                                    blJump = False
                                                    Exit For
                                                End If
                                            Next k
                                        '=================================================================================
                                        ElseIf Mbist_Repair_CompareType = "Vector" Then
                                            For k = Count To UBound(tpEvaPattCycleBlockInfor)
                                                If Mbist_repair_vector = tpEvaPattCycleBlockInfor(k).lVector Then
                                                    If tpEvaPattCycleBlockInfor(k).strCompare <> Cdata Then
                                                        PassOrFail(site) = 0
                                                        tpEvaPattCycleBlockInfor(k).strFlagName = "fail"
                                                        If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                                             theexec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicTrue
                                                        End If
                                                                                                    
                                                        '---------------------------------------------------------------------------------------------
                                                        ' SWLINZA 20181128, for MemFP DTR, C651/Si requests 2018/08/M
                                                        ' to count Fail vectors and store in ary for print later
                                                        ' Array element + Header must lower than 255, if it's over than 255 then print warning message
                                                        '----------------------------------------------------------------------------------------------
                                                        If Len("," & CStr(Mbist_repair_vector)) < CharacterNumbers Then
                                                            If Len(Pattern_Failure_Cycles(CurCount_FailAry_Element) & "," & CStr(Mbist_repair_vector)) > CharacterNumbers Then
                                                                CurCount_FailAry_Element = CurCount_FailAry_Element + 1
                                                            Else
                                                                CurCount_FailAry_Element = CurCount_FailAry_Element
                                                            End If
                                                            If CurCount_FailAry_Element > 50 Then
                                                                ReDim Preserve Pattern_Failure_Cycles(CurCount_FailAry_Element)
                                                            End If
                                                            If Pattern_Failure_Cycles(CurCount_FailAry_Element) = "" Then
                                                                Pattern_Failure_Cycles(CurCount_FailAry_Element) = CStr(Mbist_repair_vector)
                                                            Else
                                                                Pattern_Failure_Cycles(CurCount_FailAry_Element) = Pattern_Failure_Cycles(CurCount_FailAry_Element) & "," & CStr(Mbist_repair_vector)
                                                            End If
                                                        Else
                                                            theexec.Datalog.WriteComment "The Remianing Character Numbers is not enough to output failing vectors."
                                                            theexec.Datalog.WriteComment "Please check the length of header which consist of instance name and pattern name."
                                                        End If

                                                    Else
                                                        PassOrFail(site) = 1
                                                        tpEvaPattCycleBlockInfor(k).strFlagName = "pass"
                                                        If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                                            If theexec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                                                                theexec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                                                            End If
                                                        End If
                                                    End If
                                                    blJump = True
                                                    Count = k + 1
                                                Else
                                                    PassOrFail(site) = 1
                                                    tpEvaPattCycleBlockInfor(k).strFlagName = "pass"
                                                    If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                                        If theexec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                                                            theexec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                                                        End If
                                                    End If
                                                End If
                                                
                                                
                                                ReduceBlkLen tpEvaPattCycleBlockInfor(k).strBlaclName, ReviseStr
                                                LogLen = Len(ReviseStr)
                                                If LogLen Mod LogLimited <> 0 Then PrintTimes = (LogLen \ LogLimited) + 1
                                                For PrintIdx = 0 To PrintTimes - 1
                                                    ReDim Preserve DecomposeLog(PrintTimes - 1)
                                                    DecomposeLog(PrintIdx) = mid(ReviseStr, 1 + PrintIdx * LogLimited, LogLimited)
                                                    theexec.Flow.TestLimit PassOrFail(site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & m_testName, , _
                                                            DecomposeLog(PrintIdx) & " " & tpEvaPattCycleBlockInfor(k).lVector, , , , , tlForceNone
                                                Next PrintIdx
    '''                                                    TheExec.Flow.TestLimit PassOrFail(site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & m_testName, , _
    '''                                                            tpEvaPattCycleBlockInfor(k).strBlaclName & " " & tpEvaPattCycleBlockInfor(k).lVector, , , , , tlForceNone
                                                If blJump = True Then
                                                    blJump = False
                                                    Exit For
                                                End If
                                            Next k
                                        End If
                                    '=================================================================================
                                    End If
                                End If
                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            Next Pins
                        Next i
                        
                        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                        If k < UBound(tpEvaPattCycleBlockInfor) Then        '' in unread all info of  tpEvaPattCycleBlockInfor case
                            If numcap < maxDepth Then
                                PassOrFail(site) = 1
                                For k = Count To UBound(tpEvaPattCycleBlockInfor)
                                    If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                        If theexec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                                             theexec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                                        End If
                                    End If

                                    ReduceBlkLen tpEvaPattCycleBlockInfor(k).strBlaclName, ReviseStr
                                    LogLen = Len(ReviseStr)
                                    If LogLen Mod LogLimited <> 0 Then PrintTimes = (LogLen \ LogLimited) + 1
                                    For PrintIdx = 0 To PrintTimes - 1
                                        ReDim Preserve DecomposeLog(PrintTimes - 1)
                                        DecomposeLog(PrintIdx) = mid(ReviseStr, 1 + PrintIdx * LogLimited, LogLimited)
                                        theexec.Flow.TestLimit PassOrFail(site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & m_testName, , _
                                                DecomposeLog(PrintIdx) & " " & tpEvaPattCycleBlockInfor(k).lCycle, , , , , tlForceNone
                                    Next PrintIdx
'''                                            TheExec.Flow.TestLimit PassOrFail(site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & m_testName, , _
'''                                                    tpEvaPattCycleBlockInfor(k).strBlaclName & " " & tpEvaPattCycleBlockInfor(k).lCycle, , , , , tlForceNone
                                Next k
                            Else
                                '' add for HRAM is full and still have some cycles need to judge, to set all flag status = true
                                If gl_MbistFP_Binout Then
                                    For k = 0 To UBound(tpEvaPattCycleBlockInfor)
                                        If tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then theexec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicTrue
                                    Next k
                                End If
                            End If
                        End If
                        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                        If numcap >= maxDepth Then   '' HRAM is full
                            theexec.Flow.TestLimit 0, , , , , scaleNoScaling, , , "MemFP_" & m_testName, , _
                                                      "Fail_cycle_size_check", , , , , tlForceNone
                            theexec.Datalog.WriteComment ("The number of pattern fail cycles full or exceed HRAM maxDepth: " & maxDepth)
                        Else
                            theexec.Flow.TestLimit 1, , , , , scaleNoScaling, , , "MemFP_" & m_testName, , _
                                                      "Fail_cycle_size_check", , , , , tlForceNone
                        End If
                        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                        theexec.sites.item(site).TestNumber = m_tn_restore + m_tn_BurstIndex
                        Count = 0
                        k = 0
                        
                        '----------------------------------------------------------------
                        ' SWLINZA 20181128, for MemFP DTR, C651/Si requests 2018/08/M
                        ' To print final DTR,
                        ' Pattern_Failure_Cycles(AryIdx) is stored in previous procedure
                        '----------------------------------------------------------------
                        Dim AryIdx As Long
                        theexec.Datalog.WriteComment ""
                        For AryIdx = 0 To CurCount_FailAry_Element
                            theexec.Datalog.WriteComment NewFmt_Printing_Header & Pattern_Failure_Cycles(AryIdx)
                        Next AryIdx
                        theexec.Datalog.WriteComment ""
                        
                    Else    ''blPatPass(Site) = True
                        If blMbistFP_Binout Then
                            For k = 0 To UBound(tpEvaPattCycleBlockInfor)
                                If tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                    If theexec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                                        theexec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                                    End If
                                End If
                            Next k
                        End If
                    End If '' If blPatPass(Site)
                    '############################################################################################
                Next site
            End If
            '========================================================================================
        End If
        '//////////////////////////////////////////////////////////////////////////////////////////////'''''''''''finger print_block_03_end
        If glb_TesterType = "UltraFLEXplus" Then 'add for GetAlarmingSites
            Dim alarmOccurred As New SiteBoolean
            alarmOccurred = False
            alarmOccurred = TheHdw.Alarms.GetAlarmingSites(True)        ''GetAlarmingSites(clearAlarm = True)
            For Each site In theexec.sites
                If alarmOccurred(site) = True Then
                    alarmFail(site) = True      ''Update alarm info to alarmFail array for UFP
                End If
            Next site
        End If
        If mbist_loop Then
            '===================================================================
            For Each site In theexec.sites
                'testnumber(Site) = TheExec.sites.Item(Site).testnumber
                tested(site) = False
                blPatPass(site) = TheHdw.Digital.Patgen.PatternBurstPassed
                '-------------------------------------------------------------------------------------------------
                If blPatPass(site) = False Or alarmFail(site) = True Then   'pattern test fail or alarm
                    theexec.sites.item(site).FlagState(Flag_Name) = logicTrue 'pattern test fail
                    theexec.sites.item(site).testResult = siteFail
                    tested(site) = True
                    'TheExec.sites.Item(Site).testnumber = TheExec.sites.Item(Site).testnumber + 1
               '-------------------------------------------------------------------------------------------------
                Else    'blPatPass(Site) = True ; pattern test pass
                    If (tested(site) = False) Then
                        If (theexec.sites.item(site).FlagState(Flag_Name) <> logicTrue) Then 'confirm flag is true(pattern fail)
                            theexec.sites.item(site).FlagState(Flag_Name) = logicFalse       'pattern test pass
                        End If
                        theexec.sites.item(site).testResult = sitePass
                    End If
                        'Call TheExec.Datalog.WriteFunctionalResult(Site, testnumber(Site), logTestPass, , ins_new_name)
                        'TheExec.sites.Item(Site).testnumber = TheExec.sites.Item(Site).testnumber + 1
                End If  '' If blPatPass(Site) End
                '-------------------------------------------------------------------------------------------------
                'TheExec.Datalog.WriteComment "Instance                = " & ins_new_name
                'TheExec.Datalog.WriteComment "Pat Name                = " & m_pattname
                'TheExec.Datalog.WriteComment "Test Falg               =>" & flag_name & "(" & Site & ") = " & TheExec.sites.Item(Site).FlagState(flag_name) & ",     if pattern pass=> flag is logicFalse => 0" & ",     if pattern fail=> flag is logicTrue => 1"
                blPatPass(site) = False
                alarmFail(site) = False
            Next site
            '===================================================================
        End If
    Next patt
        
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_MBIST", "Finger_print_Org")
    If AbortTest Then Exit Function Else Resume Next
End Function


' [20230425][All][Carter] use for Plus
' [20230710][All][Tank] Modify VBT error, and reduce no use variable
' [20231228][T-Son][Tank] Add Finger_Print need setup HRAM When pre pattern (if PLUS and set long function "short")
Private Sub Finger_Print_NewSyntex(pattern_load As String, RunFailCycle As Boolean, Optional Flag_Name As String, Optional mbist_loop As Boolean = False)
On Error GoTo errHandler

    Dim maxDepth As Integer
    Dim patt As Variant
    Dim site As Variant

    Dim rtnPatternNames() As String, rtnPatternCount As Long
    Dim blPatPass As New SiteBoolean
    Dim tested As New SiteBoolean
    Dim AllPins As String

    Dim LogLimited As Long

    Dim blMbistFP_Binout As Boolean
    Dim PassOrFail As New SiteLong
    Dim lGetFlagIdx As Long
    Dim blJump As Boolean
    Dim m_testName As String
    Dim k As Long, p As Long, g As Long, j As Long, i As Long:: k = 0:: p = 0:: g = 0:: j = 0:: i = 0

    '----------------------------------------------------------------
    ' SWLINZA 20181128, for MemFP DTR, C651/Si requests 2018/08
    '----------------------------------------------------------------
    Dim Pattern_Desc As String
    Dim Pattern_Server As String
    Dim CurCount_FailAry_Element As Long        'Dim CurCount_FailAry_Element As New SiteLong
    Dim NewFmt_Printing_Header As String
    Dim CharacterNumbers As Long
    Dim s_Pattern_GenericName As String
    Dim MFP_pattern_idx As Long
    Dim MFP_flow_idx As Variant

    Dim sb_FingerPrint_Fail As New SiteBoolean

    'Dim Pattern_Failure_Cycles() As New SiteVariant
    Dim Pattern_Failure_Cycles() As Variant

    Dim l_patcnt As Long
    Dim l_pgVector() As Long
    Dim l_PatternSetIndex() As Long

    Dim d_pgCycle() As Double

    Dim Pattern_Ary_absolute() As String
    Dim s_TempPins() As String
    Dim Pin_Cnt As Long
        
    sb_FingerPrint_Fail = False
    AllPins = "JTAG_TDO"
    LogLimited = 255
    m_testName = Theexec.DataManager.instanceName
''    Call PATT_GetPatListFromPatternSet(pattern_load, rtnPatternNames, rtnPatternCount)
    Call GetPatsFromPatSets(pattern_load, rtnPatternNames, rtnPatternCount, False)
    Call GetPatsFromPatSets(pattern_load, Pattern_Ary_absolute, rtnPatternCount, True)

    'HRAM Setup pre pattern
    Call tl_SetInterpose(TL_C_PREPATF, "HRAMSetupInterpose")

''    TheExec.Datalog.WriteComment "*************************************************"
''    TheExec.Datalog.WriteComment "Capture Failing Cycle Info for Test Instance " & TheExec.DataManager.instancename & " (Max HRAM Size: " & maxDepth & ")"
''    TheExec.Datalog.WriteComment "*************************************************"

    Dim Cycle As IAllCapturedCycleInfo ''20221011, New syntax
''    Dim firstCycleInfo As ICapturedCycleInfo

    Theexec.DataManager.DecomposePinList AllPins, s_TempPins(), Pin_Cnt
    If Pin_Cnt <> 1 Then
        Theexec.Datalog.WriteComment "[Warning] New Syntex just can process 1 pin 1 time!!"
        Exit Sub
    Else
    End If

    If BurstYesPatDict.Exists(LCase(pattern_load)) = True Then
    ''Pattern Burst Yes
        Call TheHdw.Patterns(pattern_load).test(pfAlways, 0, tlResultModeDomain)
        blPatPass = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
        If blPatPass.Any(False) Then
            Set Cycle = TheHdw.Digital.Pins(AllPins).hram.CapturedFailCycleInfo
''            Set firstCycleInfo = TheHdw.Digital.HRAM.FirstCapturedCycleInfo
            l_patcnt = 0
            For Each patt In Pattern_Ary_absolute
            ''Find out the Result for failure site per pattern module
                s_Pattern_GenericName = GetPatName_FromAbsolute(CStr(patt))
                MFP_pattern_idx = GetPatIndex_FromFP(s_Pattern_GenericName, MatchFlag)
                tpEvaPattCycleBlockInfor = tpCycleBlockInfor(MFP_pattern_idx).tpMbistCycleBlock

                If RunFailCycle = True Then
                    'TheExec.Datalog.WriteComment ("Warning!! Pattern Name not match ")'ZYCHOUA, 220706
                    'Exit Function
                    '=========================Need to check where be used==========================================
                    If MatchFlag And blMbistFP_Binout Then
                        For k = 0 To UBound(tpEvaPattCycleBlockInfor)
                            If tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                lGetFlagIdx = GetFlagInfoArrIndex(tpEvaPattCycleBlockInfor(k).strFlagName)
                                If lGetFlagIdx >= 0 Then
                                    tyFlagInfoArr(lGetFlagIdx).CheckInfo = True
                                End If
                            End If
                        Next k
                    End If
                    '=========================Need to check where be used==========================================

                    '--------------------------------------------------------------------------
                    'SWLINZA 20181128 for MFP DTR, to get "flow-condition"(form flow) and compose DTR header
                    '--------------------------------------------------------------------------
                    MFP_flow_idx = Theexec.sites.item(site).SiteVariableValue("MFP_Flow_Idx")
                    Pattern_Desc = tpCycleBlockInfor(MFP_pattern_idx).strDecsName(MFP_flow_idx)
                    Pattern_Server = tpCycleBlockInfor(MFP_pattern_idx).strServerName(MFP_flow_idx)

                    For Each site In Theexec.sites
                        If blPatPass(site) = False Then
                            If Cycle.HasCaptureOccurred(site) Then
                                l_PatternSetIndex = Cycle.PatternSetIndex(site)

                                d_pgCycle = Cycle.pgModuleCount(site)
                                l_pgVector = Cycle.pgVector(site)
                                ReDim Pattern_Failure_Cycles(UBound(d_pgCycle)) As Variant

                                NewFmt_Printing_Header = "MemFP,1" & "," & site & "," & Pattern_Server & "," & Pattern_Desc & "," & glb_TestInstance & "," & s_Pattern_GenericName & ","
                                CharacterNumbers = LogLimited - Len(NewFmt_Printing_Header)
                                If CharacterNumbers <= 0 Then
                                    Theexec.Datalog.WriteComment "The length of header is over than" & LogLimited & "."
                                    Theexec.Datalog.WriteComment "Please  Check the length of header which consist of instance name and pattern name."
                                End If

                                For i = 0 To UBound(Pattern_Failure_Cycles())
                                    Pattern_Failure_Cycles(i) = vbNullString
                                Next i
                                CurCount_FailAry_Element = 0

                                Call Analyze_HRAMCapturedFailInfo(l_PatternSetIndex, l_pgVector, d_pgCycle, AllPins, Mbist_Repair_CompareType, l_patcnt, PassOrFail, sb_FingerPrint_Fail, _
                                                                    Pattern_Failure_Cycles(), CurCount_FailAry_Element, blMbistFP_Binout, blJump, CharacterNumbers, LogLimited, maxDepth, NewFmt_Printing_Header, site)

                            End If
                        Else
                            If blMbistFP_Binout Then
                                For k = 0 To UBound(tpEvaPattCycleBlockInfor)
                                    If tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                        If Theexec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                                            Theexec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                                        End If
                                    End If
                                Next k
                            End If
                        End If

                    Next site

                End If
                l_patcnt = l_patcnt + 1
            Next patt
        End If

    Else
    ''Pattern Burst No
        l_patcnt = 0
        For Each patt In Pattern_Ary_absolute
            s_Pattern_GenericName = GetPatName_FromAbsolute(CStr(patt))

            Call TheHdw.Patterns(patt).test(pfAlways, 0, tlResultModeDomain)
            blPatPass = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
            If blPatPass.Any(False) Then
                Set Cycle = TheHdw.Digital.Pins(AllPins).hram.CapturedFailCycleInfo
''                Set firstCycleInfo = TheHdw.Digital.HRAM.FirstCapturedCycleInfo
            End If
            ''Find out the Result for failure site per pattern module
            MFP_pattern_idx = GetPatIndex_FromFP(s_Pattern_GenericName, MatchFlag)
            tpEvaPattCycleBlockInfor = tpCycleBlockInfor(MFP_pattern_idx).tpMbistCycleBlock

            If RunFailCycle = True Then
                'TheExec.Datalog.WriteComment ("Warning!! Pattern Name not match ")'ZYCHOUA, 220706
                'Exit Function
                '=========================Need to check where be used==========================================
                If MatchFlag And blMbistFP_Binout Then
                    For k = 0 To UBound(tpEvaPattCycleBlockInfor)
                        If tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                            lGetFlagIdx = GetFlagInfoArrIndex(tpEvaPattCycleBlockInfor(k).strFlagName)
                            If lGetFlagIdx >= 0 Then
                                tyFlagInfoArr(lGetFlagIdx).CheckInfo = True
                            End If
                        End If
                    Next k
                End If
                '=========================Need to check where be used==========================================

                '--------------------------------------------------------------------------
                'SWLINZA 20181128 for MFP DTR, to get "flow-condition"(form flow) and compose DTR header
                '--------------------------------------------------------------------------
                MFP_flow_idx = Theexec.sites.item(site).SiteVariableValue("MFP_Flow_Idx")
                Pattern_Desc = tpCycleBlockInfor(MFP_pattern_idx).strDecsName(MFP_flow_idx)
                Pattern_Server = tpCycleBlockInfor(MFP_pattern_idx).strServerName(MFP_flow_idx)

                For Each site In Theexec.sites
                    If blPatPass(site) = False Then
                        If Cycle.HasCaptureOccurred(site) Then
                            l_PatternSetIndex = Cycle.PatternSetIndex(site)
                            d_pgCycle = Cycle.pgCycle(site)
                            l_pgVector = Cycle.pgVector(site)
                            ReDim Pattern_Failure_Cycles(UBound(d_pgCycle)) As Variant

                            NewFmt_Printing_Header = "MemFP,1" & "," & site & "," & Pattern_Server & "," & Pattern_Desc & "," & glb_TestInstance & "," & s_Pattern_GenericName & ","
                            CharacterNumbers = LogLimited - Len(NewFmt_Printing_Header)
                            If CharacterNumbers <= 0 Then
                                Theexec.Datalog.WriteComment "The length of header is over than " & LogLimited & "."
                                Theexec.Datalog.WriteComment "Please  Check the length of header which consist of instance name and pattern name."
                            End If

                            For i = 0 To UBound(Pattern_Failure_Cycles())
                                Pattern_Failure_Cycles(i) = ""
                            Next i
                            CurCount_FailAry_Element = 0

                            Call Analyze_HRAMCapturedFailInfo(l_PatternSetIndex, l_pgVector, d_pgCycle, AllPins, Mbist_Repair_CompareType, l_patcnt, PassOrFail, sb_FingerPrint_Fail, _
                                    Pattern_Failure_Cycles(), CurCount_FailAry_Element, blMbistFP_Binout, blJump, CharacterNumbers, LogLimited, maxDepth, NewFmt_Printing_Header, site)

                        End If
                    Else
                        If blMbistFP_Binout Then
                            For k = 0 To UBound(tpEvaPattCycleBlockInfor)
                                If tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                    If Theexec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                                        Theexec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                                    End If
                                End If
                            Next k
                        End If
                    End If

                Next site

            End If
''            l_patcnt = l_patcnt + 1

        Next patt

    End If
    '//////////////////////////////////////////////////////////////////////////////////////////////'''''''''''finger print_block_03_end
    If mbist_loop Then
        '===================================================================
        blPatPass = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
        For Each site In Theexec.sites
            'testnumber(Site) = TheExec.sites.Item(Site).testnumber
            tested(site) = False
    ''                blPatPass(site) = thehdw.Digital.Patgen.PatternBurstPassed
            '-------------------------------------------------------------------------------------------------
            If blPatPass(site) = False Or alarmFail(site) = True Then   'pattern test fail or alarm
                Theexec.sites.item(site).FlagState(Flag_Name) = logicTrue 'pattern test fail
                Theexec.sites.item(site).testResult = siteFail
                tested(site) = True
                'TheExec.sites.Item(Site).testnumber = TheExec.sites.Item(Site).testnumber + 1
           '-------------------------------------------------------------------------------------------------
            Else    'blPatPass(Site) = True ; pattern test pass
                If (tested(site) = False) Then
                    If (Theexec.sites.item(site).FlagState(Flag_Name) <> logicTrue) Then 'confirm flag is true(pattern fail)
                        Theexec.sites.item(site).FlagState(Flag_Name) = logicFalse       'pattern test pass
                    End If
                    Theexec.sites.item(site).testResult = sitePass
                End If
                    'Call TheExec.Datalog.WriteFunctionalResult(Site, testnumber(Site), logTestPass, , ins_new_name)
                    'TheExec.sites.Item(Site).testnumber = TheExec.sites.Item(Site).testnumber + 1
            End If  '' If blPatPass(Site) End
            '-------------------------------------------------------------------------------------------------
            'TheExec.Datalog.WriteComment "Instance                = " & ins_new_name
            'TheExec.Datalog.WriteComment "Pat Name                = " & m_pattname
            'TheExec.Datalog.WriteComment "Test Falg               =>" & flag_name & "(" & Site & ") = " & TheExec.sites.Item(Site).FlagState(flag_name) & ",     if pattern pass=> flag is logicFalse => 0" & ",     if pattern fail=> flag is logicTrue => 1"
            blPatPass(site) = False
            alarmFail(site) = False
        Next site
    End If
    '===================================================================

    Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_MBIST", "Finger_Print_NewSyntex")
    If AbortTest Then Exit Sub Else Resume Next
End Sub


Private Function Analyze_HRAMCapturedFailInfo(in_PatternSetIndex() As Long, in_pgVector() As Long, in_pgCycle() As Double, ByVal CapPin As String, _
                ByVal Mbist_Repair_CompareType As String, ByVal l_patcnt As Long, PassOrFail As SiteLong, sb_FingerPrint_Fail As SiteBoolean, _
                Pattern_Failure_Cycles() As Variant, CurCount_FailAry_Element As Long, blMbistFP_Binout As Boolean, blJump As Boolean, _
                CharacterNumbers As Long, LogLimited As Long, maxDepth As Integer, NewFmt_Printing_Header As String, site As Variant)
    On Error GoTo errHandler
    
    Dim ReviseStr As String
    Dim DecomposeLog() As String
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l_Count As Long
    Dim LogLen As Long
    Dim PrintIdx As Long
    Dim PrintTimes As Long
    
    Dim d_FailCycleVector As Double
    
    Dim MBISTFailBlockFlag As Boolean
    
    Dim AllPins As String
    
    Dim PinData As New PinListData
    Dim Pins As Variant
    Dim Cdata As Variant
    
    For i = 0 To UBound(in_PatternSetIndex)
        If in_PatternSetIndex(i) = l_patcnt Then
        ''check which pattern module we wanna to see
            Select Case LCase(Mbist_Repair_CompareType)
                Case "cycle":
                    d_FailCycleVector = in_pgCycle(i)
                    For j = 0 To UBound(tpEvaPattCycleBlockInfor)
                        If d_FailCycleVector = tpEvaPattCycleBlockInfor(j).lCycle Then
                            MBISTFailBlockFlag = True
                            Exit For
                        End If
                    Next j
                Case "vector":
                    d_FailCycleVector = in_pgVector(i)
                    For j = 0 To UBound(tpEvaPattCycleBlockInfor)
                        If d_FailCycleVector = tpEvaPattCycleBlockInfor(j).lVector Then
                            MBISTFailBlockFlag = True
                            Exit For
                        End If
                    Next j
            End Select
            
            If MBISTFailBlockFlag Then
                MBISTFailBlockFlag = False
                Set PinData = TheHdw.Digital.Pins(CapPin).hram.PinData(i)
                For Each Pins In PinData.Pins
                    Cdata = Pins.value
                    For k = l_Count To UBound(tpEvaPattCycleBlockInfor)
                        Select Case LCase(Mbist_Repair_CompareType)
                            Case "cycle":
                                Call Compose_HRAMCapturedFailInfo(d_FailCycleVector, tpEvaPattCycleBlockInfor(k).lCycle, Cdata, k, l_Count, PassOrFail, sb_FingerPrint_Fail, Pattern_Failure_Cycles(), CurCount_FailAry_Element, blMbistFP_Binout, blJump, CharacterNumbers, site)
                            Case "vector":
                                Call Compose_HRAMCapturedFailInfo(d_FailCycleVector, tpEvaPattCycleBlockInfor(k).lVector, Cdata, k, l_Count, PassOrFail, sb_FingerPrint_Fail, Pattern_Failure_Cycles(), CurCount_FailAry_Element, blMbistFP_Binout, blJump, CharacterNumbers, site)
                        End Select
                        
                        ReduceBlkLen tpEvaPattCycleBlockInfor(k).strBlaclName, ReviseStr
                        LogLen = Len(ReviseStr)
                        If LogLen Mod LogLimited <> 0 Then PrintTimes = (LogLen \ LogLimited) + 1
                        For PrintIdx = 0 To PrintTimes - 1
                            ReDim Preserve DecomposeLog(PrintTimes - 1)
                            DecomposeLog(PrintIdx) = mid(ReviseStr, 1 + PrintIdx * LogLimited, LogLimited)
                            Select Case LCase(Mbist_Repair_CompareType)
                                Case "cycle":
                                    theexec.Flow.TestLimit PassOrFail(site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & glb_TestInstance, , _
                                            DecomposeLog(PrintIdx) & " " & tpEvaPattCycleBlockInfor(k).lCycle, , , , , tlForceNone
                                Case "vector":
                                    theexec.Flow.TestLimit PassOrFail(site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & glb_TestInstance, , _
                                            DecomposeLog(PrintIdx) & " " & tpEvaPattCycleBlockInfor(k).lVector, , , , , tlForceNone
                            End Select
                        Next PrintIdx

                        If blJump = True Then
                            blJump = False
                            Exit For
                        End If
                    Next k
                Next Pins
            End If
        End If
    Next i
    
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If k < UBound(tpEvaPattCycleBlockInfor) And sb_FingerPrint_Fail Then        '' in unread all info of  tpEvaPattCycleBlockInfor case
        If UBound(in_PatternSetIndex) < maxDepth Then
            PassOrFail(site) = 1
            For k = l_Count To UBound(tpEvaPattCycleBlockInfor)
                If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                    If theexec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                        theexec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                    End If
                End If
                ReduceBlkLen tpEvaPattCycleBlockInfor(k).strBlaclName, ReviseStr
                LogLen = Len(ReviseStr)
                If LogLen Mod LogLimited <> 0 Then PrintTimes = (LogLen \ LogLimited) + 1
                For PrintIdx = 0 To PrintTimes - 1
                    ReDim Preserve DecomposeLog(PrintTimes - 1)
                    DecomposeLog(PrintIdx) = mid(ReviseStr, 1 + PrintIdx * LogLimited, LogLimited)
                    Select Case LCase(Mbist_Repair_CompareType)
                        Case "cycle":
                            theexec.Flow.TestLimit PassOrFail(site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & glb_TestInstance, , _
                                    DecomposeLog(PrintIdx) & " " & tpEvaPattCycleBlockInfor(k).lCycle, , , , , tlForceNone
                        Case "vector":
                            theexec.Flow.TestLimit PassOrFail(site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & glb_TestInstance, , _
                                    DecomposeLog(PrintIdx) & " " & tpEvaPattCycleBlockInfor(k).lVector, , , , , tlForceNone
                    End Select
                Next PrintIdx
            Next k
        Else
            '' add for HRAM is full and still have some cycles need to judge, to set all flag status = true
            If gl_MbistFP_Binout Then
                For k = 0 To UBound(tpEvaPattCycleBlockInfor)
                    If tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then theexec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicTrue
                Next k
            End If
        End If
    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If sb_FingerPrint_Fail Then
        If UBound(in_PatternSetIndex) >= maxDepth Then   '' HRAM is full
            theexec.Flow.TestLimit 0, , , , , scaleNoScaling, , , "MemFP_" & glb_TestInstance, , _
                                      "Fail_cycle_size_check", , , , , tlForceNone
            theexec.Datalog.WriteComment ("The number of pattern fail cycles full or exceed HRAM maxDepth: " & maxDepth)
        Else
            theexec.Flow.TestLimit 1, , , , , scaleNoScaling, , , "MemFP_" & glb_TestInstance, , _
                                      "Fail_cycle_size_check", , , , , tlForceNone
        End If
    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    l_Count = 0
    k = 0
    
    '----------------------------------------------------------------
    ' SWLINZA 20181128, for MemFP DTR, C651/Si requests 2018/08/M
    ' To print final DTR,
    ' Pattern_Failure_Cycles(AryIdx) is stored in previous procedure
    '----------------------------------------------------------------
    Dim AryIdx As Long
    If sb_FingerPrint_Fail Then
        theexec.Datalog.WriteComment ""
        For AryIdx = 0 To CurCount_FailAry_Element
            theexec.Datalog.WriteComment NewFmt_Printing_Header & Pattern_Failure_Cycles(AryIdx)
        Next AryIdx
        theexec.Datalog.WriteComment ""
    End If
                            
    Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, "LIB_MBIST", "Analyze_HRAMCapturedFailInfo")
    If AbortTest Then Exit Function Else Resume Next
End Function


Private Function Compose_HRAMCapturedFailInfo(ByVal d_FailCycleVector As Double, ByVal lVectorCycle As Long, ByVal Cdata As Variant, ByVal k As Long, l_Count As Long, _
                    PassOrFail As SiteLong, sb_FingerPrint_Fail As SiteBoolean, Pattern_Failure_Cycles() As Variant, _
                    CurCount_FailAry_Element As Long, blMbistFP_Binout As Boolean, blJump As Boolean, CharacterNumbers As Long, site As Variant)
    On Error GoTo errHandler
    
    If d_FailCycleVector = lVectorCycle Then
        If Cdata <> tpEvaPattCycleBlockInfor(k).strCompare Then
            PassOrFail(site) = 0
            sb_FingerPrint_Fail(site) = True
            tpEvaPattCycleBlockInfor(k).strFlagName = "fail"
            If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                 theexec.sites.item(tpEvaPattCycleBlockInfor(k).strFlagName) = logicTrue
            End If
                                                        
            '---------------------------------------------------------------------------------------------
            ' SWLINZA 20181128, for MemFP DTR, C651/Si requests 2018/08/M
            ' to count Fail vectors and store in ary for print later
            ' Array element + Header must lower than 255, if it's over than 255 then print warning message
            '----------------------------------------------------------------------------------------------
            If Len("," & CStr(d_FailCycleVector)) < CharacterNumbers Then
                If Len(Pattern_Failure_Cycles(CurCount_FailAry_Element) & "," & CStr(d_FailCycleVector)) > CharacterNumbers Then
                    CurCount_FailAry_Element = CurCount_FailAry_Element + 1
                Else
                    CurCount_FailAry_Element = CurCount_FailAry_Element
                End If
                If CurCount_FailAry_Element > UBound(Pattern_Failure_Cycles) Then        'If CurCount_FailAry_Element > 50 Then
                    ReDim Preserve Pattern_Failure_Cycles(CurCount_FailAry_Element)
                End If
                If Pattern_Failure_Cycles(CurCount_FailAry_Element) = "" Then
                    Pattern_Failure_Cycles(CurCount_FailAry_Element) = CStr(d_FailCycleVector)
                Else
                    Pattern_Failure_Cycles(CurCount_FailAry_Element) = Pattern_Failure_Cycles(CurCount_FailAry_Element) & "," & CStr(d_FailCycleVector)
                End If
            Else
                theexec.Datalog.WriteComment "The Remianing Character Numbers is not enough to output failing vectors/cycles."
                theexec.Datalog.WriteComment "Please check the length of header which consist of instance name and pattern name."
            End If
        Else
            PassOrFail = 1
            If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                If theexec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                     theexec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                End If
            End If
        End If
        blJump = True
        l_Count = k + 1
    Else
        PassOrFail(site) = 1
        tpEvaPattCycleBlockInfor(k).strFlagName = "pass"
        If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
            If theexec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                 theexec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
            End If
        End If
    End If
    
    Exit Function
    
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_MBIST", "Compose_HRAMCapturedFailInfo")
    If AbortTest Then Exit Function Else Resume Next
End Function


' [20231228][T-All][Tank] Add Finger_Print_NewSyntex ned use subfunction
Private Function GetPatName_FromAbsolute(s_pat As String) As String

    On Error GoTo errHandler

    Dim funcName As String:: funcName = "GetPatName_FromAbsolute"
    
    Dim strPattName As String
    
    Dim astrPattPathSplit() As String
    Dim Pattern_GenericName() As String
    
    astrPattPathSplit = Split(s_pat, "\")
    strPattName = UCase(astrPattPathSplit(UBound(astrPattPathSplit)))

    If strPattName Like "*.GZ" Then strPattName = Replace(strPattName, ".GZ", "")
    
    '---------------------------------------------
    'SWLINZA 20181128 for MFP DTR, to split patset
    '---------------------------------------------
    Pattern_GenericName() = Split(strPattName, ":")
    Pattern_GenericName(0) = UCase(Pattern_GenericName(0))
    
    If glb_TesterType = "Jaguar" Then
        If Pattern_GenericName(0) Like "*.PAT" Then Pattern_GenericName(0) = Replace(Pattern_GenericName(0), ".PAT", "")
    Else
        If Pattern_GenericName(0) Like "*.PATX" Then Pattern_GenericName(0) = Replace(Pattern_GenericName(0), ".PATX", "")
    End If
    
    GetPatName_FromAbsolute = Pattern_GenericName(0)
    Exit Function
errHandler:
    theexec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


' [20231228][T-All][Tank] Add Finger_Print_NewSyntex ned use subfunction
Private Function GetPatIndex_FromFP(in_pat_gen As String, ByRef MatchFlag) As Long
    On Error GoTo errHandler

    Dim funcName As String:: funcName = "GetPatIndex_FromFP"
    
    Dim i As Long
    Dim l_MFP_pattern_idx As Long
    
    For i = 0 To UBound(tpCycleBlockInfor)
        If UCase(tpCycleBlockInfor(i).strPattName) Like "*" & UCase(in_pat_gen) & "*" Then
''            tpEvaPattCycleBlockInfor = tpCycleBlockInfor(i).tpMbistCycleBlock
            MatchFlag = True
            l_MFP_pattern_idx = i 'SWLINZa 20180907 for MFP DTR, to indentify pattern#
            Exit For
        End If
    Next i
    GetPatIndex_FromFP = l_MFP_pattern_idx
    
Exit Function
errHandler:
    theexec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
