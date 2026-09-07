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

#Const IGXL_VER_1030 = False
'=======================20160301=======================================

Public Function MbistRampApplyLevel_AutoReadingContext(Optional ByVal ApplyPins As String = "CorePower", Optional RampingStep As Double = 10, Optional RampWaitTime As Double = 0, Optional instancename As String)

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
    
    If TheExec.enableWord("Ramping_MbistATPG") = False Then Exit Function
    
    TheExec.DataManager.DecomposePinList ApplyPins, Apply_Pins_Ary(), Apply_Pins_count
    ReDim Original_voltage(Apply_Pins_count - 1) As Double
    ReDim DiffVoltage(Apply_Pins_count - 1) As Double
    ReDim RampingVoltage(Apply_Pins_count - 1) As Double
    ReDim Apply_TargetVoltage(Apply_Pins_count - 1) As Double
    ReDim ApplyPins_Boolean(Apply_Pins_count - 1) As Boolean
    
    '----- to get target voltage from DC spec for each instance -----
    'Apply_TargetVoltage
    
    'Swlinza 20180126, to save test time in IGXL9.0, use this command instead of following two
    TheExec.DataManager.GetInstanceContext Current_DCCategory, Current_DCSelector, Dummy_tempStr, Dummy_tempStr, Dummy_tempStr, Dummy_tempStr, Dummy_tempStr, Dummy_tempStr
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
        Original_voltage(i) = FormatNumber(TheHdw.DCVS.pins(Apply_Pins_Ary(i)).Voltage.Main, 3)
        Apply_TargetVoltage(i) = TheExec.Specs.DC.item(Apply_Pins_Ary(i) & SepcSymbolic).Categories.item(Current_DCCategory).Selectors.item(Current_DCSelector).ContextValue
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
                    TheHdw.DCVS.pins(Apply_Pins_Ary(j)).Voltage.Main = Apply_TargetVoltage(j)
                Else
                    TheHdw.DCVS.pins(Apply_Pins_Ary(j)).Voltage.Main = Original_voltage(j) - RampingVoltage(j) * i
                End If
            End If
        Next j
        TheHdw.Wait Extra_RampingTime / RampingStep
    Next i
    
End Function

' [20230425][All][Carter] Add Finger print new syntex
Public Function Finger_print(pattern_load As String, RunFailCycle As Boolean, Optional Flag_Name As String, Optional mbist_loop As Boolean = False)
 
#If IGXL_VER_1030 = True Then
        If glb_TesterType = "UltraFLEXplus" Then
        Call Finger_Print_NewSyntex(pattern_load, RunFailCycle, Flag_Name, mbist_loop)      'PLUS as version 10.30.90 use function
        Else
            Call Finger_print_Org(pattern_load, RunFailCycle, Flag_Name, mbist_loop)
        End If
#Else
    Call Finger_print_Org(pattern_load, RunFailCycle, Flag_Name, mbist_loop)
#End If
End Function

Public Function ReduceBlkLen(InpStr As String, OutpStr As String)
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
End Function

Public Function Parsing_Busrt_Pattern()
''''Start, modify from T-sic, Carter, 20191106
    On Error GoTo errHandler
    
    Dim burst_pat() As String
    
    If Flag_BurstPat_INIT = False Then
        Dim i As Long
        Dim j As Long
        Dim maxcol As Long
        Dim MaxRow As Long
        Dim sheet_idx As Long
        Dim burst_idx As Long: burst_idx = 1
        Dim start_col As Integer: start_col = 2
        Dim start_row As Integer: start_row = 3
        
        Dim arr_content() As Variant
        Dim sheetnames() As String
        
        Dim Pat_sheet As Worksheet
        
        ReDim burst_pat(burst_idx)
        
        For Each Pat_sheet In Worksheets
            If Pat_sheet.name Like "Patsets_*" Then ''Patsets_CpuScan/Patsets_GfxScan/Patsets_SocScan //Patsets_*Scan
                Worksheets(Pat_sheet.name).Activate
                'Debug.Print Pat_sheet.name
                MaxRow = Worksheets(Pat_sheet.name).UsedRange.Rows.Count
                maxcol = Worksheets(Pat_sheet.name).UsedRange.Columns.Count
                arr_content = Worksheets(Pat_sheet.name).range(Cells(1, 1), Cells(MaxRow, maxcol)).value
                For i = start_row To MaxRow - 1
                    If arr_content(i, 2) <> burst_pat(burst_idx - 1) And arr_content(i, 7) Like LCase("yes") Then
                        ReDim Preserve burst_pat(burst_idx)
                        burst_pat(burst_idx) = arr_content(i, 2)
                        burst_idx = burst_idx + 1
                    ElseIf arr_content(i, 2) = "" Then
                        Exit For
                    End If
                Next i
            End If
        Next Pat_sheet
        
        For i = 0 To UBound(burst_pat)
            If burst_pat(i) <> "" Then
                If Not gl_burst_pat.Exists(burst_pat(i)) Then
                    gl_burst_pat.Add burst_pat(i), UCase(burst_pat(i))
                End If
            End If
        Next i
    End If
    
    Flag_BurstPat_INIT = True
    
Exit Function
errHandler:
        If isDebugMode Then TheExec.AddOutput "Error in the VBT Parsing_Busrt_Pattern"
        TheExec.Datalog.WriteComment "Error in the VBT Parsing_Busrt_Pattern"
        If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ATPG_offline(pattern_load As String, ResultMode As tlResultMode)
    ''Carter, 20191120
    On Error GoTo errHandler
    
    Dim pins As Variant
    Dim patt As Variant
    
    Dim PinName() As String
    Dim m_testName As String
    Dim rtnPatternNames() As String
    
    Dim offline_patallpass As Boolean
    Dim offline_pat_status As New SiteBoolean
    
    Dim NumberPins As Long
    Dim rtnPatternCount As Long
    
    Dim Core_Vmain As Double

    m_testName = TheExec.DataManager.instancename
    offline_patallpass = True
    offline_pat_status = False
    'TheHdw.DCVS.Pins("CorePower").Voltage.Output = tlDCVSVoltageMain
    'Call TheExec.DataManager.DecomposePinList("CorePower", PinName(), NumberPins)
    If InStr(pattern_load, "\") = 0 Then ''Burst Pattern, exclude walkingZ pattern
        Call PATT_GetPatListFromPatternSet(pattern_load, rtnPatternNames, rtnPatternCount)
        For Each patt In rtnPatternNames
            Call ATPG_offline_Simulation(patt, ResultMode, offline_patallpass, offline_pat_status)
        Next patt
    
    Else ''Single Pattern
        Call ATPG_offline_Simulation(pattern_load, ResultMode, offline_patallpass, offline_pat_status)
    
    End If
    
Exit Function
errHandler:
        If isDebugMode Then TheExec.AddOutput "Error in the VBT ATPG_offline"
        TheExec.Datalog.WriteComment "Error in the VBT ATPG_offline"
        If AbortTest Then Exit Function Else Resume Next
End Function


Public Function ATPG_offline_Simulation(pattern_load As Variant, ResultMode As tlResultMode, offline_patallpass As Boolean, offline_pat_status As SiteBoolean)
        
    
    If LCase(pattern_load) Like "*_in*" Then
        Call TheHdw.patterns(pattern_load).test(pfAlways, 0, ResultMode)
    Else
        offline_patallpass = True
              
        If TheExec.enableWord("Golden_Default") = False Then
            For Each site In TheExec.sites
                If offline_pat_status(site) = False Then
                    offline_pat_status(site) = IIf(Round(WorksheetFunction.Min(1, Rnd * 30), 0) = 1, True, False)
                End If
            Next site
        Else
                offline_pat_status = True
        End If
        
        For Each site In TheExec.sites
            offline_patallpass = offline_patallpass And offline_pat_status(site)
        Next site
        
        If offline_patallpass = True Then
            Call TheHdw.patterns(pattern_load).test(pfAlways, 0, ResultMode)
    
        Else
            Call TheHdw.patterns(pattern_load).test(pfNever, 0, ResultMode)
            For Each site In TheExec.sites
                If offline_pat_status(site) = False Then
                    Call TheExec.Datalog.WriteFunctionalResult(site, TheExec.sites.item(site).TestNumber, logTestFail)
    
                Else
                    Call TheExec.Datalog.WriteFunctionalResult(site, TheExec.sites.item(site).TestNumber, logTestPass)
    
                End If
            Next
        End If
    End If
    
Exit Function
errHandler:
        If isDebugMode Then TheExec.AddOutput "Error in the VBT ATPG_offline_pat"
        TheExec.Datalog.WriteComment "Error in the VBT ATPG_offline_pat"
        If AbortTest Then Exit Function Else Resume Next

End Function

Sub Sub_SetVol_toAllPins(pinAry() As String, VolAry() As Double, RampStep As Double)
    Dim i As Long
    Dim PinName As String
    Dim DropVoltage As Double
    Dim State_index As Integer
    
    State_index = MbistERT_GroupCurrentVol.Count
    Sub_GroupRamp pinAry, VolAry, State_index
    
    If UBound(VolAry) = 0 Then
        ReDim Preserve VolAry(UBound(pinAry())) As Double
        For i = 1 To UBound(VolAry)
            VolAry(i) = VolAry(0)
        Next i
    ElseIf UBound(VolAry) = UBound(pinAry) Then
        ' same count then do nothing
    Else
        TheExec.ErrorLogMessage "MBist Retention Vol Setting count is not match with Pin Count"
    End If
    
    For i = State_index To MbistERT_GroupCurrentVol.Count - 1
        PinName = Split(MbistERT_GroupPinNmame(Group_Name & i), ",")(0)
        MbistERT_OriginVol.Add Group_Name & i, lclPinListData_PreviousCorePower.pins(PinName).value
        MbistERT_TargetVol.Add Group_Name & i, MbistERT_GroupCurrentVol(Group_Name & i)
        
        DropVoltage = MbistERT_OriginVol(Group_Name & i) - MbistERT_TargetVol(Group_Name & i)
        MbistERT_DropVol.Add Group_Name & i, DropVoltage
        MbistERT_DropVol_PerSite.Add Group_Name & i, FormatNumber((DropVoltage / RampStep), 3)
    Next i
       
End Sub

Sub Sub_GroupRamp(pinAry() As String, VolAry() As Double, State_index As Integer)
    Dim Pin_index As Integer
    Dim Value_index As Integer
    Dim Group_index As Integer
    Dim PinName As String
    Dim F_diff As Boolean
    
    F_diff = False
    Group_index = State_index
    PinName = vbNullString
    
    MbistERT_GroupCurrentVol.Add Group_Name & Group_index, VolAry(0)
    MbistERT_GroupPreviousVol.Add Group_Name & Group_index, lclPinListData_PreviousCorePower.pins(pinAry(0)).value
    Group_index = Group_index + 1
    
    For Pin_index = 1 To UBound(pinAry)
        For Value_index = State_index To Group_index - 1
            If MbistERT_GroupCurrentVol(Group_Name & Value_index) = VolAry(Pin_index) And MbistERT_GroupPreviousVol(Group_Name & Value_index) = lclPinListData_PreviousCorePower.pins(pinAry(Pin_index)).value Then
                F_diff = True
                Exit For
            End If
        Next Value_index
        If Not F_diff Then
            MbistERT_GroupCurrentVol.Add Group_Name & Group_index, VolAry(Pin_index)
            MbistERT_GroupPreviousVol.Add Group_Name & Group_index, lclPinListData_PreviousCorePower.pins(pinAry(Pin_index)).value
            Group_index = Group_index + 1
        End If
        F_diff = False
    Next Pin_index
    
    For Value_index = State_index To MbistERT_GroupCurrentVol.Count - 1
        For Pin_index = 0 To UBound(pinAry)
            If MbistERT_GroupCurrentVol(Group_Name & Value_index) = VolAry(Pin_index) And MbistERT_GroupPreviousVol(Group_Name & Value_index) = lclPinListData_PreviousCorePower.pins(pinAry(Pin_index)).value Then
                PinName = IIf(PinName = "", pinAry(Pin_index), PinName & "," & pinAry(Pin_index))
            End If
        Next Pin_index
        MbistERT_GroupPinNmame.Add Group_Name & Value_index, PinName
        PinName = vbNullString
    Next Value_index
End Sub


Sub Sub_VoltageRamping(RampStep As Double, ExtraWaitTiom_forRamp As Double, RampDir As RET_RampingDir)
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
    End Select
    
    For Ramp_index = 1 To RampStep
        For Group_index = Start_index To End_index Step (RampingDir * -1)
'            If Ramp_index = RampStep - 1 Then
'                Vol_from = IIf(RampDir = RampDown, MbistERT_TargetVol(Group_Name & Group_index), MbistERT_OriginVol(Group_Name & Group_index))
'                TheHdw.DCVS.Pins(MbistERT_GroupPinNmame(Group_Name & Group_index)).Voltage.value = Vol_from
'            Else
                Vol_to = IIf(RampDir = RampDown, MbistERT_OriginVol(Group_Name & Group_index), MbistERT_TargetVol(Group_Name & Group_index))
                TheHdw.DCVS.pins(MbistERT_GroupPinNmame(Group_Name & Group_index)).Voltage.value = Vol_to - MbistERT_DropVol_PerSite(Group_Name & Group_index) * Ramp_index * RampingDir
'            End If
        Next Group_index
        TheHdw.Wait ExtraWaitTiom_forRamp / RampStep
    Next Ramp_index
    
End Sub

Sub Sub_CStrToDblAry(str As String, Ary2() As Double, Delimiter As String)
    
    Dim i As Long
    Dim StrAry() As String
    StrAry() = Split(str, Delimiter)
    ReDim Ary2(UBound(StrAry)) As Double
    For i = 0 To UBound(StrAry)
        Ary2(i) = CDbl(StrAry(i))
    Next i

End Sub

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
    Dim pins As New PinData
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
    Dim TempDigitalPin() As String
    
    m_tn_BurstIndex = 2
    AllPins = "JTAG_TDO"
    LogLimited = 255
    m_testName = TheExec.DataManager.instancename
    Call PATT_GetPatListFromPatternSet(pattern_load, rtnPatternNames, rtnPatternCount)
        
    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
    maxDepth = gl_HRAMmaxDepth                  'maxDepth = TheHdw.Digital.HRAM.maxDepth        'Flex UP1600 max depth 512      'Plus org max depth 16k, but use 512 to let speed faster
    TheHdw.Digital.hram.size = maxDepth
    TheHdw.Digital.hram.CaptureType = captFail
    
    TheHdw.Digital.hram.SetTrigger trigFail, False, 0, True
    TheHdw.Digital.Patgen.ClearFail
        
        For Each patt In rtnPatternNames
            numcap = 0
            '==================================================================='''''''''''finger print_block_01_begin
            If TheExec.enableWord("Mbist_FingerPrint") = True Then
                astrPattPathSplit = Split(CStr(patt), "\")
                strPattName = UCase(astrPattPathSplit(UBound(astrPattPathSplit)))
    '            If strPattName Like "*:*" Then
    '                astrPattPathSplit_01 = Split(strPattName, ":")
    '                strPattName = astrPattPathSplit_01(0)
    '            End If
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
                
                MatchFlag = False
                For temp = 0 To UBound(tpCycleBlockInfor)
                    If UCase(tpCycleBlockInfor(temp).strPattName) = Pattern_GenericName(0) Then
                        tpEvaPattCycleBlockInfor = tpCycleBlockInfor(temp).tpMbistCycleBlock
                        MatchFlag = True
                        MFP_pattern_idx = temp 'SWLINZa 20180907 for MFP DTR, to indentify pattern#
                        Exit For
                    End If
                Next temp
            End If
            '==================================================================='''''''''''finger print_block_01_end
            Call TheHdw.patterns(patt).test(pfAlways, 0, tlResultModeDomain)
            '//////////////////////////////////////////////////////////////////////////////////////////////////'''''''''''finger print_block_03_begin
            If RunFailCycle = True And TheExec.enableWord("Mbist_FingerPrint") = True Then

                If MatchFlag = False Then
                    'TheExec.Datalog.WriteComment ("Warning!! Pattern Name not match ")  'Wayne 241004, Remove the warning message from Kevin's request
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
                    For Each site In TheExec.sites
                        m_tn = TheExec.sites.item(site).TestNumber
                        m_tn_restore = m_tn
'                        TheExec.sites.item(site).TestNumber = m_tn * 1000 + 1
                        '############################################################################################

                        If blPatPass(site) = False Then     ''  patt fail
                            '--------------------------------------------------------------------------
                            'SWLINZA 20181128 for MFP DTR, to get "flow-condition"(form flow) and compose DTR header
                            '--------------------------------------------------------------------------
                            MFP_flow_idx = TheExec.sites.item(site).SiteVariableValue("MFP_Flow_Idx")
                            Pattern_Desc = tpCycleBlockInfor(MFP_pattern_idx).strDecsName(MFP_flow_idx)
                            Pattern_Server = tpCycleBlockInfor(MFP_pattern_idx).strServerName(MFP_flow_idx)
                            NewFmt_Printing_Header = "MemFP,1" & "," & site & "," & Pattern_Server & "," & Pattern_Desc & "," & UCase(m_testName) & "," & Pattern_GenericName(0) & ","
                            CharacterNumbers = LogLimited - Len(NewFmt_Printing_Header)
                            numcap = TheHdw.Digital.hram.CapturedCycles
                            If numcap > 0 Then
                                ReDim Pattern_Failure_Cycles(numcap - 1) As Variant
                            End If
                            If CharacterNumbers <= 0 Then
                                TheExec.Datalog.WriteComment "The length of header is over than" & LogLimited & "."
                                TheExec.Datalog.WriteComment "Please  Check the length of header which consist of instance name and pattern name."
                            End If

                            For i = 0 To UBound(Pattern_Failure_Cycles())
                                Pattern_Failure_Cycles(i) = vbNullString
                            Next i
                            CurCount_FailAry_Element = 0

                            For i = 0 To numcap - 1
                                Set PinData = TheHdw.Digital.pins(AllPins).hram.PinData(i)
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
                                
                                   For Each pins In PinData.pins
                                        Cdata = pins.value(site)
                                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                        If LCase(TheExec.DataManager.instancename) Like "*bist*" Then

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
                                                                             TheExec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicTrue
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
                                                                            TheExec.Datalog.WriteComment "The Remianing Character Numbers is not enough to output failing cycles."
                                                                            TheExec.Datalog.WriteComment "Please check the length of header which consist of instance name and pattern name."
                                                                        End If
                                                                        
                                                                    Else
                                                                        PassOrFail(site) = 1
                                                                        If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                                                            If TheExec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                                                                                 TheExec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                                                                            End If
                                                                        End If
                                                                    End If
                                                                    blJump = True
                                                                    Count = k + 1
                                                                Else
                                                                    PassOrFail(site) = 1
                                                                    If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                                                        If TheExec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                                                                             TheExec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                                                                        End If
                                                                    End If
                                                                End If
                                                                
                                                                ReduceBlkLen tpEvaPattCycleBlockInfor(k).strBlaclName, ReviseStr
                                                                LogLen = Len(ReviseStr)
                                                                If LogLen Mod LogLimited <> 0 Then PrintTimes = (LogLen \ LogLimited) + 1
                                                                For PrintIdx = 0 To PrintTimes - 1
                                                                    ReDim Preserve DecomposeLog(PrintTimes - 1)
                                                                    DecomposeLog(PrintIdx) = mid(ReviseStr, 1 + PrintIdx * LogLimited, LogLimited)
                                                                    TheExec.flow.TestLimit PassOrFail(site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & m_testName, , _
                                                                            DecomposeLog(PrintIdx) & " " & tpEvaPattCycleBlockInfor(k).lCycle, , , , , tlForceNone
                                                                Next PrintIdx
        '''                                                                TheExec.Flow.TestLimit PassOrFail(site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & m_testName, , _
        '''                                                                        tpEvaPattCycleBlockInfor(k).strBlaclName & " " & tpEvaPattCycleBlockInfor(k).lCycle, , , , , tlForceNone
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
                                                                         TheExec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicTrue
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
                                                                        TheExec.Datalog.WriteComment "The Remianing Character Numbers is not enough to output failing vectors."
                                                                        TheExec.Datalog.WriteComment "Please check the length of header which consist of instance name and pattern name."
                                                                    End If

                                                                Else
                                                                    PassOrFail(site) = 1
                                                                    tpEvaPattCycleBlockInfor(k).strFlagName = "pass"
                                                                    If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                                                        If TheExec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                                                                             TheExec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                                                                        End If
                                                                    End If
                                                                End If
                                                                blJump = True
                                                                Count = k + 1
                                                            Else
                                                                PassOrFail(site) = 1
                                                                tpEvaPattCycleBlockInfor(k).strFlagName = "pass"
                                                                If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                                                    If TheExec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                                                                         TheExec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                                                                    End If
                                                                End If
                                                            End If
                                                            
                                                            
                                                                ReduceBlkLen tpEvaPattCycleBlockInfor(k).strBlaclName, ReviseStr
                                                                LogLen = Len(ReviseStr)
                                                                If LogLen Mod LogLimited <> 0 Then PrintTimes = (LogLen \ LogLimited) + 1
                                                                For PrintIdx = 0 To PrintTimes - 1
                                                                    ReDim Preserve DecomposeLog(PrintTimes - 1)
                                                                    DecomposeLog(PrintIdx) = mid(ReviseStr, 1 + PrintIdx * LogLimited, LogLimited)
                                                                    TheExec.flow.TestLimit PassOrFail(site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & m_testName, , _
                                                                            DecomposeLog(PrintIdx) & " " & tpEvaPattCycleBlockInfor(k).lVector, , , , , tlForceNone
                                                                Next PrintIdx
        '''                                                                TheExec.Flow.TestLimit PassOrFail(site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & m_testName, , _
        '''                                                                        tpEvaPattCycleBlockInfor(k).strBlaclName & " " & tpEvaPattCycleBlockInfor(k).lVector, , , , , tlForceNone
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
                                   Next pins
                            Next i
                            
                            '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                            If k < UBound(tpEvaPattCycleBlockInfor) Then        '' in unread all info of  tpEvaPattCycleBlockInfor case
                                If numcap < maxDepth Then
                                    PassOrFail(site) = 1
                                    For k = Count To UBound(tpEvaPattCycleBlockInfor)
                                        If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                            If TheExec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                                                 TheExec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                                            End If
                                        End If
                                        
                                                ReduceBlkLen tpEvaPattCycleBlockInfor(k).strBlaclName, ReviseStr
                                                LogLen = Len(ReviseStr)
                                                If LogLen Mod LogLimited <> 0 Then PrintTimes = (LogLen \ LogLimited) + 1
                                                For PrintIdx = 0 To PrintTimes - 1
                                                    ReDim Preserve DecomposeLog(PrintTimes - 1)
                                                    DecomposeLog(PrintIdx) = mid(ReviseStr, 1 + PrintIdx * LogLimited, LogLimited)
                                                    TheExec.flow.TestLimit PassOrFail(site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & m_testName, , _
                                                            DecomposeLog(PrintIdx) & " " & tpEvaPattCycleBlockInfor(k).lCycle, , , , , tlForceNone
                                                Next PrintIdx
                                
'''                                                TheExec.Flow.TestLimit PassOrFail(site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & m_testName, , _
'''                                                        tpEvaPattCycleBlockInfor(k).strBlaclName & " " & tpEvaPattCycleBlockInfor(k).lCycle, , , , , tlForceNone
                                    Next k
                                Else
                                    '' add for HRAM is full and still have some cycles need to judge, to set all flag status = true
                                    If gl_MbistFP_Binout Then
                                        For k = 0 To UBound(tpEvaPattCycleBlockInfor)
                                            If tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then TheExec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicTrue
                                        Next k
                                    End If
                                End If
                            End If
                            '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                            If numcap >= maxDepth Then   '' HRAM is full
                                TheExec.flow.TestLimit 0, , , , , scaleNoScaling, , , "MemFP_" & m_testName, , _
                                                          "Fail_cycle_size_check", , , , , tlForceNone
                                TheExec.Datalog.WriteComment ("The number of pattern fail cycles full or exceed HRAM maxDepth: " & maxDepth)
                            Else
                                TheExec.flow.TestLimit 1, , , , , scaleNoScaling, , , "MemFP_" & m_testName, , _
                                                          "Fail_cycle_size_check", , , , , tlForceNone
                            End If
                            '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                            TheExec.sites.item(site).TestNumber = m_tn_restore + m_tn_BurstIndex
                            Count = 0
                            k = 0
                            
                            '----------------------------------------------------------------
                            ' SWLINZA 20181128, for MemFP DTR, C651/Si requests 2018/08/M
                            ' To print final DTR,
                            ' Pattern_Failure_Cycles(AryIdx) is stored in previous procedure
                            '----------------------------------------------------------------
                            Dim AryIdx As Long
                            TheExec.Datalog.WriteComment ""
                            For AryIdx = 0 To CurCount_FailAry_Element
                                    TheExec.Datalog.WriteComment NewFmt_Printing_Header & Pattern_Failure_Cycles(AryIdx)
                            Next AryIdx
                            TheExec.Datalog.WriteComment ""
                            
                        Else    ''blPatPass(Site) = True
                            If blMbistFP_Binout Then
                                For k = 0 To UBound(tpEvaPattCycleBlockInfor)
                                    If tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                        If TheExec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                                            TheExec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
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
            
            If mbist_loop Then
                '===================================================================
                For Each site In TheExec.sites
                    'testnumber(Site) = TheExec.sites.Item(Site).testnumber
                    tested(site) = False
                    blPatPass(site) = TheHdw.Digital.Patgen.PatternBurstPassed
                    '-------------------------------------------------------------------------------------------------
                    If blPatPass(site) = False Or alarmFail(site) = True Then   'pattern test fail or alarm
                        TheExec.sites.item(site).FlagState(Flag_Name) = logicTrue 'pattern test fail
                        TheExec.sites.item(site).testResult = siteFail
                        tested(site) = True
                        'TheExec.sites.Item(Site).testnumber = TheExec.sites.Item(Site).testnumber + 1
                   '-------------------------------------------------------------------------------------------------
                    Else    'blPatPass(Site) = True ; pattern test pass
                        If (tested(site) = False) Then
                            If (TheExec.sites.item(site).FlagState(Flag_Name) <> logicTrue) Then 'confirm flag is true(pattern fail)
                                TheExec.sites.item(site).FlagState(Flag_Name) = logicFalse       'pattern test pass
                            End If
                            TheExec.sites.item(site).testResult = sitePass
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
        
'        For Each patt In rtnPatternNames
'            TheExec.Flow.ForceStopOnError = True
'            If LCase(m_testName) Like "*cpumbist*" And Not LCase(patt) Like "*_flp_*" And Not LCase(patt) Like "*_efc_*" And Not LCase(patt) Like "*_pri_*" Then
'                If Not LCase(patt) Like "*mpxxx*" Then
'                    For Each site In TheExec.sites
'                        TempDigitalPin = TheHdw.Digital.FailedPins(site)
'                        If UBound(TempDigitalPin) + 1 <> 0 Then
'                            If (TheExec.sites.item(site).FlagState("F_ECPU_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_PCPU_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_GFX_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_SOC_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_SOC_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_GFX_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_PCPU_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_ECPU_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_SOC_MBIST_Alarm_Check") = logicTrue) Then
'                                TheExec.sites.item(site).FlagState("F_ECPU_MBIST_Defect_Alarm_Check") = logicTrue
'                                TheExec.sites.item(site).SortNumber = 975
'                                TheExec.sites.item(site).BinNumber = 19
'                                TheExec.sites.site(site).result = tlResultFail
'
'                            End If
'                        End If
'                                'testnumber(Site) = TheExec.sites.Item(Site).testnumber
'                    Next site
'                End If
'            ElseIf LCase(m_testName) Like "*cpumbist*" And Not LCase(patt) Like "*_flp_*" And Not LCase(patt) Like "*_efc_*" And Not LCase(patt) Like "*_pri_*" Then
'                If Not LCase(patt) Like "*mexxx*" Then
'                    For Each site In TheExec.sites
'                        TempDigitalPin = TheHdw.Digital.FailedPins(site)
'                        If UBound(TempDigitalPin) + 1 <> 0 Then
'                            If (TheExec.sites.item(site).FlagState("F_ECPU_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_PCPU_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_GFX_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_SOC_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_SOC_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_GFX_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_PCPU_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_ECPU_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_SOC_MBIST_Alarm_Check") = logicTrue) Then
'                                TheExec.sites.item(site).FlagState("F_PCPU_MBIST_Defect_Alarm_Check") = logicTrue
'                                TheExec.sites.item(site).SortNumber = 976
'                                TheExec.sites.item(site).BinNumber = 19
'                                TheExec.sites.site(site).result = tlResultFail
'
'                            End If
'                        End If
'                                'testnumber(Site) = TheExec.sites.Item(Site).testnumber
'                    Next site
'                End If
'            ElseIf LCase(m_testName) Like "*gfxmbist*" And Not LCase(patt) Like "*_flp_*" And Not LCase(patt) Like "*_efc_*" And Not LCase(patt) Like "*_pri_*" Then
'                If Not LCase(patt) Like "*_pllp_*" And Not LCase(patt) Like "*_fstp*" Then
'                    For Each site In TheExec.sites
'                        TempDigitalPin = TheHdw.Digital.FailedPins(site)
'                        If UBound(TempDigitalPin) + 1 <> 0 Then
'                            If (TheExec.sites.item(site).FlagState("F_ECPU_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_PCPU_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_GFX_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_SOC_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_SOC_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_GFX_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_PCPU_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_ECPU_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_SOC_MBIST_Alarm_Check") = logicTrue) Then
'                                TheExec.sites.item(site).FlagState("F_GFX_MBIST_Defect_Alarm_Check") = logicTrue
'                                TheExec.sites.item(site).SortNumber = 977
'                                TheExec.sites.item(site).BinNumber = 19
'                                TheExec.sites.site(site).result = tlResultFail
'
'                            End If
'                        End If
'                                'testnumber(Site) = TheExec.sites.Item(Site).testnumber
'                    Next site
'                End If
'            ElseIf LCase(m_testName) Like "*socmbist*" And Not LCase(patt) Like "*_flp_*" And Not LCase(patt) Like "*_efc_*" And Not LCase(patt) Like "*_pri_*" Then
'                    For Each site In TheExec.sites
'                        TempDigitalPin = TheHdw.Digital.FailedPins(site)
'                        If UBound(TempDigitalPin) + 1 <> 0 Then
'                            If (TheExec.sites.item(site).FlagState("F_ECPU_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_PCPU_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_GFX_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_SOC_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_SOC_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_GFX_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_PCPU_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_ECPU_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_SOC_MBIST_Alarm_Check") = logicTrue) Then
'                                TheExec.sites.item(site).FlagState("F_SOC_MBIST_Defect_Alarm_Check") = logicTrue
'                                TheExec.sites.item(site).SortNumber = 978
'                                TheExec.sites.item(site).BinNumber = 19
'                                TheExec.sites.site(site).result = tlResultFail
'
'                            End If
'                        End If
'                        'testnumber(Site) = TheExec.sites.Item(Site).testnumber
'                    Next site
'            End If
'        Next patt
        
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_MBIST", "Finger_print_Org")
    If AbortTest Then Exit Function Else Resume Next
End Function

#If IGXL_VER_1030 = True Then
' [20230425][All][Carter] use for Plus
Private Sub Finger_Print_NewSyntex(pattern_load As String, RunFailCycle As Boolean, Optional Flag_Name As String, Optional mbist_loop As Boolean = False)
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
    Dim pins As New PinData
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
    Dim CurCount_FailAry_Element As Long        'Dim CurCount_FailAry_Element As New SiteLong
    Dim NewFmt_Printing_Header As String
    Dim CharacterNumbers As Long
    Dim Pattern_GenericName() As String
    Dim s_Pattern_GenericName As String
    Dim MFP_pattern_idx As Long
    Dim MFP_flow_idx As Variant

    Dim sb_FingerPrint_Fail As New SiteBoolean

    'Dim Pattern_Failure_Cycles() As New SiteVariant
    'Dim Pattern_Failure_Cycles() As New Variant

    Dim l_patcnt As Long
    Dim l_PatIndex As Long
    Dim l_patcnt_temp As Long
    Dim l_pgVector() As Long
    Dim l_PatternSetIndex() As Long

    Dim d_pgCycle() As Double

    Dim Pattern_Ary_absolute() As String
        Dim pins() As String
        Dim Pin_Cnt As Long
        
    sb_FingerPrint_Fail = False
    m_tn_BurstIndex = 2
    AllPins = "JTAG_TDO"
    LogLimited = 255
    m_testName = TheExec.DataManager.instancename
''    Call PATT_GetPatListFromPatternSet(pattern_load, rtnPatternNames, rtnPatternCount)
    Call GetPatsFromPatSets(pattern_load, rtnPatternNames, rtnPatternCount, False)
    Call GetPatsFromPatSets(pattern_load, Pattern_Ary_absolute, rtnPatternCount, True)

    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
    ''maxDepth = TheHdw.Digital.hram.maxDepth
    TheHdw.Digital.hram.size = gl_HRAMmaxDepth
    TheHdw.Digital.hram.CaptureType = captFail

    TheHdw.Digital.hram.SetTrigger trigFail, False, 0, True
    TheHdw.Digital.Patgen.ClearFail

''    TheExec.Datalog.WriteComment "*************************************************"
''    TheExec.Datalog.WriteComment "Capture Failing Cycle Info for Test Instance " & TheExec.DataManager.instancename & " (Max HRAM Size: " & maxDepth & ")"
''    TheExec.Datalog.WriteComment "*************************************************"

    Dim Cycle As IAllCapturedCycleInfo ''20221011, New syntax
''    Dim firstCycleInfo As ICapturedCycleInfo

        TheExec.DataManager.DecomposePinList AllPins, pins(), Pin_Cnt
        If Pin_Cnt <> 1 Then
            TheExec.Datalog.WriteComment "[Warning] New Syntex just can process 1 pin 1 time!!"
                Exit Sub
        Else
        End If

    If BurstYesPatDict.Exists(LCase(pattern_load)) = True Then
    ''Pattern Burst Yes
        numcap = 0
        Call TheHdw.patterns(pattern_load).test(pfAlways, 0, tlResultModeDomain)
        blPatPass = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
        If blPatPass.Any(False) Then
            Set Cycle = TheHdw.Digital.pins(AllPins).hram.CapturedFailCycleInfo
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
                    MFP_flow_idx = TheExec.sites.item(site).SiteVariableValue("MFP_Flow_Idx")
                    Pattern_Desc = tpCycleBlockInfor(MFP_pattern_idx).strDecsName(MFP_flow_idx)
                    Pattern_Server = tpCycleBlockInfor(MFP_pattern_idx).strServerName(MFP_flow_idx)

                    For Each site In TheExec.sites
                        If blPatPass(site) = False Then
                            If Cycle.HasCaptureOccurred(site) Then
                                l_patcnt_temp = -1
                                l_PatternSetIndex = Cycle.PatternSetIndex(site)

                                d_pgCycle = Cycle.pgModuleCount(site)
                                l_pgVector = Cycle.pgVector(site)
                                ReDim Pattern_Failure_Cycles(UBound(d_pgCycle)) As New Variant

                                NewFmt_Printing_Header = "MemFP,1" & "," & site & "," & Pattern_Server & "," & Pattern_Desc & "," & glb_TestInstance & "," & s_Pattern_GenericName & ","
                                CharacterNumbers = LogLimited - Len(NewFmt_Printing_Header)
                                If CharacterNumbers <= 0 Then
                                    TheExec.Datalog.WriteComment "The length of header is over than" & LogLimited & "."
                                    TheExec.Datalog.WriteComment "Please  Check the length of header which consist of instance name and pattern name."
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
                                        If TheExec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                                            TheExec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
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
            numcap = 0
            s_Pattern_GenericName = GetPatName_FromAbsolute(CStr(patt))

            Call TheHdw.patterns(patt).test(pfAlways, 0, tlResultModeDomain)
            blPatPass = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
            If blPatPass.Any(False) Then
                Set Cycle = TheHdw.Digital.pins(AllPins).hram.CapturedFailCycleInfo
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
                MFP_flow_idx = TheExec.sites.item(site).SiteVariableValue("MFP_Flow_Idx")
                Pattern_Desc = tpCycleBlockInfor(MFP_pattern_idx).strDecsName(MFP_flow_idx)
                Pattern_Server = tpCycleBlockInfor(MFP_pattern_idx).strServerName(MFP_flow_idx)

                For Each site In TheExec.sites
                    If blPatPass(site) = False Then
                        If Cycle.HasCaptureOccurred(site) Then
                            l_PatternSetIndex = Cycle.PatternSetIndex(site)
                            d_pgCycle = Cycle.pgCycle(site)
                            l_pgVector = Cycle.pgVector(site)
                            ReDim Pattern_Failure_Cycles(UBound(d_pgCycle)) As New SiteVariant

                            NewFmt_Printing_Header = "MemFP,1" & "," & site & "," & Pattern_Server & "," & Pattern_Desc & "," & glb_TestInstance & "," & s_Pattern_GenericName & ","
                            CharacterNumbers = LogLimited - Len(NewFmt_Printing_Header)
                            If CharacterNumbers <= 0 Then
                                TheExec.Datalog.WriteComment "The length of header is over than " & LogLimited & "."
                                TheExec.Datalog.WriteComment "Please  Check the length of header which consist of instance name and pattern name."
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
                                    If TheExec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                                        TheExec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
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
        For Each site In TheExec.sites
            'testnumber(Site) = TheExec.sites.Item(Site).testnumber
            tested(site) = False
    ''                blPatPass(site) = thehdw.Digital.Patgen.PatternBurstPassed
            '-------------------------------------------------------------------------------------------------
            If blPatPass(site) = False Or alarmFail(site) = True Then   'pattern test fail or alarm
                TheExec.sites.item(site).FlagState(Flag_Name) = logicTrue 'pattern test fail
                TheExec.sites.item(site).testResult = siteFail
                tested(site) = True
                'TheExec.sites.Item(Site).testnumber = TheExec.sites.Item(Site).testnumber + 1
           '-------------------------------------------------------------------------------------------------
            Else    'blPatPass(Site) = True ; pattern test pass
                If (tested(site) = False) Then
                    If (TheExec.sites.item(site).FlagState(Flag_Name) <> logicTrue) Then 'confirm flag is true(pattern fail)
                        TheExec.sites.item(site).FlagState(Flag_Name) = logicFalse       'pattern test pass
                    End If
                    TheExec.sites.item(site).testResult = sitePass
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
#End If

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
    Dim pins As Variant
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
                Set PinData = TheHdw.Digital.pins(CapPin).hram.PinData(i)
                For Each pins In PinData.pins
                    Cdata = pins.value
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
                                                                TheExec.flow.TestLimit PassOrFail(site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & glb_TestInstance, , _
                                            DecomposeLog(PrintIdx) & " " & tpEvaPattCycleBlockInfor(k).lCycle, , , , , tlForceNone
                                Case "vector":
                                                                TheExec.flow.TestLimit PassOrFail(site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & glb_TestInstance, , _
                                            DecomposeLog(PrintIdx) & " " & tpEvaPattCycleBlockInfor(k).lVector, , , , , tlForceNone
                            End Select
                        Next PrintIdx

                        If blJump = True Then
                            blJump = False
                            Exit For
                        End If
                    Next k
                Next pins
            End If
        End If
    Next i
    
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If k < UBound(tpEvaPattCycleBlockInfor) And sb_FingerPrint_Fail Then        '' in unread all info of  tpEvaPattCycleBlockInfor case
        If UBound(in_PatternSetIndex) < maxDepth Then
            PassOrFail(site) = 1
            For k = l_Count To UBound(tpEvaPattCycleBlockInfor)
                If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                    If TheExec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                         TheExec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
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
                                                TheExec.flow.TestLimit PassOrFail(site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & glb_TestInstance, , _
                                    DecomposeLog(PrintIdx) & " " & tpEvaPattCycleBlockInfor(k).lCycle, , , , , tlForceNone
                        Case "vector":
                                                TheExec.flow.TestLimit PassOrFail(site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & glb_TestInstance, , _
                                    DecomposeLog(PrintIdx) & " " & tpEvaPattCycleBlockInfor(k).lVector, , , , , tlForceNone
                    End Select
                Next PrintIdx
            Next k
        Else
            '' add for HRAM is full and still have some cycles need to judge, to set all flag status = true
            If gl_MbistFP_Binout Then
                For k = 0 To UBound(tpEvaPattCycleBlockInfor)
                    If tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then TheExec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicTrue
                Next k
            End If
        End If
    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If sb_FingerPrint_Fail Then
        If UBound(in_PatternSetIndex) >= maxDepth Then   '' HRAM is full
            TheExec.flow.TestLimit 0, , , , , scaleNoScaling, , , "MemFP_" & glb_TestInstance, , _
                                      "Fail_cycle_size_check", , , , , tlForceNone
            TheExec.Datalog.WriteComment ("The number of pattern fail cycles full or exceed HRAM maxDepth: " & maxDepth)
        Else
            TheExec.flow.TestLimit 1, , , , , scaleNoScaling, , , "MemFP_" & glb_TestInstance, , _
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
        TheExec.Datalog.WriteComment ""
        For AryIdx = 0 To CurCount_FailAry_Element
                TheExec.Datalog.WriteComment NewFmt_Printing_Header & Pattern_Failure_Cycles(AryIdx)
        Next AryIdx
        TheExec.Datalog.WriteComment ""
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
                 TheExec.sites.item(tpEvaPattCycleBlockInfor(k).strFlagName) = logicTrue
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
                If CurCount_FailAry_Element > UBound(Pattern_Failure_Cycles) Then               'If CurCount_FailAry_Element > 50 Then
                    ReDim Preserve Pattern_Failure_Cycles(CurCount_FailAry_Element)
                End If
                If Pattern_Failure_Cycles(CurCount_FailAry_Element) = "" Then
                    Pattern_Failure_Cycles(CurCount_FailAry_Element) = CStr(d_FailCycleVector)
                Else
                    Pattern_Failure_Cycles(CurCount_FailAry_Element) = Pattern_Failure_Cycles(CurCount_FailAry_Element) & "," & CStr(d_FailCycleVector)
                End If
            Else
                TheExec.Datalog.WriteComment "The Remianing Character Numbers is not enough to output failing vectors/cycles."
                TheExec.Datalog.WriteComment "Please check the length of header which consist of instance name and pattern name."
            End If
        Else
            PassOrFail = 1
            If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                If TheExec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                     TheExec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                End If
            End If
        End If
        blJump = True
        l_Count = k + 1
    Else
        PassOrFail(site) = 1
        tpEvaPattCycleBlockInfor(k).strFlagName = "pass"
        If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
            If TheExec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                 TheExec.sites.item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
            End If
        End If
    End If
    
    Exit Function
    
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_MBIST", "Compose_HRAMCapturedFailInfo")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function MbistRetentionLevelWait_SpecifiedPinGrp(Ret_WaitTime As Double, Retention_SRAMPins As PinList, Retention_LogicPins As PinList, OriCategory As String, RampCategory As String, RampWaitTime As Double, Ramp_Step As Double)
    On Error GoTo errHandler
    Dim j As Long
    Dim k As Long
    Dim RampDown_Step As Double: RampDown_Step = Ramp_Step
    Dim RampDown_Time As Double: RampDown_Time = RampWaitTime 'RampDown_Time = 0
    
    Dim pin_type As String
    Dim Retention_RampUp_Seq As String
    Dim Retention_RampDown_Seq As String
    
    Dim Ret_SRAMPins_Ary() As String
    Dim Ret_LogicPins_Ary() As String
    
    Dim Retention_Pins_Ary() As String
    Dim Retention_Pins_count As Long

    Dim Voltage_from_HW As String
    Dim RET_Pins() As String
    Dim RET_Volts() As String
    
    'Dim Retention_Pins As String
    
    '----------------------------------------------------------------------------------------'
    '------ To check counts number between pins and vol-settings, Add Vol in Dictionary
    '------ Expand Vol to all pins, if only one voltage set
    '------ ERROR while count is not 1 but also matched
    '----------------------------------------------------------------------------------------'
    If Retention_SRAMPins <> "" Then
        Ret_SRAMPins_Ary() = Split(Retention_SRAMPins, ",")
        Call Sub_SetVol_toAllPins_By_Category(Ret_SRAMPins_Ary(), OriCategory, RampCategory, RampDown_Step)
    End If
    
    If Retention_LogicPins <> "" Then
        Ret_LogicPins_Ary() = Split(Retention_LogicPins, ",")
        Call Sub_SetVol_toAllPins_By_Category(Ret_LogicPins_Ary(), OriCategory, RampCategory, RampDown_Step)
    End If
    '----------------------------------------------------------------------------------------'
    'Due to SELSRM, we need to avoid negative current/alarm happens during ramping
    'We have to consider about Different Retention Sequence for RampUp/Down
    '----------------------------------------------------------------------------------------'
    If Retention_LogicPins = "" Then
        Retention_RampDown_Seq = Retention_SRAMPins
        Retention_RampUp_Seq = Retention_SRAMPins
    ElseIf Retention_SRAMPins = "" Then
        Retention_RampDown_Seq = Retention_LogicPins
        Retention_RampUp_Seq = Retention_LogicPins
    Else
        Retention_RampDown_Seq = Retention_LogicPins & "," & Retention_SRAMPins
        Retention_RampUp_Seq = Retention_SRAMPins & "," & Retention_LogicPins
    End If
    '------------------------------------------------'
    '--------- Ramp down for retention voltage ------'
    '------------------------------------------------'
    TheExec.DataManager.DecomposePinList Retention_RampDown_Seq, Retention_Pins_Ary(), Retention_Pins_count
    Call Sub_VoltageRamping_SpecifiedPinGrp(Retention_Pins_Ary(), RampDown_Step, RampDown_Time, RampDown)
    '--------------------------------------------------------------------'
    '----------- Wait time Procedure/ Chcking HW setting ----------------'
    '--------------------------------------------------------------------'
    Voltage_from_HW = ""
    '--------- Read back retention voltage from HW ------'
    For j = 0 To Retention_Pins_count - 1
        If UCase(TheExec.DataManager.ChannelType(Retention_Pins_Ary(j))) <> "N/C" Then
    '                Retention_Pins_Ary(j) = GetUsedPinNameForDCVI(Retention_Pins_Ary(j))
            pin_type = UCase(GetInstrument(Retention_Pins_Ary(j), 0))
            Select Case (pin_type)
                Case "DC-07":
                    If j = 0 Then
                        Voltage_from_HW = CStr(FormatNumber(TheHdw.DCVI.pins(Retention_Pins_Ary(j)).Voltage.value, 3)) & " V"
                    Else
                        Voltage_from_HW = Voltage_from_HW & ", " & CStr(FormatNumber(TheHdw.DCVI.pins(Retention_Pins_Ary(j)).Voltage.value, 3)) & " V"
                    End If
                Case Else:
                    If j = 0 Then
                        Voltage_from_HW = CStr(FormatNumber(TheHdw.DCVS.pins(Retention_Pins_Ary(j)).Voltage.value, 3)) & " V"
                    Else
                        Voltage_from_HW = Voltage_from_HW & ", " & CStr(FormatNumber(TheHdw.DCVS.pins(Retention_Pins_Ary(j)).Voltage.value, 3)) & " V"
                    End If
            End Select
        End If
    Next j
        
    RET_Pins() = Split(Retention_RampDown_Seq, ",")
    RET_Volts() = Split(Voltage_from_HW, ",")
    
    '----- Retention Wait time 100 ms ------
    TheHdw.Wait Ret_WaitTime * 0.001
    TheExec.flow.TestLimit Ret_WaitTime, PinName:="Wait_Time", unit:=unitCustom, customUnit:="mSec"
    TheExec.Datalog.WriteComment "*************************************************"
    For k = 0 To UBound(RET_Pins)
        TheExec.Datalog.WriteComment "*print: MbistRetention Pins " & RET_Pins(k) & " " & RET_Volts(k)
    Next k
    TheExec.Datalog.WriteComment "*print: MbistRetention wait " & Ret_WaitTime & " ms*"
    
    '        TheExec.Datalog.WriteComment "*print: MbistRetention Volt " & Voltage_from_HW
    TheExec.Datalog.WriteComment "*************************************************"
    DebugPrintFunc ""
    
    '------------------------------------------------'
    '--------- Ramp Up for retention voltage --------'
    '------------------------------------------------'
    TheExec.DataManager.DecomposePinList Retention_RampUp_Seq, Retention_Pins_Ary(), Retention_Pins_count
    Call Sub_VoltageRamping_SpecifiedPinGrp(Retention_Pins_Ary(), RampDown_Step, RampDown_Time, RampUp)

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_MBIST", "MbistRetentionLevelWait_SpecifiedPinGrp")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Sub_SetVol_toAllPins_By_Category(pinAry() As String, OriSpec As String, TargetSpec As String, RampStep As Double)
    Dim i As Long
    Dim PinName As String
    Dim DropVoltage As Double
    Dim Spec_Var As String
    
    
    For i = 0 To UBound(pinAry)
        PinName = pinAry(i)
        Spec_Var = PinName + "_VAR"
        If TheExec.DataManager.ChannelType(PinName) <> "N/C" Then
            MbistERT_OriginVol.Add PinName, FormatNumber(TheExec.Specs.DC.item(Spec_Var).Categories.item(OriSpec).Typ.value, 3)
            MbistERT_TargetVol.Add PinName, FormatNumber(TheExec.Specs.DC.item(Spec_Var).Categories.item(TargetSpec).Min.value, 3)
    
            DropVoltage = MbistERT_OriginVol(PinName) - MbistERT_TargetVol(PinName)
            MbistERT_DropVol.Add PinName, DropVoltage
            MbistERT_DropVol_PerSite.Add PinName, FormatNumber((DropVoltage / RampStep), 3)
        End If
    Next i

End Function

Public Function Sub_VoltageRamping_SpecifiedPinGrp(pinAry() As String, RampStep As Double, ExtraWaitTiom_forRamp As Double, RampDir As RET_RampingDir)
    Dim i, j As Integer
    Dim RampingDir As Double
    Dim Vol_from As Double
    Dim Vol_to As Double
    Dim pin_type As String
    
    Select Case RampDir
        Case RampDown:
            RampingDir = 1
        Case RampUp:
            RampingDir = -1
    End Select
    
    For i = 0 To RampStep - 1
        For j = 0 To UBound(pinAry)
            pin_type = UCase(GetInstrument(pinAry(j), 0))
            If i = RampStep - 1 Then
                Vol_from = IIf(RampDir = RampDown, MbistERT_TargetVol(pinAry(j)), MbistERT_OriginVol(pinAry(j)))
                Select Case (pin_type)
                    Case "DC-07":
                        TheHdw.DCVI.pins(pinAry(j)).Voltage.value = Vol_from
                    Case Else:
                        TheHdw.DCVS.pins(pinAry(j)).Voltage.value = Vol_from
                End Select
            Else
                Vol_to = IIf(RampDir = RampDown, MbistERT_OriginVol(pinAry(j)), MbistERT_TargetVol(pinAry(j)))
                Select Case (pin_type)
                    Case "DC-07":
                        TheHdw.DCVI.pins(pinAry(j)).Voltage.value = Vol_to - MbistERT_DropVol_PerSite(pinAry(j)) * i * RampingDir
                    Case Else:
                        TheHdw.DCVS.pins(pinAry(j)).Voltage.value = Vol_to - MbistERT_DropVol_PerSite(pinAry(j)) * i * RampingDir
                End Select
            End If
            'TheExec.Datalog.WriteComment "Pin:" & PinAry(j) & ",Vol: " & FormatNumber(TheHdw.DCVS.Pins(PinAry(j)).Voltage.Value, 3)
        Next j
        'TheExec.Datalog.WriteComment ""
        TheHdw.Wait ExtraWaitTiom_forRamp / RampStep
    Next i
    
End Function
