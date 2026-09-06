Attribute VB_Name = "VBT_LIB_Digital_Shmoo"
Option Explicit
'Revision History:
'V0.0 initial bring up

Public Interpose_PrePat_GLB As String
Public ReadHWPowerValue_GLB As String
Public PL_DC_conditions_GLB As New SiteVariant
Public Vbump_for_Interpose As Boolean
Public InitialPatCnt As Long
Public PayloadPatCnt As Long
Public PatAmount As Integer
Public gl_BY_PASS_SHMOO_HOLE As Boolean

' ============
' Private Data
' ============

' Context values on the Test Instances sheet
Private m_TimeSetSheet As String, m_LevelsSheet As String

' States of driver features which are saved and restored
Private m_OldPatThreading As Boolean
Private m_OldFlagMatchEnable As Boolean
Private m_OldWaitFlagsHigh As Long
Private m_OldWaitFlagsLow As Long
Private m_OldMatchAllSites As Boolean

' Cached parameters for PostTest POSTPATBPF interpose function. This
' is needed for the pattern set breakpoint feature.
Private m_DrivePins As String
Private m_FloatPins As String
Private m_EndOfBodyF As String
Private m_EndOfBodyFArgs As String

Private m_InterposeFunctionsSet As Boolean

Private Const TL_E_AT_PATSET_BREAKPT = &HC0000014
Public g_INIT_PAT_DONE As Boolean
Public Sweep_cnt As Integer

'Public Type Tracking_Axis_Para
'
'End Type
Public Type Axis_Para
    X_axis As New SiteDouble ' only store X_axis info
    Y_axis As New SiteDouble ' only store Y_axis info
    Z_axis As New SiteDouble ' only store Z_axis info
    X_axis_Tracking() As New SiteDouble  ' only store X_axis Tracking info
    Y_axis_Tracking() As New SiteDouble ' only store Y_axis Tracking info
    Z_axis_Tracking() As New SiteDouble ' only store Z_axis Tracking info
    CurrResult As New SiteVariant
End Type

Public Type g_3DShmooCurrPointResult
    Axis_CurrPoint() As Axis_Para
End Type

Public g_ShmooResult As g_3DShmooCurrPointResult

Public g_TestNum As Long
Public Xaxis_index As Long
Public Yaxis_index As Long
Public Zaxis_index As Long
'Public Count_Point As Long
Public MaxArrIndex As Long

Public X_Tracking_Point As Long
Public Y_Tracking_Point As Long
Public Z_Tracking_Point As Long
Public X_dimemsion As Boolean
Public Y_dimemsion As Boolean
Public Z_dimemsion As Boolean
Public Multi_Axis_PTR As Boolean
Public TPModeAsCharz_GLB As Boolean

Public X_Point As Long
Public Y_Point As Long
Public Z_Point As Long
Public LVCC_flag As Boolean
Public HVCC_flag As Boolean
Public StartPoint As Double
Public StopPoint As Double
Public StepSize As Double
Public RangeSeq(2)  As Boolean '0:X-axis, 1:Y-axis, 2:Z-axis

Public g_ShmooPin As New PinListData
Public Type TypePatternInfo
    testType As String
    Pattern As New Pattern
    Sequence As Long
    IsInitPattern As Boolean
    DynamicSourceBit As String
    WaitTime As Double
    RampStep As Integer
    RamppingTime As Double
    RampOffset As Double
    RampOffsetSymbol As String
    DigSrc_BitSize As String
    DigSrc_Seg As String
    DigSrc_pin As String
    digSrc_EQ As String
    SegDict As New Dictionary
    DictApplyVol As New Dictionary
    'EqDict As New Dictionary
    PowerRunCond As String
    SelSramMatchIdx As Long
    DigSrcType As String
    PatternLoopCnt As Long
    ForceVoltage As New PinListData
    GuardBandSymbol As String
    GuardBandVal As Double
End Type

Public g_CharPattInfoAry() As TypePatternInfo
Public g_MergeVDD As String
Public g_MergeCond As New SiteVariant
Public g_AllDCVSPin As String
'Public g_BitsDef As String
Public Const g_BitsDef = "VDD_DISP,VDD_AVE,VDD_GPU,VDD_ECPU,VDD_PCPU,VDD_DCS_DDR,VDD_SOC"


''//////// Put before the function call/////////
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
Public g_PR_Scenario As String
Public g_Retention_Info As String
Public g_instSSNinfo As Inst_SSN
Public CZ_inst_info As Instance_Info ''20240521 MFSTP
Public inst_info As Instance_Info ''20240521 apply BV
Public gb_ApplyBV As Boolean 'apply BV
Dim PrintString As String
'' //////////////////////////////////////////////

'Public axis_val() As New SiteVariant 'XYZ
'Public axis_pin() As String 'XYZ

' ===============
' Private Helpers
' ===============

' This template needs to know timing and levels sheet names.
' Fetch them from the Context Manager
Private Sub FetchContext()
On Error GoTo errHandler
    Dim a(0 To 4) As String

    ' For compatibility with 7.01.01 and earlier:
    ' In earlier versions, a contextmgr bug made using a MemberIndex > 0 act like the CurrentlyAppliedContext parameter was False.
    ' This caused "" to be returned for the output parameters...so that ApplyLevelsTiming was NOT called for 2nd & later members of a test group
    
    Dim MemberIndex As Long
    MemberIndex = TheExec.DataManager.MemberIndex
    
    Dim UseCurrentContext As Boolean
    UseCurrentContext = (MemberIndex = 0)
    
    Call m_STDSvcClient.Dmgr.ContextMgr.GetInstanceContextInformation(TheExec.DataManager.instancename, MemberIndex, _
                a(0), a(1), m_TimeSetSheet, a(2), a(3), a(4), m_LevelsSheet, True, UseCurrentContext)

    Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Shmoo", "FetchContext")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

' Restore tester state to the default
Private Sub PostBody(DrivePins As String, FloatPins As String, WaitTimeDomain As String, WaitFlagA As tlWaitVal, _
                    WaitFlagB As tlWaitVal, WaitFlagC As tlWaitVal, WaitFlagD As tlWaitVal)
On Error GoTo errHandler

    If TheExec.Flow.IsRunning = False Then Exit Sub
    
    ' Clear previously registered interpose function names
    Call tl_ClearInterpose(TL_C_PREPATF, TL_C_POSTPATF, TL_C_PRETESTF, TL_C_POSTTESTF, TL_C_POSTPATBPF)
    m_InterposeFunctionsSet = False

    ' Return channels to the default start-state condition, as needed
    If NonBlank(DrivePins) Then Call tl_SetStartState(DrivePins, chstartNone)

    ' Return specified DUT pins, if any, to connection with tester pin-electronics & power
    If NonBlank(FloatPins) Then Call tl_ConnectTester(FloatPins)
    
    ' Restore flag match feature
    ' for compatibility, the flag set/restore should be conditional if asynchronous pattern start not disabled and not suspended
    If ((TheHdw.Patterns.EnableAsyncPatternStart <> tlAsyncPatternModeDisabled) And (TheHdw.Patterns.SuspendAsyncPatternStart = False)) Then
       ' If the flag match settings are defaults then should not call SetFlagMatch
       If ((WaitFlagA <> waitoff) And (WaitFlagB <> waitoff) And (WaitFlagC <> waitoff) And (WaitFlagD <> waitoff)) Then
          Call TheHdw.Digital.TimeDomains(WaitTimeDomain).Patgen.SetFlagMatch( _
                   m_OldFlagMatchEnable, m_OldWaitFlagsHigh, m_OldWaitFlagsLow, _
                   m_OldMatchAllSites)
       End If
    Else
        Call TheHdw.Digital.TimeDomains(WaitTimeDomain).Patgen.SetFlagMatch( _
         m_OldFlagMatchEnable, m_OldWaitFlagsHigh, m_OldWaitFlagsLow, _
                 m_OldMatchAllSites)
    End If
    ' Restore pattern threading
    TheHdw.Patterns().Threading.Enable = m_OldPatThreading
    Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Shmoo", "PostBody")
    If AbortTest Then Exit Sub Else Resume Next
End Sub


' Run the pattern and see if it passed or failed
Private Sub Body(FloatPins As PinList, PatternTimeout As Double, Patterns As Pattern, _
                 ReportResult As PFType, ResultMode As tlResultMode)
On Error GoTo errHandler
    ' Remove specified DUT pins, if any, from connection to tester pin-electronics and other resources
    If NonBlank(FloatPins) Then Call tl_SetFloatState(FloatPins)
    m_FloatPins = FloatPins.value
    
    ' Enable the pattern timeout counter
    TheHdw.Digital.Patgen.TimeoutEnable = True
    TheHdw.Digital.Patgen.TimeOut = PatternTimeout
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Shmoo", "Body")
    If AbortTest Then Exit Sub Else Resume Next
End Sub


' =====================
' Private work routines
' =====================

' Do test setup.  This involves setting the timing and levels, connecting the pins, and
' various other functions in preparation for running the pattern.
Private Sub PreBody(DriveHiPins As PinList, DriveLoPins As PinList, DriveZPins As PinList, DisablePins As PinList, _
                    Util1Pins As PinList, Util0Pins As PinList, WaitFlagA As tlWaitVal, _
                    WaitFlagB As tlWaitVal, WaitFlagC As tlWaitVal, WaitFlagD As tlWaitVal, MatchAllSites As Boolean, _
                    PatThreading As Boolean, RelayMode As tlRelayMode, _
                    WaitTimeDomain As String, Interpose_PrePat As String)
On Error GoTo errHandler

    Dim ConnectAllPins As Boolean, LoadLevels As Boolean, LoadTiming As Boolean

    ' Save previous state of pattern threading and set according to parameter.
    m_OldPatThreading = TheHdw.Patterns().Threading.Enable
    TheHdw.Patterns().Threading.Enable = PatThreading

    ' Set drive state on specified utility pins
    If NonBlank(Util0Pins) Then Call tl_SetUtilState(Util0Pins, 0)
    If NonBlank(Util1Pins) Then Call tl_SetUtilState(Util1Pins, 1)
    
    
    ' Instruct functional voltages/currents hardware drivers to acquire
    '   drive/receive values from the DataManager and apply them.
    If NonBlank(m_LevelsSheet) Then LoadLevels = True
    
    ' Instruct functional timing hardware drivers to acquire timing values
    '   from the DataManager and apply them.
    If NonBlank(m_TimeSetSheet) Then LoadTiming = True
    
    ' Close Pin-Electronics, High-Voltage, & Power Supply Relays,
    '   of pins noted on the active levels sheet, if needed
    ConnectAllPins = True
    If (RelayMode <> TL_C_RELAYPOWERED) Then
        LoadLevels = True   'If levels are powered down, they must be powered up again
    End If
        
    ' ApplyLevelTiming will
    '   Optionally power down instruments and power supplies
    '   Optionally Close Pin-Electronics, High-Voltage, & Power Supply Relays,
    '       of pins noted on the active levels sheet
    '   Optionally load Timing and Levels information
    '   Set init-state driver conditions on specified pins
    '       Setting init state causes the pin to drive the specified value.  Init
    '       state is set once, during the prebody, before the first pattern burst.
    '       Default is to leave the pin driving whatever value it last drove during
    '       the previous pattern burst.

    '     thehdw.DCVS.pins("AllUvsCP,VDD_CPU").Alarm(tlDCVSAlarmAll) = tlAlarmOff
    Call TheHdw.Digital.ApplyLevelsTiming(ConnectAllPins, LoadLevels, LoadTiming, RelayMode)
    
    
      '' 20150625 - Apply Char setup
'    If UCase(TheExec.CurrentJob) Like "*CHAR*" Then
'        If Interpose_PrePat <> "" Then
'            Call SetForceCondition(Interpose_PrePat)
'        End If
'    End If

    
''    Call StartSBClock(24000000)
''    Call ReStartFRC
    'add wait time here
    'Call thehdw.Wait(5 * 0.001)
    'theexec.Datalog.WriteComment ("add 5ms wait time for level switch")
    'end add wait time
    
    'thehdw.DCVS.pins("AllUvsCP,VDD_CPU").Alarm(tlDCVSAlarmAll) = tlAlarmOff
    If NonBlank(DriveLoPins) Then Call tl_SetInitState(DriveLoPins, chInitLo)
    If NonBlank(DriveHiPins) Then Call tl_SetInitState(DriveHiPins, chInitHi)
    If NonBlank(DriveZPins) Then Call tl_SetInitState(DriveZPins, chInitoff)
    
    If NonBlank(DisablePins) Then Call tl_SetDisableState(DisablePins)
    
    ' Set start-state driver conditions on specified pins.
    ' Start state determines the driver value the pin is set to as each pattern burst starts.
    ' Default is to have start state automatically selected appropriately
    '   depending on the Format of the first vector of each pattern burst.
    If NonBlank(DriveLoPins) Then Call tl_SetStartState(DriveLoPins, chStartLo)
    If NonBlank(DriveHiPins) Then Call tl_SetStartState(DriveHiPins, chStartHi)
    If NonBlank(DriveZPins) Then Call tl_SetStartState(DriveZPins, chStartOff)
    m_DrivePins = tl_tm_CombineCslStrings(DriveHiPins, DriveLoPins)
    m_DrivePins = tl_tm_CombineCslStrings(DriveZPins, m_DrivePins)
    
    ' Read back state of flag feature for later restoration
    ' for compatibility, the flag set/restore should be conditional if asynchronous pattern start not disabled and not suspended
    If ((TheHdw.Patterns.EnableAsyncPatternStart <> tlAsyncPatternModeDisabled) And (TheHdw.Patterns.SuspendAsyncPatternStart = False)) Then
       ' If the flag match settings are defaults then should not call GetFlagMatch
       If ((WaitFlagA <> waitoff) And (WaitFlagB <> waitoff) And (WaitFlagC <> waitoff) And (WaitFlagD <> waitoff)) Then
          Call TheHdw.Digital.TimeDomains(WaitTimeDomain).Patgen.GetFlagMatch( _
                   m_OldFlagMatchEnable, m_OldWaitFlagsHigh, m_OldWaitFlagsLow, _
                   m_OldMatchAllSites)
       End If
    Else
        Call TheHdw.Digital.TimeDomains(WaitTimeDomain).Patgen.GetFlagMatch( _
                 m_OldFlagMatchEnable, m_OldWaitFlagsHigh, m_OldWaitFlagsLow, _
         m_OldMatchAllSites)
    End If

    ' Set desired state according to arguments.
    Call SetFlagMatch(WaitFlagA, WaitFlagB, WaitFlagC, WaitFlagD, _
        MatchAllSites, WaitTimeDomain)
    Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Shmoo", "PreBody")
    If AbortTest Then Exit Sub Else Resume Next
End Sub



Public Function TPmode_Char_on()
On Error GoTo errHandler
TPModeAsCharz_GLB = True
Multi_Axis_PTR = False

''======char shmoo error code initial=====''

F_shmoo_abnormal_counter = True ''default turn on shmoo_abnormal_counter
shmoohole_count = 0
shmooallfail_count = 0
shmooalarm_count = 0
total_shmoo_count = 0
included_shmoo_count = 0
excluded_shmoo_count = 0
        
        
 Parse_EMA_DigSrcInfo
        
  ''======char shmoo error code initial=====''
        
    TheExec.Datalog.Setup.Shared.ascii.Columns.EnableCustomWidths = True
    TheExec.Datalog.Setup.Shared.ascii.Columns.Parametric.TestName.Width = 75
    TheExec.Datalog.Setup.Shared.ascii.Columns.Functional.TestName.Width = 110
    TheExec.Datalog.Setup.Shared.ascii.Columns.Functional.Pattern.Width = 70

    TheExec.Datalog.ApplySetup  'must need to apply after datalog setup
        
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Shmoo", "TPmode_Char_on")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function TPmode_Char_off()
On Error GoTo errHandler
    TPModeAsCharz_GLB = False
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Shmoo", "TPmode_Char_off")
    If AbortTest Then Exit Function Else Resume Next
End Function



Private Function Validate_Char(PatternString As String, PatThreading As Boolean, _
                          DriveLoPins As PinList, DriveHiPins As PinList, _
                          DriveZPins As PinList, DisablePins As PinList, FloatPins As PinList, _
                          Util1Pins As PinList, Util0Pins As PinList, _
                          PatternTimeout As String, WaitTimeDomain As String) As Boolean
On Error GoTo errHandler
    Dim Patterns As New Pattern
    Dim PatternStringArr() As String
    Dim Pat As Variant
    Dim patArr() As String
    
    PatternStringArr = Split(PatternString, ",")
    
    Validate_Char = True ' Assume the best and override if trouble is found
    
    For Each Pat In PatternStringArr
        If Pat <> "" Then
           If InStr(CStr(Pat), ":") > 0 Then
              patArr = Split(Pat, ":")
              Patterns.value = patArr(0)
           Else
              Patterns.value = Pat
           End If
           If Not ValidatePatternThreading(Patterns, PatThreading, 1, True, 26) Then Validate_Char = False
              If Validate_Char Then Call PrLoadPattern(Patterns.value)
        Else
        End If
    Next Pat
    
    ' Validate the pin state parameters.
    If Not ValidatePinStates(DriveLoPins, DriveHiPins, DriveZPins, DisablePins, _
                             FloatPins, Util1Pins, Util0Pins) Then Validate_Char = False
        
    If ValidateNumeric(PatternTimeout, "PatternTimeout", 33) Then
        ' Validate  0.0 <= PatternTimeout
        If Not ValidateInRange(StrToDbl(PatternTimeout), "PatternTimeout", 0#, , , , 33) Then Validate_Char = False
    Else
        Validate_Char = False
    End If
    
    'validate timedomain
    If Not ValidateTimeDomain(WaitTimeDomain, "WaitTimeDomain", 34) Then Validate_Char = False
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Shmoo", "Validate_Char")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function freerunclk_set_XY(argc As Integer, argv() As String) As Long
On Error GoTo errHandler
    ''argv(0) is shmoo axis
    ''argv(1) is FRC nWire port name
    ''argv(2) is FRC AC spec symbol
    
    '' 20151029 - Add tracking info for second nWire-FRC
    '' argv(3) is tracking FRC nWire port name >> Clock_Port1
    '' argv(4) is tracking FRC AC spec symbol >> XI0_Freq_1_S
    Dim site As Variant
    Dim suspended As Boolean
    Dim specval As Long
    Dim pointval As Double
    Dim pointss As Long
    Dim axis_type As tlDevCharShmooAxis
    
    
    '' 20151029 - If Shmoo XI0_Freq and tracking condition as XI0_Freq_1 that should be vary the XI0_Freq_1 by using nWire-FRC
    Dim TrackingStepName() As String
    Dim Trackingcount As Integer
    Dim TeackingPointVal() As Double
    Dim b_IsTracking As Boolean
    
    If argc > 3 Then
        b_IsTracking = True
    Else
        b_IsTracking = False
    End If
    
    Select Case UCase(argv(0))
        Case "X":
            axis_type = tlDevCharShmooAxis_X
            
        Case "Y":
            axis_type = tlDevCharShmooAxis_Y
            
        Case "Z":
            axis_type = tlDevCharShmooAxis_Z
            
        Case Else:
            axis_type = tlDevCharShmooAxis_Invalid
    End Select
    
    suspended = TheExec.Datalog.DatalogSuspended
    
    TheExec.Datalog.DatalogSuspended = False
    
    For Each site In TheExec.sites
        pointval = TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(axis_type).value
        
        '' 20151029 - Get Tracking point value
        If b_IsTracking = True Then
            TrackingStepName = TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(axis_type).TrackingParameters.list
            Trackingcount = TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(axis_type).TrackingParameters.Count
            Dim i As Double
            ReDim TeackingPointVal(Trackingcount)
            For i = 1 To Trackingcount
                TeackingPointVal(i) = TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(axis_type).TrackingParameters(TrackingStepName(i - 1)).value
            Next i
        End If
    Next site
    
    Dim FRC_PortName As String
    Dim FRC_ACSpec As String
    FRC_PortName = argv(1)
    FRC_ACSpec = argv(2)
    
    Call VaryFreq(FRC_PortName, pointval, FRC_ACSpec)
    
    TheHdw.Wait 0.005
    
    Dim Track_FRC_PortName As String
    Dim Track_FRC_ACSpec As String
        If b_IsTracking = True Then
            Dim j As Double
            Dim k As Double
            k = 1
            For j = 3 To argc - 1 Step 2
                Track_FRC_PortName = argv(j)
                Track_FRC_ACSpec = argv(j + 1)
                    Call VaryFreq(Track_FRC_PortName, TeackingPointVal(k), Track_FRC_ACSpec)
                TheHdw.Wait 0.005
                k = k + 1
            Next j
        End If
     
    TheExec.Datalog.DatalogSuspended = False
    
'    g_RTOSNwireChar = True

    Exit Function
errHandler:
    If isDebugMode = True Then Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Shmoo", "freerunclk_set_XY")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function freerunclk_stop(argc As Integer, argv() As String) As Long
On Error GoTo errHandler
    ''argv(0) is FRC nWire port name
    ''argv(1) is tracking FRC nWire port name >> Clock_Port1
    
    Dim FRC_PortName As String
    Dim site As Variant
    
    '' 20151029 - Stop tracking nWireFRC
    Dim b_IsTracking As Boolean
    Dim Track_FRC_PortName() As String
    Dim FRC_PortName_1 As String
    Dim FRC_PortName_2 As String
    If argc > 1 Then
        b_IsTracking = True
    Else
        b_IsTracking = False
    End If
    
    FRC_PortName = argv(0)
'    FRC_PortName = Replace(FRC_PortName, "+", ",")
    
    '' 20151029 - Stop tracking nWireFRC
    If b_IsTracking = True Then
        Dim i As Double
        ReDim Track_FRC_PortName(argc - 1)
        For i = 1 To argc - 1
            Track_FRC_PortName(i) = argv(i)
        Next i
    End If
    
    If LCase(glb_TesterType) = "jaguar" Then
        For Each site In TheExec.sites.Active
            If TheHdw.Protocol.ports(FRC_PortName).Enabled = True Then
                TheHdw.Protocol.ports(FRC_PortName).Halt
               ' TheHdw.Protocol.Ports(FRC_PortName).Enabled = False   ' marked for shmoo XI0 at PA mode
            End If
            If b_IsTracking = True Then
                For i = 1 To argc - 1
                    If TheHdw.Protocol.ports(Track_FRC_PortName(i)).Enabled = True Then
                        TheHdw.Protocol.ports(Track_FRC_PortName(i)).Halt
                    End If
                Next i
            End If
        Next site
    Else
        Dim FreeRunName As String
        FreeRunName = Replace(LCase(FRC_PortName), "_port", vbNullString)
        If TheHdw.Digital.Pins(FreeRunName).FreeRunningClock.IsRunning Then
            TheHdw.Digital.Pins(FreeRunName).FreeRunningClock.stop '' Stop the clock.
            TheHdw.Digital.Pins(FreeRunName).FreeRunningClock.Enabled = False  ''Disable the clock (optional).
        End If
    End If

    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Shmoo", "freerunclk_stop")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function CharStoreResultsUntilNextRun()
On Error GoTo errHandler
    TheExec.DevChar.Configuration.Features.item(tlDevCharFeature_StoreResultsUntilNextRun).Enabled = False
    m_STDSvcClient.SelfTest.MemoryCollectRunInterval = 1
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Shmoo", "CharStoreResultsUntilNextRun")
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function run_shmoo(shmoo_setup As String)
On Error GoTo errHandler
    If TheExec.DevChar.Setups.IsRunning = True Then Exit Function
        With TheExec.DevChar.Setups(shmoo_setup)
            .SaveState ("current")
            .Execute False
            .RestoreState ("current")
        End With
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Shmoo", "run_shmoo")
    If AbortTest Then Exit Function Else Resume Next
End Function
' ==============
' Public Methods
' ==============
' Perform a digital functional test.
' Return TL_SUCCESS if the test executes without problems, else TL_ERROR.
'201612 Add DigSrc Arguments
Public Function Functional_T_char(StartOfBodyF As InterposeName, _
                             PrePatF As InterposeName, PreTestF As InterposeName, _
                             PostTestF As InterposeName, PostPatF As InterposeName, EndOfBodyF As InterposeName, _
                             ReportResult As PFType, ResultMode As tlResultMode, DriveLoPins As PinList, DriveHiPins As PinList, _
                             DriveZPins As PinList, DisablePins As PinList, FloatPins As PinList, StartOfBodyFArgs As String, _
                             PrePatFArgs As String, PreTestFArgs As String, PostTestFArgs As String, _
                             PostPatFArgs As String, EndOfBodyFArgs As String, Util1Pins As PinList, _
                             Util0Pins As PinList, PatFlagF As InterposeName, _
                             PatFlagFArgs As String, Validating_ As Boolean, _
                             Optional PatternTimeout As String = "30", Optional Step_ As SubType, _
                             Optional WaitTimeDomain As String, _
                             Optional ConcurrentMode As tlPatConcurrentMode = tlPatConcurrentModeCached, _
                             Optional Interpose_PrePat As String, Optional INIT_Patset As Pattern, Optional PL_Patset As Pattern, _
                             Optional Power_Run_Scenario As String, Optional Wait As String, Optional Ret_Ramp_Setting As String, _
                             Optional DigSrc_BitSize As String, Optional DigSrc_Seg As String, Optional DigSrc_DigSrcPin As String, Optional digSrc_EQ As String, Optional Order_LSB As Boolean = False, _
                             Optional BlockType As String, Optional SELSRAM_DSSC As String, Optional pmode As String, Optional One_Time_INIT As Boolean = False, _
                             Optional BypassShmooHole As Boolean = False, Optional Harvest_Header As String, Optional RetentionMeasVIF As String = vbNullString, Optional MeasureAutoRange As Boolean = False, Optional Harv_FailFlag As String, Optional SSN_SpecifyFlag As String, Optional SSN_EnableCoreMask As Boolean = True, _
                             Optional UserFunction As String, Optional ApplyVoltageFromBinCut As String = vbNullString) As Long

' EDITFORMAT1 1,,Pattern,,,Patterns|7,,InterposeName,Interpose Functions,,StartOfBodyF|9,,InterposeName,,,PrePatF|11,,InterposeName,,,PreTestF|13,,InterposeName,,,PostTestF|15,,InterposeName,,,PostPatF|17,,InterposeName,,,EndOfBodyF|2,,PFType,,,ReportResult|6,,tlResultMode,,,ResultMode|19,,pinlist,Pin States,,DriveLoPins|20,,pinlist,,,DriveHiPins|21,,pinlist,,,DriveZPins|22,,pinlist,,,DisablePins|23,,pinlist,,,FloatPins|8,,String,,,StartOfBodyFArgs|10,,String,,,PrePatFArgs|12,,String,,,PreTestFArgs|14,,String,,,PostTestFArgs|16,,String,,,PostPatFArgs|18,,String,,,EndOfBodyFArgs|24,,pinlist,,,Util1Pins|25,,pinlist,,,Util0Pins|31,,InterposeName,,,PatFlagF|32,,String,,,PatFlagFArgs|5,,tlRelayMode,,,RelayMode|3,,Boolean,,,PatThreading|30,,Boolean,,,MatchAllSites|26,,tlWaitVal,Flag Match,,WaitFlagA|27,,tlWaitVal,,,WaitFlagB|28,,tlWaitVal,,,WaitFlagC|29,,tlWaitVal,,,WaitFlagD|0,,Boolean,,,Validating_|4,,String,,0 <= PatternTimeout,PatternTimeout|6,,tlPatStartConcurrentMode,,,ConcurrentMode

    '-------------------------------------------------------------------------------------------
    'argument format instruction
    'Interpose_PrePat(Force Condition),support  "V" - force voltage
    '                                           "VRS" - PL pat run by context NV condition(INIT_NV)
    '                                           "USL" / "LSL" - LVCC limit
    '                                   ex: VDD_ECPU:V:0.7;VDD_FABRIC:V:0.75;VDD_SOC:V:0.5
    'Power_Run_Scenario                 ex: INIT_NV_PL_SWEEP
    'Pmode                              ex: Bincut_X_X_X:NV
    'Wait                               ex: PL02:0.02:,PL03:0.01:+0.025
    'Ret_Ramp_Setting                   ex: PL1:5:0.01,PL2:3:0.05
    'DigSrc_BitSize                     ex: INIT1:10
    'DigSrc_Seg                         ex: INIT1=SELSRAM, INIT6=sgmt0_7+sgmt1_19+sgmt3_7
    'DigSrc_DigSrcPin                   ex: INIT1:JTAG_TDI
    'digSrc_EQ                          ex: INIT6:sgmt0=0x41+sgmt1=0x00+sgmt2=0x0000
    '-------------------------------------------------------------------------------------------
    Dim PatString As String: PatString = vbNullString
    
    Dim RelayMode As tlRelayMode
    Dim PatThreading As Boolean
    Dim MatchAllSites As Boolean
    Dim WaitFlagA As tlWaitVal
    Dim WaitFlagB As tlWaitVal
    Dim WaitFlagC As tlWaitVal
    Dim WaitFlagD As tlWaitVal
    
    Dim tmpPatAry() As New Pattern
    Dim PatternAry_1() As String
    Dim PatternAry_2() As String
    Dim tmp_patAry() As String
    Dim tmp_patAry_spilt() As String
    Dim tmp_patAry_spiltCnt As Long, j As Long
    Dim tmp() As String
    Dim ShmooTestIdx As Long
    Dim CurConcurrentContext As Long
    Dim tempendofbody As String
    Dim tempendofbodyfargs As String
    Dim tempdrivepins As String
    Dim tempfloatpins As String
    Dim tmp_Shmoo_Pattern As String
    
    Dim i As Long
    Dim site As Variant
    Dim pin_count As Long
    Dim First_InitV As New PinListData
    Dim ShmooPinAry() As String
    Dim shmooPinCnt As Long
    Dim tmpShmooPin As Variant
    Dim StrPat As String
    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)
    
    RelayMode = tlPowered
    WaitFlagA = waitoff
    WaitFlagB = waitoff
    WaitFlagC = waitoff
    WaitFlagD = waitoff

'==============Add by Leslie 200915 for bypass shmoo hole to find the LVCC============'
    If BypassShmooHole Then
        gl_BY_PASS_SHMOO_HOLE = True
    Else
        gl_BY_PASS_SHMOO_HOLE = False
    End If
'==============Add by Leslie 200915 for bypass shmoo hole to find the LVCC============'

    Dim Test_Pattern As String
    Functional_T_char = TL_SUCCESS   ' be optimistic
    If Not TheExec.Flow.IsRunning Then Exit Function
    
    ' Cache parameters for PostTest
    m_EndOfBodyF = EndOfBodyF
    m_EndOfBodyFArgs = EndOfBodyFArgs
    
    ' Apply default values to parameters whose values were not specified.
    ApplyDefaults PatternTimeout
   
    If Validating_ Then
        If INIT_Patset <> "" Then PatString = INIT_Patset.value
        If PL_Patset <> "" Then PatString = PatString & "," & PL_Patset.value
        If Not Validate_Char(PatString, PatThreading, DriveLoPins, DriveHiPins, DriveZPins, DisablePins, _
            FloatPins, Util1Pins, Util0Pins, PatternTimeout, WaitTimeDomain) Then Functional_T_char = TL_ERROR
        Exit Function
    End If
    
    If Step_ = subAllBody Or Step_ = subPrebody Or m_InterposeFunctionsSet = False Then
        ' Register certain interpose function names with flow controller
        Call tl_SetInterpose(TL_C_PREPATF, PrePatF.value, PrePatFArgs, _
                             TL_C_POSTPATF, PostPatF.value, PostPatFArgs, _
                             TL_C_PRETESTF, PreTestF.value, PreTestFArgs, _
                             TL_C_POSTTESTF, PostTestF.value, PostTestFArgs, _
                             TL_C_FLAGMATCHF, PatFlagF.value, PatFlagFArgs, _
                             TL_C_POSTPATBPF, "PostTest", "")

        m_InterposeFunctionsSet = True
    End If

    ' PreBody
    If Step_ = subAllBody Or Step_ = subPrebody Then
        FetchContext
            
        ' Set up the test
        Call PreBody(DriveHiPins, DriveLoPins, DriveZPins, DisablePins, Util1Pins, Util0Pins, _
                 WaitFlagA, WaitFlagB, WaitFlagC, WaitFlagD, MatchAllSites, _
                 PatThreading, RelayMode, WaitTimeDomain, Interpose_PrePat)

        'DC_LEVEL Powers Stored
        Shmoo_Save_Power

        g_FirstSetp = True
        g_Vbump_function = True
        g_shmoo_ret = False
        Vbump_for_Interpose = True
        g_INIT_PAT_DONE = False
        Sweep_cnt = 0
                
        'Initial Shmoo Hole array
        ReDim Preserve Find_shmoo_hole_low_power_twice_hvcc_step(0) As New SiteLong
        ReDim Preserve Find_shmoo_hole_low_power_twice_hvcc_voltage(0) As New SiteDouble
        For Each site In TheExec.sites
            Find_shmoo_hole_low_power_twice_hvcc_step(0)(site) = 0
            Find_shmoo_hole_low_power_twice_hvcc_voltage(0)(site) = 0
        Next site
        '20240327 applyBV
        If TheExec.Flow.enableWord("Enable_ApplyBVToCZ") = False Then
            gb_ApplyBV = False
            ApplyVoltageFromBinCut = vbNullString
        Else
            gb_ApplyBV = True
        End If
        If ApplyVoltageFromBinCut <> "" Then
            'init bincut data structure
            Call initialize_inst_info(CZ_inst_info, ApplyVoltageFromBinCut)
            
            'search bincut voltage
            Call bincut_power_Setting_VT(CZ_inst_info, CurrentPassBinCutNum, BinCut_Payload_Voltage, True)
            
            PrintString = vbNullString
            For Each site In TheExec.sites
                For i = 0 To UBound(pinGroup_CorePower)
                    PrintString = PrintString & pinGroup_CorePower(i) & ": " & CStr(BinCut_Payload_Voltage(VddBinStr2Enum(pinGroup_CorePower(i)))) & "mV, "
                Next i
                For i = 0 To UBound(pinGroup_OtherRail)
                    PrintString = PrintString & pinGroup_OtherRail(i) & ": " & CStr(BinCut_Payload_Voltage(VddBinStr2Enum(pinGroup_OtherRail(i)))) & "mV, "
                Next i
                TheExec.Datalog.WriteComment "Site: " & site & ", " & PrintString
                PrintString = vbNullString
            Next site
        End If
        If RetentionMeasVIF <> "" Then
            Enable_AutoRange = MeasureAutoRange
            RetentionMeasureParser RetentionMeasVIF
            Glb_RetentionMeasurement = True
        Else
            Glb_RetentionMeasurement = False
        End If

        If TheExec.Flow.enableWord("BringUp_Shmoo") = True Then
            InitialPatCnt = 0
            PayloadPatCnt = 0
            Power_Run_Scenario = "INIT_NV_PL_SWEEP"
            DecomposePattSet_Mod INIT_Patset, tmpPatAry
        Else
            tmp_patAry = Split(INIT_Patset, ",")
            For i = 0 To UBound(tmp_patAry)
                tmp_patAry_spilt = TheExec.DataManager.Raw.GetPatternsInSet(tmp_patAry(i), tmp_patAry_spiltCnt)
                For j = 0 To tmp_patAry_spiltCnt - 1
                    ReDim Preserve PatternAry_1(j + InitialPatCnt)
                    PatternAry_1(j + InitialPatCnt) = tmp_patAry_spilt(j)
                Next j
                InitialPatCnt = InitialPatCnt + tmp_patAry_spiltCnt
            Next i
            
            tmp_patAry = Split(PL_Patset, ",")
            For i = 0 To UBound(tmp_patAry)
                tmp_patAry_spilt = TheExec.DataManager.Raw.GetPatternsInSet(tmp_patAry(i), tmp_patAry_spiltCnt)
                For j = 0 To tmp_patAry_spiltCnt - 1
                    ReDim Preserve PatternAry_2(j + PayloadPatCnt)
                    PatternAry_2(j + PayloadPatCnt) = tmp_patAry_spilt(j)
                Next j
                PayloadPatCnt = PayloadPatCnt + tmp_patAry_spiltCnt
            Next i
            For i = 0 To InitialPatCnt - 1
                ReDim Preserve tmpPatAry(i)
                tmpPatAry(i) = PatternAry_1(i)
            Next i
            For i = 0 To PayloadPatCnt - 1
                ReDim Preserve tmpPatAry(InitialPatCnt + i)
                tmpPatAry(InitialPatCnt + i) = PatternAry_2(i)
            Next i
            StrPat = Join(PatternAry_1, ",") & "," & Join(PatternAry_2, ",") '20240521 MFSTP
        End If
        PatAmount = UBound(tmpPatAry)
        Shmoo_Pattern = vbNullString
        'Retention case
        If UCase(Power_Run_Scenario) Like "*_SWEEP*" Then
            g_PLSWEEP = True
        Else
            g_PLSWEEP = False
        End If
        'Procss voltage, pattern type, source bit,...
        ProcessPattInfo tmpPatAry, SELSRAM_DSSC, DigSrc_BitSize, DigSrc_Seg, DigSrc_DigSrcPin, digSrc_EQ, Interpose_PrePat, Wait, Ret_Ramp_Setting, "", "Temp", Power_Run_Scenario, pmode
        'ProcessPattInfo tmpPatAry, SELSRAM_DSSC, DigSrc_BitSize, DigSrc_Seg, DigSrc_DigSrcPin, digSrc_EQ, Interpose_PrePat, Wait, "", "Temp", Power_Run_Scenario, pmode
        If TheExec.DevChar.Setups.IsRunning = True Then
            If SSN_EnableCoreMask = True Then
                Dim suspend As Boolean
                With TheExec.Datalog
                    suspend = .DatalogSuspended
                    .DatalogSuspended = False
                    .WriteComment "<SSN_CoreMask>"
                    ''20240222: Added argument in char for manual control SSN flag
                    SSN_CoreMask SSN_SpecifyFlag
                    .DatalogSuspended = suspend
                End With
            Else
                TheExec.Datalog.WriteComment "SSN_EnableCoreMask = False, no SSN core is masked!!"
            End If
        End If
        'If (TheExec.Flow.IsCharacterizing = True) Then SSN_CoreMask
        ''20240521 MFSTP
        If UserFunction <> "" Then
            CZ_inst_info.MultiFSTP_Enable = True
            CZ_inst_info.Harvest_Core_DigSrc_Pin = "JTAG_TDI"
            CZ_inst_info.Harvest_Core_DigSrc_SignalName = "Harvest_Core_DigSrcSignal"
            Call CheckInstForUserFunction(UserFunction, CZ_inst_info.digSrcLabel, CZ_inst_info.digSrcPatterns, StrPat)
        Else
                CZ_inst_info.MultiFSTP_Enable = False
        End If
    End If ' PreBody
    
    CurConcurrentContext = m_STDSvcClient.FlowDomainService.ConcurrentContext
    
    ' Body
    If Step_ = subAllBody Or Step_ = subBody Then
        
        ' cache member variables
        ' there are statements below which can cause us to jump to the next subflow if we're running with concurrent test.
        ' if the next test in the next subflow runs this function then it will overwrite the below member variables, such
        ' that when we get back to this call they will have different values.  so we cache the values here and then
        ' restore them right after the code that can cause us to jump to the next subflow.  then later on in
        ' postbody and posttest when they're used they'll have the proper values.
        If CurConcurrentContext Then
            tempendofbody = m_EndOfBodyF
            tempendofbodyfargs = m_EndOfBodyFArgs
            tempdrivepins = m_DrivePins
            tempfloatpins = m_FloatPins
        End If
                
        ' Perform the test
        Call Interpose(StartOfBodyF, StartOfBodyFArgs)

        If g_FirstSetp = True Then
            For Each site In TheExec.sites
            Interpose_PrePat = g_MergeCond
            g_ForceCond_VDD = g_MergeVDD
               ' =======================3D Char 20190705========================
               If InStr(UCase(Interpose_PrePat), "USL") > 0 Then
                  PL_DC_conditions_GLB = mid(Interpose_PrePat, 1, InStr(UCase(Interpose_PrePat), "USL") - 2)
               ' =======================3D Char========================
               Else
                  PL_DC_conditions_GLB = Interpose_PrePat
               End If
            Next site
            Dim PatArray() As String
            Dim Pat As Variant
            Dim PatCount As Long
            ReDim PatArray(PatAmount)
            For i = 0 To PatAmount
                PatArray(i) = g_CharPattInfoAry(i).Pattern.value
            Next
            Harv_FailFlag = "HarvestPinFlag_Table;EnableCoreHarvest:TRUE;EnableCoreMask:TRUE"
            '20231211: Decide SSN enable/disable
            If Flag_HarvPinFlag_Mapping_Table_Parsed = True And ssnPatternsDict.Count > 0 Then
                Dim instSSNinfo As Inst_SSN
                Call CheckInstForSSN(PatArray, instSSNinfo, SSNMapping, Harv_FailFlag)
            End If
            g_FirstSetp = False
            End If
            Call Body(FloatPins, StrToDbl(PatternTimeout), g_CharPattInfoAry(UBound(g_CharPattInfoAry)).Pattern, ReportResult, ResultMode)
        
        ' Run the pattern.  Perform functional test.
        If TheExec.sites.ActiveCount > 0 Then
        
            Get_Shmoo_Set_Pin Shmoo_Apply_Pin, g_ForceCond_VDD, pin_count
            TheExec.Datalog.WriteComment Power_Run_Scenario
                           
            For ShmooTestIdx = 0 To UBound(g_CharPattInfoAry)
                If g_CharPattInfoAry(ShmooTestIdx).Pattern.value <> "" Then
                    With g_CharPattInfoAry(ShmooTestIdx)
                    If .IsInitPattern = True And g_INIT_PAT_DONE = True Then GoTo skip_pat_run
                        
                        'reset 1st pattern voltage status (set context value)
                        'reset 1st pl pattern (switch to Vmain) for avoid hot switch
                        If .Sequence = 1 And TheExec.DevChar.Setups.IsRunning = True Then
                            If ALL_Power_DCVS_pins <> "" Then
                                TheExec.DataManager.DecomposePinList ALL_Power_DCVS_pins, ShmooPinAry, shmooPinCnt
                                For Each tmpShmooPin In ShmooPinAry
                                    TheHdw.DCVS.Pins(tmpShmooPin).Voltage.Main.value = g_ContextVmainValue.Pins(tmpShmooPin).value
                                    TheHdw.DCVS.Pins(tmpShmooPin).Voltage.Alt.value = g_ContextVAltValue.Pins(tmpShmooPin).value
                                Next tmpShmooPin
                                TheHdw.DCVS.Pins(ALL_Power_DCVS_pins).Voltage.Output = tlDCVSVoltageMain
                            End If
                        End If

                        Shmoo_Test_Pattern .Pattern, ReportResult, CLng(TL_C_YES), ConcurrentMode, .PowerRunCond, _
                        Shmoo_Apply_Pin, .IsInitPattern, .Sequence, .ForceVoltage, .WaitTime, _
                        instSSNinfo, CZ_inst_info, _
                        .DigSrc_BitSize, .DigSrc_Seg, .DigSrc_pin, .digSrc_EQ, .DynamicSourceBit, .DigSrcType, _
                        .SelSramMatchIdx, .PatternLoopCnt, .testType, .SegDict, Order_LSB, .DictApplyVol
                        '.SelSramMatchIdx, .PatternLoopCnt, .testType, .EqDict, Order_LSB
                        '.SelSramMatchIdx, .PatternLoopCnt, .testType
                        
                        If tmp_Shmoo_Pattern = "" Then
                            tmp_Shmoo_Pattern = .Pattern.value
                        Else
                            tmp_Shmoo_Pattern = tmp_Shmoo_Pattern & "," & .Pattern.value
                        End If
                    End With
                End If
skip_pat_run:
            Next ShmooTestIdx
            
            Shmoo_Pattern = tmp_Shmoo_Pattern
            
            If TheExec.DevChar.Setups.IsRunning = False And CharSetName_GLB <> "" Then
                Dim p As Variant, p_ary() As String, p_cnt As Long, ApplyPins As String, Setup_mode As String
                If TheExec.DevChar.Setups(CharSetName_GLB).TestMethod.value = tlDevCharTestMethod_Reburst Then TheExec.Datalog.WriteComment "[PrintCharCondition:" & PrintCharSetup(Interpose_PrePat) & ",Test]"
                Setup_mode = TheExec.DevChar.Setups(CharSetName_GLB).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.name
                If (LCase(Setup_mode) <> "vid" And LCase(Setup_mode) <> "vicm") Then
                    ApplyPins = TheExec.DevChar.Setups(CharSetName_GLB).Shmoo.Axes(tlDevCharShmooAxis_X).ApplyTo.Pins
                    TheExec.DataManager.DecomposePinList ApplyPins, p_ary, p_cnt
                    For Each p In p_ary
                        TheExec.DevChar.Setups(CharSetName_GLB).Shmoo.Axes(tlDevCharShmooAxis_X).ApplyTo.Pins = p
                        run_shmoo CharSetName_GLB
                    Next p
                    TheExec.DevChar.Setups(CharSetName_GLB).Shmoo.Axes(tlDevCharShmooAxis_X).ApplyTo.Pins = ApplyPins
                Else
                    run_shmoo CharSetName_GLB
                End If
            End If

            If err.number <> 0 Then
                If err.number = TL_E_AT_PATSET_BREAKPT Then
                    Exit Function
                Else
                    GoTo errHandler
                End If
            End If
            On Error GoTo errHandler
        End If
                
        ' restore the member variables for posttest
        If CurConcurrentContext Then
            m_EndOfBodyF = tempendofbody
            m_EndOfBodyFArgs = tempendofbodyfargs
        End If
    
        ' Calls End of Body Interpose Function, anything from here to the end of the Body
        ' should be added to PostTest()
        Dim argv() As String
        PostTest 0, argv
        
        DebugPrintFunc Shmoo_Pattern
        
        ' restore the member variables for postbody (do this here instead of a couple of lines above since posttest could
        ' possibly jump to the next subflow in a concurrent test and cause the below memeber variables to change again.
        If CurConcurrentContext Then
            m_DrivePins = tempdrivepins
            m_FloatPins = tempfloatpins
        End If
        TheHdw.DCVS.Pins(ALL_Power_DCVS_pins).Voltage.Output = tlDCVSVoltageMain
        If One_Time_INIT And TheExec.DevChar.Setups.IsRunning = True Then
            g_INIT_PAT_DONE = True
            Sweep_cnt = Sweep_cnt + 1
        End If
    End If ' Body
    
    ' PostBody
    If Step_ = subAllBody Or Step_ = subPostbody Then
        ReadHWPowerValue_GLB = PrintCharSetup(Interpose_PrePat)
        
        g_Vbump_function = False
        Shmoo_Restore_Power_per_site_Vbump_NV True
        'PowerGRP MOD 210601
        TheHdw.DCVS.Pins(ALL_Power_DCVS_pins).Voltage.Output = tlDCVSVoltageMain
        TheHdw.Wait 0.001
         If Interpose_PrePat <> "" Then
             Call SetForceCondition("RESTOREPREPAT")
         End If
        
        Call PostBody(m_DrivePins, m_FloatPins, WaitTimeDomain, WaitFlagA, WaitFlagB, WaitFlagC, WaitFlagD)
    End If ' PostBody
    
    ' There shouldn't be any code below this line. Any other necessary
    ' code should be added to the PostTest method to support pattern set
    ' breakpoints.
    
    
    Exit Function
    
errHandler:
        Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Shmoo", "Functional_T_char", "Test " & TL_C_ERRORSTR & ", Instance: " & TheExec.DataManager.instancename)
    'Call theexec.ErrorLogMessage("Test " & TL_C_ERRORSTR & ", Instance: " & theexec.DataManager.instancename)
    Call TheExec.ErrorReport
    ' Clear previously registered interpose function names
    Call tl_ClearInterpose(TL_C_PREPATF, TL_C_POSTPATF, TL_C_PRETESTF, TL_C_POSTTESTF)
    m_InterposeFunctionsSet = False
    Functional_T_char = TL_ERROR
        Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Shmoo", "Functional_T_char")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Sub ApplyDefaults(ByRef PatternTimeout As String)
On Error GoTo errHandler
    ' If the worksheet doesn't have a value then apply 30 as the default.
    If Not NonBlank(PatternTimeout) Then
        PatternTimeout = "30"
    End If
    Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Shmoo", "ApplyDefaults")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

'removed by SY, no one use this function and this will causing compiler fail for non-AP project since different efuse structure

''Public Function Read_Waferdata_char()
''
''    Dim site As Variant
''    If UCase(TheExec.CurrentJob) Like "*FT*" Then
''        For Each site In TheExec.sites
''            HramLotId(site) = ECIDFuse.Category(ECIDIndex("Lot_ID")).Read.ValStr(site) + "-" + ECIDFuse.Category(ECIDIndex("Wafer_ID")).Read.ValStr(site)
''            'HramWaferId(Site) = TheExec.Datalog.Setup.WaferSetup.ID
''            XCoord(site) = ECIDFuse.Category(ECIDIndex("X_Coordinate")).Read.ValStr(site)
''            YCoord(site) = ECIDFuse.Category(ECIDIndex("Y_Coordinate")).Read.ValStr(site)
''        Next site
''    ElseIf UCase(TheExec.CurrentJob) Like "*CP*" Then
''        For Each site In TheExec.sites
''            HramLotId(site) = TheExec.Datalog.setup.LotSetup.LotID
''            HramWaferId(site) = TheExec.Datalog.setup.WaferSetup.ID
''            'HramLotId(Site) = HramLotId(Site) & "-" & HramWaferId(Site)
''            XCoord(site) = TheExec.Datalog.setup.WaferSetup.GetXCoord(site)
''            YCoord(site) = TheExec.Datalog.setup.WaferSetup.GetYCoord(site)
''
''            TheExec.Datalog.WriteComment "[XY_Coordinate_Read,Site:" & site & ",X:" & XCoord(site) & ",Y:" & YCoord(site) & ",LotId:" & HramLotId(site) & "]"
''            'TheExec.AddOutput "[XY_Coordination_Read,Site:" & site & ",X:" & XCoord(site) & ",Y:" & YCoord(site) & ",LotId:" & HramLotId(site) & "]"
''        Next site
''    End If
''
''
''
''End Function

Public Function PrintShmooInfo(argc As Long, argv() As String)
On Error GoTo errHandler
    Dim SetupName As String
    Dim method As String
    Dim shmoo_axis As Variant

    '20180118 Refresh shmoo overlay count
    If TheExec.Overlays.Count > 10000 Then
        TheExec.Overlays.RemoveAll
    End If

    SetupName = TheExec.DevChar.Setups.ActiveSetupName
    'Prevent Algorithm "List" Error
    With TheExec.DevChar.Setups(SetupName).Shmoo
        For Each shmoo_axis In .Axes.list
            If UCase(.Axes(shmoo_axis).Algorithm.name) = "LIST" Then
                TheExec.Datalog.WriteComment "*** The Algorithm: List is not Support ***"
                Exit Function
            End If
        Next shmoo_axis
    End With
    
    With TheExec.DevChar.Setups(SetupName)
        If .Shmoo.Axes.Count > 1 And .Shmoo.Axes.Count < 3 Then
            Call ShmooPostStep2Dto1D(argc, argv)
            Call ShmooPostStep2D(argc, argv)
        ElseIf .Shmoo.Axes.Count = 1 Then
            TheExec.Datalog.WriteComment "[Start_Shmoo]"
            Call ShmooPostStep1D(argc, argv)
            TheExec.Datalog.WriteComment "[End_Shmoo]"
        End If
        
        If Multi_Axis_PTR = True Then
'        If TheExec.DevChar.Setups.Item(SetupName).output.SuspendDatalog = False And .Shmoo.Axes.count = 2 Then
            If .Shmoo.Axes.Count = 2 Then
                StoreEachPointResult_2D
                TheExec.Datalog.WriteComment "<<<<< 2D Shmoo Print Start >>>>>"
                Call Print2DShmooInfo(argc, argv)
                TheExec.Datalog.WriteComment "<<<<< 2D Shmoo Print Stop >>>>>"
    '            Count_Point = 0
            End If
        'If TheExec.DevChar.Setups.item(SetupName).output.SuspendDatalog = False And .Shmoo.Axes.count = 3 Then
            If .Shmoo.Axes.Count = 3 Then
                Call ShmooPostStep3D(argc, argv)
                StoreEachPointResult_3D
                TheExec.Datalog.WriteComment "<<<<< 3D Shmoo Print Start >>>>>"
                Dim AxisOrder As Variant
                AxisOrder = TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.AxisOrder
                If AxisOrder = tlDevCharAxisOrder_ZXY Then  ''Case ZXY
                   Call Print3DShmooInfo_ZXY(argc, argv)
                ElseIf AxisOrder = tlDevCharAxisOrder_XYZ Then ''  Case XYZ
                   Call Print3DShmooInfo_XYZ(argc, argv)
                End If
                TheExec.Datalog.WriteComment "<<<<< 3D Shmoo Print Stop >>>>>"
    '            Count_Point = 0
            End If
        End If
        ''''' 20180710 Initialize GLlobal power condition
        ReadHWPowerValue_GLB = vbNullString
        Charz_Force_Power_condition = vbNullString
        
        ''''' 20180710 Add initialize value ''''''''''''
        CHAR_USL_HVCC = 9999
        CHAR_USL_LVCC = 9999
        CHAR_LSL_HVCC = 9999
        CHAR_LSL_LVCC = 9999
        
        
    End With

    Dim AcCat As String
    Dim site As Variant
    Dim SetSite As Integer
    AcCat = TheExec.Contexts.ActiveSelection.ACCategory

    'For turn off TName Sweep point
    gl_flag_end_shmoo = True
    gl_flag_CZ_Nominal_Measured_1st_Point = False
    
    'If TheExec.DevChar.Setups.item(TheExec.DevChar.Setups.ActiveSetupName).Output.SuspendDatalog = False Then    '20180718 add
    '    Call TheExec.sites(site).IncrementTestNumber
    'End If
    
  '''''''''''''''''Support multiple nWire port 20170503'''''''''''''
    Dim nWire_port_ary() As String
    Dim nwp As Variant ', all_ports As String, all_pins As String
    Dim port_pa As String, ac_spec_pa As String, pin_pa As String, global_spec_pa As String

    nWire_port_ary = Split(nWire_Ports_GLB, ",")

    With TheExec.DevChar.Setups(SetupName)
       For Each nwp In nWire_port_ary
            Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
            If UCase(.Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.name) = UCase(ac_spec_pa) Then
                Call VaryFreq(port_pa, TheExec.Specs.AC(ac_spec_pa).ContextValue, ac_spec_pa)
            End If
            If .Shmoo.Axes.Count > 1 Then
                If UCase(.Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.name) = UCase(ac_spec_pa) Then
                    Call VaryFreq(port_pa, TheExec.Specs.AC(ac_spec_pa).ContextValue, ac_spec_pa)
                End If
            End If
        Next nwp
        'If .Shmoo.Axes.Count > 1 Then
        '    If TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.type.value = "AC Spec" Or TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.type.value = "AC Spec" Or TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.type.value = "Global Spec" Or TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.type.value = "Global Spec" Then
        ''    If TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.Type.Value = "AC Spec" Or TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.Type.Value = "Global Spec" Then
        '        For Each nwp In nWire_port_ary
        '            Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
        '            Call VaryFreq(port_pa, TheExec.Specs.AC(ac_spec_pa).ContextValue, ac_spec_pa)
        '        Next nwp
        '    End If
        'Else
        '    For Each nwp In nWire_port_ary
        '        Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
        '        Call VaryFreq(port_pa, TheExec.Specs.AC(LCase(ac_spec_pa)).ContextValue, ac_spec_pa)
        '    Next nwp
        'End If
    End With
    Shmoo_End = True
    g_TestNum = -1
    Multi_Axis_PTR = False
    'alarm check
    Dim AlarmOccur As Boolean
    Dim MonitorPower As String: MonitorPower = ALL_Power_DCVS_pins + IIf(ALL_Power_DCVS_pins <> "" And ALL_Power_DCVI_pins <> "", ",", "") + ALL_Power_DCVI_pins
    Dim PowerPinAry() As String
    Dim PowerPinCnt As Long
    Dim pin As Variant
    AlarmOccur = TheHdw.Alarms.Check
    If AlarmOccur = True Then
        TheExec.DataManager.DecomposePinList MonitorPower, PowerPinAry, PowerPinCnt
        For Each site In TheExec.sites
            For Each pin In PowerPinAry
                If TheExec.sites.item(site).FlagState("F_Shmoo_Alarm") = logicTrue Then Exit For
                If gl_GetInstrument_Dic.Exists(LCase(pin)) Then
                    If LCase(gl_GetInstrumentType_Dic(LCase(pin))) Like "*dcvs*" Then
                        If TheHdw.DCVS.Pins(pin).Gate = False Then
                            TheExec.Datalog.WriteComment "<Warning> Alarm on Site: " & site & " , Pin: " & pin & ", Occur Alarm. Set Alarm Bin Out!"
                            TheExec.sites.item(site).FlagState("F_Shmoo_Alarm") = logicTrue
                        End If
                    ElseIf LCase(gl_GetInstrumentType_Dic(LCase(pin))) Like "*dcvi*" Then
                        If TheHdw.DCVI.Pins(pin).Gate = False Then
                            TheExec.Datalog.WriteComment "<Warning> Alarm on Site: " & site & " , Pin: " & pin & ", Occur Alarm. Set Alarm Bin Out!"
                            TheExec.sites.item(site).FlagState("F_Shmoo_Alarm") = logicTrue
                        End If
                    End If
                End If
            Next pin
        Next site
    End If
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Shmoo", "PrintShmooInfo")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Flow_Shmoo_Setup()
On Error GoTo errHandler
    Dim DevChar_Setup As String
    Dim shmoo_axis As Variant, Shmoo_Tracking_Item As Variant
    Dim axis_name As Variant, axis_type As String, Tracking_num As Long
    Dim i As Long, Shmoo_Spec As String, Shmoo_StepSize As Double, shmoo_step As Long
    Dim StepSize As Double
    Dim arg_ary() As String
    Dim site As Variant
    Shmoo_setup_str = vbNullString
    Flow_Shmoo_Axis_Count = 0
    
    Flow_Shmoo_X_Current_Step = -1
    Flow_Shmoo_Y_Current_Step = -1
    Flow_Shmoo_X_Last_Value = -99
    Flow_Shmoo_Y_Last_Value = -99
    Flow_Shmoo_X_Fast = False
    
    For Each site In TheExec.sites
        DevChar_Setup = TheExec.sites.item(site).SiteVariableValue("Flow_Shmoo_DevCharSetup")
        With TheExec.DevChar.Setups(DevChar_Setup)
            For Each shmoo_axis In .Shmoo.Axes.list
                Select Case shmoo_axis
                    Case tlDevCharShmooAxis_X:
                        axis_type = "X"
                    Case tlDevCharShmooAxis_Y:
                        axis_type = "Y"
                End Select
                With TheExec.DevChar.Setups(DevChar_Setup).Shmoo.Axes(shmoo_axis)
                    TheExec.sites.item(site).SiteVariableValue("Flow_Shmoo_" & axis_type & "_Start") = .Parameter.range.from
                    TheExec.sites.item(site).SiteVariableValue("Flow_Shmoo_" & axis_type & "_Stop") = .Parameter.range.to
                    TheExec.sites.item(site).SiteVariableValue("Flow_Shmoo_" & axis_type & "_StepSize") = .Parameter.range.StepSize
                    shmoo_step = Abs(Floor((.Parameter.range.to - .Parameter.range.from) / .Parameter.range.StepSize))
                    If axis_type = "X" Then
                        Flow_Shmoo_X_Step = shmoo_step
                        TheExec.sites(site).SiteVariableValue("Flow_Shmoo_X_Step") = shmoo_step
                    Else
                        Flow_Shmoo_Y_Step = shmoo_step
                        TheExec.sites(site).SiteVariableValue("Flow_Shmoo_Y_Step") = shmoo_step
                    End If
                End With
            Next shmoo_axis
        End With
    Next site
     With TheExec.DevChar.Setups(DevChar_Setup)
        For Each shmoo_axis In .Shmoo.Axes.list
            Select Case shmoo_axis
                Case tlDevCharShmooAxis_X:
                    axis_type = "X"
                Case tlDevCharShmooAxis_Y:
                    axis_type = "Y"
            End Select
            With TheExec.DevChar.Setups(DevChar_Setup).Shmoo.Axes(shmoo_axis)
                Select Case .Parameter.type.value
                    Case "Level": Shmoo_Spec = .ApplyTo.Pins & "(" & .Parameter.name & ")"
                    Case "AC Spec": Shmoo_Spec = .Parameter.name
                    Case "Global Spec":
                        arg_ary = Split(.InterposeFunctions.PrePoint.Arguments, ",")
                        If LCase(.InterposeFunctions.PrePoint.name) Like "freerunclk_set_xy" Then
                            Shmoo_Spec = arg_ary(2)
                        Else
                            Shmoo_Spec = .Parameter.name
                        End If
                    
                End Select

                Flow_Shmoo_Axis(Flow_Shmoo_Axis_Count) = axis_type
                Flow_Shmoo_Axis_Count = Flow_Shmoo_Axis_Count + 1
                If .Parameter.range.from < .Parameter.range.to Then
                    StepSize = .Parameter.range.StepSize
                Else
                    StepSize = -(.Parameter.range.StepSize)
                End If
                If Shmoo_setup_str = "" Then
                    Shmoo_setup_str = "Shmoo_Setup(" & DevChar_Setup & ")" & axis_type & ":" & Shmoo_Spec & "=" & .Parameter.range.from & "," & .Parameter.range.to & "," & StepSize & "; "
                Else
                    Shmoo_setup_str = Shmoo_setup_str & axis_type & ":" & Shmoo_Spec & "=" & .Parameter.range.from & "," & .Parameter.range.to & "," & StepSize & "; "
                End If
            End With
            With TheExec.DevChar.Setups(DevChar_Setup).Shmoo.Axes(shmoo_axis).TrackingParameters
                For Each Shmoo_Tracking_Item In .list
                    Shmoo_Spec = .item(Shmoo_Tracking_Item).ApplyTo.Pins
                    Shmoo_StepSize = (.item(Shmoo_Tracking_Item).range.to - .item(Shmoo_Tracking_Item).range.from) / shmoo_step
                    Shmoo_setup_str = Shmoo_setup_str & axis_type & ":" & Shmoo_Spec & "=" & .item(Shmoo_Tracking_Item).range.from & "," & .item(Shmoo_Tracking_Item).range.to & "," & Shmoo_StepSize & "; "
                    Flow_Shmoo_Axis(Flow_Shmoo_Axis_Count) = axis_type & Tracking_num
                    Flow_Shmoo_Axis_Count = Flow_Shmoo_Axis_Count + 1
                Next Shmoo_Tracking_Item
            End With
        Next shmoo_axis
    End With
    TheExec.Datalog.WriteComment "******************************    " & Shmoo_setup_str & "   ******************************"

    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Shmoo", "Flow_Shmoo_Setup")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function CheckCharErrorCount(shmoo_abnormal_type As String, shmoo_abnormal_ratio_hi_lim As Double) As Long
On Error GoTo errHandler

    'Dim site As Long
    Dim shmoo_abnormal_ratio As New SiteDouble
    Dim test_name As String
    Dim site As Variant 'Carter, 20240304
    
    '' select one of the abnormal type [alarm, shmoo_hole, all_fail]
    For Each site In TheExec.sites
        If included_shmoo_count(site) <> 0 Then
            Select Case LCase(Trim(shmoo_abnormal_type))
                Case "alarm":
                    'For Each site In TheExec.sites
                    shmoo_abnormal_ratio(site) = Format(shmooalarm_count(site) / included_shmoo_count(site), "0.000")
                    'Next site
                    test_name = "CheckChar_Alarm"
                    shmooalarm_count = 0
                Case "shmoo_hole":
                    'For Each site In TheExec.sites
                    shmoo_abnormal_ratio(site) = Format(shmoohole_count(site) / included_shmoo_count(site), "0.000")
                   ' Next site
                    test_name = "CheckChar_Hole"
                    shmoohole_count = 0
                
                Case "all_fail":
                    'For Each site In TheExec.sites
                    shmoo_abnormal_ratio(site) = Format(shmooallfail_count(site) / included_shmoo_count(site), "0.000")
                    'Next site
                    test_name = "CheckChar_AllFail"
                    shmooallfail_count = 0
                Case Else:
                    TheExec.Datalog.WriteComment " [Warning] Unrecognized shmoo_abnormal_type = " & shmoo_abnormal_type + vbCrLf
                    Exit Function
            End Select
            
            '' Apply Limit
            TheExec.Datalog.WriteComment vbNullString
            'For Each site In TheExec.sites
            TheExec.Datalog.WriteComment "site " & site & ", Total_shmoo_count = " & total_shmoo_count(site)
            TheExec.Datalog.WriteComment "site " & site & ", Included_shmoo_count = " & included_shmoo_count(site)
            TheExec.Datalog.WriteComment "site " & site & ", Excluded_shmoo_count = " & excluded_shmoo_count(site)
            'Next site
            TheExec.Datalog.WriteComment vbNullString
            TheExec.Flow.TestLimit resultVal:=shmoo_abnormal_ratio, hiVal:=shmoo_abnormal_ratio_hi_lim, Tname:=test_name, scaletype:=scalePercent
        Else
            'For Each site In TheExec.sites
            TheExec.Datalog.WriteComment vbNullString
            TheExec.Datalog.WriteComment " site " & site & " " & shmoo_abnormal_type & ", [Warning] Included shmoo count = 0 "
            'Next site
        End If
    Next site
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Shmoo", "CheckCharErrorCount")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function EnableShmooAbnormalCounter()
On Error GoTo errHandler

    F_shmoo_abnormal_counter = True
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Shmoo", "EnableShmooAbnormalCounter")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function DisableShmooAbnormalCounter()
On Error GoTo errHandler

    F_shmoo_abnormal_counter = False
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Shmoo", "DisableShmooAbnormalCounter")
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Re_PowerOn_WhileSweep(argc As Long, argv() As String)
On Error GoTo errHandler

'argv(0)=PowerPin
'argv(1)=wait time
'argv(2)=cool down voltage

    Dim pin As Variant
    Dim Pins() As String
    Dim Pin_Cnt As Long
    Dim InstName As String
    Dim site As Variant
    Dim gate_off As Boolean
    Dim PowerVolage As New PinListData
    
    
    
    'Initialize
    Dim powerPin, wait_time, cool_down_voltage As String


'exit function

    powerPin = "CorePower"
    wait_time = "0.1"
    'cool_down_voltage = "0.1"
    gate_off = False
    
    If argc = 2 Then
        powerPin = CStr(argv(0))
        wait_time = CDbl(argv(1))
        'cool_down_voltage = CDbl(argv(2))
    Else
        TheExec.Datalog.WriteComment "Using default string, Mointor Pin: CorePower, Wait_Time 0.1s"
        TheExec.Datalog.WriteComment "Please fill argument in this format to change default value: Power_pins, wait_time"
        'Exit Function
    End If


    TheExec.DataManager.DecomposePinList powerPin, Pins(), Pin_Cnt


    For Each pin In Pins
        PowerVolage.AddPin CStr(pin)
    Next pin

    For Each site In TheExec.sites.Active
        For Each pin In Pins
            If TheExec.DataManager.ChannelType(pin) <> "N/C" Then
                PowerVolage.Pins(CStr(pin)).value = TheHdw.DCVS.Pins(CStr(pin)).Voltage.value
            End If
        Next pin
    Next site

    For Each site In TheExec.sites.Active
        For Each pin In Pins
            If TheExec.DataManager.ChannelType(pin) <> "N/C" Then
                InstName = GetInstrument(CStr(pin), 0)
                Select Case InstName
                Case "VHDVS"
                    If TheHdw.DCVS.Pins(pin).Gate = False Then
                        gate_off = True
                        Exit For
                    End If
                Case "HexVS"
                    If TheHdw.DCVS.Pins(pin).Gate = False Then
                        gate_off = True
                        Exit For
                    End If
                Case Else
                    TheExec.Datalog.WriteComment "Error in Re_PowerOn_WhileSweep()"
                End Select
            End If
        Next pin
         
        If gate_off Then
        
            For Each pin In Pins
                If TheExec.DataManager.ChannelType(pin) <> "N/C" Then
                    InstName = GetInstrument(CStr(pin), 0)
                    Select Case InstName
                    Case "VHDVS"
                        PowerVolage.Pins(CStr(pin)).value = TheHdw.DCVS.Pins(CStr(pin)).Voltage.value
                       'thehdw.DCVS.Pins(Pin).Voltage.Main.Value = thehdw.DCVS.Pins(Pin).Voltage.Main.Value - cool_down_voltage
                    Case "HexVS"
                        PowerVolage.Pins(CStr(pin)).value = TheHdw.DCVS.Pins(CStr(pin)).Voltage.value
                       'thehdw.DCVS.Pins(Pin).Voltage.Main.Value = thehdw.DCVS.Pins(Pin).Voltage.Main.Value - cool_down_voltage
                    Case Else
                        TheExec.Datalog.WriteComment "Error in Re_PowerOn_WhileSweep()"
                        TheExec.Datalog.WriteComment "Pin: " & pin
                    End Select
                End If
            Next pin
        End If
    Next site

    If gate_off = False Then
        Exit Function
    End If

    DCVS_PowerDown_Parallel_Interpose AllPowerPinlist, 0.001, False
    
    
    TheHdw.Wait wait_time
    
    DCVS_PowerUp_Parallel_Interpose AllPowerPinlist, vbNullString, 0.001, False


    For Each site In TheExec.sites.Active
        For Each pin In Pins
            If TheExec.DataManager.ChannelType(pin) <> "N/C" Then
                InstName = GetInstrument(CStr(pin), 0)
                Select Case InstName
                Case "VHDVS"
                    TheHdw.DCVS.Pins(pin).Voltage.Main.value = PowerVolage.Pins(pin).value 'thehdw.DCVS.Pins(Pin).Voltage.Main.Value + cool_down_voltage
                Case "HexVS"
                    TheHdw.DCVS.Pins(pin).Voltage.Main.value = PowerVolage.Pins(pin).value 'thehdw.DCVS.Pins(Pin).Voltage.Main.Value + cool_down_voltage
                Case Else
                    TheExec.Datalog.WriteComment "Error in Re_PowerOn_WhileSweep()"
                End Select
            End If
         Next pin
    Next site

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Shmoo", "Re_PowerOn_WhileSweep")
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function DCVS_PowerUp_Parallel_Interpose(PowerPinList As String, DisconnectPinList As String, Optional WaitConnectTime As Double = 0.001, Optional DebugFlag As Boolean = False)
On Error GoTo errHandler
'power up sequence at flow start
    Dim CurrentChans As String
    Dim site As Variant
    Dim Pins() As String, PinCnt As Long
    Dim powerPin As Variant
    Dim PowerName As String
    Dim TempString As String
    Dim VMain As Double
    Dim Irange As Double
    Dim step As Integer
    Dim RiseTime As Double
    Dim PowerSequence As Double
    Dim nwire_port1 As Double
    Dim nwire_port2 As Double
    Dim i As Long
    Dim PowerSequencePin() As String
    Dim TempMaxSequence As Long:: TempMaxSequence = 0
    
    Dim XO0_Port As New PinList
    Dim CLK32768_Port As New PinList
    Dim nwire01_name As String
    Dim nwire02_name As String
    
    Dim XI0_Pin As String
    Dim XI0_SeqName As String
    Dim XI0_Seq As Long
    Dim RTCLK_Pin As String
    Dim RTCLK_SeqName As String
    Dim RTCLK_Seq As Long

    
    nwire01_name = vbNullString
    nwire02_name = vbNullString
    
    '/////1226///
    ''  ------------------- 20180305 nWire pin form nWire_Ports_GLB ---------------------------
    Dim nWire_port_ary() As String
'    Dim i As Integer
    nWire_port_ary = Split(nWire_Ports_GLB, ",")
    For i = 0 To UBound(nWire_port_ary)
        If LCase(nWire_port_ary(i)) Like "*diff*" Then ' Diff nWire pin
            If LCase(nWire_port_ary(i)) Like "rt*" Then
                RTCLK_Pin = nWire_port_ary(i)                                   '//nWire port name
                nwire01_name = "RT_CLK32768_Diff_Port_PowerSequence_GLB"
            Else
                XI0_Pin = nWire_port_ary(i)
                nwire02_name = "XO0_Diff_Port_PowerSequence_GLB"      '//GB sequence number
            End If
        Else 'SE nWire pin
            If LCase(nWire_port_ary(i)) Like "rt*" Then
                RTCLK_Pin = nWire_port_ary(i)                                   '//nWire port name
                nwire01_name = "RT_CLK32768_Port_PowerSequence_GLB"   '//GB sequence number
            Else
                XI0_Pin = nWire_port_ary(i)
                nwire02_name = "XO0_Port_PowerSequence_GLB"           '//GB sequence number
            End If
        End If
    Next i

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
     '///1226
     TheHdw.Utility.Pins("k0,k1").State = tlUtilBitOn
     '///1226
    
    Call Print_Header("Power up sequence_Interpose")

    TheHdw.DCVS.Pins(PowerPinList).Voltage.Main = 0  'reset to 0V
    
    TheExec.DataManager.DecomposePinList PowerPinList, Pins(), PinCnt

    ReDim PowerSequencePin(PinCnt)  'get pin count numbers to arrange array's memory

    For Each powerPin In Pins
        TempString = vbNullString
        PowerName = CStr(powerPin)

        'get power sequence global spec
        TempString = PowerName & "_PowerSequence_GLB"
        PowerSequence = TheExec.Specs.Globals(TempString).ContextValue

        If TheExec.DataManager.ChannelType(powerPin) <> "N/C" Then
            If PowerSequence <> 99 Then
                'string power sequence pin
                If PowerSequencePin(PowerSequence) = "" Then
                    PowerSequencePin(PowerSequence) = PowerName
                Else
                    PowerSequencePin(PowerSequence) = PowerSequencePin(PowerSequence) & "," & PowerName
                End If
                If PowerSequence >= TempMaxSequence Then TempMaxSequence = PowerSequence
            'sequence 99, means disconnect pins
            Else
                'TheHdw.DCVS.Pins(PowerPin).Disconnect ' it cause voltage spike, removed it
                VMain = TheHdw.DCVS.Pins(powerPin).Voltage.Main.value
                Irange = TheHdw.DCVS.Pins(powerPin).CurrentRange.value
                If DebugFlag = True Then    'debugprint
                End If
            End If
        Else
            VMain = 0   'Can not read from DCVS
            Irange = 0
            If DebugFlag = True Then    'debugprint
            End If
        End If
    Next powerPin
    
    '////1226
    nwire_port1 = TheExec.Specs.Globals(nwire01_name).ContextValue
    nwire_port2 = TheExec.Specs.Globals(nwire02_name).ContextValue
    '///1226
    Dim clk_value As Double
    Dim sites As Variant
    For Each sites In TheExec.sites
        If TheExec.DevChar.Setups.item(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes.Contains(tlDevCharShmooAxis_Y) Then
            If UCase(TheExec.DevChar.Setups.item(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes.item(tlDevCharShmooAxis_Y).Parameter.name) Like "X?#*" Then
                clk_value = TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_Y).value
            End If
        End If
        
        If UCase(TheExec.DevChar.Setups.item(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes.item(tlDevCharShmooAxis_X).Parameter.name) Like "X?#*" Then
            clk_value = TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_X).value
    End If
    Next sites
    
    For i = 0 To PinCnt
        If PowerSequencePin(i) <> "" Then
        ''power up
        'theexec.Datalog.WriteComment vbCrLf & "print: power up action(" & i & ")" & vbCrLf & RepeatChr("*", 120)
        DCVS_PowerOn_I_Meter_Parallel PowerSequencePin(i), WaitConnectTime, WaitConnectTime, i, DebugFlag
        
        '///1226
            If nwire_port1 = i Then
                TheExec.Datalog.WriteComment vbCrLf & "print: power up for nwire(" & i & ")" & vbCrLf & RepeatChr("*", 120)
                PowerUp_Interpose CLK32768_Port, DebugFlag
            End If
            If nwire_port2 = i Then
                TheExec.Datalog.WriteComment vbCrLf & "print: power up for nwire(" & i & ")" & vbCrLf & RepeatChr("*", 120)
                PowerUp_Interpose XO0_Port, DebugFlag
                '//1226
                Call VaryFreq("XO0_Port", clk_value, "XO0_Freq_Var")
                '//1226
            End If
        '////1226
        
        ''power up
        End If
    Next i
    
    
    Call Print_Footer("Power up sequence_Interpose")
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Shmoo", "DCVS_PowerUp_Parallel_Interpose")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function DCVS_PowerDown_Parallel_Interpose(PowerPinList As String, Optional WaitConnectTime As Double = 0.001, Optional DebugFlag As Boolean = True) ', _
                Optional DriveLowPinList As PinList, Optional ClockPat As Pattern, Optional RTCLK_Relay As PinList, Optional XI0_Relay As PinList)
On Error GoTo errHandler
    Dim CurrentChans As String
    Dim Pins() As String, PinCnt As Long
    Dim powerPin As Variant
    Dim PowerName As String
    Dim TempString As String
    Dim VMain As Double
    Dim Irange As Double
    Dim step As Integer
    Dim FallTime As Double
    Dim PowerSequence As Double
    Dim i As Long
    Dim site As Variant
    
    Dim RTCLK_Relay As New PinList
    Dim XI0_Relay As New PinList
    
    Dim PowerSequencePin() As String
    Dim seqnum As Long
    Dim TempMaxSequence As Long:: TempMaxSequence = 0
    
    Dim XI0_Pin As String
    Dim XI0_SeqName As String
    Dim XI0_Seq As Long
    Dim RTCLK_Pin As String
    Dim RTCLK_SeqName As String
    Dim RTCLK_Seq As Long
    
    Call Print_Header("Power down sequence_Interpose")
    
    TheExec.DataManager.DecomposePinList PowerPinList, Pins(), PinCnt
    ReDim PowerSequencePin(PinCnt)

    For Each powerPin In Pins
        TempString = vbNullString
        PowerName = CStr(powerPin)
        
        'get power sequence global spec
        TempString = PowerName & "_PowerSequence_GLB"
        PowerSequence = TheExec.Specs.Globals(TempString).ContextValue

        If TheExec.DataManager.ChannelType(powerPin) <> "N/C" Then 'check CP or FT NC pins
            If PowerSequence <> 99 Then
                If PowerSequencePin(PowerSequence) = "" Then
                    PowerSequencePin(PowerSequence) = PowerName
                Else
                    PowerSequencePin(PowerSequence) = PowerSequencePin(PowerSequence) & "," & PowerName
                End If
                If PowerSequence >= TempMaxSequence Then TempMaxSequence = PowerSequence
            'sequence 99, means disconnect pins
            Else
                TheHdw.DCVS.Pins(powerPin).Disconnect
                VMain = TheHdw.DCVS.Pins(powerPin).Voltage.Main.value
                Irange = TheHdw.DCVS.Pins(powerPin).CurrentRange.value
                
                If DebugFlag = True Then    'debugprint
                
                End If
            End If
        'NC pins, does not need to power off
        Else
            VMain = 0   'Can not read from DCVS
            Irange = 0
            If DebugFlag = True Then    'debugprint
            End If
        End If
    Next powerPin
    
'/////////////////////////////////////////////////////////////////////////////////////////
    
    For i = PinCnt To 0 Step -1
        If PowerSequencePin(i) <> "" Then
            ''power up
            DCVS_PowerOff_I_Meter_Parallel PowerSequencePin(i), WaitConnectTime, WaitConnectTime, i, DebugFlag 'WaitConnectTime, WaitConnectTime, i, DebugFlag
            ''power up
        End If
    Next i
'/////////////////////////////////////////////////////////////////////////////////////////
    
    Call Print_Footer("Power down sequence")

    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Shmoo", "DCVS_PowerDown_Parallel_Interpose")
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function StoreMaxNum(argc As Long, argv() As String)
On Error GoTo errHandler

    Dim DevSetupName As String
    Dim RangeFrom(2) As Double, RangeTo(2) As Double, RangeStepSize(2) As Double, RangeSteps(2) As Long
    Dim RangeLow(2) As Double
    Dim RangeCalcType(2) As tlDevCharRangeField
    Dim curr_axis As Variant
    Dim Index_val As Long
    Dim sheetName As String
    Dim TNum_column As Long
    Dim i, j, k As Long
    Dim Tracking_Item As Variant
    Dim tmpstr As String
    Dim site As Variant
    
    
'    Count_Point = 0
    X_Point = 0
    Y_Point = 0
    Z_Point = 0
    Xaxis_index = 0
    Yaxis_index = 0
    Zaxis_index = 0
    X_Tracking_Point = 0
    Y_Tracking_Point = 0
    Z_Tracking_Point = 0
    
    LVCC_flag = False
    HVCC_flag = False
'    Exit Function
    X_dimemsion = False
    Y_dimemsion = False
    Z_dimemsion = False
    Multi_Axis_PTR = True
    
    tmpstr = UCase(Join(argv))
    
    ' Enable which axis need to print out HVCC/LVCC
    If InStr(tmpstr, "X") <> 0 Then X_dimemsion = True
    If InStr(tmpstr, "Y") <> 0 Then Y_dimemsion = True
    If InStr(tmpstr, "Z") <> 0 Then Z_dimemsion = True
    
    sheetName = TheExec.Flow.Raw.SheetInRun
    TNum_column = TheExec.Flow.Raw.GetCurrentLineNumber + 5
'    g_TestNum = Sheets(SheetName).Cells(TNum_column, 10)
    g_TestNum = TheExec.sites.item(site).TestNumber

    
    DevSetupName = TheExec.DevChar.Setups.ActiveSetupName
    
    For Each curr_axis In TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes.list
        With TheExec.DevChar
            RangeFrom(curr_axis) = .Setups(DevSetupName).Shmoo.Axes(curr_axis).Parameter.range.from
            RangeTo(curr_axis) = .Setups(DevSetupName).Shmoo.Axes(curr_axis).Parameter.range.to
            RangeSteps(curr_axis) = .Setups(DevSetupName).Shmoo.Axes(curr_axis).Parameter.range.Steps
            RangeStepSize(curr_axis) = .Setups(DevSetupName).Shmoo.Axes(curr_axis).Parameter.range.StepSize

            If RangeFrom(curr_axis) < RangeTo(curr_axis) Then
                RangeSeq(curr_axis) = True ' small---->lager
            Else
                RangeSeq(curr_axis) = False
            End If
            Index_val = RangeSteps(curr_axis)
            
            Select Case curr_axis
                Case 0
                    Xaxis_index = Index_val + 1 'Index Num = Steps + 1 --- ex:0.4,0.3,0.2,0.1 ---> Steps = 3, Index Num = 4
                    X_Tracking_Point = TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Count
                Case 1
                    Yaxis_index = Index_val + 1
                    Y_Tracking_Point = TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Count
                Case 2
                    Zaxis_index = Index_val + 1
                    Z_Tracking_Point = TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Count
            End Select
        End With
    Next curr_axis
    If TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes.Count = 3 Then '  3D shmoo only
        MaxArrIndex = Xaxis_index * Yaxis_index * Zaxis_index
        ReDim g_ShmooResult.Axis_CurrPoint(MaxArrIndex - 1)
        If X_Tracking_Point = 0 And Y_Tracking_Point = 0 And Z_Tracking_Point = 0 Then ' means no any tracking parameters.
            'Do nothing
        Else
            For k = 0 To MaxArrIndex - 1
                If Not X_Tracking_Point = 0 Then
                    ReDim g_ShmooResult.Axis_CurrPoint(k).X_axis_Tracking(X_Tracking_Point - 1)
                End If

                If Not Y_Tracking_Point = 0 Then
                    ReDim g_ShmooResult.Axis_CurrPoint(k).Y_axis_Tracking(Y_Tracking_Point - 1)
                End If

                If Not Z_Tracking_Point = 0 Then
                    ReDim g_ShmooResult.Axis_CurrPoint(k).Z_axis_Tracking(Z_Tracking_Point - 1)
                End If
            Next k
        End If
    ElseIf TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes.Count = 2 Then ' 2D shmoo case
        MaxArrIndex = Xaxis_index * Yaxis_index ' only X-Y case
        ReDim g_ShmooResult.Axis_CurrPoint(MaxArrIndex - 1)
        If X_Tracking_Point = 0 And Y_Tracking_Point = 0 Then  ' means no any tracking parameters.
            'Do nothing
        Else
            For k = 0 To MaxArrIndex - 1
                If Not X_Tracking_Point = 0 Then
                    ReDim g_ShmooResult.Axis_CurrPoint(k).X_axis_Tracking(X_Tracking_Point - 1)
                End If

                If Not Y_Tracking_Point = 0 Then
                    ReDim g_ShmooResult.Axis_CurrPoint(k).Y_axis_Tracking(Y_Tracking_Point - 1)
                End If
            Next k
        End If
    End If

    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Shmoo", "StoreMaxNum")
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function InverStr(SELSRAM_DSSC As String) As String
On Error GoTo errHandler

    Dim outputStr As String
    Dim i, j As Integer
    Dim StrLength As Long
    Dim tempChar As String

  
  outputStr = vbNullString
  
  If SELSRAM_DSSC <> "" Then
    StrLength = Len(SELSRAM_DSSC)
    
    For i = 1 To StrLength
         tempChar = mid(SELSRAM_DSSC, i, 1)
         
         If tempChar = "0" Then
            tempChar = "1"
         ElseIf tempChar = "1" Then
            tempChar = "0"
         ElseIf tempChar = "S" Then
            tempChar = "S"
         Else
            TheExec.Datalog.WriteComment "<Warning> The SELESRAM String has an unrecognizable character!!!"
         End If
         
         outputStr = outputStr & tempChar
         
    Next i
    SELSRAM_DSSC = outputStr
    InverStr = outputStr
  End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Shmoo", "InverStr")
    If AbortTest Then Exit Function Else Resume Next
End Function



Public Function CharacterizationDigSrcPrePoint(argc As Long, argv() As String)
On Error GoTo errHandler
    ' Dylan Edited by 20190726
    Dim site As Variant
    Dim i, j As Integer
    Dim Dec2BinInt As Integer
    Dim StringTemp As String
    Dim StringSplit() As String
    Dim GlobalInt As New SiteLong
    Dim DigSrcDSPWave() As New DSPWave
    ReDim DigSrcDSPWave(argc - 1)
    
    For i = 0 To argc - 1
        StringTemp = argv(i)
        StringTemp = Replace(StringTemp, "[", ":")
        StringTemp = Replace(StringTemp, "]", vbNullString)
        StringSplit = Split(StringTemp, ":")
        DigSrcDSPWave(i).CreateConstant 0, StringSplit(1), DspLong
        For Each site In TheExec.sites.Active
        
            If TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.type.value = "Global Spec" And _
                LCase(TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.name) Like "*digsrc*" Then
                GlobalInt = TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_X).value
            Else
                GlobalInt = TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_Y).value
            End If
            
            For j = 0 To CInt(StringSplit(1)) - 1
                Dec2BinInt = GlobalInt Mod 2
                DigSrcDSPWave(i).Element(j) = Dec2BinInt
                GlobalInt = GlobalInt \ 2
            Next j
            
            AddStoredCaptureData CStr(StringSplit(2)), DigSrcDSPWave(i)
        Next site
    Next i
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Shmoo", "CharacterizationDigSrcPrePoint")
    If AbortTest Then Exit Function Else Resume Next
End Function
