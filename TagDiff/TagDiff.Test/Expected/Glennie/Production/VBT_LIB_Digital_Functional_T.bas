Attribute VB_Name = "VBT_LIB_Digital_Functional_T"

Option Explicit
'Revision History:
'V0.0 initial bring up
'V0.1 Add Mbist finger print function

' Digital Functional Test

' (c) Teradyne, Inc, 1997-2008
'     All Rights Reserved
' Inclusion of a copyright notice does not imply that this software has been
' published.
' This software is the trade secret information of Teradyne, Inc.
' Use of this software is only in accordance with the terms of a license
' agreement from Teradyne, Inc.
'
' Revision History:
' Date                  Description
' 12-jun-12 s.bullock           tersw00187233 - locally save/restored some member variables in functioal_t to handle concurrent test use cases.
' 27-Mar-12 Venkata Kotireddy tersw00184533 - Fixed the issue to not save/restore flag match if the user did not specify flags to match when APS is enabled, and not suspended
' 13-Mar-12 R.Stimson   tersw00184562 - Restore error handler after Patterns.Test.
' 03-Jul-11 Obula Reddy   tersw00172420 - added getdefaults() to set the defaule value for template arguments
' 08-Jul-10 Pavan         tersw00166339  Added validation support for WaitTimeDomain.
' 09/10/09  David Sanders tersw00146124 Template code for patgen flag matching
'                         does not work with multiple time domains
' 01/06/09  Tim Orr     tersw001334426, slowdown in template test groups: see FetchContext
' 09/1/2005             Ported from Flex


' ============
' Private Data
' ============

' Context values on the Test Instances sheet
Private m_TimeSetSheet As String, m_LevelsSheet As String
Private m_InstanceName As String

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

' ============
' Public Enums
' ============

' CPU flag wait conditions
'Public Enum tlWaitVal
'    waitoff = -2    ' default value is first
'    waitLo = 0
'    waitHi = -1
'End Enum

Public Enum CusWaitVal
    waitoff = 0    ' default value is first
    waitLo = -1
    waitHi = -2
End Enum

Private Const TL_E_AT_PATSET_BREAKPT = &HC0000014

''''20151106 Set Variable for Functional_T_char_Mbist()
Private gm_Patterns As String
Private gm_bistType As String
Private gm_Power_Run_Scenario As String
Private gm_CharInputString As String
Private gm_AI_fail_point As String
Private gm_testName As String
Private gm_patcnt As Long
Private gm_rtnINITpattArr() As String
Private gm_rtnPLLPpattArr() As String
Private gm_wait_time_ary() As String
Public gB_shmooAccumResult As New SiteLong ''''20151110 New for the Accumlated shmoo result for the multi-patterns
Private gm_freqPattSet As New Pattern ''''20151111 New
Public Block As String
Public mbist_flag_set_placement As Long

Public HarvFailCnt As New SiteLong
''''''Public HarvOtherFailCnt As New SiteLong

Public HarvPinsFailCnt As New SiteLong

Private m_ScanLogging As Boolean

Public Enum tlTemplateScanFailDataLogging
    templateScanFailDataLoggingDisabled ' default value is first
    templateScanFailDataLoggingEnabled
    templateScanFailDataLoggingUseDataLogValue
End Enum
Public Enum tlTemplateScanPinListSource
    templateScanPinListSourceUseDatalogValue ' default value is first
    templateScanPinListSourceUseTemplateValue
    templateScanPinListSourceUsePatternValue
End Enum
Public Enum tlTemplateScanCaptureFormat
    templateScanCaptureFormatUseDatalogValue ' default value is first
    templateScanCaptureFormatCycle
    templateScanCaptureFormatVector
    templateScanCaptureFormatCycleAndVector
End Enum
Public Enum tlTemplateScanCaptureDataType
    templateScanCaptureDataTypePassFail ' default value is first
    templateScanCaptureDataTypePassFailExpect
    templateScanCaptureDataTypePassFailExpectActual
End Enum
Public Enum tlTemplateScanUserCommentSource
    templateScanUserCommentSourceUseDatalogValue ' default value is first
    templateScanUserCommentSourceUseTemplateValue
End Enum
Public Enum tlTemplateATPGPinMapSource
    templateATPGPinMapSourceNone    ' default value is first
    templateATPGPinMapSourceUseDatalogValue
    templateATPGPinMapSourceUseTemplateValue
End Enum
Private m_ScanLoggingEnabled As Boolean
Private m_ScanPinList As String
Private m_ScanCaptureFormat As tlDatalogScanCaptureFormat
Private m_CaptureSource As tlCMEMCaptureSource
Private m_ScanUserComment As String
Private m_ScanATPGPinmapUsed As Boolean
Private m_ScanATPGPinmapSourcePattern As String
Private m_ScanBurstEnabled As Boolean
Public g_ScanLoggingFeature As Boolean

Private Const TL_SCAN_CAPTURE_LIMIT_DEFAULT = 1000

Private Sub ScanLoggingSetup(ScanFailDataLogging As tlTemplateScanFailDataLogging, ScanCaptureLimitMode As tlDigitalCMEMCaptureLimitMode, _
                    ScanCaptureLimitPerPin As Long, ScanPinListSource As tlTemplateScanPinListSource, _
                    ScanPinList As PinList, ScanCaptureFormat As tlTemplateScanCaptureFormat, _
                    ScanCaptureDataType As tlTemplateScanCaptureDataType, ScanUserCommentSource As tlTemplateScanUserCommentSource, _
                    ScanUserComment As String, ATPGPinMapSource As tlTemplateATPGPinMapSource, ATPGPinMapSourcePattern As String, ScanTimeDomain As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    If glb_TesterType = "UltraFLEXplus" Then
        m_ScanLoggingEnabled = TheExec.Datalog.Setup.ScanSetup.EnableScanLogging
        m_ScanPinList = TheExec.Datalog.Setup.ScanSetup.PinList
        m_ScanCaptureFormat = TheExec.Datalog.Setup.ScanSetup.CaptureFormat
        m_ScanUserComment = TheExec.Datalog.Setup.ScanSetup.COMMENT
        m_ScanATPGPinmapUsed = TheExec.Datalog.Setup.ScanSetup.ATPGPinmap.Used
        m_ScanATPGPinmapSourcePattern = TheExec.Datalog.Setup.ScanSetup.ATPGPinmap.SourcePattern
        m_ScanBurstEnabled = TheHdw.Digital.TimeDomains(ScanTimeDomain).Patgen.ScanBurstEnabled
        TheExec.Datalog.Setup.ScanSetup.EnableScanLogging = True
        TheHdw.Digital.TimeDomains(ScanTimeDomain).Patgen.ScanBurstEnabled = True
        TheHdw.Digital.TimeDomains(ScanTimeDomain).CMEM.CaptureLimitMode = ScanCaptureLimitMode
        If ScanCaptureLimitPerPin = 0 Then
            TheHdw.Digital.TimeDomains(ScanTimeDomain).CMEM.CaptureLimit = TheExec.Datalog.Setup.ScanSetup.CMEMCaptureLimit
        Else
            TheHdw.Digital.TimeDomains(ScanTimeDomain).CMEM.CaptureLimit = ScanCaptureLimitPerPin
        End If
        If ScanPinListSource = templateScanPinListSourceUseTemplateValue Then
            TheExec.Datalog.Setup.ScanSetup.PinList = ScanPinList
        Else
            If ScanPinListSource = templateScanPinListSourceUsePatternValue Then
                TheExec.Datalog.Setup.ScanSetup.PinList = vbNullString
            End If
        End If
        Select Case ScanCaptureFormat
            Case tlTemplateScanCaptureFormat.templateScanCaptureFormatUseDatalogValue
            Case tlTemplateScanCaptureFormat.templateScanCaptureFormatCycle
                TheExec.Datalog.Setup.ScanSetup.CaptureFormat = tl_DCScanCaptureFormat_Cycle
            Case tlTemplateScanCaptureFormat.templateScanCaptureFormatVector
                TheExec.Datalog.Setup.ScanSetup.CaptureFormat = tl_DCScanCaptureFormat_Vector
            Case tlTemplateScanCaptureFormat.templateScanCaptureFormatCycleAndVector
                TheExec.Datalog.Setup.ScanSetup.CaptureFormat = tl_DCScanCaptureFormat_CycleVector
        End Select
        Select Case ScanCaptureDataType
            Case tlTemplateScanCaptureDataType.templateScanCaptureDataTypePassFail
                m_CaptureSource = tlCMEMCaptureSourcePassFailData
            Case tlTemplateScanCaptureDataType.templateScanCaptureDataTypePassFailExpect
                m_CaptureSource = tlCMEMCaptureSource_PatPassFailData
            Case tlTemplateScanCaptureDataType.templateScanCaptureDataTypePassFailExpectActual
                m_CaptureSource = tlCMEMCaptureSource_PatDutData
        End Select
        Call TheHdw.Digital.TimeDomains(ScanTimeDomain).CMEM.SetCaptureConfig(-1, CmemCaptFail, m_CaptureSource)
        If ScanUserCommentSource = templateScanUserCommentSourceUseTemplateValue Then
            TheExec.Datalog.Setup.ScanSetup.COMMENT = ScanUserComment
        End If
        If ATPGPinMapSource = tlTemplateATPGPinMapSource.templateATPGPinMapSourceUseTemplateValue Then
            TheExec.Datalog.Setup.ScanSetup.ATPGPinmap.Used = True
            TheExec.Datalog.Setup.ScanSetup.ATPGPinmap.SourcePattern = ATPGPinMapSourcePattern
        Else
            If ATPGPinMapSource = tlTemplateATPGPinMapSource.templateATPGPinMapSourceNone Then
                TheExec.Datalog.Setup.ScanSetup.ATPGPinmap.Used = False
            End If
        End If
        TheExec.Datalog.Setup.ScanSetup.Conditions.RemoveAll
    End If
Exit Sub 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Functional_T", "ScanLoggingSetup") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/08/18
End Sub


Private Sub ScanLoggingRestore(ScanFailDataLogging As tlTemplateScanFailDataLogging, ScanCaptureLimitMode As tlDigitalCMEMCaptureLimitMode, _
                    ScanCaptureLimitPerPin As Long, ScanPinListSource As tlTemplateScanPinListSource, _
                    ScanPinList As PinList, ScanCaptureFormat As tlTemplateScanCaptureFormat, _
                    ScanCaptureDataType As tlTemplateScanCaptureDataType, ScanUserCommentSource As tlTemplateScanUserCommentSource, _
                    ScanUserComment As String, ATPGPinMapSource As tlTemplateATPGPinMapSource, ATPGPinMapSourcePattern As String, ScanTimeDomain As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    If glb_TesterType = "UltraFLEXplus" Then
        Call TheHdw.Digital.TimeDomains(ScanTimeDomain).CMEM.SetCaptureConfig(-1, CmemCaptNone, m_CaptureSource)
        TheExec.Datalog.Setup.ScanSetup.EnableScanLogging = m_ScanLoggingEnabled
        TheHdw.Digital.TimeDomains(ScanTimeDomain).CMEM.CaptureLimitMode = tlDigitalCMEMCaptureLimitMode_Disable
        TheHdw.Digital.TimeDomains(ScanTimeDomain).CMEM.CaptureLimit = TL_SCAN_CAPTURE_LIMIT_DEFAULT
        TheExec.Datalog.Setup.ScanSetup.PinList = m_ScanPinList
        TheExec.Datalog.Setup.ScanSetup.CaptureFormat = m_ScanCaptureFormat
        TheExec.Datalog.Setup.ScanSetup.COMMENT = m_ScanUserComment
        TheExec.Datalog.Setup.ScanSetup.ATPGPinmap.Used = m_ScanATPGPinmapUsed
        TheExec.Datalog.Setup.ScanSetup.ATPGPinmap.SourcePattern = m_ScanATPGPinmapSourcePattern
        TheHdw.Digital.TimeDomains(ScanTimeDomain).Patgen.ScanBurstEnabled = m_ScanBurstEnabled
    End If
Exit Sub 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Functional_T", "ScanLoggingRestore") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/08/18
End Sub


' Perform a digital functional test.
' Return TL_SUCCESS if the test executes without problems, else TL_ERROR.
' [20230420][All][Wyatt] Fix use "like" compare null pattern name
' [20230525][All][Jim] Remove ForceResults Argument and add MultiFSTP_Enable Argument to Functional_T_updated (For new MTFSTP)
' [20230621][All][Jim] Set Result Mode to tlResultModeDomain for UFLEX, because of tlResultModeModule mode only return the last failed site result when pattern fail over 2 site
' [20230801][All][Jim] Change ResultMode to tlResultModeDomain, just for Harvest.
' [20230809][All][Wyatt] move set BinCut_Payload_Voltage from pre body to body, due to shmoo setup problem.
' [20230907][All][Oliver] modify for only DCVS pin use tlDCVSVoltageMain
' [20231228][T-Tah][Tank] Add process SFC
Public Function Functional_T_updated(patterns As Pattern, StartOfBodyF As InterposeName, _
                             PrePatF As InterposeName, PreTestF As InterposeName, _
                             PostTestF As InterposeName, PostPatF As InterposeName, EndOfBodyF As InterposeName, _
                             ReportResult As PFType, ResultMode As tlResultMode, DriveLoPins As PinList, DriveHiPins As PinList, _
                             DriveZPins As PinList, DisablePins As PinList, FloatPins As PinList, StartOfBodyFArgs As String, _
                             PrePatFArgs As String, PreTestFArgs As String, PostTestFArgs As String, _
                             PostPatFArgs As String, EndOfBodyFArgs As String, Util1Pins As PinList, _
                             Util0Pins As PinList, PatFlagF As InterposeName, _
                             PatFlagFArgs As String, RelayMode As tlRelayMode, PatThreading As Boolean, _
                             MatchAllSites As Boolean, WaitFlagA As CusWaitVal, WaitFlagB As CusWaitVal, _
                             WaitFlagC As CusWaitVal, WaitFlagD As CusWaitVal, Validating_ As Boolean, _
                             Optional PatternTimeout As String = "30", Optional Step_ As SubType, _
                             Optional WaitTimeDomain As String, Optional ConcurrentMode As tlPatConcurrentMode = tlPatConcurrentModeCached, _
                             Optional Interpose_PrePat As String, _
                             Optional RunFailCycle As Boolean = False, Optional EnableBinOut As Boolean = False, Optional PFAMultiFunc As Boolean = False, _
                             Optional DigSource As String = vbNullString, Optional DSSCSetup As String, Optional ApplyVoltageFromBinCut As String = vbNullString, _
                             Optional MbistMatchLoopCountValue As Long = 0, _
                             Optional ScanFailDataLogging As tlTemplateScanFailDataLogging = templateScanFailDataLoggingDisabled, Optional ScanCaptureLimitMode As tlDigitalCMEMCaptureLimitMode = tlDigitalCMEMCaptureLimitMode_Enable, _
                             Optional ScanCaptureLimitPerPin As Long = 0, Optional ScanPinListSource As tlTemplateScanPinListSource = templateScanPinListSourceUseDatalogValue, _
                             Optional ScanPinList As PinList, Optional ScanCaptureFormat As tlTemplateScanCaptureFormat = templateScanCaptureFormatUseDatalogValue, _
                             Optional ScanCaptureDataType As tlTemplateScanCaptureDataType = templateScanCaptureDataTypePassFail, Optional ScanUserCommentSource As tlTemplateScanUserCommentSource = templateScanUserCommentSourceUseDatalogValue, _
                             Optional ScanUserComment As String, Optional ATPGPinMapSource As tlTemplateATPGPinMapSource = templateATPGPinMapSourceNone, _
                             Optional ATPGPinMapSourcePattern As String, Optional ScanTimeDomain As String, Optional HarvestPinGrpOtherFail As String = vbNullString, Optional Harv_DigSrc As String = vbNullString, Optional Harv_FailFlag As String = vbNullString, Optional MultiFSTP_Enable As Boolean = False) As Long
' EDITFORMAT1 1,,Pattern,,,Patterns|7,,InterposeName,Interpose Functions,,StartOfBodyF|9,,InterposeName,,,PrePatF|11,,InterposeName,,,PreTestF|13,,InterposeName,,,PostTestF|15,,InterposeName,,,PostPatF|17,,InterposeName,,,EndOfBodyF|2,,PFType,,,ReportResult|6,,tlResultMode,,,ResultMode|19,,pinlist,Pin States,,DriveLoPins|20,,pinlist,,,DriveHiPins|21,,pinlist,,,DriveZPins|22,,pinlist,,,DisablePins|23,,pinlist,,,FloatPins|8,,String,,,StartOfBodyFArgs|10,,String,,,PrePatFArgs|12,,String,,,PreTestFArgs|14,,String,,,PostTestFArgs|16,,String,,,PostPatFArgs|18,,String,,,EndOfBodyFArgs|24,,pinlist,,,Util1Pins|25,,pinlist,,,Util0Pins|31,,InterposeName,,,PatFlagF|32,,String,,,PatFlagFArgs|5,,tlRelayMode,,,RelayMode|3,,Boolean,,,PatThreading|30,,Boolean,,,MatchAllSites|26,,CusWaitVal,Flag Match,,WaitFlagA|27,,CusWaitVal,,,WaitFlagB|28,,CusWaitVal,,,WaitFlagC|29,,CusWaitVal,,,WaitFlagD|0,,Boolean,,,Validating_|4,,String,,0 <= PatternTimeout,PatternTimeout|6,,tlPatStartConcurrentMode,,,ConcurrentMode
    ''Optional MbistIndicator As Boolean = False ''20220308
    Dim i As Variant
    Dim inst_info As Instance_Info '''www
    Dim voltage_forBinCut() As SiteDouble
    Dim result_modetmp As tlResultMode
    Dim PrintString As String
    Dim VoltageString As String
    
    Functional_T_updated = TL_SUCCESS   ' be optimistic
    If Not TheExec.flow.IsRunning Then Exit Function
    
    On Error GoTo errHandler
        
    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)
    
    ' Cache parameters for PostTest
    m_EndOfBodyF = EndOfBodyF
    m_EndOfBodyFArgs = EndOfBodyFArgs
    If TheExec.enableWord("Enable_ApplyVoltageFromBinCut") = False Then ApplyVoltageFromBinCut = vbNullString
    ' Apply default values to parameters whose values were not specified.
    ApplyDefaults PatternTimeout
    
    If Validating_ Then
        ' Perform additional parameter validation
        If Not Validate(patterns, PatThreading, DriveLoPins, DriveHiPins, DriveZPins, DisablePins, _
            FloatPins, Util1Pins, Util0Pins, PatternTimeout, WaitTimeDomain) Then Functional_T_updated = TL_ERROR
        If patterns.value <> "" Then Call PrLoadPattern(patterns.value)
'''        If PrePatF.Value <> "" Then Call PrLoadPattern(PrePatF.Value)
        Exit Function
    End If
    
    Call SFC_Main(StartOfBodyF, PostPatF, PostPatFArgs)
    
#If LCD = True Then
    If DSSCSetup <> "" Then myUtility.Initialize DSSCSetup:=DSSCSetup
#End If
    
    
    If ApplyVoltageFromBinCut <> "" Then
        Call initialize_inst_info(inst_info, ApplyVoltageFromBinCut)
        inst_info.selsrm_DigSrc_Pin = "JTAG_TDI"
        inst_info.selsrm_DigSrc_SignalName = "FUNC_SRC"
        'inst_info.inst_name = inst_info.inst_name & "_outsidebincut_bv"
    End If
    
    '''20240124: Added for new MFSTP
    inst_info.Harvest_Core_DigSrc_Pin = "JTAG_TDI"
    inst_info.Harvest_Core_DigSrc_SignalName = "Harvest_Core_DigSrcSignal"
    
    If Step_ = subAllBody Or Step_ = subPrebody Or _
        m_InterposeFunctionsSet = False Then

        ' Register certain interpose function names with flow controller
        Call tl_SetInterpose(TL_C_PREPATF, PrePatF.value, PrePatFArgs, _
                             TL_C_POSTPATF, PostPatF.value, PostPatFArgs, _
                             TL_C_PRETESTF, PreTestF.value, PreTestFArgs, _
                             TL_C_POSTTESTF, PostTestF.value, PostTestFArgs, _
                             TL_C_FLAGMATCHF, PatFlagF.value, PatFlagFArgs, _
                             TL_C_POSTPATBPF, "PostTest", vbNullString)

        m_InterposeFunctionsSet = True

    End If

    ' PreBody
    If Step_ = subAllBody Or Step_ = subPrebody Then
        FetchContext
        '==============================20180226 Vramp to prevent voltage spike==============================
        m_InstanceName = LCase(TheExec.DataManager.instancename)
        
        If UCase(m_InstanceName) Like UCase("SocMbist*") Or UCase(m_InstanceName) Like UCase("CpuMbist*") Or UCase(m_InstanceName) Like UCase("GfxMbist*") Then
            Call MbistRampApplyLevel_AutoReadingContext(, , , m_InstanceName)
        End If
        '===================================================================================================
         
        ' Set up the test
        Call PreBody(DriveHiPins, DriveLoPins, DriveZPins, DisablePins, Util1Pins, Util0Pins, _
                 WaitFlagA, WaitFlagB, WaitFlagC, WaitFlagD, MatchAllSites, _
                 PatThreading, RelayMode, WaitTimeDomain, vbNullString, _
                 ScanFailDataLogging, ScanCaptureLimitMode, ScanCaptureLimitPerPin, ScanPinListSource, _
                 ScanPinList, ScanCaptureFormat, ScanCaptureDataType, ScanUserCommentSource, ScanUserComment, _
                 ATPGPinMapSource, ATPGPinMapSourcePattern, ScanTimeDomain)
    End If ' PreBody
        
    Dim CurConcurrentContext As Long
    CurConcurrentContext = m_STDSvcClient.FlowDomainService.ConcurrentContext
    
'    If ApplyVoltageFromBinCut <> "" Then ''20220310
'        Call bincut_power_Setting_VT(inst_info, CurrentPassBinCutNum, BinCut_Payload_Voltage, True)
'        Call Set_PayloadVoltage_to_DCVS(Flag_Enable_Rail_Switch, pinGroup_BinCut, BinCut_Payload_Voltage)
'        PrintString = vbNullString
'        For Each site In theexec.sites
'            For i = 0 To UBound(pinGroup_CorePower)
'                'For Each site In theexec.sites
'                    'theexec.Datalog.WriteComment pinGroup_CorePower(i) & "   :  " & CStr(BinCut_Payload_Voltage(VddBinStr2Enum(pinGroup_CorePower(i))))
'                   ' theexec.Datalog.WriteComment pinGroup_OtherRail(i) & "   :  " & CStr(BinCut_Payload_Voltage(i + UBound(pinGroup_CorePower)))
'                PrintString = PrintString & pinGroup_CorePower(i) & ": " & CStr(BinCut_Payload_Voltage(VddBinStr2Enum(pinGroup_CorePower(i)))) & "mV, "
                    
'                'Next site
'            Next i
'            For i = 0 To UBound(pinGroup_OtherRail)
'                PrintString = PrintString & pinGroup_OtherRail(i) & ": " & CStr(BinCut_Payload_Voltage(VddBinStr2Enum(pinGroup_OtherRail(i)))) & "mV, "
'            Next i
'            theexec.Datalog.WriteComment "Site: " & site & ", " & PrintString
'            PrintString = vbNullString
'        Next site
        
'    End If
    
    
    ' Body
    If Step_ = subAllBody Or Step_ = subBody Then
        
        ' cache member variables
        ' there are statements below which can cause us to jump to the next subflow if we're running with concurrent test.
        ' if the next test in the next subflow runs this function then it will overwrite the below member variables, such
        ' that when we get back to this call they will have different values.  so we cache the values here and then
        ' restore them right after the code that can cause us to jump to the next subflow.  then later on in
        ' postbody and posttest when they're used they'll have the proper values.
        
        Dim tempendofbody As String
        Dim tempendofbodyfargs As String
        Dim tempdrivepins As String
        Dim tempfloatpins As String
        Dim site As Variant 'Carter, 20240304

        '20230808 Wyatt due to shmoo setup problem , this code can't get shmoo information in pre body . We moved to body
        If ApplyVoltageFromBinCut <> "" Then ''20220310
            Call bincut_power_Setting_VT(inst_info, CurrentPassBinCutNum, BinCut_Payload_Voltage, True)
            Call Set_PayloadVoltage_to_DCVS(Flag_Enable_Rail_Switch, pinGroup_BinCut, BinCut_Payload_Voltage)
            PrintString = vbNullString
            For Each site In TheExec.sites
                For i = 0 To UBound(pinGroup_CorePower)
                    'For Each site In theexec.sites
                        'theexec.Datalog.WriteComment pinGroup_CorePower(i) & "   :  " & CStr(BinCut_Payload_Voltage(VddBinStr2Enum(pinGroup_CorePower(i))))
                       ' theexec.Datalog.WriteComment pinGroup_OtherRail(i) & "   :  " & CStr(BinCut_Payload_Voltage(i + UBound(pinGroup_CorePower)))
                    PrintString = PrintString & pinGroup_CorePower(i) & ": " & CStr(BinCut_Payload_Voltage(VddBinStr2Enum(pinGroup_CorePower(i)))) & "mV, "
                    
                    'Next site
                Next i
                For i = 0 To UBound(pinGroup_OtherRail)
                    PrintString = PrintString & pinGroup_OtherRail(i) & ": " & CStr(BinCut_Payload_Voltage(VddBinStr2Enum(pinGroup_OtherRail(i)))) & "mV, "
                Next i
                TheExec.Datalog.WriteComment "Site: " & site & ", " & PrintString
                PrintString = vbNullString
            Next site
        End If

        If CurConcurrentContext Then
            tempendofbody = m_EndOfBodyF
            tempendofbodyfargs = m_EndOfBodyFArgs
            tempdrivepins = m_DrivePins
            tempfloatpins = m_FloatPins
        End If
                
        ' Perform the test
        Call Interpose(StartOfBodyF, StartOfBodyFArgs)
        
        '''20180621 for shmoo PTR high/low limit
        'Add for force condition.
        '2017/11/02 Add STORE Pre Pat String
        If Interpose_PrePat <> "" Then
            Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
        End If
        
        If (RelayMode = tlUnpowered) Then MsgBox "Please change relay mode to powered", vbOKOnly, "IG-XL alarm"
     
        Call Body(FloatPins, StrToDbl(PatternTimeout), patterns, ReportResult, ResultMode)
        
                
        ' Run the pattern.  Perform functional test.
        If TheExec.sites.ActiveCount > 0 Then
            On Error Resume Next
            
            Shmoo_Pattern = patterns.value

            '20240124: New MFSTP/UserFunction DigSrc defined by Alan
            Call CheckInstForUserFunction(Harv_DigSrc, inst_info.digSrcLabel, inst_info.digSrcPatterns, CStr(patterns))
            
            '20240124: Exist MFSTP pattern(s)
            If Harv_DigSrc <> "" And inst_info.digSrcPatterns(0) <> "" Then
                For i = 0 To UBound(inst_info.digSrcPatterns)
                    Call Calculate_Harvest_Core_DSSC_Source_For_UserFunction(CStr(inst_info.digSrcPatterns(i)), CStr(inst_info.digSrcLabel(i)), inst_info.Harvest_Core_DigSrc_Pin, inst_info.Harvest_Core_DigSrc_SignalName)
                Next i
            End If
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''2021Harvest
            'Call Sub_SourceEMA_SelSRM(Shmoo_Pattern, m_InstanceName, DigSource)
            Dim PatArray() As String
            Dim pat As Variant
            Dim PatCount As Long
            PatArray = TheExec.DataManager.Raw.GetPatternsInSet(Shmoo_Pattern, PatCount)
            If ApplyVoltageFromBinCut <> "" Then
                Call Check_and_Decompose_PrePatt_FuncPat(inst_info, result_modetmp, "true", vbNullString, patterns.value)
                'inst_info.patt_SelsrmDigSrc_single = CStr(Pat)
                Call Calculate_Selsrm_DSSC_For_BinCut(inst_info, CurrentPassBinCutNum)
                For Each site In TheExec.sites
                    If inst_info.str_Selsrm_DSSC_Info(site) <> "" Then
                        TheExec.Datalog.WriteComment inst_info.str_Selsrm_DSSC_Info(site)
                    End If
                Next site
            Else
                'PatArray = TheExec.DataManager.Raw.GetPatternsInSet(Shmoo_Pattern, PatCount)
                For Each pat In PatArray
                    If UCase(CStr(pat)) Like "*_DSRMDSSC_*" Or UCase(CStr(pat)) Like "*_DSRMDSSC*" Or UCase(CStr(pat)) Like "*_SRMDSSC*" Then
                        If DigSource <> "" Then
                            If CStr(pat) Like ".\*" Then
                                Call Sub_SourceEMA_SelSRM(Shmoo_Pattern, glb_TestInstance, DigSource)
                            Else
                                Call Sub_SourceEMA_SelSRM(CStr(pat), glb_TestInstance, DigSource)
                            End If
                        End If
                    ElseIf UCase(CStr(pat)) Like "*_DSSC*OR" Or UCase(CStr(pat)) Like "*ORDSSC" Then
                        If Harv_DigSrc <> "" Then
                            'Call Harvest_DigSrc(Harv_DigSrc, Pat)
                        End If
                    End If
                Next pat
            End If
            
            '20231211: Decide SSN enable/disable
            If Flag_HarvPinFlag_Mapping_Table_Parsed = True And ssnPatternsDict.Count > 0 Then
                Dim instSSNinfo As Inst_SSN
                Call CheckInstForSSN(PatArray, instSSNinfo, SSNMapping, Harv_FailFlag)
            End If
 
            If MbistMatchLoopCountValue > 0 Then
                Dim MatchLoopNum As Long
    
                MatchLoopNum = MbistMatchLoopCountValue
                TheHdw.Digital.Patgen.counter(tlPgCounter10) = MatchLoopNum
             
                '''====================================
                If TheExec.TesterMode = testModeOffline Then
                    TheHdw.Digital.Patgen.counter(tlPgCounter10) = 1
                End If
                '''====================================
            End If
    
            Dim Flag_find_Mbist_pl As Boolean
            Flag_find_Mbist_pl = False
            
            If ApplyVoltageFromBinCut <> "" Then
                If inst_info.Test_Type = testType.Mbist Then
                '''ex: "*cpu*bist*", "*gfx*bist*", "*gpu*bist*", "*soc*bist*".
                    '''====================================================================================================='''
                    '''C651 didn't implement Vbump op-code in MBIST init pattern for project with rail-switch.
                    '''So that we have to switch DCVS to Valt by VBT code here before running FuncPat for Mbist instances.
                    '''====================================================================================================='''
                    For Each pat In PatArray
                            'strAry_PatNameSplit = Split(LCase(ary_patt_decompose(i)), "_")
                            '''ToDo: Maybe we can check patterns with keywords "*_pllp*", "*_fulp*", and "*_pl*"...
                        If LCase(pat) Like "*pl*" Then
                            If Flag_Enable_Rail_Switch Then '''For projects with Rail Switch, BinCut payload voltages are applied to DCVS Valt.
                                select_DCVS_output_for_powerDomain tlDCVSVoltageAlt
                                inst_info.currentDcvsOutput = tlDCVSVoltageAlt
                                Flag_find_Mbist_pl = True
                            Else '''For conventional projects without Rail Switch, BinCut payload voltages(BV) are applied to DCVS Vmain.
                                select_DCVS_output_for_powerDomain tlDCVSVoltageMain
                                inst_info.currentDcvsOutput = tlDCVSVoltageMain
                            End If
                        End If
                        If Flag_find_Mbist_pl = True Then Exit For
                    Next pat
                End If
            End If
            
         
            If UCase(glb_TesterType) = UCase("Jaguar") And ResultMode <> tlResultModeDomain And Harv_FailFlag <> "" Then
                Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Functional_T", "Functional_T_updated", "[Result Mode Setting]:Ultra Flex Result Mode Change to tlResultModeDomain from " & "ResultMode:" & ResultMode)
                ResultMode = tlResultModeDomain
            End If
            Call pattern_module_test(patterns.value, RunFailCycle, EnableBinOut, ReportResult, TL_C_YES, ResultMode, ConcurrentMode, instSSNinfo, _
                                     Harv_FailFlag, HarvestPinGrpOtherFail, PFAMultiFunc, ApplyVoltageFromBinCut)   ' test chip block loop function
        
            If MbistMatchLoopCountValue > 0 Then
        
                Dim RealLoopNum As Long
                
                RealLoopNum = (MatchLoopNum - TheHdw.Digital.Patgen.counter(tlPgCounter10))
                TheExec.Datalog.WriteComment "Set C10 of " & TheExec.DataManager.instancename & " : " & MatchLoopNum & " run down to " & " : " & TheHdw.Digital.Patgen.counter(tlPgCounter10) & " Total Loop Counts " & " : " & RealLoopNum
        
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
        DebugPrintFunc patterns.value, False, , ApplyVoltageFromBinCut
        
        Dim PatSetArray() As String
        Dim PrintPatSet As Variant
        Dim patt_ary_debug() As String
        Dim pat_count_debug As Long
        Dim patt As Variant
        If False Then
            PatSetArray = Split(patterns.value, ",")
        
            For Each PrintPatSet In PatSetArray
                If LCase(PrintPatSet) Like "*.pat*" Then
                    TheExec.Datalog.WriteComment "  Pattern : " & PrintPatSet
                Else
                    GetPatListFromPatternSet CStr(PrintPatSet), patt_ary_debug, pat_count_debug
                    For Each patt In patt_ary_debug
                        If patt <> "" Then TheExec.Datalog.WriteComment "  Pattern : " & patt
                    Next patt
                End If
            Next PrintPatSet
        End If
        
        
        ' restore the member variables for postbody (do this here instead of a couple of lines above since posttest could
        ' possibly jump to the next subflow in a concurrent test and cause the below memeber variables to change again.
        If CurConcurrentContext Then
            m_DrivePins = tempdrivepins
            m_FloatPins = tempfloatpins
        End If
        
    End If ' Body
    
    ' PostBody
    If Step_ = subAllBody Or Step_ = subPostbody Then
    
    '2017/11/02 Add RESTORE Pre Pat String in post body
        If Interpose_PrePat <> "" Then
            Call SetForceCondition("RESTOREPREPAT")
        End If
        
        Call PostBody(m_DrivePins, m_FloatPins, WaitTimeDomain, WaitFlagA, WaitFlagB, WaitFlagC, WaitFlagD, _
                        ScanFailDataLogging, ScanCaptureLimitMode, ScanCaptureLimitPerPin, ScanPinListSource, _
                        ScanPinList, ScanCaptureFormat, ScanCaptureDataType, ScanUserCommentSource, ScanUserComment, _
                        ATPGPinMapSource, ATPGPinMapSourcePattern, ScanTimeDomain)
                 
        TheHdw.DCVS.pins(Core_Power_DCVS_pins).Voltage.Output = tlDCVSVoltageMain
    End If ' PostBody
    
    ' There shouldn't be any code below this line. Any other necessary
    ' code should be added to the PostTest method to support pattern set
    ' breakpoints.
        

   
    Exit Function
    
errHandler:
    Call TheExec.ErrorLogMessage("Test " & TL_C_ERRORSTR & ", Instance: " & TheExec.DataManager.instancename)
    Call TheExec.ErrorReport
    ' Clear previously registered interpose function names
    Call tl_ClearInterpose(TL_C_PREPATF, TL_C_POSTPATF, TL_C_PRETESTF, TL_C_POSTTESTF)
    m_InterposeFunctionsSet = False

    Functional_T_updated = TL_ERROR
    If AbortTest Then Exit Function Else Resume Next
End Function

' Perform a digital functional test.
' Return TL_SUCCESS if the test executes without problems, else TL_ERROR.




Public Function DatalogType() As Integer
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    DatalogType = logFunctional
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Functional_T", "DatalogType") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

' =====================
' Private work routines
' =====================

' Do test setup.  This involves setting the timing and levels, connecting the pins, and
' various other functions in preparation for running the pattern.
Private Sub PreBody(DriveHiPins As PinList, DriveLoPins As PinList, DriveZPins As PinList, DisablePins As PinList, _
                    Util1Pins As PinList, Util0Pins As PinList, WaitFlagA As CusWaitVal, _
                    WaitFlagB As CusWaitVal, WaitFlagC As CusWaitVal, WaitFlagD As CusWaitVal, MatchAllSites As Boolean, _
                    PatThreading As Boolean, RelayMode As tlRelayMode, _
                    WaitTimeDomain As String, CharInputString As String, _
                    Optional ScanFailDataLogging As tlTemplateScanFailDataLogging, Optional ScanCaptureLimitMode As tlDigitalCMEMCaptureLimitMode, _
                    Optional ScanCaptureLimitPerPin As Long, Optional ScanPinListSource As tlTemplateScanPinListSource, _
                    Optional ScanPinList As PinList, Optional ScanCaptureFormat As tlTemplateScanCaptureFormat, _
                    Optional ScanCaptureDataType As tlTemplateScanCaptureDataType, Optional ScanUserCommentSource As tlTemplateScanUserCommentSource, _
                    Optional ScanUserComment As String, Optional ATPGPinMapSource As tlTemplateATPGPinMapSource, Optional ATPGPinMapSourcePattern As String, Optional ScanTimeDomain As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim ConnectAllPins As Boolean, LoadLevels As Boolean, LoadTiming As Boolean

    ' Save previous state of pattern threading and set according to parameter.
    m_OldPatThreading = TheHdw.patterns().Threading.Enable
    TheHdw.patterns().Threading.Enable = PatThreading

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
    If UCase(TheExec.DataManager.instancename) Like "*_EVS_*" Then
        TheHdw.DCVS.pins("VDD_SOC").CurrentRange.value = 30
        TheHdw.DCVS.pins("VDD_PCPU").CurrentRange.value = 30
        TheHdw.DCVS.pins("VDD_GPU").CurrentRange.value = 30
    End If
    
    Call TheHdw.Digital.ApplyLevelsTiming(ConnectAllPins, LoadLevels, LoadTiming, RelayMode)
    
        If TheExec.DataManager.instancename Like "*WalkingZ_DC_Continuity*" And InStr(UCase(TheExec.CurrentChanMap), UCase("ChannelMap_CP_5_site")) <> 0 Then         'Modify for CP DCVI     2021/12/21
       ''With TheHdw.DCVI.Pins("VDD_FIXED_LPDP_RX_DCVI,VDD_FIXED_LPDP_TX_DCVI,VDD_FIXED_PCIE_DCVI,VDD12_LPDP_RX_DCVI,VDD12_PCIE_DCVI")
        With TheHdw.DCVI.pins("VDD_FIXED_LPDP_RX,VDD_FIXED_LPDP_TX,VDD_FIXED_PCIE_DCVI,VDD12_LPDP_RX_DCVI,VDD12_PCIE_DCVI")
            .mode = tlDCVIModeVoltage
            .SetCurrentAndRange 800 * mA, 1000 * mA
            TheHdw.Wait 2 * ms
            .Voltage = 0
            .Gate = True
            .Connect
        End With
    TheHdw.Wait 0.005
    End If
    
    If TheExec.DataManager.instancename Like "*WalkingZ_DC_Continuity*" And InStr(UCase(TheExec.CurrentChanMap), UCase("ChannelMap_WLFT_6_site")) <> 0 Then         'Modify for CP DCVI     2021/12/21
        With TheHdw.DCVI.pins("VDD_FIXED_LPDP_RX,VDD_FIXED_LPDP_TX")
            .mode = tlDCVIModeVoltage
            .SetCurrentAndRange 800 * mA, 1000 * mA
            TheHdw.Wait 2 * ms
            .Voltage = 0
            .Gate = True
            .Connect
        End With
    TheHdw.Wait 0.005
    End If
    
      If TheExec.DataManager.instancename Like "*WalkingZ_DC_Continuity*" And InStr(UCase(TheExec.CurrentChanMap), UCase("ChannelMap_FT_6_site")) <> 0 Then         'Modify for CP DCVI     2021/12/21
        With TheHdw.DCVI.pins("VDD_FIXED_LPDP_RX,VDD_FIXED_LPDP_TX")
            .mode = tlDCVIModeVoltage
            .SetCurrentAndRange 800 * mA, 1000 * mA
            TheHdw.Wait 2 * ms
            .Voltage = 0
            .Gate = True
            .Connect
        End With
    TheHdw.Wait 0.005
    End If
        
      '' 20150625 - Apply Char setup
'    If UCase(TheExec.CurrentJob) Like "*CHAR*" Then
'        If CharInputString <> "" Then
'            Call SetCharPower(CharInputString)
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
    If ((TheHdw.patterns.EnableAsyncPatternStart <> tlAsyncPatternModeDisabled) And (TheHdw.patterns.SuspendAsyncPatternStart = False)) Then
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
    g_ScanLoggingFeature = False
    If glb_TesterType = "UltraFLEXplus" Then
        If g_ScanLoggingFeature Then
            Select Case ScanFailDataLogging
                Case templateScanFailDataLoggingDisabled
                    m_ScanLogging = False
                Case templateScanFailDataLoggingEnabled
                    If TheExec.Datalog.Setup.DatalogSetup.DatalogOn Then
                        m_ScanLogging = True
                    Else
                        m_ScanLogging = False
                    End If
                Case templateScanFailDataLoggingUseDataLogValue
                    If (TheExec.Datalog.Setup.ScanSetup.EnableScanLogging And _
                        TheExec.Datalog.Setup.DatalogSetup.DatalogOn) Then
                        m_ScanLogging = True
                    Else
                        m_ScanLogging = False
                    End If
            End Select
            If m_ScanLogging Then
                Call ScanLoggingSetup(ScanFailDataLogging, ScanCaptureLimitMode, _
                    ScanCaptureLimitPerPin, ScanPinListSource, _
                    ScanPinList, ScanCaptureFormat, _
                    ScanCaptureDataType, ScanUserCommentSource, _
                    ScanUserComment, ATPGPinMapSource, ATPGPinMapSourcePattern, ScanTimeDomain)
            End If
        End If
    End If
Exit Sub 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Functional_T", "PreBody") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/08/18
End Sub


' Run the pattern and see if it passed or failed
Private Sub Body(FloatPins As PinList, PatternTimeout As Double, patterns As Pattern, _
                 ReportResult As PFType, ResultMode As tlResultMode)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    ' Remove specified DUT pins, if any, from connection to tester pin-electronics and other resources
    If NonBlank(FloatPins) Then Call tl_SetFloatState(FloatPins)
    m_FloatPins = FloatPins.value
    
    ' Enable the pattern timeout counter
    TheHdw.Digital.Patgen.TimeoutEnable = True
    TheHdw.Digital.Patgen.TimeOut = PatternTimeout
Exit Sub 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Functional_T", "Body") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/08/18
End Sub


' Restore tester state to the default
Private Sub PostBody(DrivePins As String, FloatPins As String, WaitTimeDomain As String, WaitFlagA As CusWaitVal, _
                    WaitFlagB As CusWaitVal, WaitFlagC As CusWaitVal, WaitFlagD As CusWaitVal, _
                    Optional ScanFailDataLogging As tlTemplateScanFailDataLogging, Optional ScanCaptureLimitMode As tlDigitalCMEMCaptureLimitMode, _
                    Optional ScanCaptureLimitPerPin As Long, Optional ScanPinListSource As tlTemplateScanPinListSource, _
                    Optional ScanPinList As PinList, Optional ScanCaptureFormat As tlTemplateScanCaptureFormat, _
                    Optional ScanCaptureDataType As tlTemplateScanCaptureDataType, Optional ScanUserCommentSource As tlTemplateScanUserCommentSource, _
                    Optional ScanUserComment As String, Optional ATPGPinMapSource As tlTemplateATPGPinMapSource, Optional ATPGPinMapSourcePattern As String, Optional ScanTimeDomain As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    If TheExec.flow.IsRunning = False Then Exit Sub
    
    ' Clear previously registered interpose function names
    Call tl_ClearInterpose(TL_C_PREPATF, TL_C_POSTPATF, TL_C_PRETESTF, TL_C_POSTTESTF, TL_C_POSTPATBPF)
    m_InterposeFunctionsSet = False

    ' Return channels to the default start-state condition, as needed
    If NonBlank(DrivePins) Then Call tl_SetStartState(DrivePins, chstartNone)

    ' Return specified DUT pins, if any, to connection with tester pin-electronics & power
    If NonBlank(FloatPins) Then Call tl_ConnectTester(FloatPins)
    
    ' Restore flag match feature
    ' for compatibility, the flag set/restore should be conditional if asynchronous pattern start not disabled and not suspended
    If ((TheHdw.patterns.EnableAsyncPatternStart <> tlAsyncPatternModeDisabled) And (TheHdw.patterns.SuspendAsyncPatternStart = False)) Then
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
    TheHdw.patterns().Threading.Enable = m_OldPatThreading
    If m_ScanLogging Then
        Call ScanLoggingRestore(ScanFailDataLogging, ScanCaptureLimitMode, _
                    ScanCaptureLimitPerPin, ScanPinListSource, _
                    ScanPinList, ScanCaptureFormat, _
                    ScanCaptureDataType, ScanUserCommentSource, _
                    ScanUserComment, ATPGPinMapSource, ATPGPinMapSourcePattern, ScanTimeDomain)
    End If
Exit Sub 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Functional_T", "PostBody") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/08/18
End Sub


' PostPat Breakpoint interpose function. This is need to support the pattern set
' breakpoint feature.
Public Sub PostTest(argc As Long, argv() As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Call Interpose(m_EndOfBodyF, m_EndOfBodyFArgs)
Exit Sub 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Functional_T", "PostTest") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/08/18
End Sub


' ===============
' Private Helpers
' ===============

' This template needs to know timing and levels sheet names.
' Fetch them from the Context Manager
Private Sub FetchContext()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
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

Exit Sub 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Functional_T", "FetchContext") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/08/18
End Sub


Private Function Validate(patterns As Pattern, PatThreading As Boolean, _
                          DriveLoPins As PinList, DriveHiPins As PinList, _
                          DriveZPins As PinList, DisablePins As PinList, FloatPins As PinList, _
                          Util1Pins As PinList, Util0Pins As PinList, _
                          PatternTimeout As String, WaitTimeDomain As String) As Boolean
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    
    Validate = True ' Assume the best and override if trouble is found
    
    If Not ValidatePatternThreading(patterns, PatThreading, 1, True, 26) Then Validate = False
    
    ' Validate the pin state parameters.
    If Not ValidatePinStates(DriveLoPins, DriveHiPins, DriveZPins, DisablePins, _
                             FloatPins, Util1Pins, Util0Pins) Then Validate = False
        
    If ValidateNumeric(PatternTimeout, "PatternTimeout", 33) Then
        ' Validate  0.0 <= PatternTimeout
        If Not ValidateInRange(StrToDbl(PatternTimeout), "PatternTimeout", 0#, , , , 33) Then Validate = False
    Else
        Validate = False
    End If
    
    'validate timedomain
    If Not ValidateTimeDomain(WaitTimeDomain, "WaitTimeDomain", 34) Then Validate = False
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Functional_T", "Validate") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Private Sub ApplyDefaults(ByRef PatternTimeout As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    ' If the worksheet doesn't have a value then apply 30 as the default.
    If Not NonBlank(PatternTimeout) Then
        PatternTimeout = "30"
    End If
Exit Sub 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Functional_T", "ApplyDefaults") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/08/18
End Sub

'this function is used by test instance sheet,to write the default value for the argument when new test instance is created
'the array can be redimed to hold more values.
Public Function getdefaults() As Variant
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim argdefaults(4) As Variant
    argdefaults(0) = "waitflaga,0" 'argument name and value .
    argdefaults(1) = "waitflagb,0"
    argdefaults(2) = "waitflagc,0"
    argdefaults(3) = "waitflagd,0"
    argdefaults(4) = "patterntimeout,30"
    getdefaults = argdefaults
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Functional_T", "getdefaults") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' Return TL_SUCCESS if the test executes without problems, else TL_ERROR.
'=============================20160413==================================
' [20230524][All][Si] add TTR judge need Alarms.Check or not
' [20230525][All][Jim] add 1. Initial HarvPinsFailCnt  to 0, 2. Add else case Datalog
' 3. Add comment for All Pass, 4. Binout Other case
' [20231228][T-Tah][Tank] Add process SFC
Public Function pattern_module_test(pattern_load As String, RunFailCycle As Boolean, EnableBinOut As Boolean, ReportResult As PFType, TL_C_YES As Long, ResultMode As tlResultMode, ConcurrentMode As tlPatConcurrentMode, instSSNinfo As Inst_SSN, _
                                    Optional Harv_FailFlag As String, Optional HarvestPinGrpOtherFail As String, Optional PFAMultiFunc As Boolean = False, Optional ApplyVoltageFromBinCut As String = vbNullString)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim ins_name As String
    Dim i As Long:: i = 0
    Dim site As Variant
    Dim site_BK_loop_count As Long
    Dim pattern_name(8) As String
    Dim Flag_Name As String
    Dim c As Boolean
    Dim confirm_inst As Boolean
    Dim ws_def As Worksheet
    Dim wb As Workbook

    Dim maxDepth As Integer
    Dim patt As Variant
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
    
    Dim blMbistFP_Binout As Boolean
    Dim MBISTFailBlockFlag As Boolean
    Dim PassOrFail As New SiteLong
    Dim lGetFlagIdx As Long
    Dim blJump As Boolean
    Dim m_testName As String
    Dim k As Long, p As Long, g As Long, j As Long:: k = 0:: p = 0:: g = 0:: j = 0
    Dim PrintFailPat_Flag As Boolean
    Dim Bool_CheckInitPat As Boolean
    Dim sBool_PatternPass As New SiteBoolean
    Dim TempDigitalPin() As String
    Dim CustHarvCondition() As Boolean
    Dim Current_FlowSheet_Name As String
    Dim h As Long, RowCnt As Long
    Dim ATPG_Pin_Table_Row As Long
    Dim Search_ATPG_Harvest_Flag As Boolean
    Dim strAry_PathSplit As Variant, strAry_PatNameSplit As Variant
    Dim Sbln_PatternPass As New SiteBoolean
    Dim Harvest_Argument_Info() As String
    Dim Harvest_enableword_Info() As String
    Dim SCAN_Site_Blooean As New SiteBoolean
    
    On Error GoTo errHandler
    
    For Each site In TheExec.sites
        SCAN_Site_Blooean(site) = True
    Next
    Harvest_Pin_From_Table_Flag = False 'Initial Flag status
    HarvFailCnt = 0 'Initial HarvFailCnt
    'LogLimited = 255
    AllPins = "JTAG_TDO"
    m_testName = TheExec.DataManager.instancename
    'Dim pattern_load As String
    'pattern_load = ".\Patterns\vreg_test_pop_student.pat"
    '-----------------------------------------------------------------------------------------
    For Each site In TheExec.sites
        If (TheExec.sites(site).SiteVariableValue("LP_BM") = TheExec.sites(site).SiteVariableValue("Lcount_BM")) Then
            confirm_in_loop = False
            Exit For
        End If
    Next site
    '-----------------------------------------------------------------------------------------
    If (confirm_in_loop = True) Then
        confirm_inst = False
        '============================================================================init setting
        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
        
        ins_name = TheExec.DataManager.instancename
        
        For Each site In TheExec.sites
            site_BK_loop_count = TheExec.sites(site).SiteVariableValue("LP_BM")
            'pattern_name(Site) = mbist_dynamic.Block_dynamic(0).pat_name_dynamic(site_BK_loop_count)
            
            If UCase(ins_name) Like "*IVDM*" Or UCase(ins_name) Like "*SNSUHS*" Then
            Else
                TheExec.sites.item(site).TestNumber = TheExec.sites.item(site).TestNumber + site_BK_loop_count * 100001
            End If
            
            'Exit For
        Next site
        '============================================================================excute test
        For i = 0 To UBound(mbist_dynamic.Block_dynamic(0).block_type_pat_dynamic(0).instance_dynamic)
            If (UCase(ins_name) Like UCase("*" + mbist_dynamic.Block_dynamic(0).block_type_pat_dynamic(0).instance_dynamic(i) + "*") And (mbist_dynamic.Block_dynamic(0).block_type_pat_dynamic(0).instance_dynamic(i) <> "")) Then
                confirm_inst = True
                Exit For
            End If
        Next i
 
        If (confirm_inst = True) Then
           Call auto_Mbist_Block_loop_inst_match(ins_name, pattern_load, site_BK_loop_count, Flag_Name, EnableBinOut, RunFailCycle)
        Else
           Call auto_Mbist_Block_loop_inst_non_match(ins_name, pattern_load, site_BK_loop_count, Flag_Name, EnableBinOut, RunFailCycle)
        End If
        '============================================================================
    Else
        Dim patset() As String
        Dim patcnt As Long
        Dim pat As Variant
        Call PatternBurstCheckAndSplit(pattern_load, patset, patcnt)
        
        If LCase(m_testName) Like "*bist*" And TheExec.enableWord("Mbist_FingerPrint") = True Then
            Call Finger_print(pattern_load, RunFailCycle, Flag_Name, False)
        Else

            Dim e As Long
            Dim HarvCnt As Long
            Dim HarvFailFlagArr() As String
            Dim HarvConditionArr() As String

            If BurstYesPatDict.Exists(LCase(pattern_load)) Then
                If TheExec.TesterMode = testModeOffline Then
                    Call ATPG_offline(pattern_load, ResultMode)
                Else
                    'Condition_and_FailFlag = P:ECPU_CORE0(F_ECPU_CORE0);P:ECPU_CORE1(F_ECPU_CORE1)
                    If Harv_FailFlag <> "" Then
                        Call Harvest_FailFlagSplit(Harv_FailFlag, HarvConditionArr, HarvFailFlagArr, CustHarvCondition)
                        HarvFailCnt = UBound(HarvFailFlagArr) + 1
                    End If
                    
                    If gl_bTTRDisableAlarm = False Then     'T-Col TTR approve by Si -- 230413
                        TheHdw.Alarms.Check
                    End If
                   
                    If (Harv_FailFlag <> "") Or (glb_SFC_Scan_Check = True) Then
                        Call Harvest_CMEM_InitSetup
                    End If
                    
                    Call TheHdw.patterns(pattern_load).test(ReportResult, CLng(TL_C_YES), ResultMode, ConcurrentMode)
                    
                    If Harv_FailFlag <> "" Then
                        For HarvCnt = 0 To UBound(HarvFailFlagArr)
                            Call Harvest_Decision(HarvFailFlagArr(HarvCnt), HarvConditionArr(HarvCnt))
                        Next HarvCnt
                        For Each site In TheExec.sites
                            If HarvFailCnt(site) = 0 Then
                                TheExec.sites.item(site).FlagState(HarvestPinGrpOtherFail) = logicTrue
                            End If
                        Next site
                    Else
'                        PrintFailPat_Flag = theexec.Datalog.Setup.DatalogSetup.SelectSetupFile
                        If TheExec.enableWord("PatternFailInfo") = True And TheHdw.Digital.hram.size <> 0 Then
                            Call Printing_StandalonePat(pattern_load, patset, patcnt)
                        End If
                    End If
                    
                    If Harv_FailFlag <> "" Then
                        Call Harvest_CMEM_Stop(HarvConditionArr, CustHarvCondition, HarvFailFlagArr)
                    ElseIf glb_SFC_Scan_Check = True Then
                        Call SFC_CMEM_Stop
                    End If
                    
                End If
            Else
                'Condition_and_FailFlag = P:ECPU_CORE0(F_ECPU_CORE0);P:ECPU_CORE1(F_ECPU_CORE1)
                If PFAMultiFunc Then        'OTC2-TER-ZY. call PFA multi function
''                    If LCase(TheExec.DataManager.instancename) Like "*chain*" Or LCase(TheExec.DataManager.instancename) Like "*sa*" Then
''                        Call RunFailCntShm(pattern_load, ScanCaptureLimitPerPin, "ShiftIn")
''                    Else
''                        Call RunFailCntShm(pattern_load, ScanCaptureLimitPerPin, "FRC")
''                    End If
                Else
                    glb_CheckSSNToSFC = False
                    If Harv_FailFlag = "" Then ''Flat pattern
                        Call PatternExecution(patset, ReportResult, TL_C_YES, ResultMode, ConcurrentMode, SCAN_Site_Blooean, ApplyVoltageFromBinCut)
                    ElseIf Harv_FailFlag <> "" And instSSNinfo.bSSNTest = False Then ''Traditional_Harvesting
                        Call HarvestingMainProcedure(patset, ReportResult, TL_C_YES, ResultMode, ConcurrentMode, SCAN_Site_Blooean, ApplyVoltageFromBinCut, Harv_FailFlag, HarvestPinGrpOtherFail)
                    ElseIf Harv_FailFlag <> "" And instSSNinfo.bSSNTest = True Then ''SSN
                        'If glb_TesterType = "UltraFLEXplus" Then
                            glb_CheckSSNToSFC = True
                            Call SSNMainProcedure(patset, ReportResult, TL_C_YES, ResultMode, ConcurrentMode, SCAN_Site_Blooean, instSSNinfo, ApplyVoltageFromBinCut, Harv_FailFlag, HarvestPinGrpOtherFail)
                        'End If
                    Else
                    End If
                End If
            End If
        End If
    End If
    
    '--------------------------------------------------------------------------------print out flag sheet for txt/csv file
    If (create_flag_sheet And confirm_in_loop = True) Then
        Dim FileExists As Boolean
        Dim string_store As String, string_store01 As String
        For Each site In TheExec.sites
            FileExists = (Dir(File_path) <> "")
            If FileExists = False Then
                Open File_path For Output As #1
            End If
            If (TheExec.sites(site).SiteVariableValue("LP_BM") = 0 And index_flag_y = 1) Then
                string_store = vbNullString
                string_store01 = vbNullString
                string_store = "Flag-" + bist_type + "_" + mbist_dynamic.Block_dynamic(0).block_name_dynamic
                string_store01 = "Binout-" + bist_type + "_" + mbist_dynamic.Block_dynamic(0).block_name_dynamic
                Print #1, "===============================================================,======================="
                Print #1, string_store + "," + string_store01
                'Write #1, "Flag-SOC_ SOC,Binout-SOC_ SOC"
                index_flag_y = index_flag_y + 1
            End If
            Print #1, Flag_Name + ",";
            If (TheExec.sites.item(site).FlagState(Flag_Name) = logicTrue) Then
                Write #1, "Fail"
            ElseIf (TheExec.sites.item(site).FlagState(Flag_Name) = logicFalse) Then
                Write #1, "Pass"
            Else
                Write #1, "Clean"
            End If
            'Close #1
            Exit For
        Next site
    End If
    '--------------------------------------------------------------------------------
    ' Disable Fringer Pint
'    TheExec.Flow.ForceStopOnError = True
'    For Each Pat In patset
'        If LCase(m_testName) Like "*cpumbist*" And Not LCase(Pat) Like "*_flp_*" And Not LCase(Pat) Like "*_efc_*" And Not LCase(Pat) Like "*_pri_*" Then
'            If Not LCase(Pat) Like "*_inlp_*" And Not LCase(Pat) Like "*mpxxx*" Then
'                For Each site In TheExec.sites
'                    TempDigitalPin = TheHdw.Digital.FailedPins(site)
'                    If UBound(TempDigitalPin) + 1 <> 0 Then
'                        If (TheExec.sites.item(site).FlagState("F_ECPU_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_PCPU_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_GFX_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_SOC_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_SOC_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_GFX_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_PCPU_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_ECPU_SCAN_Alarm_Check") = logicTrue) Then
'                            TheExec.sites.item(site).FlagState("F_ECPU_MBIST_Defect_Alarm_Check") = logicTrue
'                            TheExec.sites.item(site).SortNumber = 975
'                            TheExec.sites.item(site).BinNumber = 19
'                            TheExec.sites.site(site).result = tlResultFail
'                        End If
'                    End If
'                            'testnumber(Site) = TheExec.sites.Item(Site).testnumber
'                Next site
'            End If
'        ElseIf LCase(m_testName) Like "*cpumbist*" And Not LCase(Pat) Like "*_flp_*" And Not LCase(Pat) Like "*_efc_*" And Not LCase(Pat) Like "*_pri_*" Then
'            If Not LCase(Pat) Like "*_inlp_*" And Not LCase(Pat) Like "*mexxx*" Then
'                For Each site In TheExec.sites
'                    TempDigitalPin = TheHdw.Digital.FailedPins(site)
'                    If UBound(TempDigitalPin) + 1 <> 0 Then
'                        If (TheExec.sites.item(site).FlagState("F_ECPU_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_PCPU_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_GFX_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_SOC_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_SOC_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_GFX_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_PCPU_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_ECPU_SCAN_Alarm_Check") = logicTrue) Then
'                            TheExec.sites.item(site).FlagState("F_PCPU_MBIST_Defect_Alarm_Check") = logicTrue
'                            TheExec.sites.item(site).SortNumber = 976
'                            TheExec.sites.item(site).BinNumber = 19
'                            TheExec.sites.site(site).result = tlResultFail
'                        End If
'                    End If
'                            'testnumber(Site) = TheExec.sites.Item(Site).testnumber
'                Next site
'            End If
'        ElseIf LCase(m_testName) Like "*gfxmbist*" And Not LCase(Pat) Like "*_flp_*" And Not LCase(Pat) Like "*_efc_*" And Not LCase(Pat) Like "*_pri_*" Then
'            If Not LCase(Pat) Like "*_pllp_*" And Not LCase(Pat) Like "*_fstp*" Then
'                For Each site In TheExec.sites
'                    TempDigitalPin = TheHdw.Digital.FailedPins(site)
'                    If UBound(TempDigitalPin) + 1 <> 0 Then
'                        If (TheExec.sites.item(site).FlagState("F_ECPU_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_PCPU_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_GFX_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_SOC_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_SOC_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_GFX_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_PCPU_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_ECPU_SCAN_Alarm_Check") = logicTrue) Then
'                            TheExec.sites.item(site).FlagState("F_GFX_MBIST_Defect_Alarm_Check") = logicTrue
'                            TheExec.sites.item(site).SortNumber = 977
'                            TheExec.sites.item(site).BinNumber = 19
'                            TheExec.sites.site(site).result = tlResultFail
'                        End If
'                    End If
'                            'testnumber(Site) = TheExec.sites.Item(Site).testnumber
'                Next site
'            End If
'        ElseIf LCase(m_testName) Like "*socmbist*" And Not LCase(Pat) Like "*_flp_*" And Not LCase(Pat) Like "*_efc_*" And Not LCase(Pat) Like "*_pri_*" Then
'            If Not LCase(Pat) Like "*_inlp_*" And Not LCase(Pat) Like "*_fstp*" Then
'                For Each site In TheExec.sites
'                    TempDigitalPin = TheHdw.Digital.FailedPins(site)
'                    If UBound(TempDigitalPin) + 1 <> 0 Then
'                        If (TheExec.sites.item(site).FlagState("F_ECPU_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_PCPU_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_GFX_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_SOC_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_SOC_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_GFX_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_PCPU_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_ECPU_SCAN_Alarm_Check") = logicTrue) Then
'                            TheExec.sites.item(site).FlagState("F_SOC_MBIST_Defect_Alarm_Check") = logicTrue
'                            TheExec.sites.item(site).SortNumber = 978
'                            TheExec.sites.item(site).BinNumber = 19
'                            TheExec.sites.site(site).result = tlResultFail
'                        End If
'                    End If
'                            'testnumber(Site) = TheExec.sites.Item(Site).testnumber
'                Next site
'            End If
'        ElseIf LCase(m_testName) Like "*cpusa*" And LCase(Pat) Like "*_mexxxx_*" Then
'             For Each site In TheExec.sites
'                TempDigitalPin = TheHdw.Digital.FailedPins(site)
'                If UBound(TempDigitalPin) + 1 <> 0 Then
'                    If (TheExec.sites.item(site).FlagState("F_ECPU_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_PCPU_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_GFX_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_SOC_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_SOC_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_GFX_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_PCPU_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_ECPU_SCAN_Alarm_Check") = logicTrue) Then
'                        TheExec.sites.item(site).FlagState("F_ECPU_SCAN_Defect_Alarm_Check") = logicTrue
'                        TheExec.sites.item(site).SortNumber = 971
'                        TheExec.sites.item(site).BinNumber = 19
'                        TheExec.sites.site(site).result = tlResultFail
'                    End If
'                End If
'              Next site
'        ElseIf LCase(m_testName) Like "*cpusa*" And LCase(Pat) Like "*_mpxxxx_*" Then
'            For Each site In TheExec.sites
'               TempDigitalPin = TheHdw.Digital.FailedPins(site)
'               If UBound(TempDigitalPin) + 1 <> 0 Then
'                   If (TheExec.sites.item(site).FlagState("F_ECPU_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_PCPU_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_GFX_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_SOC_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_SOC_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_GFX_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_PCPU_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_ECPU_SCAN_Alarm_Check") = logicTrue) Then
'                       TheExec.sites.item(site).FlagState("F_PCPU_SCAN_Defect_Alarm_Check") = logicTrue
'                       TheExec.sites.item(site).SortNumber = 972
'                       TheExec.sites.item(site).BinNumber = 19
'                       TheExec.sites.site(site).result = tlResultFail
'                   End If
'               End If
'            Next site
'        ElseIf LCase(m_testName) Like "*gfxSa*" And Not LCase(Pat) Like "*_cfxx_*" Then
'            For Each site In TheExec.sites
'               TempDigitalPin = TheHdw.Digital.FailedPins(site)
'               If UBound(TempDigitalPin) + 1 <> 0 Then
'                   If (TheExec.sites.item(site).FlagState("F_ECPU_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_PCPU_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_GFX_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_SOC_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_SOC_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_GFX_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_PCPU_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_ECPU_SCAN_Alarm_Check") = logicTrue) Then
'                       TheExec.sites.item(site).FlagState("F_GFX_SCAN_Defect_Alarm_Check") = logicTrue
'                       TheExec.sites.item(site).SortNumber = 973
'                       TheExec.sites.item(site).BinNumber = 19
'                       TheExec.sites.site(site).result = tlResultFail
'                   End If
'               End If
'            Next site
'        ElseIf LCase(m_testName) Like "*socsa*" Then
'            For Each site In TheExec.sites
'               TempDigitalPin = TheHdw.Digital.FailedPins(site)
'               If UBound(TempDigitalPin) + 1 <> 0 Then
'                   If (TheExec.sites.item(site).FlagState("F_ECPU_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_PCPU_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_GFX_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_SOC_MBIST_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_SOC_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_GFX_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_PCPU_SCAN_Alarm_Check") = logicTrue) Or (TheExec.sites.item(site).FlagState("F_ECPU_SCAN_Alarm_Check") = logicTrue) Then
'                       TheExec.sites.item(site).FlagState("F_SOC_SCAN_Defect_Alarm_Check") = logicTrue
'                       TheExec.sites.item(site).SortNumber = 974
'                       TheExec.sites.item(site).BinNumber = 19
'                       TheExec.sites.site(site).result = tlResultFail
'                   End If
'               End If
'            Next site
'        End If
'    Next Pat
    

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Functional_T", "pattern_module_test") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' [20231106][T-ALL][Oliver] add for UltraFlexPlus will not change alarmfail flag state
Public Function auto_Mbist_Block_loop_inst_match(instance_name As String, m_pattname As String, bk_loop_count As Long, ByRef Flag_Name As String, EnableBinOut As Boolean, RunFailCycle As Boolean) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim funcName As String:: funcName = "auto_Mbist_Block_loop_inst_match"
    Dim site As Variant
    'Dim flag_name As String
    Dim ins_array_name_pp() As String
    Dim ins_array_name_type() As String
    Dim ins_array_name_others() As String
    
    Dim flag_array_string_match() As String
    Dim flag_array_string_inst() As String
    Dim match_begin As Long, match_end As Long:: match_begin = match_end = 0
    Dim flag_spilt As String
    Dim ins_array_name_long() As String

    Dim ins_array_name_perf_v() As String
    Dim i As Long, k As Long, p As Long, g As Long, j As Long:: i = 0:: k = 0:: p = 0:: g = 0:: j = 0
    Dim confirm As Boolean
    
    Dim LNH_V As String
    Dim perofmrance As String
    Dim maxDepth As Integer
    Dim patt As Variant
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
    Dim flag_match As Boolean
    Dim temp As Long
    Dim AllPins As String
    Dim PinData As New PinListData
    
    Dim match_string_1st As String
    
    Dim blMbistFP_Binout As Boolean
    Dim MBISTFailBlockFlag As Boolean
    Dim PassOrFail As New SiteLong
    Dim lGetFlagIdx As Long
    Dim blJump As Boolean
    
    Dim m_testName As String
    Dim for_confirm_ins_name As String
    for_confirm_ins_name = vbNullString
    Dim for_confirm_ins_name_array() As String
    Dim alarmOccurred As New SiteBoolean
    
    
    AllPins = "JTAG_TDO"
    ins_array_name_perf_v = Split(instance_name, "_")
    m_testName = ins_array_name_perf_v(0)
    '=================================================================================================test flag
    For i = 0 To UBound(ins_array_name_perf_v)
        If (UCase(ins_array_name_perf_v(i)) Like "NV" Or UCase(ins_array_name_perf_v(i)) Like "LV" Or UCase(ins_array_name_perf_v(i)) Like "HV" Or UCase(ins_array_name_perf_v(i)) Like "MNV" Or UCase(ins_array_name_perf_v(i)) Like "MLV" Or UCase(ins_array_name_perf_v(i)) Like "MHV") Then
            LNH_V = "_" + ins_array_name_perf_v(i)            ''''''''''N/L/HV
        Else
            LNH_V = vbNullString
        End If

        If (UCase(ins_array_name_perf_v(i)) Like "MC*" Or UCase(ins_array_name_perf_v(i)) Like "MS*" Or UCase(ins_array_name_perf_v(i)) Like "MG*" Or UCase(ins_array_name_perf_v(i)) Like "MA*") Then
            If (IsNumeric(mid(ins_array_name_perf_v(i), 3, 1))) Then
                perofmrance = "_" + ins_array_name_perf_v(i) '''''''''''performance name
            Else
                perofmrance = vbNullString
            End If
        End If
    Next i
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''match instance name, prepare conbine flag name
    For i = 0 To mbist_match(type_nu).inst_count - 1
        match_begin = match_end = 0
        confirm = False
        p = 0:: g = 0
        If (UCase(instance_name) Like UCase("*" + mbist_match(type_nu).inst_nu(i).binflag_match_name + "*")) Then
            flag_match = True
            Flag_Name = vbNullString
            flag_array_string_match = Split(mbist_match(type_nu).inst_nu(i).binflag_match_name, "_")
            flag_array_string_inst = Split(instance_name, "_")
            match_string_1st = vbNullString
            If (flag_array_string_match(0) = "" Or flag_array_string_match(0) = " ") Then
                match_string_1st = flag_array_string_match(1)
                g = 1
            Else
                match_string_1st = flag_array_string_match(0)
                g = 0
            End If
            For k = 0 To UBound(flag_array_string_inst)
                If (UCase(flag_array_string_inst(k)) Like UCase(match_string_1st)) Then
                    match_begin = k
                    confirm = True
                    g = g + 1
                ElseIf (confirm = True And k <= UBound(flag_array_string_inst) And g <= UBound(flag_array_string_match) And flag_array_string_match(g) <> "") Then
                    If (UCase(flag_array_string_inst(k)) Like UCase(flag_array_string_match(g))) Then
                        confirm = True
                    Else
                        confirm = False
                    End If
                
                    If (flag_array_string_match(UBound(flag_array_string_match)) = "") Then
                        If (UCase(flag_array_string_inst(k)) Like UCase(flag_array_string_match(UBound(flag_array_string_match) - 1)) And confirm = True) Then
                            match_end = k
                            Exit For
                        End If
                    Else
                        If (UCase(flag_array_string_inst(k)) Like UCase(flag_array_string_match(UBound(flag_array_string_match))) And confirm = True) Then
                            match_end = k
                            Exit For
                        End If
                    End If
                
                    g = g + 1
                End If
            Next k

            If (confirm = True And match_end <> 0) Then
                For k = 0 To UBound(flag_array_string_inst)
                    If (k >= match_begin And k <= match_end) Then
                        If (p = 0 And Flag_Name <> "") Then
                            Flag_Name = Flag_Name + "_" + mbist_match(type_nu).inst_nu(i).binflag_mid_name  '//check
                        Else
                            Flag_Name = mbist_match(type_nu).inst_nu(i).binflag_mid_name
                        End If
                        p = p + 1
                    Else
                        If (Flag_Name <> "") Then
                            Flag_Name = Flag_Name + "_" + flag_array_string_inst(k)
                        Else
                            Flag_Name = flag_array_string_inst(k)
                        End If
                
                    End If
                Next k

            End If
            Exit For
        End If
    Next i
    
    If (instance_name Like "*Mbist_*") Then
        ins_array_name_type = Split(instance_name, "_")
    End If
    
    If (instance_name Like "CpuMbist_*") Then
        ins_array_name_others = Split(instance_name, "CpuMbist_")
    ElseIf (instance_name Like "SocMbist_*") Then
        ins_array_name_others = Split(instance_name, "SocMbist_")
    End If
    
    '=================================================================================================instance name
    ins_new_name = ins_array_name_type(0) + "_" + mbist_dynamic.Block_dynamic(0).block_count_name_dynamic(bk_loop_count) + "_" + ins_array_name_others(1)
    Block = mbist_dynamic.Block_dynamic(0).block_count_name_dynamic(bk_loop_count)
    '===================================================Print debug information===============================================================

    'TheExec.Datalog.WriteComment "~~~~~~~ Instance match ~~~~~~~"
    If (Flag_Name <> "" And flag_match = True) Then
        'TheExec.Datalog.WriteComment "~~~~~~~ Flag match ~~~~~~~"
        If (instance_name Like "*_PP_*") Then
            ins_array_name_pp = Split(Flag_Name, "_PP_")
            for_confirm_ins_name = "PP_" + ins_array_name_pp(1)
        ElseIf (instance_name Like "*_DD_*") Then
            ins_array_name_pp = Split(Flag_Name, "_DD_")
            for_confirm_ins_name = "DD_" + ins_array_name_pp(1)
        ElseIf (instance_name Like "*_CZ_*") Then
            ins_array_name_pp = Split(Flag_Name, "_CZ_")
            for_confirm_ins_name = "CZ_" + ins_array_name_pp(1)
        Else
            ins_array_name_pp = Split(Flag_Name, "_")
            ins_array_name_pp(0) = vbNullString
            for_confirm_ins_name = "" + ins_array_name_pp(1)
        End If
        Flag_Name = ins_array_name_pp(0) + LNH_V + "_" + mbist_dynamic.Block_dynamic(0).block_count_name_dynamic(bk_loop_count)
    ElseIf (flag_match = False) Then
        'TheExec.Datalog.WriteComment "~~~~~~~ Flag no match ~~~~~~~"
        If (ins_new_name Like "*_PP_*") Then
            ins_array_name_pp = Split(ins_new_name, "_PP_")
            for_confirm_ins_name = "PP_" + ins_array_name_pp(1)
        ElseIf (ins_new_name Like "*_DD_*") Then
            ins_array_name_pp = Split(ins_new_name, "_DD_")
            for_confirm_ins_name = "DD_" + ins_array_name_pp(1)
        ElseIf (ins_new_name Like "*_CZ_*") Then
            ins_array_name_pp = Split(ins_new_name, "_CZ_")
            for_confirm_ins_name = "CZ_" + ins_array_name_pp(1)
        Else
            ins_array_name_pp = Split(ins_new_name, "_")
            ins_array_name_pp(0) = vbNullString
            for_confirm_ins_name = "" + ins_array_name_pp(1)
        End If

        ins_array_name_pp = Split(ins_new_name, "_")
        flag_spilt = ins_array_name_pp(mbist_flag_set_placement + 1)
        ins_array_name_long = Split(ins_new_name, "_" + flag_spilt)
        Flag_Name = ins_array_name_long(0) + LNH_V
    Else
        'TheExec.Datalog.WriteComment "~~~~~~~ Flag conbine Erro ~~~~~~~"
        'TheExec.Flow.TestLimit -1, 0, 1, , , , unitNone, , "Test_Falg"
    End If
    
    '=========================================================================================================================================
    
    Flag_Name = "F_" + Flag_Name
    
    for_confirm_ins_name = Replace(for_confirm_ins_name, "_NV", "")
    for_confirm_ins_name = Replace(for_confirm_ins_name, "_LV", "")
    for_confirm_ins_name = Replace(for_confirm_ins_name, "_HV", "")
    for_confirm_ins_name = Replace(for_confirm_ins_name, "_MNV", "")
    for_confirm_ins_name = Replace(for_confirm_ins_name, "_MLV", "")
    for_confirm_ins_name = Replace(for_confirm_ins_name, "_MHV", "")
    
'''    For Each Site In theExec.sites
'''        theExec.sites.Item(Site).FlagState(flag_name) = logicFalse ''''mean Pass
'''    Next Site

    Call TheExec.Datalog.SetDynamicTestName(ins_new_name, False)
    '=================================================================================================pattern
    For k = 0 To UBound(mbist_dynamic.Block_dynamic(0).block_type_pat_dynamic(bk_loop_count).pat_dynamic)
        If (Trim(for_confirm_ins_name) = Trim(mbist_dynamic.Block_dynamic(0).block_type_pat_dynamic(bk_loop_count).instance_dynamic(k))) Then
            m_pattname = mbist_dynamic.Block_dynamic(0).block_type_pat_dynamic(bk_loop_count).pat_dynamic(k)
        End If
    Next k
    '=================================================================================================patt test
    blMbistFP_Binout = EnableBinOut And gl_MbistFP_Binout
    TheHdw.patterns(m_pattname).Load
    
    If TheExec.enableWord("Mbist_FingerPrint") = True Then
        Call Finger_print(m_pattname, RunFailCycle, Flag_Name, True)
    Else
        Call PATT_GetPatListFromPatternSet(m_pattname, rtnPatternNames, rtnPatternCount)
        For Each patt In rtnPatternNames
            TheExec.Datalog.WriteComment "<" & ins_new_name & ">" & " dummy "
        For Each site In TheExec.sites
            tested(site) = False 'swinza move to out of pattern-loop
        Next site
            Call TheHdw.patterns(patt).test(pfAlways, 0, tlResultModeDomain)

            If glb_TesterType = "UltraFLEXplus" Then 'add for GetAlarmingSites
                alarmOccurred = False
                alarmOccurred = TheHdw.Alarms.GetAlarmingSites(True)        ''GetAlarmingSites(clearAlarm = True)
                For Each site In TheExec.sites
                    If alarmOccurred(site) = True Then
                        alarmFail(site) = True      ''Update alarm info to alarmFail array for UFP
                    End If
                Next site
            End If
            '===================================================================
            For Each site In TheExec.sites
                'testnumber(Site) = TheExec.sites.Item(Site).testnumber
                'tested(Site) = False
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
        Next patt
    End If

Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Functional_T", "auto_Mbist_Block_loop_inst_match") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18

End Function


Public Function auto_Mbist_Block_loop_inst_non_match(instance_name As String, m_pattname As String, bk_loop_count As Long, ByRef Flag_Name As String, EnableBinOut As Boolean, RunFailCycle As Boolean) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim funcName As String:: funcName = "auto_Mbist_Block_loop_inst_non_match"
    Dim site As Variant
    'Dim flag_name As String
    Dim ins_array_name_pp() As String
    Dim ins_array_name_long() As String
    Dim ins_array_name_type() As String
    Dim ins_array_name_others() As String
    Dim flag_spilt As String
    
    Dim flag_array_string_match() As String
    Dim flag_array_string_inst() As String
    Dim match_begin As Long, match_end As Long:: match_begin = match_end = 0
    
    Dim ins_array_name_perf_v() As String
    Dim i As Long, k As Long, p As Long, g As Long, j As Long:: i = 0:: k = 0:: p = 0:: g = 0:: j = 0
    Dim confirm As Boolean
    
    Dim LNH_V As String
    Dim perofmrance As String
    Dim maxDepth As Integer
    Dim patt As Variant
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
    
    Dim blMbistFP_Binout As Boolean
    Dim MBISTFailBlockFlag As Boolean
    Dim PassOrFail As New SiteLong
    Dim lGetFlagIdx As Long
    Dim blJump As Boolean
    Dim m_testName As String
    
    
    
    AllPins = "JTAG_TDO"
    ins_array_name_perf_v = Split(instance_name, "_")
    m_testName = ins_array_name_perf_v(0)
    '=================================================================================================test flag
    For i = 0 To UBound(ins_array_name_perf_v)
        If (UCase(ins_array_name_perf_v(i)) Like "NV" Or UCase(ins_array_name_perf_v(i)) Like "LV" Or UCase(ins_array_name_perf_v(i)) Like "HV" Or UCase(ins_array_name_perf_v(i)) Like "MNV" Or UCase(ins_array_name_perf_v(i)) Like "MLV" Or UCase(ins_array_name_perf_v(i)) Like "MHV") Then
            LNH_V = "_" + ins_array_name_perf_v(i)            ''''''''''N/L/HV
        Else
            LNH_V = vbNullString
        End If

        If (UCase(ins_array_name_perf_v(i)) Like "MC*" Or UCase(ins_array_name_perf_v(i)) Like "MS*" Or UCase(ins_array_name_perf_v(i)) Like "MG*" Or UCase(ins_array_name_perf_v(i)) Like "MA*") Then
            If (IsNumeric(mid(ins_array_name_perf_v(i), 3, 1))) Then
                perofmrance = "_" + ins_array_name_perf_v(i) '''''''''''performance name
            Else
                perofmrance = vbNullString
            End If
        End If
    Next i
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'TheExec.Datalog.WriteComment "~~~~~~~ Instance no match ~~~~~~~"
    
    If (instance_name Like "*Mbist_*") Then
        ins_array_name_type = Split(instance_name, "_")
    End If
    
    If (instance_name Like "CpuMbist_*") Then
        ins_array_name_others = Split(instance_name, "CpuMbist_")
    ElseIf (instance_name Like "SocMbist_*") Then
        ins_array_name_others = Split(instance_name, "SocMbist_")
    End If
    
    '=================================================================================================instance name
    ins_new_name = ins_array_name_type(0) + "_" + mbist_dynamic.Block_dynamic(0).block_count_name_dynamic(bk_loop_count) + "_" + ins_array_name_others(1)
    ins_array_name_pp = Split(ins_new_name, "_")
    flag_spilt = ins_array_name_pp(mbist_flag_set_placement + 1)
    ins_array_name_long = Split(ins_new_name, "_" + flag_spilt)
    
    Flag_Name = ins_array_name_long(0) + LNH_V
    Flag_Name = "F_" + Flag_Name

    Call TheExec.Datalog.SetDynamicTestName(ins_new_name, False)
    '=================================================================================================patt test
    blMbistFP_Binout = EnableBinOut And gl_MbistFP_Binout
    TheHdw.patterns(m_pattname).Load
    
    If TheExec.enableWord("Mbist_FingerPrint") = True Then
        Call Finger_print(m_pattname, RunFailCycle, Flag_Name, True)
    Else
        Call PATT_GetPatListFromPatternSet(m_pattname, rtnPatternNames, rtnPatternCount)
        For Each patt In rtnPatternNames
            TheExec.Datalog.WriteComment "<" & ins_new_name & ">" & " dummy "
            For Each site In TheExec.sites
                tested(site) = False 'swinza move to out of pattern-loop
            Next site
            Call TheHdw.patterns(patt).test(pfAlways, 0, tlResultModeDomain)
            '===================================================================
            For Each site In TheExec.sites
                'testnumber(Site) = TheExec.sites.Item(Site).testnumber
                'tested(Site) = False
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
        Next patt
        
    End If
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Functional_T", "auto_Mbist_Block_loop_inst_non_match") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


'[20230712][T-Pal][Tank] for TTR change "Shmoo_Save_core_power_per_site_for_Vbump" function use position
'[20231106][T-Tah][Oliver] add multiple SRAM pin compares with one logic pin method
Public Function Sub_SourceEMA_SelSRM(PatternName As String, instancename As String, DigSource As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    
    ' ======DSSC======'
    'Add one more condition to enable DigSrc, SWLINZA 20190602
    Dim F_DigSrcEnable As Boolean
    Dim DigSrc_pin As New PinList
    Dim DigSource_Arr() As String
    Dim TestCase As String
    
    Dim funcName As String:: funcName = "Sub_SourceEMA_SelSRM"
    
    
    'Dim Site As Variant 'already defined in public variant
    Dim Store_PinList As New PinListData
    Dim BlockType As String, BlockHeader As String
    BlockType = Split(instancename, "_")(0)
    BlockHeader = left(BlockType, 3)
    
''''    Hard code "blockheader" for Cpu Scan to check "SELSRM_Mapping_Table", SSYANGI 20190617
    If UCase(BlockHeader) = "CPU" Then
        If UCase(PatternName) Like "*E*" Then
            BlockHeader = "ecpu"
        ElseIf UCase(PatternName) Like "*P*" Then
            BlockHeader = "pcpu"
        Else
            TheExec.ErrorLogMessage ("Please check the CpuScan Type")
        End If
    End If
    
    If UCase(BlockType) Like "*SA*" Or UCase(BlockType) Like "*TD*" Then
        BlockType = "SCAN"
    ElseIf UCase(BlockType) Like "*MBIST*" Then
        BlockType = "MBIST"
    End If
    
    F_DigSrcEnable = False  'initialize
    
    If DigSource <> "" Then  'digsrc triggered by instance argument
        F_DigSrcEnable = True  'enable digsrc
        DigSource_Arr() = Split(DigSource, ":")
        TestCase = DigSource_Arr(0)
        DigSrc_pin.value = DigSource_Arr(1)
    ElseIf UCase(instancename) Like UCase("*DSSC*") Then  'digsrc triggered by pattern name
        Dim Current_DCCategory As String
        Dim Current_DCSelector As String
        Dim Dummy_tempStr As String
        
        F_DigSrcEnable = True
        
        If BlockType = "SCAN" Then
            TestCase = "Test_AutoSwitch"
        ElseIf BlockType = "MBIST" Then
            TheExec.DataManager.GetInstanceContext Current_DCCategory, Current_DCSelector, Dummy_tempStr, Dummy_tempStr, Dummy_tempStr, Dummy_tempStr, Dummy_tempStr, Dummy_tempStr
            Select Case UCase(Current_DCSelector)
                Case "MIN":
                    Dummy_tempStr = "LV"
                Case "MAX":
                    Dummy_tempStr = "HV"
                Case "TYP":
                    Dummy_tempStr = "NV"
            End Select
            TestCase = "Test_" & Current_DCCategory & "_" & Dummy_tempStr
        Else
            TheExec.ErrorLogMessage "Pattern Named DSSC but Instance without EMA Setting: " & instancename
        End If
        DigSrc_pin.value = "JTAG_TDI"
    Else
        'TheExec.ErrorLogMessage "Please indicate DigSource in Test Instance Arugment: " & m_InstanceName
    End If
    
    BlockType = UCase(BlockHeader) & BlockType
    
    'PC modified for DSSC Mapping
    'SW modified for one more case, 20190602
    If F_DigSrcEnable = True Then
       
        Dim Pattern_Ary_generic() As String
        Dim Pattern_Ary_absolute() As String
        Dim patcnt As Long
        Dim DSSC_Pattern As String
        Dim DSSC_Pattern_Count As Long
        Dim i As Long, j As Long
        Dim digSrc_EQ As String
        Dim DigSrc_Size As Double
        Dim DigSrc_Ary() As String
        Dim DigSrc_LngAry() As Long
        Dim DigSrc_wav As New DSPWave
        Dim PattArray() As String
         
        
        'Pattern_Decompose = TheExec.DataManager.Raw.GetPatternsInSet(Shmoo_Pattern, PatCnt)
        Call GetPatsFromPatSets(PatternName, Pattern_Ary_generic(), patcnt, False)
        Call GetPatsFromPatSets(PatternName, Pattern_Ary_absolute(), patcnt, True)
        
        DSSC_Pattern_Count = 0
        Dim DecodeBit_Str As String
        
        Dim sSrcSigName As String
        Dim tempVarArray As Variant
        Dim site As Variant 'Carter, 20240304
        
        For i = 0 To UBound(Pattern_Ary_generic)
            If Dic_SrcStockIndex.Exists(Pattern_Ary_generic(i)) Then
                DSSC_Pattern = Pattern_Ary_absolute(i)
                DSSC_Pattern_Count = DSSC_Pattern_Count + 1 'Prevent DSSC patterns more than one
                
                tempVarArray = TheHdw.DSSC.pins(DigSrc_pin).Pattern(DSSC_Pattern).Source.Labels.list
                sSrcSigName = tempVarArray(0)
                If sSrcSigName = "" Then
                    sSrcSigName = "FUNC_SRC"
                End If
                
                Call GetSrcString_fromEMAArray(Pattern_Ary_generic(i), TestCase, digSrc_EQ, DigSrc_Size, DigSrc_Ary)
                
                If UCase(digSrc_EQ) Like "*S*" Then
                    Shmoo_Save_Power False
                    Set DigSrc_wav = Nothing
                    DigSrc_wav.CreateConstant 0, CLng(DigSrc_Size)
                    Dim tempStr As String
''                    TempStr = Decide_Switching_Bit(digSrc_EQ, DigSrc_wav, g_ApplyLevelTimingValt, BlockType, DecodeBit_Str)
                    tempStr = Decide_Switching_Bit(digSrc_EQ, DigSrc_wav, g_ApplyLevelTimingValt, BlockType, DecodeBit_Str, , , , , Pattern_Ary_generic(i))
                    For Each site In TheExec.sites.Active
                        DigSrc_LngAry = DigSrc_wav.ConvertDataTypeTo(DspLong).data
                        Exit For
                    Next site
                    
                Else
                    ReDim DigSrc_LngAry(UBound(DigSrc_Ary())) As Long
                    For j = 0 To UBound(DigSrc_Ary())
                        DigSrc_LngAry(j) = CLng(DigSrc_Ary(j))
                    Next j
                    tempStr = digSrc_EQ
                    
                End If
                
                'Add start, Leon Li, when running char, use site loop to get per site value.
                If TheExec.DevChar.Setups.IsRunning = False Then
                    Call DSSC_SetupDigSrcArr_allSites(DSSC_Pattern, DigSrc_pin, sSrcSigName, CLng(DigSrc_Size), DigSrc_LngAry())
                    If DecodeBit_Str = "" Then DecodeBit_Str = tempStr
                    For Each site In TheExec.sites
                        TheExec.Datalog.WriteComment "Site" & site & " " & "DigSrc pattern = " & "DSSC_Pattern" & ": " & DSSC_Pattern & "," & "Src Bits = " & Len(tempStr) & "," & "SourceCode [ First(L) ==> Last(R) ] " & tempStr & ", SelSram :" & DecodeBit_Str
                    Next site
                Else
                    For Each site In TheExec.sites
                        Call DSSC_SetupDigSrcWave_TTR(DSSC_Pattern, DigSrc_pin, "FUNC_SRC", CLng(DigSrc_Size), DigSrc_wav)
    
                        tempStr = vbNullString: DecodeBit_Str = vbNullString
                        For j = 0 To DigSrc_wav.SampleSize - 1
                            tempStr = tempStr & DigSrc_wav.Element(j)
                        Next j
                        DecodeBit_Str = DecodingRealSourceBit(tempStr, BlockType)
                        TheExec.Datalog.WriteComment "Site" & site & " " & "DigSrc pattern = " & "DSSC_Pattern" & ": " & DSSC_Pattern & "," & "Src Bits = " & Len(tempStr) & "," & "SourceCode [ First(L) ==> Last(R) ] " & tempStr & ", SelSram :" & DecodeBit_Str
                    Next site
                End If
                
''                For Each site In TheExec.sites
''                    If DecodeBit_Str = "" Then DecodeBit_Str = TempStr
''                    TheExec.Datalog.WriteComment "Site" & site & " " & "DigSrc pattern = " & "DSSC_Pattern" & ": " & DSSC_Pattern & "," & "Src Bits = " & Len(TempStr) & "," & "SELSRAM_SEND [ First(L) ==> Last(R) ] " & TempStr & "," & DecodeBit_Str
''                Next site
            End If
        Next i
        
        If DSSC_Pattern_Count > 1 Then TheExec.ErrorLogMessage "Number of DSSC Patterns more than one   "
        If TheExec.flow.enableWord("Dummy") <> True Then
            If DSSC_Pattern = "" Then TheExec.ErrorLogMessage " Can not find corresponding DSSC pattern from DSSC Mapping table"
        End If
        
        digSrc_EQ_GB = digSrc_EQ:: BlockType_GB = BlockType:: DigSrcSize_GB = DigSrc_Size:: dssc_pat_init_GB = DSSC_Pattern:: DigSrc_pin_GB = DigSrc_pin
        
    End If
''======================================================================DSSC Capture Set up================================================================
    'Dim OutDspWave(0) As New DSPWave
    
    ''ReDim OutDspWave(2)
    '
    ''PC modified for DSSC Capture
'    If DigCapture <> "" Then
'        Dim OutDspWave As New DSPWave
'        Dim DigCap_Pin As New PinList
'        Dim DigCap_Sample_Size As Long
'
'        Dim DigCap_DataWidth As Long
'        Dim DSSC_Capture_Out As String
'        Dim DigCap_Arr() As String
'
'        Set OutDspWave = Nothing
'
'        DigCap_Arr() = Split(DigCapture, ":")
'        DigCap_Sample_Size = DigCap_Arr(0)
'        DigCap_Pin = DigCap_Arr(1)
'        Call GeneralDigCapSetting(PattArray(0), DigCap_Pin, DigCap_Sample_Size, OutDspWave)
'        TheExec.Datalog.WriteComment ("Cap Bits = " & DigCap_Sample_Size)
'        TheExec.Datalog.WriteComment ("Cap Pin = " & DigCap_Pin)
'        TheExec.Datalog.WriteComment ("======== Setup Dig Cap Test End   ========")
'     End If
' =========================================================================DSSC============================================================================'
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Functional_T", "Sub_SourceEMA_SelSRM") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function




'20231207: Added to mask ssn core using an individual instance
Public Function SSN_CoreMask(Optional specifyFailFlag As String = vbNullString)
On Error GoTo errHandler
Dim funcName As String: funcName = "SSN_CoreMask"
Dim i As Long
Dim pat As Variant
Dim isAnyCoreMasked As Boolean: isAnyCoreMasked = False

If ssnPatternsDict.Count = 0 Then
    TheExec.Datalog.WriteComment "No SSN pattern was found. Exit Function " & funcName
    Exit Function
End If

'Added to support core mask manually
Dim failFlagArr() As String, Failflag As Variant
If specifyFailFlag <> "" Then

    ''20240222: Clear all existing SSN mask if user define which core to mask
    TheHdw.Digital.ScanNetworks.ClearAllMasks
    TheExec.Datalog.WriteComment "SSN clear all mask to ignore real SSN harvest conditon!!"
    
    ''Clear all SSN related flag from table
    For i = 0 To UBound(SSNMapping)
        failFlagArr = Split(SSNMapping(i).Failflag, ",")
        For Each Failflag In failFlagArr
            For Each site In TheExec.sites
                TheExec.sites(site).FlagState(Failflag) = logicFalse
            Next site
        Next Failflag
    Next i
    
    ''Raise the flag that user defined in argument
    failFlagArr = Split(specifyFailFlag, ";")
    For Each Failflag In failFlagArr
        For Each site In TheExec.sites
            TheExec.sites(site).FlagState(Failflag) = logicTrue
        Next site
    Next Failflag
End If

'Find each matched keyword from table
For Each pat In ssnPatternsDict.Keys
'    If InStr(Pat, "\") = 0 Then
        For i = 0 To UBound(SSNMapping)
            If UCase(pat) Like UCase(SSNMapping(i).patternKeyword) Then
                ''Find and save SSN mapping name
                Dim patternCount As Long, patternPath As Variant, ssnMappingName As String
                patternPath = TheExec.DataManager.Raw.GetPatternsInSet(pat, patternCount)
                ssnMappingName = Split(UCase(patternPath(0)), ".PAT")(0)
                ''Mask SSN core using flags of each SSN set
                Call MaskCoreForSSN(CStr(SSNMapping(i).coreName), CStr(SSNMapping(i).Failflag), CStr(ssnMappingName), isAnyCoreMasked)
                Exit For
            End If
        Next i
'    Else
'        Stop
'    End If
Next

'No core needs to be masked, print datalog (to ensure <TestInst> shows on datalog)
If isAnyCoreMasked = False Then
    TheExec.flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:="No SSN core is masked", ForceResults:=tlForceNone
End If

'Added to support core mask manually
If specifyFailFlag <> "" Then
    For Each Failflag In failFlagArr
        For Each site In TheExec.sites
            TheExec.sites(site).FlagState(Failflag) = logicFalse
        Next site
    Next Failflag
End If

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
