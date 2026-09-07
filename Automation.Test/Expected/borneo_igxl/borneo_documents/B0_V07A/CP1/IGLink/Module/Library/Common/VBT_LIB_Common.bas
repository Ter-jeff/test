Attribute VB_Name = "VBT_LIB_Common"
#Const isUFP = True
Option Explicit
'Revision History:
'V0.0 initial bring up
'V0.1 add keep alive function
'V0.2 add disable compare and enable compare function.
'variable declaration
Public DicDiffPairs As New Scripting.Dictionary  'relocation for minimum VBT with RF code'*****************************************
'*****************************************
'******               Relay controls******
'*****************************************
'20211007 Add because current profile
Public PowerPinCnt_mapping As Integer
Public fs As FileSystemObject
Public mapping_textStream As TextStream
Public create_folderName As Boolean
Public Profile_Folder As String
Public glsRAKerrMsg As String
Public glbRAKBinout As Boolean
Public glbRAKisLoad As Boolean
Private Const moduleName As String = "VBT_LIB_Common"
Private functionName As String


Public Function Relay_Control(Optional relay_on As PinList, Optional relay_off As PinList, Optional WaitTime As Double = 0.003)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
'control relay on off, will auto trim NC pins Tto cover CP, FT both stages
    Dim Pins_On() As String, Pin_Cnt_On As Long
    Dim Pins_Off() As String, Pin_Cnt_Off As Long
    Dim p As Variant
    Dim relayOnStr As String, relayOffStr As String

    
    relayOnStr = vbNullString
    relayOffStr = vbNullString

    theexec.DataManager.DecomposePinList relay_on, Pins_On(), Pin_Cnt_On
    theexec.DataManager.DecomposePinList relay_off, Pins_Off(), Pin_Cnt_Off

    Trim_NC_Pin Pins_On, Pin_Cnt_On
    Trim_NC_Pin Pins_Off, Pin_Cnt_Off

    If Pin_Cnt_On <> 0 Then
        TheHdw.Utility.Pins(relay_on).State = tlUtilBitOn
        For Each p In Pins_On
            If relayOnStr = "" Then
                relayOnStr = relayOnStr & p
            Else
                relayOnStr = relayOnStr & ", " & p
            End If
        Next p
        theexec.Datalog.WriteComment "Relay On : " & relayOnStr
    End If

    If Pin_Cnt_Off <> 0 Then
        TheHdw.Utility.Pins(relay_off).State = tlUtilBitOff
        For Each p In Pins_Off
            If relayOffStr = "" Then
                relayOffStr = relayOffStr & p
            Else
                relayOffStr = relayOffStr & ", " & p
            End If
        Next p
        theexec.Datalog.WriteComment "Relay off : " & relayOffStr
    End If

    Wait WaitTime
    
    Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Relay_Control") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

'*****************************************
'******         free run clk, nWire ******
'*****************************************


Public Function StartSBClock(SBFreq As Double) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim SBC_Enable As Long
    'Dim SBFreq As Double
    'TheExec.Datalog.WriteComment "******************  Enable Support BD clock ****************"
    'SBFreq = TheExec.specs.Globals("SBC_Freq_Var").ContextValue
    If glb_TesterType = "Jaguar" Then
        With TheHdw.DIB.SupportBoardClock
            .Connect
            .Frequency = SBFreq
            .Vih = XI0_ref_VOH ' Max is 6V
            .Vil = 0 ' Min is -1V
            .start
        End With
    ElseIf glb_TesterType = "UltraFLEXplus" Then
    
    End If
    SBC_Enable = 1
    theexec.Flow.TestLimit SBC_Enable, 1, 1, tlSignGreaterEqual, tlSignLessEqual, Tname:="SBC enable" 'BurstResult=1:Pass
    'printing in data log
    theexec.Datalog.WriteComment "********** support board clock = " & Format(SBFreq / 1000000, "0.000") & " Mhz, Clock_Vih = " _
                             & XI0_ref_VOH & " V, Clock_Vil = " & XI0_ref_VOL & " V  *******"
    
    Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "StartSBClock") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function StopSBClock()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim SBC_Enable As Long
    ' Stop and disconnect the support board clock.
    With TheHdw.DIB.SupportBoardClock
        .stop
        .Disconnect
    End With
    SBC_Enable = 0
    theexec.Flow.TestLimit SBC_Enable, 1, 1, tlSignGreaterEqual, tlSignLessEqual, Tname:="SBC disable" 'BurstResult=1:Pass
    'printing in data log
    theexec.Datalog.WriteComment "******************  Disable Support BD clock ****************"
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "StopSBClock") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function FreeRunClk_Disable(PortName As String, Optional PowerDown_Flag As Boolean = False) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim site As Variant
    
    If glb_TesterType = "Jaguar" Then
        
    ElseIf glb_TesterType = "UltraFLEXplus" Then
        If InStr(PortName, "_Port") <> 0 Then
            PortName = Replace(PortName, "_Port", "")
        End If
    End If
    
    Call Disable_FRC(PortName)  '''''Support multiple nWire port 20170718'''''''''''''
''    ' Disable the nWire engine.
''    'If TheExec.Flow.EnableWord("XI0_nWire") = True Then removed
''        TheHdw.Protocol.ports(PortName).Halt
''        TheHdw.Protocol.ports(PortName).Enabled = False     'scope out point
''    'End If
    If PowerDown_Flag = False Then theexec.Flow.TestLimit 0, 0, 0, tlSignGreaterEqual, tlSignLessEqual, Tname:="nWire halt" 'BurstResult=1:Pass
    'printing to data log
    'TheExec.Datalog.WriteComment "******************  Disable freerunning clock ****************"
    
    ''upload to global constant
    FreeRunFreq_debug = 0
    clock_Vih_debug = 0
    clock_Vil_debug = 0
    
    glb_ApplyLevelTiming_FRC_Flag = False
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "FreeRunClk_Disable") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


'[20230809][T-All][Oliver] remove the *.Reinitialize for DCVS and DCVI instrument setting
Function Start_Profile_DebugOnly(PinName As PinList, WhatToCapture As String, SampleRate As Double, sampleSize As Long, Optional CapSignalName As String = "Capture_signal")
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
'start current or voltage profile capturing

' Wait if another capture is running
    Do While TheHdw.DCVS.Pins(PinName).Capture.IsRunning = True
    Loop
    
    'Create a SIGNAL to set up instrument
    TheHdw.DCVS.Pins(PinName).Capture.Signals.Add CapSignalName
    
    'Set this as the default signal
    TheHdw.DCVS.Pins(PinName).Capture.Signals.DefaultSignal = CapSignalName
    
    'Define the signal used for the capture
    With TheHdw.DCVS.Pins(PinName).Capture.Signals.item(CapSignalName)
'        .Reinitialize
        If (WhatToCapture = "I") Then
            .mode = tlDCVSMeterCurrent
            .range = TheHdw.DCVS.Pins(PinName).CurrentRange.max '2
        Else
            .mode = tlDCVSMeterVoltage
            .range = 10
        End If
        .SampleRate = SampleRate
        .sampleSize = sampleSize
    
    End With
    
    ' Setup the hardware by loading the signal
    TheHdw.DCVS.Pins(PinName).Capture.Signals.item(CapSignalName).LoadSettings
    
    ' Start the capture
    TheHdw.DCVS.Pins(PinName).Capture.Signals.item(CapSignalName).Trigger

    Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Start_Profile_DebugOnly") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


'[20230809][T-All][Oliver] remove the *.Reinitialize for DCVS and DCVI instrument setting
Public Function start_profile_DCVI(PinName As String, WhatToCapture As String, SampleRate As Double, sampleSize As Long)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Do While TheHdw.DCVI.Pins(PinName).Capture.IsCaptureDone = False        ' Wait if another capture is running
    Loop
    TheHdw.DCVI.Pins(PinName).Capture.Signals.Add "Capture_signal"              'Create a SIGNAL to set up instrument
    TheHdw.DCVI.Pins(PinName).Capture.Signals.DefaultSignal = "Capture_signal"  'Set this as the default signal
    
    With TheHdw.DCVI.Pins(PinName)
        .Gate = False
        .mode = tlDCVIModeCurrent
        .Voltage = 6
        .VoltageRange.Autorange = True
        .CurrentRange.Autorange = True
        .Current = 0
        .Connect tlDCVIConnectDefault
        .Gate = True
    End With
     
    With TheHdw.DCVI.Pins(PinName).Capture.Signals.item("Capture_signal")    ' Define the signal used for the capture
    '    .Reinitialize
        If (WhatToCapture = "I") Then
            .mode = tlDCVIMeterCurrent
            .range = 0.02
        Else
            .mode = tlDCVIMeterVoltage
             .range = 7
        End If
        .SampleRate = SampleRate
        .sampleSize = sampleSize
    End With
    
    TheHdw.DCVI.Pins(PinName).Capture.Signals.item("Capture_signal").LoadSettings  ' Setup the hardware by loading the signal
    TheHdw.DCVI.Pins(PinName).Capture.Signals.item("Capture_signal").Trigger            ' Start the capture
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "start_profile_DCVI") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

' [20230407][RF] Add DCVI plot current profile
' [20230420][All][Tank] change include pin string to LCase
' [20230511][All][Tank] Fix VBT error about dictionary .Exist to .Exists
' [20240419][All][Clyde] Merge function to one Print_IProfileValue function
Public Function Plot_Profile(PinName As PinList, Optional CapSignalName As String = "Capture_signal", Optional ExportWaveform As Boolean = False, Optional PlotWaveform As Boolean = False, Optional Calculate_ProfileInfo As Boolean = False, Optional percent_control As Double)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
'Plot profiles

    Dim dspw As New DSPWave
    Dim Label As String
    Dim site As Variant
    Dim Pin_Ary() As String
    Dim Pin_Cnt As Long
    Dim p As Variant
    Dim filename As String
    Dim lastBurstPat As New SiteVariant
    Dim isGrp As New SiteBoolean
    Dim lastLabel As New SiteVariant
    Dim day_code As String
    Dim Current_Insatance As String
    Current_Insatance = m1_InstanceName
    If Current_Insatance = "" Then theexec.Datalog.WriteComment "<ERROR> Instance name is empty.Please check instance global name is defined."
    
    
    Dim Current_Dirctory As String
    Dim Profile_Split() As String
    Dim Mapping As String
    Dim original_InstName As String
    Dim Instance_mappingName As String
    Dim mapping_txt As String
    Dim TestName As String
    Dim CurrentProfile As New DSPWave
    Dim Plot_Profile_Label() As String
    ReDim Plot_Profile_Label(0)
    Dim T1 As Double
    Dim T2 As Double
    Dim sampleR As String
    Dim sampleSize As String
    Dim sPowerPin As String
    Dim sDCVS_Pin As String
    Dim sDCVI_Pin As String
    'Currentprofile new feature
    
    Dim TempStr As String
    Dim DspwTempAry() As Double
    Dim DSPTempStr As String
    Dim ForceV_Val As Double
    Dim s_RawdataItem As String
    
    DSPTempStr = vbNullString
    Dim x As Double
    Dim Folder_Bk As String
    Dim Arr_TP_Name() As String
    Dim TP_Name As String
    
    Arr_TP_Name = Split(Replace(theexec.TestProgram.PathAndName, CurDir & "\", ""), "_")
    TP_Name = Arr_TP_Name(0) & "_" & Arr_TP_Name(1)
    Folder_Bk = "D:\Local_TP\Profile_" & TP_Name
    
    If glb_Boolean_export = False Then
        Call Export_All
        glb_Boolean_export = True
    End If
    
    If gl_EnableVoltageProfile = False And gl_EnableCurrentProfile = False Then
    
        If UCase(glb_WhatToCapture) = "V" Then
           gl_EnableVoltageProfile = True
           Profile_byflow = True
        ElseIf UCase(glb_WhatToCapture) = "I" Then
           gl_EnableCurrentProfile = True
           Profile_byflow = True
        Else
        End If
     
    End If
    
    If percent_control = 0 Then percent_control = 0.99
    If theexec.enableWord("DisablePlot_IProfile") = False Or theexec.enableWord("Print_IProfileValue") = True Then
        Call theexec.DataManager.DecomposePinList(PinName, Pin_Ary(), Pin_Cnt)
        TestName = theexec.DataManager.instancename
        If create_folderName = False Then
            Profile_Folder = "X" & CStr(XCoord(0)) & "Y" & CStr(YCoord(0)) & "_" & right("0" & CStr(Month(Now)), 2) & right("0" & CStr(Day(Now)), 2) & right("0" & CStr(Hour(Now)), 2) & right("0" & CStr(Minute(Now)), 2)
            create_folderName = True
        End If
        If theexec.enableWord("DisablePlot_IProfile") = False Then
            
            If UCase(CurDir) Like "*X:\*" Then
                Call Print_Error_Message(Warning_Info, "VBT_LIB_Common", "TP is in X:\, change Profile to D:\Local_TP\Profile_" & TP_Name)
                
                If Dir("X:\Local_TP", vbDirectory) = Empty Then
                    Call Print_Error_Message(Warning_Info, "VBT_LIB_Common", "D:\Local_TP folder is not exist skip export profile")
                    GoTo Skip_Write_profile
                Else
                    If Dir(Folder_Bk, vbDirectory) = Empty Then
                        MkDir Folder_Bk
                    Else
                    End If
                End If
                
                Current_Dirctory = Folder_Bk & "\" & Profile_Folder
            Else
                Current_Dirctory = CurDir & "\" & Profile_Folder
            End If
            
            If Dir(Current_Dirctory, vbDirectory) = Empty Then
                If ExportWaveform Then
                     MkDir Current_Dirctory
                End If
            End If
        End If

        sPowerPin = CStr(PinName)
        Call SortAllPinInstrumentType(sPowerPin, sDCVS_Pin, sDCVI_Pin)
        
        If sDCVS_Pin <> "" Then
            Do While TheHdw.DCVS.Pins(sDCVS_Pin).Capture.IsRunning = True
            Loop
        End If
        
        If sDCVI_Pin <> "" Then
            Do While TheHdw.DCVI.Pins(sDCVI_Pin).Capture.IsCaptureDone = False
            Loop
        End If

        day_code = right("0" & CStr(Month(Now)), 2) & right("0" & CStr(Day(Now)), 2)
        day_code = day_code & right("0" & CStr(Hour(Now)), 2) & right("0" & CStr(Minute(Now)), 2) & right("0" & CStr(Second(Now)), 2)
        ' Get the captured samples from the instrument
        
        
        'original instance name ex:SocSaChain_MXXXXX_SRMDSSC_MN_CREA0_S_PL00_CH_cxsa_cxsb_cxsc_cxsd_cxse_cxsg_saa_unc_aut_MSPXXX_MSXXXX_MSAXXX_MSIXXX_MSLXXX_MSPXXX_DM_LV
        original_InstName = Profile_Header
        Profile_Split = Split(Profile_Header, "_")
        If UBound(Profile_Split) < 6 Then
            Instance_mappingName = original_InstName
        Else
            PowerPinCnt_mapping = PowerPinCnt_mapping + 1
            Instance_mappingName = Profile_Split(0) & "_" & Profile_Split(1) & "_" & Profile_Split(2) & "_" & Profile_Split(3) & "_" & Profile_Split(4) & "_" & Profile_Split(5) & "_" & Format(PowerPinCnt_mapping, "000000")
        End If
          
        For Each p In Pin_Ary
            sPowerPin = LCase(p)
            If gl_GetInstrumentType_Dic.Exists(sPowerPin) Then
                If LCase(gl_GetInstrumentType_Dic(sPowerPin)) Like "*dcvs*" Then
                    dspw = TheHdw.DCVS.Pins(sPowerPin).Capture.Signals(CapSignalName).DSPWave
                    For Each site In theexec.sites
                        If TheHdw.DCVS.Pins(sPowerPin).Voltage.Output = tlDCVSVoltageMain Then
                            ForceV_Val = TheHdw.DCVS.Pins(sPowerPin).Voltage.Main
                        Else
                            ForceV_Val = TheHdw.DCVS.Pins(sPowerPin).Voltage.Alt
                        End If
                        Exit For
                    Next site
                    
                ElseIf LCase(gl_GetInstrumentType_Dic(sPowerPin)) Like "*dcvi*" Then
                    dspw = TheHdw.DCVI.Pins(sPowerPin).Capture.Signals(CapSignalName).DSPWave
                    ForceV_Val = TheHdw.DCVI.Pins(sPowerPin).Voltage.value
                Else
                    Call Print_Error_Message(Error_Warning_Info.Warning_Info, "VBT_LIB_Common", "Plot_Profile", "Use wrong instrument to plot current profile !!")
                End If
                
                
                If theexec.TesterMode = testModeOffline Then
                    If LCase(gl_GetInstrumentType_Dic(sPowerPin)) Like "*dcvs*" Then
                        sampleR = CStr(TheHdw.DCVS.Pins(sPowerPin).Capture.Signals.item(CapSignalName).SampleRate)
                        sampleSize = CStr(TheHdw.DCVS.Pins(sPowerPin).Capture.Signals.item(CapSignalName).sampleSize)
                    ElseIf LCase(gl_GetInstrumentType_Dic(sPowerPin)) Like "*dcvi*" Then
                        sampleR = CStr(TheHdw.DCVI.Pins(sPowerPin).Capture.Signals.item(CapSignalName).SampleRate)      'CStr(TheHdw.DCVI.pins(p).Capture.SampleRate)
                        sampleSize = CStr(TheHdw.DCVI.Pins(sPowerPin).Capture.Signals.item(CapSignalName).sampleSize)
                    End If
                    Set dspw = Nothing
                    dspw.CreateRamp 0, 1 / sampleSize, sampleSize, DspDouble
                Else
                End If
                
                If glb_ProfileFilter Then
                    Call ProfileData_Filter(p, dspw, 0.02, ForceV_Val) ''0928 debug
                Else
                End If
                
                For Each site In theexec.sites
                    If Calculate_ProfileInfo Then
                        Call Print_IProfileValue(sPowerPin, TestName, dspw, site, percent_control)
                    End If
                    If theexec.enableWord("DisablePlot_IProfile") = False Then
                        'new instance name ex:SocSaChain_MXXXX_SRMDSSC_MN_CREA0_S_00001
        '                TheHdw.Digital.Patgen.ReadLastStart lastBurstPat, isGrp, lastLabel
                        If LCase(gl_GetInstrumentType_Dic(sPowerPin)) Like "*dcvs*" Then
                            sampleR = CStr(TheHdw.DCVS.Pins(sPowerPin).Capture.Signals.item(CapSignalName).SampleRate)
                            sampleSize = CStr(TheHdw.DCVS.Pins(sPowerPin).Capture.Signals.item(CapSignalName).sampleSize)
                            If TheHdw.DCVS.Pins(sPowerPin).Meter.mode = tlDCVSMeterCurrent Then
                                Label = "Current Profile for Site: " & site & " " & " " & CapSignalName & "Pin :" & " " & sPowerPin
                                filename = "CurrentProfile-Site" & site & "-" & sPowerPin & "-" & sampleR & "-" & sampleSize & "-" & Instance_mappingName & "-" & day_code & ".txt"
                            Else
                                Label = "Voltage Profile for Site: " & site & " " & " " & CapSignalName & "Pin :" & " " & sPowerPin
                                filename = "VoltageProfile-Site" & site & "-" & sPowerPin & "-" & sampleR & "-" & sampleSize & "-" & Instance_mappingName & "-" & day_code & ".txt"
                            End If
                        ElseIf LCase(gl_GetInstrumentType_Dic(sPowerPin)) Like "*dcvi*" Then
                            sampleR = CStr(TheHdw.DCVI.Pins(sPowerPin).Capture.Signals.item(CapSignalName).SampleRate)      'CStr(TheHdw.DCVI.pins(p).Capture.SampleRate)
                            sampleSize = CStr(TheHdw.DCVI.Pins(sPowerPin).Capture.Signals.item(CapSignalName).sampleSize)       'CStr(TheHdw.DCVI.pins(p).Capture.SampleSize)
                            If TheHdw.DCVI.Pins(sPowerPin).Capture.Signals.item(CapSignalName).mode = tlDCVIMeterCurrent Then     'TheHdw.DCVI.pins(p).Meter.mode
                                Label = "Current Profile for Site: " & site & " " & " " & CapSignalName & "Pin :" & " " & sPowerPin
                                filename = "CurrentProfile-Site" & site & "-" & sPowerPin & "-" & sampleR & "-" & sampleSize & "-" & Instance_mappingName & "-" & day_code & ".txt"
                            Else
                                Label = "Voltage Profile for Site: " & site & " " & " " & CapSignalName & "Pin :" & " " & sPowerPin
                                filename = "VoltageProfile-Site" & site & "-" & sPowerPin & "-" & sampleR & "-" & sampleSize & "-" & Instance_mappingName & "-" & day_code & ".txt"
                            End If
                        Else
                        End If
                        
                        If PlotWaveform Then dspw.Plot Label   'for pliot
                            If ExportWaveform Then
    
                                 ReDim DspwTempAry(dspw.sampleSize)
                                 DspwTempAry = dspw.data
    
                                 s_RawdataItem = Replace(filename, ".txt", "")
    
                                 Print #13, s_RawdataItem
                                    For x = 0 To UBound(DspwTempAry)
    
                                        If DSPTempStr <> "" Then
                                             DSPTempStr = DSPTempStr & "," & CStr(DspwTempAry(x))
                                        Else
                                             DSPTempStr = CStr(DspwTempAry(x))
                                        End If
    
                                        If (x + 1) Mod 10 = 0 Then
                                            Print #13, DSPTempStr
                                            DSPTempStr = vbNullString
                                        End If
                                    Next x
                                    If DSPTempStr <> "" Then
                                        Print #13, DSPTempStr
                                        DSPTempStr = vbNullString
                                    End If
                                    Call ProfileRecord("Action_Plot")
                            End If
    
                            
                            
                            If Calculate_ProfileInfo Then
                                If UBound(Plot_Profile_Label) = 0 Then
                                    Plot_Profile_Label(0) = Label
                                Else
                                    Plot_Profile_Label(UBound(Plot_Profile_Label)) = Label
                                End If
                                ReDim Preserve Plot_Profile_Label(UBound(Plot_Profile_Label) + 1)
                                Call AddStoredCaptureData(Label, dspw(site))
                            End If
                        End If
                        
                        'If LCase(GetInstrument(CStr(p), 0)) <> "hexvs" Then DSPW.Clear
                Next site
                If glb_DisablePlot_IProfile = False Then
                    TempStr = Current_Dirctory & "\"
                    Set fs = CreateObject("Scripting.FileSystemObject")
                    
                    mapping_txt = CStr(sPowerPin) + "    " + original_InstName + " : " + Instance_mappingName
                    If fs.FileExists(TempStr & "mapping.txt") = False Then
                        
                        Set mapping_textStream = fs.CreateTextFile(TempStr & "mapping.txt", True)
                        mapping_textStream.WriteLine "ProjectName : " & theexec.TestProgram.name
                        mapping_textStream.WriteLine "JobName : " & theexec.CurrentJob
                        mapping_textStream.WriteLine ""
                        mapping_textStream.WriteLine "---------------------------------------------------------"
                        mapping_textStream.WriteLine "| PowerPin | OriginalInstanceName : MappingInstanceName |"
                        mapping_textStream.WriteLine "---------------------------------------------------------"
                        If UBound(Profile_Split) >= 6 Then mapping_textStream.WriteLine mapping_txt
                        'mapping_textStream.WriteLine mapping_txt
                    Else
                        If ExportWaveform Then
                            'mapping_textStream.WriteLine mapping_txt
                            If UBound(Profile_Split) >= 6 Then mapping_textStream.WriteLine mapping_txt
                        End If
                    End If
                End If
                Set dspw = Nothing
                'End If
            End If
        Next p

    End If
    
Skip_Write_profile:
    glb_WhatToCapture = vbNullString
    If Profile_byflow = True Then
        gl_EnableVoltageProfile = False
        gl_EnableCurrentProfile = False
        Profile_byflow = False
    Else
    End If
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Plot_Profile") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Plot_profile_DCVI(PinName As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim dspw As New DSPWave
    Dim Label As String
    Dim site As Variant
    
    
    ' Get the captured samples from the instrument
    dspw = TheHdw.DCVI.Pins(PinName).Capture.Signals.item("Capture_signal").DSPWave
    
    For Each site In theexec.sites.Active
        If TheHdw.DCVI.Pins(PinName).Meter.mode = tlDCVIMeterCurrent Then
            Label = "Current Profile for Site: " & site
        Else
            Label = "Voltage Profile for Site: " & site
        End If
        
        dspw.Plot Label
        
    Next site
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    If isDebugMode Then theexec.AddOutput "Error in the Plot Profile"
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Plot_profile_DCVI") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

' [20230407][RF] Add DCVI current profile
' [20230420][All][Tank] modify instrument name
' [20230809][All][Oliver] modify for current profile other sheet method
Function Start_Profile_AutoResolution(PinName As String, WhatToCapture As String, Optional CapSignalName As String = "Capture_signal", Optional Plottime As Double = 0, Optional ByFlow As Boolean = False, Optional nSampleRate As Long = 1)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
'start current or voltage profile capturing

    Dim Profile_AllPin() As String
    Dim PinCnt As Long
    Dim HexPins As String
    Dim UVSPins As String
    Dim VSMPins As String
    Dim VS5AMPins As String
    Dim VS800mAPins As String
    Dim pin As Variant
    Dim tempsplitary() As String
    
    Dim Profile_SampleRate_Hex As Double
    Dim Profile_SampleSize_Hex As Double
    Dim Down_SampleRatio_Hex As Long

    Dim Profile_SampleRate_UVS As Double
    Dim Profile_SampleSize_UVS As Double
    Dim Down_SampleRatio_UVS As Long

    Dim Profile_SampleRate_VSM As Double
    Dim Profile_SampleSize_VSM As Double
    Dim Down_SampleRatio_VSM As Long

    Dim Profile_SampleRate_VS5A As Double
    Dim Profile_SampleSize_VS5A As Double
    Dim Down_SampleRatio_VS5A As Long
    
    Dim Profile_SampleRate_VS800mA As Double
    Dim Profile_SampleSize_VS800mA As Double
    Dim Down_SampleRatio_VS800mA As Long
    
    Dim Profile_SampleRate_DC07 As Double
    Dim Profile_SampleSize_DC07 As Double
    Dim Down_SampleRatio_DC07 As Long
    
    Dim Profile_SampleRate_DC30 As Double
    Dim Profile_SampleSize_DC30 As Double
    Dim Down_SampleRatio_DC30 As Long
    
    Dim Profile_SampleRate_DC75 As Double
    Dim Profile_SampleSize_DC75 As Double
    Dim Down_SampleRatio_DC75 As Long
    
    glb_WhatToCapture = WhatToCapture
    
    If gl_EnableCurrentProfile = False And gl_EnableVoltageProfile = False Then
        Profile_byflow = True
    Else
        Profile_byflow = False
    End If
    
    If ByFlow Then
        Profile_Header = UCase(theexec.Flow.CurrentFlowSheetName)
    Else
        Profile_Header = UCase(theexec.DataManager.instancename)
    End If
    ''Add buffer to avoid missing the capture data
    If Plottime <> 0 Then
        If Plottime > 1 Then
            Plottime = Plottime + 0.1
        Else
            Plottime = Plottime + 0.01
        End If
    Else ''Minimum capture time
        Plottime = 0.001
    End If
        
    'SplitPinByinstrument PinName, HexPins, UVSPins, VSMPins, VS5AMPins, VS800mAPins
        
    Dim sPin_HEXVS As String
    Dim sPin_VHDVS As String
    Dim sPin_VSM As String
    Dim sPin_VS800MA As String
    Dim sPin_VS5A As String
    Dim sPin_DC07 As String
    Dim sPin_DC30 As String
    Dim sPin_DC75 As String
    
    Call GetAllPinInstrument(PinName, sPin_HEXVS, sPin_VHDVS, sPin_VSM, sPin_VS800MA, sPin_VS5A, sPin_DC07, sPin_DC30, sPin_DC75)
    
    If theexec.enableWord("DownSample_IProfile") = True Then
    
        tempsplitary = Split(HexPins, ",")
        Down_SampleRatio_Hex = GetDownSampleRatio(glbConstIns_HEXVS, UBound(tempsplitary))
        tempsplitary = Split(UVSPins, ",")
        Down_SampleRatio_UVS = GetDownSampleRatio(glbConstIns_VHDVS, UBound(tempsplitary))
        tempsplitary = Split(VSMPins, ",")
        Down_SampleRatio_VSM = GetDownSampleRatio(glbConstIns_VSM, UBound(tempsplitary))
        tempsplitary = Split(VS5AMPins, ",")
        Down_SampleRatio_VS5A = GetDownSampleRatio(glbConstIns_VS5A, UBound(tempsplitary))
        tempsplitary = Split(VS800mAPins, ",")
        Down_SampleRatio_VS800mA = GetDownSampleRatio(glbConstIns_VS800MA, UBound(tempsplitary))
        tempsplitary = Split(sPin_DC07, ",")
        Down_SampleRatio_DC07 = GetDownSampleRatio(glbConstIns_DC07, UBound(tempsplitary))
        tempsplitary = Split(sPin_DC30, ",")
        Down_SampleRatio_DC30 = GetDownSampleRatio(glbConstIns_DC30, UBound(tempsplitary))
        tempsplitary = Split(sPin_DC75, ",")
        Down_SampleRatio_DC75 = GetDownSampleRatio(glbConstIns_DC75, UBound(tempsplitary))
    End If
         
    If sPin_HEXVS <> "" Then
        Profile_SampleRate_Hex = nSampleRate
        Call ProfileAutoResolution(glbConstIns_HEXVS, Plottime, Profile_SampleSize_Hex, Profile_SampleRate_Hex, _
                                    Down_SampleRatio_Hex, nSampleRate)
        StartProfile sPin_HEXVS, WhatToCapture, Profile_SampleRate_Hex, Profile_SampleSize_Hex, CapSignalName, glbConstIns_HEXVS
    End If
    
    If sPin_VHDVS <> "" Then
        Profile_SampleRate_UVS = nSampleRate
        Call ProfileAutoResolution(glbConstIns_VHDVS, Plottime, Profile_SampleSize_UVS, Profile_SampleRate_UVS, _
                                    Down_SampleRatio_UVS, nSampleRate)
        StartProfile sPin_VHDVS, WhatToCapture, Profile_SampleRate_UVS, Profile_SampleSize_UVS, CapSignalName, glbConstIns_VHDVS
    End If
    
    If sPin_VSM <> "" Then
        Profile_SampleRate_VSM = nSampleRate
        Call ProfileAutoResolution(glbConstIns_VSM, Plottime, Profile_SampleSize_VSM, Profile_SampleRate_VSM, _
                                    Down_SampleRatio_VSM, nSampleRate)
        StartProfile sPin_VSM, WhatToCapture, Profile_SampleRate_VSM, Profile_SampleSize_VSM, CapSignalName, glbConstIns_VSM
    End If
    
    If sPin_VS5A <> "" Then
        Profile_SampleRate_VS5A = nSampleRate
        Call ProfileAutoResolution(glbConstIns_VS5A, Plottime, Profile_SampleSize_VS5A, Profile_SampleRate_VS5A, _
                                    Down_SampleRatio_VS5A, nSampleRate)
        StartProfile sPin_VS5A, WhatToCapture, Profile_SampleRate_VS5A, Profile_SampleSize_VS5A, CapSignalName, glbConstIns_VS5A
    End If
    
    If sPin_VS800MA <> "" Then
        Profile_SampleRate_VS800mA = nSampleRate
        Call ProfileAutoResolution(glbConstIns_VS800MA, Plottime, Profile_SampleSize_VS800mA, Profile_SampleRate_VS800mA, _
                                    Down_SampleRatio_VS800mA, nSampleRate)
        StartProfile sPin_VS800MA, WhatToCapture, Profile_SampleRate_VS800mA, Profile_SampleSize_VS800mA, CapSignalName, glbConstIns_VS800MA
    End If
    
    If sPin_DC07 <> "" Then
        Profile_SampleRate_DC07 = nSampleRate
        Call ProfileAutoResolution_DCVI(glbConstIns_DC07, Plottime, Profile_SampleSize_DC07, Profile_SampleRate_DC07, Down_SampleRatio_DC07)
        StartProfile_DCVI sPin_DC07, WhatToCapture, Profile_SampleRate_DC07, Profile_SampleSize_DC07, CapSignalName, glbConstIns_DC07
    End If
    
    If sPin_DC30 <> "" Then
        Profile_SampleRate_DC30 = nSampleRate
        Call ProfileAutoResolution_DCVI(glbConstIns_DC30, Plottime, Profile_SampleSize_DC30, Profile_SampleRate_DC30, Down_SampleRatio_DC30)
        StartProfile_DCVI sPin_DC30, WhatToCapture, Profile_SampleRate_DC30, Profile_SampleSize_DC30, CapSignalName, glbConstIns_DC30
    End If
    
    If sPin_DC75 <> "" Then
        Profile_SampleRate_DC75 = nSampleRate
        Call ProfileAutoResolution_DCVI(glbConstIns_DC75, Plottime, Profile_SampleSize_DC75, Profile_SampleRate_DC75, Down_SampleRatio_DC75)
        StartProfile_DCVI sPin_DC75, WhatToCapture, Profile_SampleRate_DC75, Profile_SampleSize_DC75, CapSignalName, glbConstIns_DC75
    End If
    
   
   '==============for  DownSample_IProfile  debug==================================
    theexec.Datalog.WriteComment "DownSample_IProfile:" & theexec.enableWord("DownSample_IProfile")
    theexec.Datalog.WriteComment "Profile_SampleSize_Hex : " & Profile_SampleSize_Hex & "     Profile_SampleRate_Hex : " & Profile_SampleRate_Hex & "    Down_SampleRatio_Hex : " & Down_SampleRatio_Hex
    theexec.Datalog.WriteComment "Profile_SampleSize_UVS : " & Profile_SampleSize_UVS & "     Profile_SampleRate_UVS : " & Profile_SampleRate_UVS & "    Down_SampleRatio_UVS : " & Down_SampleRatio_UVS
    theexec.Datalog.WriteComment "Profile_SampleSize_VSM : " & Profile_SampleSize_VSM & "     Profile_SampleRate_VSM : " & Profile_SampleRate_VSM & "    Down_SampleRatio_VSM : " & Down_SampleRatio_VSM
    theexec.Datalog.WriteComment "Profile_SampleSize_VS5A : " & Profile_SampleSize_VS5A & "     Profile_SampleRate_VS5A : " & Profile_SampleRate_VS5A & "    Down_SampleRatio_VS5A : " & Down_SampleRatio_VS5A
    theexec.Datalog.WriteComment "Profile_SampleSize_VS800mA : " & Profile_SampleSize_VS800mA & "     Profile_SampleRate_VS800mA : " & Profile_SampleRate_VS800mA & "    Down_SampleRatio_VS800mA : " & Down_SampleRatio_VS800mA
    
    theexec.Datalog.WriteComment "====DCVI===="
    theexec.Datalog.WriteComment "Profile_SampleSize_DC07 : " & Profile_SampleSize_DC07 & "     Profile_SampleRate_DC07 : " & Profile_SampleRate_DC07 & "    Down_SampleRatio_DC07 : " & Down_SampleRatio_DC07
    theexec.Datalog.WriteComment "Profile_SampleSize_DC30 : " & Profile_SampleSize_DC30 & "     Profile_SampleRate_DC30 : " & Profile_SampleRate_DC30 & "    Down_SampleRatio_DC30 : " & Down_SampleRatio_DC30
    theexec.Datalog.WriteComment "Profile_SampleSize_DC75 : " & Profile_SampleSize_DC75 & "     Profile_SampleRate_DC75 : " & Profile_SampleRate_DC75 & "    Down_SampleRatio_DC75 : " & Down_SampleRatio_DC75
    theexec.Datalog.WriteComment "====DCVI===="
   '==============for  DownSample_IProfile  debug==================================
   
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Start_Profile_AutoResolution") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function Print_Footer(PrintInfo As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    theexec.Datalog.WriteComment "******************************"
    theexec.Datalog.WriteComment "*print: " & PrintInfo & " end*"
    theexec.Datalog.WriteComment "******************************"

    ''20210504 reset the datalog format
    Init_Datalog_Setup
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Print_Footer") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function Print_Header(PrintInfo As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    theexec.Datalog.WriteComment "********************************"
    theexec.Datalog.WriteComment "*print: " & PrintInfo & " start*"
    theexec.Datalog.WriteComment "********************************"

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Print_Header") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

'Public Function Start_Current_Profile(PinName As PinList, SampleRate As Double, SampleSize As Long)
'    TheExec.EnableWord("Profile_Voltage") = False
'    Call Start_Profile(PinName, "I", SampleRate, SampleSize)
'End Function
'
'Public Function Start_Voltage_Profile(PinName As PinList, SampleRate As Double, SampleSize As Long)
'    TheExec.EnableWord("Profile_Current") = False
'    Call Start_Profile(PinName, "V", SampleRate, SampleSize)
'End Function

Public Function Print_PgmInfo(separate As Boolean)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    theexec.Datalog.WriteComment "********************************"
    theexec.Datalog.WriteComment "*print: Program information    *"
    theexec.Datalog.WriteComment "********************************"
    
    OS_Info
    IGXL_Info
    RunOptions_Info
    Version_Info
    Print_Version
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Print_PgmInfo") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
'*****************************************
'******            Read/Write EPPROM******
'*****************************************
Public Function Write_DIB_EEPROM(Optional DIB_SerialNumber As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim CurrJob As String
    Dim config(2) As Long
    config(0) = 32768
    config(1) = 0
    config(2) = 0
    
    TheHdw.DIB.PIBEEPROM.program (config)
'''    CurrJob = TheExec.CurrentJob    'CP/FT judge
'''    change to LCase to prevent ft1 FT1 issue

    If theexec.Flow.enableWord("Write_EEPROM_DIBID") = True Then
'''     If (CurrJob Like "cp*") Then
        If LCase(theexec.CurrentJob) Like "cp*" Then
            If RegKeyRead("PROBECARD_ID") <> "" Then
                DIB_SerialNumber = RegKeyRead("PROBECARD_ID")
            Else
                DIB_SerialNumber = vbNullString
            End If
'''     ElseIf (CurrJob Like "ft*") Then
        ElseIf LCase(theexec.CurrentJob) Like "ft*" Then
            If RegKeyRead("LOADBOARD_ID") <> "" Then
                DIB_SerialNumber = RegKeyRead("LOADBOARD_ID")
            Else
                DIB_SerialNumber = vbNullString
            End If
        Else
            Call Print_Error_Message(Error_Warning_Info.Warning_Info, "VBT_LIB_Common", "Write_DIB_EEPROM", "This process Job name not define !!")
        End If
        
        TheHdw.DIB.PIBEEPROM.Record("DIB_SerialNum") = DIB_SerialNumber
        TheHdw.DIB.PIBEEPROM.Record.WriteToHW   'write to HW
        theexec.Datalog.WriteComment ("print: Write DIB ID " & DIB_SerialNumber)
    End If  'end enable wd block
    'only fuse once
    theexec.Flow.enableWord("Write_EEPROM_DIBID") = False
    
    Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Write_DIB_EEPROM") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

'Public Function Write_SPIROM_Check() As Boolean
'
'    Write_SPIROM_Check = False
'
'    Dim Site As Variant
'
'    For Each Site In TheExec.Sites
'        If (write_SPIROM_CheckSum And (2 ^ Site)) = 0 Then
'            Write_SPIROM_Check = True
'            write_SPIROM_CheckSum = (write_SPIROM_CheckSum Or (2 ^ Site))
'        End If
'    Next Site
'
'End Function
Public Function Read_DIB_EEPROM() As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    
    Dim rec() As IDIB_EEPROM_RecordObj
    Dim i As Integer
    Dim DIB_ID As String
    
    If TheHdw.DIB.PIBEEPROM.IsProgrammed Then
        'Debug.Print "The PIB EEPROM is programmed"
        rec = TheHdw.DIB.PIBEEPROM.Record.list
    
'''        For i = 0 To UBound(rec)
'''            Debug.Print "Record " + rec(i).ID + " = " + rec(i).Value
'''        Next i

        DIB_ID = rec(0).value
        theexec.Datalog.WriteComment ("print: Read DIB ID is " & DIB_ID)
    Else
        'Debug.Print "The PIB EEPROM is not programmed"
        theexec.Datalog.WriteComment ("print: Read DIB ID is not programmed, read fail")
    End If
    
    Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Read_DIB_EEPROM") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' [20240110][All][Riko] Add get FT handler temp
Public Function ReadProberTemp(Temp_Hilimit As Double, Temp_Lolimit As Double)

Dim Prober_Temp As Double
Dim Prober_Temp_str As String
Dim Handler_Temp As Double
Dim tempHandler_Temp() As String
Dim Handler_Temp_str As String
Dim Handler_temp_Arr() As String
Dim Handler_Temp_Site As New SiteDouble
Dim site As Variant
On Error GoTo errHandler

    If UCase(theexec.CurrentChanMap) Like "CHANNELMAP_FT*" Then
        theexec.Datalog.WriteComment ""
        'TheExec.Datalog.WriteComment "/********* Read Handler Temperature *********/"
        theexec.Datalog.WriteComment "********************************"
        theexec.Datalog.WriteComment "*print: Read Handler Temperature*"
        theexec.Datalog.WriteComment "********************************"
    
        If CStr(RegKeyRead("Dut_Temperature")) <> vbNullString Then
            Handler_Temp_str = RegKeyRead("Dut_Temperature")
            If InStr(1, Handler_Temp_str, "_") Then
                ''Case: _Site0Temp_Site1Temp___
                Handler_Temp_str = mid(Handler_Temp_str, 2, Len(Handler_Temp_str) - 2) ''Site0Temp_Site1Temp__
                Handler_temp_Arr = Split(Handler_Temp_str, "_")
                For Each site In theexec.sites.Active
                    If IsNumeric(Handler_temp_Arr(site)) Then
                        Handler_Temp_Site = CDbl(Handler_temp_Arr(site))
                        'gL_ProductionTemp = Handler_temp_Arr(Site) 'upload to global variant
                    Else
                        Handler_Temp_Site = 999
                        theexec.Datalog.WriteComment "Registry Read Back For Active site: " & site & " Is Not Numberic, Please Check"
                    End If
                Next
            Else
                theexec.Datalog.WriteComment ("Please Check Handler RegKey Dut_Temperature")
            End If
        Else
            theexec.Datalog.WriteComment "Handler Registry is empty"
            For Each site In theexec.sites.Active
                Handler_Temp_Site = 0
            Next
            'Handler_Temp = "00000"
        End If
        
        'gL_ProductionTemp = Handler_temp_Arr(site) 'upload to global variant
        
        If theexec.TesterMode = testModeOffline Then    ''' offline mode
            theexec.Datalog.WriteComment "Handler_Temp(Registry) : offline_mode_fake_temperature_value_bypass_check" ''offline_mode_bypass_check --> offline_mode_fake_temperature_value_bypass_check
            theexec.Flow.TestLimit 1, lowVal:=1, hiVal:=1, Tname:="Handler_Temp_test" & CStr(mid(theexec.DataManager.instancename, 16, 18))
        Else                                            ''' online mode
            If theexec.RunMode = runModeDebug Then      ''' debug mode
                theexec.Datalog.WriteComment "Handler_Temp(Registry) : engineering_mode_fake_temperature_value_bypass_check" ''debug_mode_bypass_check --> engineering_mode_fake_temperature_value_bypass_check
                theexec.Flow.TestLimit 1, lowVal:=1, hiVal:=1, Tname:="Handler_Temp_test" & CStr(mid(theexec.DataManager.instancename, 16, 18))
            Else                                        ''' production mode
                theexec.Datalog.WriteComment "Handler_Temp(Registry) : " & RegKeyRead("Dut_Temperature")
                theexec.Flow.TestLimit Handler_Temp_Site, lowVal:=Temp_Lolimit, hiVal:=Temp_Hilimit, Tname:="Handler_Temp_" & CStr(mid(theexec.DataManager.instancename, 16, 18))
            End If
        End If
    ElseIf UCase(theexec.CurrentChanMap) Like "CHANNELMAP_CP*" Or UCase(theexec.CurrentChanMap) Like "CHANNELMAP_WLFT*" Then
        theexec.Datalog.WriteComment ""
        'TheExec.Datalog.WriteComment "/********* Read Prober Temperature *********/"
        theexec.Datalog.WriteComment "********************************"
        theexec.Datalog.WriteComment "*print: Read Prober Temperature*"
        theexec.Datalog.WriteComment "********************************"
        
        If IsNumeric(RegKeyRead("Prober_Temp")) = True Then
            Prober_Temp_str = RegKeyRead("Prober_Temp")

        Else
            If RegKeyRead("Prober_Temp") = "" Then
                theexec.Datalog.WriteComment "Registry is empty"
                Prober_Temp_str = "00000"
            Else
                theexec.Datalog.WriteComment "Registry is not empty nor number."
                Prober_Temp_str = "99999"
            End If
        End If
        
        If mid(Prober_Temp_str, 1, 1) Like "+" Then
            Prober_Temp = CDbl(mid(Prober_Temp_str, 2, 5))
        Else
            Prober_Temp = CDbl(Prober_Temp_str)
        End If
    
        gL_ProductionTemp = Prober_Temp 'upload to global variant
    
        If theexec.TesterMode = testModeOffline Then    ''' offline mode
            theexec.Datalog.WriteComment "Prober_Temp(Registry) : offline_mode_fake_temperature_value_bypass_check" ''offline_mode_bypass_check --> offline_mode_fake_temperature_value_bypass_check
            theexec.Flow.TestLimit 1, lowVal:=1, hiVal:=1, Tname:="Prober_Temp_test" & CStr(mid(theexec.DataManager.instancename, 16, 18))
        Else                                            ''' online mode
            If theexec.RunMode = runModeDebug Then      ''' debug mode
                theexec.Datalog.WriteComment "Prober_Temp(Registry) : engineering_mode_fake_temperature_value_bypass_check" ''debug_mode_bypass_check --> engineering_mode_fake_temperature_value_bypass_check
                theexec.Flow.TestLimit 1, lowVal:=1, hiVal:=1, Tname:="Prober_Temp_test" & CStr(mid(theexec.DataManager.instancename, 16, 18))
            Else                                        ''' production mode
                theexec.Datalog.WriteComment "Prober_Temp(Registry) : " & RegKeyRead("Prober_Temp")
                theexec.Flow.TestLimit Prober_Temp, lowVal:=Temp_Lolimit, hiVal:=Temp_Hilimit, Tname:="Prober_Temp_" & CStr(mid(theexec.DataManager.instancename, 16, 18))
            End If
        End If
    Else
         theexec.Datalog.WriteComment theexec.DataManager.instancename & " Can not Judge the Channel Map Name, Please Check the Channel Map Naming "
    End If
       
    Exit Function
errHandler:
    If UCase(theexec.CurrentChanMap) Like "CHANNELMAP_FT*" Then
        theexec.Datalog.WriteComment "Read Handler Temp VBT function is error "
        theexec.Datalog.WriteComment "Registry String :" & RegKeyRead("Dut_Temperature")
        theexec.Datalog.WriteComment ("Error #: " & str(err.number) & " " & err.Description)
    Else
        theexec.Datalog.WriteComment "Read Prober Temp VBT function is error "
        theexec.Datalog.WriteComment "Registry String :" & RegKeyRead("Prober_Temp")
        theexec.Datalog.WriteComment ("Error #: " & str(err.number) & " " & err.Description)
    End If
        Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "ReadProberTemp") 'Add ErrHandler 2023/08/18
        If AbortTest Then Exit Function Else Resume Next
End Function



Public Function Set_PPMU_Clamp(Pin_GP1 As PinList, Pin_GP1_Vch As Double, _
                               Pin_GP2 As PinList, Pin_GP2_Vch As Double, _
                               Pin_GP3 As PinList, Pin_GP3_Vch As Double, _
                               Pin_GP4 As PinList, Pin_GP4_Vch As Double, _
                               Pin_GP5 As PinList, Pin_GP5_Vch As Double, _
                               Pin_GP6 As PinList, Pin_GP6_Vch As Double, _
                               Pin_GP7 As PinList, Pin_GP7_Vch As Double, _
                               Pin_GP8 As PinList, Pin_GP8_Vch As Double, _
                               Pin_GP9 As PinList, Pin_GP9_Vcl As Double, _
                               Pin_GP10 As PinList, Pin_GP10_Vcl As Double)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    If Pin_GP1.value <> "" Then
        TheHdw.PPMU.Pins(Pin_GP1).ClampVHi = Pin_GP1_Vch
    End If
    
    If Pin_GP2.value <> "" Then
        TheHdw.PPMU.Pins(Pin_GP2).ClampVHi = Pin_GP2_Vch
    End If
    
    If Pin_GP3.value <> "" Then
        TheHdw.PPMU.Pins(Pin_GP3).ClampVHi = Pin_GP3_Vch
    End If

    If Pin_GP4.value <> "" Then
        TheHdw.PPMU.Pins(Pin_GP4).ClampVHi = Pin_GP4_Vch
    End If
 
    If Pin_GP5.value <> "" Then
        TheHdw.PPMU.Pins(Pin_GP5).ClampVHi = Pin_GP5_Vch
    End If
 
    If Pin_GP6.value <> "" Then
        TheHdw.PPMU.Pins(Pin_GP6).ClampVHi = Pin_GP6_Vch
    End If
 
    If Pin_GP7.value <> "" Then
        TheHdw.PPMU.Pins(Pin_GP7).ClampVHi = Pin_GP7_Vch
    End If
 
    If Pin_GP8.value <> "" Then
        TheHdw.PPMU.Pins(Pin_GP8).ClampVHi = Pin_GP8_Vch
    End If
 
    If Pin_GP9.value <> "" Then
        TheHdw.PPMU.Pins(Pin_GP9).ClampVLo = Pin_GP9_Vcl
    End If
    
    If Pin_GP10.value <> "" Then
        TheHdw.PPMU.Pins(Pin_GP10).ClampVLo = Pin_GP10_Vcl
    End If

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Set_PPMU_Clamp") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' [20231108][All][Carter] Restore delete function
Public Function Disable_compare(DisableCompare_Pin As PinList)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    TheHdw.Digital.Pins(DisableCompare_Pin).DisableCompare = True
    
    theexec.Datalog.WriteComment "*************************************************"
    theexec.Datalog.WriteComment "*Disable Compare Pin:" & DisableCompare_Pin
    theexec.Datalog.WriteComment "*************************************************"
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Disable_compare") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function Enble_compare(EnableCompare_Pin As PinList)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    TheHdw.Digital.Pins(EnableCompare_Pin).DisableCompare = False

    theexec.Datalog.WriteComment "*************************************************"
    theexec.Datalog.WriteComment "*Enable Compare Pin:" & EnableCompare_Pin
    theexec.Datalog.WriteComment "*************************************************"
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Enble_compare") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
''20230223: Modidfied to let function can set FRC frequency in PLUS .
' [20230906][T-Spa][Jim] SBC Free Running Clock gating
Public Function FreeRunclk_Enable(PortName As String, Optional Freq As Double = 0, Optional Threshold_Range As String)


    On Error GoTo errHandler
    
    If theexec.DataManager.instancename <> "" Then
       If glb_ApplyLevelTiming_FRC_Flag = False Then
          TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
          glb_ApplyLevelTiming_FRC_Flag = True
       End If
    End If
    
    ' Add New argument to set nWire Frequency 221013
    If glb_TesterType = "Jaguar" Then
        If Freq <> 0 Then theexec.Datalog.WriteComment "This Argument is not support on UltraFlex"
    ElseIf glb_TesterType = "UltraFLEXplus" Then
        gl_nWireFreq = Freq
        If InStr(PortName, "_Port") <> 0 Then
            PortName = Replace(PortName, "_Port", "")
        End If
    End If
    
    Call Enable_FRC(PortName, False, Freq, Threshold_Range)
    TheHdw.Wait 0.001
    Exit Function
    
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "FreeRunclk_Enable")
    If AbortTest Then Exit Function Else Resume Next

End Function


''20230223: Modidfied to reduce find PowerPin_range_ary loop .
' [20230524][All][Si] add TTR add glb_PD_parsing_once flag let parsing pin seq one time.
' [20230731][All][Carter] exclude InitIO LO/HI/HZ in All digital pins
' [20230906][All][Carter] Modify fix process powerdown when Conti binout
' [20231003][All][Tank] modify after Chihome review
' [20231124][All][Tank] Fix exclude InitIO use not exist group
' [20240312][All][Clyde] TTR and modularize
Public Function PowerDown_Parallel(PowerPins As String, DisconnectPinList As String, Optional WaitConnectTime As Double = 0.001, Optional debugF As Boolean = False)
On Error GoTo errHandler
    functionName = "PowerDown_Parallel"
    
    Call Print_Header("Power down sequence")
    theexec.Datalog.WriteComment "print: Power down start, Power pins: " & PowerPins
    '------------------------------------------------------------------'set vt mode , then set digital pin to 0v , at last disconnect
    With TheHdw.Digital.Pins(DisconnectPinList)
        .Levels.DriverMode = tlDriverModeVt
        .Levels.value(chVt) = 0
        .Disconnect
    End With
    
    theexec.Datalog.WriteComment "print: Power down digital disconnect, Digital pins: " & DisconnectPinList
    theexec.Datalog.WriteComment RepeatChr("*", 120)
    '------------------------------------------------------------------
    
    If Not powerDownEnable Then
        Call PowerUpDownClockSeq(pwrDownSeq, False)
        Call PowerUpDownPinSeq(PowerPins, pwrDownSeq, False)
        powerDownEnable = True
    End If
    
    If debugF Then
        theexec.Datalog.WriteComment pwrDownSeq.PowerSeq99Log
        theexec.Datalog.WriteComment pwrDownSeq.PowerSeqNCLog
    End If
    
    PowerUpDown_Process WaitConnectTime, debugF, False
    
    '20230406 disconnect digital pins after power pin 0v and disconnect avoid hot switch
    TheHdw.Digital.Pins(DisconnectPinList).Disconnect
    Call Print_Footer("Power down sequence")
    
    Exit Function
    
errHandler:
    Call Print_Error_Message(Error_Info, moduleName, functionName)
    If AbortTest Then Exit Function Else Resume Next
End Function

''20230223: Modidfied to reduce find PowerPin_range_ary loop .
' [20230524][All][Si] add TTR add power_up_en flag let parsing pin seq one time.
' [20230906][All][Carter] Modify fix process powerdown when Conti binout
' [20231003][All][Tank] modify after Chihome review
' [20240312][All][Clyde] TTR and modularize
Public Function PowerUp_Parallel(PowerPins As String, ioPinGrp As String, DisconnectPinList As String, Optional WaitConnectTime As Double = 0.001, Optional debugF As Boolean = False)
On Error GoTo errHandler
    functionName = "PowerUp_Parallel"
    
    Dim PowerUp_WaitTime As Double
    Dim v_site As Variant
    Dim instSet As New InstrumentUtility
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    Call Print_Header("Power up sequence")
    theexec.Datalog.WriteComment "print: Power up start, Power pins: " & PowerPins
    instSet.Initialize PowerPins
    Call instSet.ApplyPower(0, fVoltage)
    '------------------------------------------------------------------disconnect I/O pins
    TheHdw.Digital.Pins(DisconnectPinList).Connect
    TheHdw.Digital.Pins(DisconnectPinList).initState = chInitoff
    
    theexec.Datalog.WriteComment "print: Power up digital disconnect, Digital pins: " & DisconnectPinList
    theexec.Datalog.WriteComment RepeatChr("*", 120)
    '------------------------------------------------------------------
    
    If Not powerUpEnable Then
        Call PowerUpDownClockSeq(pwrUpSeq, True)
        Call PowerUpDownIOSeq(ioPinGrp)
        Call PowerUpDownPinSeq(PowerPins, pwrUpSeq, True)
        powerUpEnable = True
    End If
    
    If debugF Then
        theexec.Datalog.WriteComment pwrUpSeq.PowerSeq99Log
        theexec.Datalog.WriteComment pwrUpSeq.PowerSeqNCLog
    End If
    
    PowerUpDown_Process WaitConnectTime, debugF, True
    TheHdw.Digital.Pins(DisconnectPinList).Disconnect
    Call Print_Footer("Power up sequence")
    
    ' Record which site has do power up
    For Each v_site In theexec.sites
        powerUpDone(v_site) = True
    Next v_site
    
    If gl_EnableCurrentProfile Or gl_EnableVoltageProfile Then
        DoAll_save = theexec.RunOptions.DoAll
        OverRide_FailStop = theexec.RunOptions.OverrideFailStop
        theexec.RunOptions.DoAll = True
        theexec.RunOptions.OverrideFailStop = True
        theexec.Datalog.WriteComment "Enable Word:CurrentProfile = Enable"
        theexec.Datalog.WriteComment "DoAll = Enable"
        theexec.Datalog.WriteComment "OverrideFailStop = Enable"
    End If
    
    'Print Error; Message
    If isDebugMode Then TheHdw.Alarms.DumpState

    'Wait for timeout alarm before testlimit
    PowerUp_WaitTime = TheHdw.DCVS.Pins(instSet.GetDCVSPins).CurrentLimit.Source.FoldLimit.TimeOut
    TheHdw.Wait PowerUp_WaitTime
    TheHdw.Alarms.Check
    Exit Function
    
errHandler:
    Call Print_Error_Message(Error_Info, moduleName, functionName)
    If AbortTest Then Exit Function Else Resume Next
End Function

' [20230407][T-Ibi] Add soure and sink alarm set
' [20230524][All][Si] add TTR no need to set alarm state pin by pin.
' [20230906][All][Tank] Modify that different stage set different pin alarm behavior(ex: CP:VDD_A,VDD_B,VDD_C;FT:VDD_D,VDD_E)
' [20230907][T-Don][Carter] Source/Sink Timeout Alarm behavior gate off setting
Public Function Set_Power_Alarm(Pin_DCVS As String, SourceFoldAlarmTime_DCVS As Double, SinkFoldLevelRatio_DCVS As Double, SinkFoldAlarmTime_DCVS As Double, Pin_DCVI As String, AlarmTime_DCVI As Double, Pin_DCVS_GP2 As String, SourceFlodAlarmTime_GP2 As Double, SinkFoldLevelRatio_GP2 As Double, SinkFoldAlarmTime_GP2 As Double, Optional b_AlarmForceBin As Boolean = False)        'add boolean to set force_bin or force_fail

On Error GoTo errHandler
    
''    Pin_DCVS.value = "All_Power"
''    Pin_DCVI.value = "ALL_DCVI"
    
    Dim i As Long
    Dim n_DCVSPinCnt As Long
    Dim n_DCVIPinCnt As Long
    Dim n_PinGP2Cnt As Long
    Dim sa_DCVSPinAry() As String
    Dim sa_DCVIPinAry() As String
    Dim sa_PinGP2Ary() As String
    Dim s_DCVSPin As String
    Dim s_DCVIPin As String
    Dim s_DCVS_GP2 As String
    
    If SourceFoldAlarmTime_DCVS = 0 Then SourceFoldAlarmTime_DCVS = 0.05
    If AlarmTime_DCVI = 0 Then AlarmTime_DCVI = 0.05
    If SourceFlodAlarmTime_GP2 = 0 Then SourceFlodAlarmTime_GP2 = 0.05
    If SinkFoldLevelRatio_DCVS = 0 Then SinkFoldLevelRatio_DCVS = 1
    If SinkFoldAlarmTime_DCVS = 0 Then SinkFoldAlarmTime_DCVS = 1
    If SinkFoldLevelRatio_DCVS > 1 Then SinkFoldLevelRatio_DCVS = 1
    If SinkFoldLevelRatio_GP2 > 1 Then SinkFoldLevelRatio_GP2 = 1
    
    If Pin_DCVS = "" And Pin_DCVI = "" And Pin_DCVS_GP2 = "" Then
        theexec.Datalog.WriteComment "*************************************************************************"
        theexec.Datalog.WriteComment "Error on Set_Power_pins_alarm, please fill in pin name to set alarm time."
        theexec.Datalog.WriteComment "*************************************************************************"
    End If
    
    If Pin_DCVS <> "" Then
        If (UCase(Pin_DCVS) Like "*CP*=*" Or UCase(Pin_DCVS) Like "*FT*=*") Then
            s_DCVSPin = Select_MeasPin(Pin_DCVS, UCase(currentJobName))
        Else
            s_DCVSPin = Pin_DCVS
        End If
        theexec.DataManager.DecomposePinList s_DCVSPin, sa_DCVSPinAry(), n_DCVSPinCnt
    End If
    If Pin_DCVI <> "" Then
        If (UCase(Pin_DCVI) Like "*CP*=*" Or UCase(Pin_DCVI) Like "*FT*=*") Then
            s_DCVIPin = Select_MeasPin(Pin_DCVI, UCase(currentJobName))
        Else
            s_DCVIPin = Pin_DCVI
        End If
        theexec.DataManager.DecomposePinList s_DCVIPin, sa_DCVIPinAry(), n_DCVIPinCnt
    End If
    If Pin_DCVS_GP2 <> "" Then
        If (UCase(Pin_DCVS_GP2) Like "*CP*=*" Or UCase(Pin_DCVS_GP2) Like "*FT*=*") Then
            s_DCVS_GP2 = Select_MeasPin(Pin_DCVS_GP2, UCase(currentJobName))
        Else
            s_DCVS_GP2 = Pin_DCVS_GP2
        End If
        theexec.DataManager.DecomposePinList s_DCVS_GP2, sa_PinGP2Ary(), n_PinGP2Cnt
    End If
    If theexec.Flow.enableWord("CurrentProfile") Or theexec.Flow.enableWord("VoltageProfile") Or Profile_byflow = True Then
        '==================================================================
        If n_DCVSPinCnt > 0 Then
            For i = 0 To n_DCVSPinCnt - 1
                With TheHdw.DCVS.Pins(sa_DCVSPinAry(i))
                    .Gate = False
                    TheHdw.Wait 1 * ms '20170209 Add to wait gate off to avoid set ifold time out error
                    .mode = tlDCVSModeVoltage
                    .Voltage.Main = 0
                    .Alarm(tlDCVSAlarmOverrange) = tlAlarmOff
                    '.CurrentLimit.Sink.FoldLimit.TimeOut = AlarmTime_DCVS
                    ' 20220113  modify sink limit for negative current
                    .CurrentLimit.Sink.FoldLimit.TimeOut = SinkFoldAlarmTime_DCVS
                    ' 20220526  modify sink limit for negative current level
                    .CurrentLimit.Sink.FoldLimit.level = SinkFoldLevelRatio_DCVS * .CurrentLimit.Sink.FoldLimit.level.max
                    .CurrentLimit.Source.FoldLimit.TimeOut = SourceFoldAlarmTime_DCVS
                    ' 20230602 Source/Sink set gate off at SetPowerAlarm
                    .CurrentLimit.Sink.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorGateOff
                    .CurrentLimit.Source.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorGateOff
                    If glb_TesterType = "UltraFLEXplus" Then
                        .Alarm(tlDCVSAlarmSinkFoldcurrentLimitLevel) = tlAlarmOff
                        .Alarm(tlDCVSAlarmSourceFoldCurrentLimitLevel) = tlAlarmOff
                    End If

                End With
            Next i
        End If
        '==================================================================
        theexec.enableWord("HardIP_Autorange") = False
    Else
    
        '==================================================================
        If n_DCVSPinCnt > 0 Then
            With TheHdw.DCVS.Pins(s_DCVSPin)
                .Gate = False
                TheHdw.Wait 1 * ms '20170209 Add to wait gate off to avoid set ifold time out error
                .mode = tlDCVSModeVoltage
                .Voltage.Main = 0
                
                If b_AlarmForceBin = False Then
                    .Alarm(tlDCVSAlarmAll) = tlAlarmForceFail
                Else
                    .Alarm(tlDCVSAlarmAll) = tlAlarmForceBin
                End If
                If CStr(SourceFoldAlarmTime_DCVS) = "" Or SourceFoldAlarmTime_DCVS < 0.0001 Then
                    theexec.Datalog.WriteComment "************************************************************"
                    theexec.Datalog.WriteComment "Error on Set_DCVS_alarm, please put a reasonable alarm time."
                    theexec.Datalog.WriteComment "************************************************************"
                End If
                '.CurrentLimit.Sink.FoldLimit.TimeOut = AlarmTime_DCVS
                ' 20220113  modify sink limit for negative current
                .CurrentLimit.Sink.FoldLimit.TimeOut = SinkFoldAlarmTime_DCVS
                .CurrentLimit.Source.FoldLimit.TimeOut = SourceFoldAlarmTime_DCVS
                ' 20230602 Source/Sink set gate off at SetPowerAlarm
                .CurrentLimit.Sink.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorGateOff
                .CurrentLimit.Source.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorGateOff
                If glb_TesterType = "UltraFLEXplus" Then
                    .Alarm(tlDCVSAlarmSinkFoldcurrentLimitLevel) = tlAlarmOff
                    .Alarm(tlDCVSAlarmSourceFoldCurrentLimitLevel) = tlAlarmOff
                End If
            End With
            ' 20220526  modify sink limit for negative current level
            For i = 0 To n_DCVSPinCnt - 1
                TheHdw.DCVS.Pins(sa_DCVSPinAry(i)).CurrentLimit.Sink.FoldLimit.level = SinkFoldLevelRatio_DCVS * TheHdw.DCVS.Pins(sa_DCVSPinAry(i)).CurrentLimit.Sink.FoldLimit.level.max
            Next i
        End If
        '==================================================================
        If n_DCVIPinCnt > 0 Then
            With TheHdw.DCVI.Pins(s_DCVIPin)
                .Gate = False
                .mode = tlDCVIModeVoltage
                .Voltage = 0
    '''            .Alarm(tlDCVIAlarmAll) = tlAlarmForceFail
                If b_AlarmForceBin = False Then
                    .Alarm(tlDCVIAlarmAll) = tlAlarmForceFail
                Else
                    .Alarm(tlDCVIAlarmAll) = tlAlarmForceBin
                End If
    '            .FoldCurrentLimit.TimeOut = AlarmTime             'UVI80 spec 100us ~ 100ms
                .FoldCurrentLimit.TimeOut = AlarmTime_DCVI         '0.05
                .FoldCurrentLimit.Behavior = tlDCVIFoldCurrentLimitBehaviorGateOff
            End With
        End If
        '==================================================================
        If n_PinGP2Cnt > 0 Then
            With TheHdw.DCVS.Pins(s_DCVS_GP2)
                .Gate = False
                TheHdw.Wait 1 * ms '20170209 Add to wait gate off to avoid set ifold time out error
                .mode = tlDCVSModeVoltage
                .Voltage.Main = 0
                
                If b_AlarmForceBin = False Then
                    .Alarm(tlDCVSAlarmAll) = tlAlarmForceFail
                Else
                    .Alarm(tlDCVSAlarmAll) = tlAlarmForceBin
                End If
                If CStr(SourceFlodAlarmTime_GP2) = "" Or SourceFlodAlarmTime_GP2 < 0.0001 Then
                    theexec.Datalog.WriteComment "*******************************************************************"
                    theexec.Datalog.WriteComment "Error on Set_DCVS_alarm, please put a reasonable alarm time on GP2."
                    theexec.Datalog.WriteComment "*******************************************************************"
                End If
                '.CurrentLimit.Sink.FoldLimit.TimeOut = AlarmTime
                ' 20220113  modify sink limit for negative current
                .CurrentLimit.Sink.FoldLimit.TimeOut = SinkFoldAlarmTime_GP2
                .CurrentLimit.Source.FoldLimit.TimeOut = SourceFlodAlarmTime_GP2
                ' 20230602 Source/Sink set gate off at SetPowerAlarm
                .CurrentLimit.Sink.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorGateOff
                .CurrentLimit.Source.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorGateOff
                If glb_TesterType = "UltraFLEXplus" Then
                    .Alarm(tlDCVSAlarmSinkFoldcurrentLimitLevel) = tlAlarmOff
                    .Alarm(tlDCVSAlarmSourceFoldCurrentLimitLevel) = tlAlarmOff
                End If
            End With
            ' 20220526  modify sink limit for negative current level
            For i = 0 To n_PinGP2Cnt - 1
                TheHdw.DCVS.Pins(sa_PinGP2Ary(i)).CurrentLimit.Sink.FoldLimit.level = SinkFoldLevelRatio_GP2 * TheHdw.DCVS.Pins(sa_PinGP2Ary(i)).CurrentLimit.Sink.FoldLimit.level.max
            Next i
        End If
        
    End If
    
    ' 20210112 temporary add for open Kelvin Alarm bypass for open socket check
    If theexec.Flow.enableWord("OpenKelvinAlarmOff") = True Then
        TheHdw.DCVS.Pins("All_Power,SPI_PWR").Alarm(tlDCVSAlarmAll) = tlAlarmOff
    End If

    Exit Function
 
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Set_Power_Alarm")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Search_UnExistPin() As Long
    Dim Pinmap_Sheet, Group_Name, All_Groups, Tset_Pins, All_TsetPins, All_PinGroup_Pins As String
    Dim cnt, i, j, k, colcnt As Long
    Dim Activate_Sheet As Worksheet
    Dim Export_sheet As Worksheet
    Dim Tset_Sheets() As String
    Dim All_Groups_arr() As String
    Dim All_PinGroup_Pins_arr() As String
    Dim All_TsetPins_arr() As String
    Dim PinAry() As String
    Dim PinCnt As Long
    Dim Not_exist_pins As String
    Dim Not_exist_pins_arr() As String
    
    On Error Resume Next
    
    Not_exist_pins = vbNullString
    All_Groups = vbNullString
    cnt = 0
    Application.ScreenUpdating = False
    
    Set Export_sheet = ThisWorkbook.Sheets("UnExistPin")
    If Export_sheet Is Nothing Then
        Sheets.Add after:=Sheets(Sheets.Count)
        Sheets(Sheets.Count).name = "UnExistPin"
        Set Export_sheet = ThisWorkbook.Sheets("UnExistPin")
    Else
        Application.DisplayAlerts = False
        Sheets("UnExistPin").delete
        Sheets.Add after:=Sheets(Sheets.Count)
        Sheets(Sheets.Count).name = "UnExistPin"
        Application.DisplayAlerts = True
        Set Export_sheet = ThisWorkbook.Sheets("UnExistPin")
    End If
    
    'Active sheets
    For Each Activate_Sheet In ThisWorkbook.Sheets
        If Activate_Sheet.Visible = xlSheetVisible Then
            If LCase(Activate_Sheet.name) Like "*timeset*" Or LCase(Activate_Sheet.name) Like "*tsb*" Then
                Activate_Sheet.Activate
                ReDim Preserve Tset_Sheets(cnt)
                Tset_Sheets(cnt) = Activate_Sheet.name
                cnt = cnt + 1
            ElseIf LCase(Activate_Sheet.name) Like "*pinmap*" Then
                Activate_Sheet.Activate
                Pinmap_Sheet = Activate_Sheet.name
            End If
        End If
    Next Activate_Sheet
    
    
    ' Parsing all pin groups in pinmap
    cnt = Worksheets(Pinmap_Sheet).Cells(4, 3).End(xlDown).Row 'row count in pinmap
    For i = 4 To cnt
        Group_Name = Worksheets(Pinmap_Sheet).Cells(i, 2)
        If Group_Name <> "" Then
            If LCase(Group_Name) Like "pins_*" Or LCase(Group_Name) Like "*_pa*" Then
                If InStr(All_Groups, Group_Name) <> 0 Then
                Else
                    If All_Groups = "" Then
                        All_Groups = Group_Name
                    Else
                        All_Groups = All_Groups & "," & Group_Name
                    End If
                End If
            End If
        End If
    Next i
    
    All_Groups_arr = Split(All_Groups, ",")
    ' Parsing all pins in pin group
    For i = 0 To UBound(All_Groups_arr)
        theexec.DataManager.DecomposePinList All_Groups_arr(i), PinAry(), PinCnt
        For j = 0 To PinCnt - 1
            If InStr(All_PinGroup_Pins, PinAry(j) & ",") = 0 Then
                If All_PinGroup_Pins = "" Then
                    All_PinGroup_Pins = PinAry(j) & ","
                Else
                    All_PinGroup_Pins = All_PinGroup_Pins & PinAry(j) & ","
                End If
            End If
        Next j
    Next i
    
    ''''TheExec.AddOutput "*********Search Start*********"
    ''''TheExec.AddOutput " "
    ''''TheExec.AddOutput "Search pin Groups:" & All_Groups
    ''''TheExec.AddOutput "Search Timset Sheets:" & Join(Tset_Sheets, ",")
    ''''TheExec.AddOutput " "
    
    For i = 1 To UBound(All_Groups_arr) + 2
        If i = 1 Then
            Export_sheet.Cells(i, 1) = "Search pin Groups"
        Else
            Export_sheet.Cells(i, 1) = All_Groups_arr(i - 2)
        End If
    Next i
    For i = 1 To UBound(Tset_Sheets) + 2
        If i = 1 Then
            Export_sheet.Cells(i, 2) = "Search TimSet Sheets"
        Else
            Export_sheet.Cells(i, 2) = Tset_Sheets(i - 2)
        End If
    Next i
    
    colcnt = 3
    For i = 0 To UBound(Tset_Sheets)
        All_TsetPins = vbNullString
        cnt = Worksheets(Tset_Sheets(i)).Cells(7, 4).End(xlDown).Row 'row count in TimeSet sheet
        For j = 8 To cnt
            Tset_Pins = Worksheets(Tset_Sheets(i)).Cells(j, 4)
            If InStr(All_TsetPins, Tset_Pins & ",") = 0 Then
                If All_TsetPins = "" Then
                    All_TsetPins = Tset_Pins & ","
                Else
                    All_TsetPins = All_TsetPins & Tset_Pins & ","
                End If
            End If
        Next j
        
        All_TsetPins_arr = Split(All_TsetPins, ",")
        'Search start
        For j = 0 To UBound(All_TsetPins_arr)
            If InStr(All_PinGroup_Pins, All_TsetPins_arr(j) & ",") = 0 Then
                If Not_exist_pins = "" Then
                    Not_exist_pins = All_TsetPins_arr(j) & ","
                Else
                    Not_exist_pins = Not_exist_pins & All_TsetPins_arr(j) & ","
                End If
            End If
        Next j
        If Not_exist_pins <> "" Then
            Not_exist_pins_arr = Split(Not_exist_pins, ",")
            Not_exist_pins = mid(Not_exist_pins, 1, Len(Not_exist_pins) - 1)
    '''        TheExec.AddOutput " "
    '''        TheExec.AddOutput Tset_Sheets(i) & " not exist pins in pinmap:"
    '''        TheExec.AddOutput Not_exist_pins
    '''        TheExec.AddOutput "----------------------------------------------"
    '''        TheExec.AddOutput " "
            
            Export_sheet.Cells(1, colcnt) = "Not exist pins in: " & Tset_Sheets(i)
            For k = 0 To UBound(Not_exist_pins_arr)
                Export_sheet.Cells(k + 2, colcnt) = Not_exist_pins_arr(k)
            Next k
            colcnt = colcnt + 1
            Not_exist_pins = vbNullString
        End If
    Next i
    
    Application.ScreenUpdating = True
    If isDebugMode Then theexec.AddOutput "*********ExportUnExistPins Search End*********"
    
End Function


Public Function auto_ReadHandlerOCRData() 'VBT function
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim funcName As String:: funcName = "auto_ReadHandlerOCRData"

    'SWLINZA 20170406 for OCR reading
    Dim site As Variant
    Dim i As Integer
    Dim SiteCount As Integer
    Dim OCR_FullString_32Sites As String
    Dim OCR_StringBySite() As String
    Dim OCR_StringBySite_Sort() As String
    Dim OCR_available_chk As New SiteDouble
    
    '------------------------------------------------------------------------------
    '--- define string name as '32sites' is because the OCR maximun is 32sites ----'
    '------------------------------------------------------------------------------
    OCR_FullString_32Sites = vbNullString
    OCR_FullString_32Sites = RegKeyRead("HandlerBarCodeString")
    OCR_StringBySite = Split(OCR_FullString_32Sites, ",")
    
    'To judge how manys sites automatically
    SiteCount = UBound(OCR_StringBySite)
    
    If SiteCount >= 0 Then
        ReDim OCR_StringBySite_Sort(SiteCount) As String
        For i = 0 To SiteCount
            OCR_StringBySite_Sort(SiteCount - i) = OCR_StringBySite(i)
        Next i
    Else
        SiteCount = theexec.sites.Existing.Count
        ReDim OCR_StringBySite_Sort(SiteCount - 1) As String
        For Each site In theexec.sites
            OCR_StringBySite_Sort(site) = vbNullString
        Next site
    End If
    
    theexec.Datalog.WriteComment ""
    
    'To write OCR data in STDF file
    For Each site In theexec.sites
        theexec.Datalog.WriteComment ("<@OCR_Data=" & site & "|" & OCR_StringBySite_Sort(site) & ">")
    Next site
    
    'To check OCR is available, string or something else but not empty
    For Each site In theexec.sites
        OCR_available_chk = 0
        If (theexec.TesterMode = testModeOffline) Then
            OCR_available_chk = 1
        Else
            If IsEmpty(OCR_StringBySite_Sort(site)) Or OCR_StringBySite_Sort(site) = "" Then
                OCR_available_chk = 0
            Else
                OCR_available_chk = 1 '1 means OCR is not ""
            End If
        End If
        theexec.Flow.TestLimit OCR_available_chk, 1, 1, Tname:="OCR_Checking"
    Next site
    
    
    theexec.Datalog.WriteComment ""

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "auto_ReadHandlerOCRData") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

''------------------------------------------------------------
''20210504
''move the function from Exec IP module.
''Data log format default setting.
''PartResult                : Part result data consists of bin and sort numbers for each site.
''XYCoordinates             : gets or sets whether to include X/Y coordinates in part results.
''DisableChannelNumberInPTR : use a channel number in a Parametric Test Record
''OutputWidth               : gets or sets the maximum number of characters to be displayed per line
''DisableInstanceNameInPTR  : use of a test instance name in a Parametric Test Record (PTR) in datalogging
''DisablePinNameInPTR       : use of a pin name in a Parametric Test Record (PTR) in datalogging
''PTR_InstanceNameIsTINameOnly :  gets or sets which test name to use in a Parametric Test Record (PTR) test text
''------------------------------------------------------------
Public Function Init_Datalog_Setup()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim funcName As String:: funcName = "Init_Datalog_Setup"
    
    With theexec.Datalog
        .WriteComment "----------------------------------------"
        .WriteComment "Reset to default datalog format!!!!!!!!!"
        .WriteComment "----------------------------------------"
    End With
    
    With theexec.Datalog.Setup.DatalogSetup
        .PartResult = True                      'Default: TRUE
        .XYCoordinates = True                   'Default: TRUE
        .DisableChannelNumberInPTR = True       'Default: FALSE 'disable channel name to stdf, PE's datalog request -- 131225, chihome
        .OutputWidth = 0                        'Default: 0
        .DisableInstanceNameInPTR = False       'Default: FALSE '20210519 add
        .DisablePinNameInPTR = False            'Default: FALSE '20210519 add
        .PTR_InstanceNameIsTINameOnly = True    'Default: FALSE '20210519 add   '[20240207][All][Brian] conment for fix and optimize flag

    End With
    
    With theexec.Datalog.Setup.Shared.ascii.Columns
        .EnableCustomWidths = True
        .Parametric.TestName.Width = 200
        .Parametric.Measured.Width = 16
        .Parametric.number.Width = 12
        .Parametric.pin.Width = 30
        .Functional.TestName.Width = 200
        .Functional.Pattern.Width = 100
        .Functional.number.Width = 12
        .Functional.Cycle.Width = 9
        
    End With
    
    theexec.Datalog.ApplySetup  'must need to apply after datalog setup

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Init_Datalog_Setup") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


''------------------------------------------------------------
''20210504
''Customize datalog format
''------------------------------------------------------------
' [20231108][All][Carter] Restore delete function
Public Function Customize_Datalog_Setup(BlockName As String, _
                                        Optional P_TestNameWidth As Long = 200, _
                                        Optional P_MeasuredWidth As Long = 16, _
                                        Optional P_NumberWidth As Long = 12, _
                                        Optional P_PinWidth As Long = 30, _
                                        Optional F_TestNameWidth As Long = 200, _
                                        Optional F_PatternWidth As Long = 100, _
                                        Optional F_NumberWidth As Long = 12, _
                                        Optional F_CycleWidth As Long = 9) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim funcName As String:: funcName = "Custom_Datalog_Setup"
    
    With theexec.Datalog
        .WriteComment "---------------------------------------------------"
        .WriteComment "Set a customized datalog format : " & BlockName
        .WriteComment "---------------------------------------------------"
    End With
    
    If P_TestNameWidth = 0 Then P_TestNameWidth = 200
    If P_MeasuredWidth = 0 Then P_MeasuredWidth = 16
    If P_NumberWidth = 0 Then P_NumberWidth = 12
    If P_PinWidth = 0 Then P_PinWidth = 30
    If F_TestNameWidth = 0 Then F_TestNameWidth = 200
    If F_PatternWidth = 0 Then F_PatternWidth = 100
    If F_NumberWidth = 0 Then F_NumberWidth = 12
    If F_CycleWidth = 0 Then F_CycleWidth = 9

    With theexec.Datalog.Setup.DatalogSetup
        .PartResult = True
        .XYCoordinates = True
        .DisableChannelNumberInPTR = True 'disable channel name to stdf, PE's datalog request -- 131225, chihome
        .OutputWidth = 0
        .DisableInstanceNameInPTR = False
        .DisablePinNameInPTR = False
        '.PTR_InstanceNameIsTINameOnly = False          '[20240207][All][Brian] conment for fix and optimize flag

    End With
    With theexec.Datalog.Setup.Shared.ascii.Columns
        .EnableCustomWidths = True
        .Parametric.TestName.Width = P_TestNameWidth
        .Parametric.Measured.Width = P_MeasuredWidth
        .Parametric.number.Width = P_NumberWidth
        .Parametric.pin.Width = P_PinWidth
        .Functional.TestName.Width = F_TestNameWidth
        .Functional.Pattern.Width = F_PatternWidth
        .Functional.number.Width = F_NumberWidth
        .Functional.Cycle.Width = F_CycleWidth
    End With
    
    theexec.Datalog.ApplySetup  'must need to apply after datalog setup
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Customize_Datalog_Setup") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Print_Version()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    theexec.Datalog.WriteComment "-----------------------------------------------"
    theexec.Datalog.WriteComment "      DATE : " & CStr(DATE_VER) & "   SVN_VER : " & CStr(SVN_VER) & "     "
    theexec.Datalog.WriteComment "-----------------------------------------------"

    theexec.Datalog.WriteComment "CZ            : " & CStr(CZ_VER)
    theexec.Datalog.WriteComment "COMMON        : " & CStr(COMMON_VER)
    theexec.Datalog.WriteComment "DC            : " & CStr(DC_VER)
    theexec.Datalog.WriteComment "DIGITAL       : " & CStr(DIGITAL_VER)
    theexec.Datalog.WriteComment "EFUSE         : " & CStr(EFUSE_VER)
    theexec.Datalog.WriteComment "HIP           : " & CStr(HIP_VER)
    theexec.Datalog.WriteComment "MBIST         : " & CStr(MBIST_VER)
    theexec.Datalog.WriteComment "SPIROM        : " & CStr(SPIROM_VER)
    theexec.Datalog.WriteComment "VDDBINNING    : " & CStr(VDDBINNING_VER)
    theexec.Datalog.WriteComment "RF_FUNC       : " & CStr(RF_FUNC_VER)
    theexec.Datalog.WriteComment "RF_DVDC       : " & CStr(RF_DVDC_VER)
    theexec.Datalog.WriteComment "LCD_OTP       : " & CStr(LCD_OTP_VER)
    theexec.Datalog.WriteComment "LCD_DVDC      : " & CStr(LCD_DVDC_VER)
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Print_Version") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' [20230620][T-All] add for check FRC frequency
Public Function show_FRC() As Long
    On Error GoTo errHandler
    Dim key As Variant
    For Each key In gl_nWireFreq_Value_Dict.Keys
        theexec.Datalog.WriteComment "********** freerunning clock pin = " & key & ", clock =" & gl_nWireFreq_Value_Dict(key) & " Hz "
    Next
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "show_FRC")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function FixedTuneCWRT(TunePins As String, InterPosePrePat As String, ApplyPinsVoltage As Double, TunePinVoltages As String, CurrentRanges As String, Optional Pattern As String = vbNullString, _
                                    Optional SQ_M As String = vbNullString, Optional AccuracyRatio As Double = 1, Optional NoRunBelowCurrentRange As Double = 0, Optional SmoothCnt As Integer = 10, _
                                    Optional WriteCurrentProfileFile As Boolean = False, Optional LoadTuneResult As Boolean = False)
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "FixedTuneCWRT"
    Dim CRWT As New Class_CRWT
    Dim sq_cnt As Integer
    Dim sq_array() As String
    Dim i As Integer
    
    sq_array = Split(SQ_M, ",")
    
    If Pattern <> "" Then
        TheHdw.Patterns(Pattern).Load
        TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD
        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    End If
    
    Call CRWT.Initialize(TunePins)
    If LoadTuneResult Then Call CRWT.LoadCRWTResultToDict
    If Pattern <> "" Then
        Call TheHdw.Patterns(Pattern).start
        Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
    
        For i = 0 To UBound(sq_array)
            If sq_array(i) <> "I" Then
                TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD
                TheHdw.Wait 0.01
            Else
                Call CRWT.FixedVCR_WaitTime(TunePinVoltages, CurrentRanges, AccuracyRatio, NoRunBelowCurrentRange, SmoothCnt, WriteCurrentProfileFile)
                TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD
            End If
        Next i
        TheHdw.Digital.Patgen.HaltWait
        Call HardIP_WriteFuncResult
    Else
        Call CRWT.FixedVCR_WaitTime(TunePinVoltages, CurrentRanges, AccuracyRatio, NoRunBelowCurrentRange, SmoothCnt, WriteCurrentProfileFile)
    End If
    
    Call CRWT.TuneClose
    Exit Function
errHandler:
    theexec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function LoopTuneCRWT(TunePins As String, InterPosePrePat As String, ApplyPinsVoltage As Double, StartVoltage As Double, EndVoltage As Double, StepVoltage As Double, _
                                Optional Pattern As String = vbNullString, Optional SQ_M As String, Optional NoRunBelowCurrentRange As Double = 0, Optional AccuracyRatio As Double = 1, _
                                Optional SmoothCnt As Integer = 10, Optional WriteCurrentProfileFile As Boolean = False, Optional LoadTuneResult As Boolean = False)
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "LoopTuneCRWT"
    Dim CRWT As New Class_CRWT
    Dim sq_cnt As Integer
    Dim sq_array() As String
    Dim i As Integer
    
    sq_array = Split(SQ_M, ",")
    
    If Pattern <> "" Then
        TheHdw.Patterns(Pattern).Load
        TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD
        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    End If
    
    Call CRWT.Initialize(TunePins)
    If LoadTuneResult Then Call CRWT.LoadCRWTResultToDict
    If Pattern <> "" Then
        Call TheHdw.Patterns(Pattern).start
        Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
    
        For i = 0 To UBound(sq_array)
            If sq_array(i) <> "I" Then
                TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD
                TheHdw.Wait 0.01
            Else
                Call CRWT.LoopVCR_WaitTime(StartVoltage, EndVoltage, StepVoltage, AccuracyRatio, NoRunBelowCurrentRange, SmoothCnt, WriteCurrentProfileFile)
                TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD
            End If
        Next i
        TheHdw.Digital.Patgen.HaltWait
        Call HardIP_WriteFuncResult
    Else
        Call CRWT.LoopVCR_WaitTime(StartVoltage, EndVoltage, StepVoltage, AccuracyRatio, NoRunBelowCurrentRange, SmoothCnt, WriteCurrentProfileFile)
    End If
    Call CRWT.TuneClose
    Exit Function
errHandler:
    theexec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function onProgramStartedBinOutFunction(Optional HarvestInstName As String = vbNullString)
    Call HardIP_RAK_Init(True)
    Call ParseIDSMappingTable(True)
        Call CharSetUpCheck
    If HarvestInstName = vbNullString Then
        theexec.Datalog.WriteComment "Check TestInst_Harvest Sheet"
        HarvestInstName = "TestInst_Harvest"
    Else
        theexec.Datalog.WriteComment "Check " & HarvestInstName & " Sheet"
    End If
    Call Check_HarvestEfuse_Arg(HarvestInstName)
End Function

