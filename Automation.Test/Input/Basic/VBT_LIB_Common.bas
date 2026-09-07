Attribute VB_Name = "VBT_LIB_Common"
Option Explicit
'Revision History:
'V0.0 initial bring up
'V0.1 add keep alive function
'V0.2 add disable compare and enable compare function.
'variable declaration
Public BarcodeString As String
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
'''220125 Add for skipping BC alarm items
Public Bringup_ByPass_Dic As New Scripting.Dictionary
''' Add for TMPS Adaptive coolling @William 230202
Public glb_MTRRecord As String
Public glb_MTRBTSCnt As String

Public glb_Power_down_Flag As Boolean           ''Add for fix Tunm shift when shut down site  @ Brian 20231221
Private Const moduleName As String = "VBT_LIB_Common"
Private functionName As String


Public Function Relay_Control(Optional Relay_on As PinList, Optional relay_off As PinList, Optional WaitTime As Double = 0.003)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
'control relay on off, will auto trim NC pins Tto cover CP, FT both stages
    Dim Pins_On() As String, Pin_Cnt_On As Long
    Dim Pins_Off() As String, Pin_Cnt_Off As Long
    Dim p As Variant
    Dim relayOnStr As String, relayOffStr As String

    
    relayOnStr = vbNullString
    relayOffStr = vbNullString

    TheExec.DataManager.DecomposePinList Relay_on, Pins_On(), Pin_Cnt_On
    TheExec.DataManager.DecomposePinList relay_off, Pins_Off(), Pin_Cnt_Off

    Trim_NC_Pin Pins_On, Pin_Cnt_On
    Trim_NC_Pin Pins_Off, Pin_Cnt_Off

    If Pin_Cnt_On <> 0 Then
        TheHdw.Utility.Pins(Relay_on).State = tlUtilBitOn
        For Each p In Pins_On
            If relayOnStr = "" Then
                relayOnStr = relayOnStr & p
            Else
                relayOnStr = relayOnStr & ", " & p
            End If
        Next p
        TheExec.Datalog.WriteComment "Relay On : " & relayOnStr
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
        TheExec.Datalog.WriteComment "Relay off : " & relayOffStr
    End If

    Wait WaitTime
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Relay_Control") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function

'*****************************************
'******         free run clk, nWire ******
'*****************************************


Public Function StartSBClock(SBFreq As Double) As Long
    On Error GoTo errHandler
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
    TheExec.Flow.TestLimit SBC_Enable, 1, 1, tlSignGreaterEqual, tlSignLessEqual, Tname:="SBC enable" 'BurstResult=1:Pass
    'printing in data log
    TheExec.Datalog.WriteComment "********** support board clock = " & Format(SBFreq / 1000000, "0.000") & " Mhz, Clock_Vih = " _
                             & XI0_ref_VOH & " V, Clock_Vil = " & XI0_ref_VOL & " V  *******"
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "StartSBClock") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function StopSBClock()
    Dim SBC_Enable As Long
    ' Stop and disconnect the support board clock.
    With TheHdw.DIB.SupportBoardClock
        .stop
        .Disconnect
    End With
    SBC_Enable = 0
    TheExec.Flow.TestLimit SBC_Enable, 1, 1, tlSignGreaterEqual, tlSignLessEqual, Tname:="SBC disable" 'BurstResult=1:Pass
    'printing in data log
    TheExec.Datalog.WriteComment "******************  Disable Support BD clock ****************"
End Function
Public Function FreeRunclk_Enable_ori(PortName As String)

    Dim site As Variant
    Dim i As Long
    Dim PLLLockChecked As New SiteLong
    Dim measf As New PinListData
    
    Dim NotLocked As Boolean
    Dim XI0_REFCLK As String
    Dim PortMode As String
    Dim PLL_Lock As New SiteLong
    Dim port_level_name As String
    Dim port_level_value As Double
    Dim FreeRunFreq As Double
    On Error GoTo errHandler
    
    'Default to stop nWire before apply new spec
    'If TheExec.Flow.EnableWord("XI0_nWire") = True Then ' remove this enable word, must trigger N-wire
        TheHdw.Protocol.ports(PortName).Halt
        TheHdw.Protocol.ports(PortName).Enabled = False
    'End If
    
    'CurrentXi0Freq = FreeRunFreq
    'If TheExec.Flow.EnableWord("XI0_nWire") = True Then  remove this enable word, must trigger N-wire
        If TheExec.DataManager.instanceName <> "" Then
            TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
        Else
            TheHdw.Digital.ApplyLevelsTiming False, True, True, tlPowered, , , , Level_nWire, "common", "typ", TSB_nWire, "XI0_24M_TCK_24M", "typ"
        End If

        
        '' 20151028 - Setup FRCPath by input port name
        TheHdw.Protocol.ports(PortName).Enabled = True
        TheHdw.Protocol.ports(PortName).NWire.ResetPLL
        TheHdw.Wait 0.001
        
        '///////////////////////////////////////////////////////////////////////
        If LCase(PortName) Like "*diff*" Then
               port_level_name = Replace(LCase(PortName), "port", "pa")
               port_level_value = TheHdw.Digital.Pins(port_level_name).DifferentialLevels.value(chVid)
        Else
               port_level_name = Replace(LCase(PortName), "port", "pa")
               port_level_value = TheHdw.Digital.Pins(port_level_name).Levels.value(chVih)
        End If
        '///////////////////////////////////////////////////////////////////////
        
        ' Test the PLL lock status and datalog results.
        For Each site In TheExec.sites.Selected
            If TheHdw.Protocol.ports(PortName).NWire.IsPLLLocked = False Then
                PLLLockChecked = 1
                'TheExec.Datalog.WriteComment "print: site(" & Site & "), PortName(" & PortName & "), IsPLLLocked = False"
                PLL_Lock(site) = 0
            Else
                PLLLockChecked = 0
                'TheExec.Datalog.WriteComment "print: site(" & Site & "), PortName(" & PortName & "), IsPLLLocked = True"   'comment only
                'TheExec.Datalog.WriteComment "print: site(" & Site & "), PortName(" & PortName & "), Wake up finished"
                PLL_Lock(site) = 1
            End If
        Next site
        
        ' Start the nWire engine.
        Call TheHdw.Protocol.ports(PortName).NWire.Frames("RunFreeClock").Execute
        
        TheHdw.Protocol.ports(PortName).IdleWait
        TheHdw.Wait 0.001
        
        'offline mode simulation
        If TheExec.TesterMode = testModeOffline Then
            For Each site In TheExec.sites.Selected
                PLL_Lock(site) = 1
            Next site
        End If

        TheExec.Flow.TestLimit PLL_Lock, 1, 1, tlSignGreaterEqual, tlSignLessEqual, Tname:="nWirePLL_Lock" 'BurstResult=1:Pass

        'TheExec.Datalog.WriteComment ""
        'TheExec.Datalog.WriteComment "********** Enable freerunning clock *********"
                      
        '****print out to data log about nWire clock condition
        FreeRunFreq_debug = 1 / TheHdw.Digital.Timing.period(PortName) / 1000000

        'offline mode simulation
        If TheExec.TesterMode = testModeOffline Then
            FreeRunFreq = 24000000  'offline
            FreeRunFreq_debug = FreeRunFreq / 1000000
        End If
        
        TheExec.Datalog.WriteComment "********** freerunning clock = " & Format(FreeRunFreq_debug, "0.000") & " Mhz, port_level_value" = " & port_level_value"

    Exit Function
    
errHandler:
    If AbortTest Then Exit Function Else Resume Next


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
    If PowerDown_Flag = False Then TheExec.Flow.TestLimit 0, 0, 0, tlSignGreaterEqual, tlSignLessEqual, Tname:="nWire halt" 'BurstResult=1:Pass
    'printing to data log
    'TheExec.Datalog.WriteComment "******************  Disable freerunning clock ****************"
    
    ''upload to global constant
    FreeRunFreq_debug = 0
    clock_Vih_debug = 0
    clock_Vil_debug = 0

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "FreeRunClk_Disable") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Function Start_Profile_DebugOnly(PinName As PinList, WhatToCapture As String, SampleRate As Double, sampleSize As Long, Optional CapSignalName As String = "Capture_signal")
'start current or voltage profile capturing

On Error GoTo errHandler
' Wait if another capture is running
    Do While TheHdw.DCVS.Pins(PinName).Capture.IsRunning = True
    Loop
    
    'Create a SIGNAL to set up instrument
    TheHdw.DCVS.Pins(PinName).Capture.Signals.Add CapSignalName
    
    'Set this as the default signal
    TheHdw.DCVS.Pins(PinName).Capture.Signals.DefaultSignal = CapSignalName
    
    'Define the signal used for the capture
    With TheHdw.DCVS.Pins(PinName).Capture.Signals.item(CapSignalName)
        .Reinitialize
        If (WhatToCapture = "I") Then
            .mode = tlDCVSMeterCurrent
            .range = TheHdw.DCVS.Pins(PinName).CurrentRange.Max '2
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
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Start_Profile_DebugOnly") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next

End Function
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
    .Reinitialize
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
    If Current_Insatance = "" Then TheExec.Datalog.WriteComment "<ERROR> Instance name is empty.Please check instance global name is defined."
    
    
    Dim Current_Dirctory As String
    Dim Plot_Profile_Label() As String
    Dim Profile_Split() As String
    Dim Mapping As String
    Dim original_InstName As String
    Dim Instance_mappingName As String
    Dim mapping_txt As String
    Dim TestName As String
    Dim CurrentProfile As New DSPWave
    ReDim Plot_Profile_Label(0)
    Dim t1 As Double
    Dim t2 As Double
    Dim sampleR As String
    Dim sampleSize As String
    Dim sPowerPin As String
    Dim sDCVS_Pin As String
    Dim sDCVI_Pin As String
    
    If percent_control = 0 Then percent_control = 0.99
    If TheExec.enableWord("DisablePlot_IProfile") = False Or TheExec.enableWord("Print_IProfileValue") = True Then
        Call TheExec.DataManager.DecomposePinList(PinName, Pin_Ary(), Pin_Cnt)
        TestName = TheExec.DataManager.instanceName
        If create_folderName = False Then
            Profile_Folder = "X" & CStr(XCoord(0)) & "Y" & CStr(YCoord(0)) & "_" & right("0" & CStr(Month(Now)), 2) & right("0" & CStr(Day(Now)), 2) & right("0" & CStr(Hour(Now)), 2) & right("0" & CStr(Minute(Now)), 2)
            create_folderName = True
        End If
        If TheExec.enableWord("DisablePlot_IProfile") = False Then
            Current_Dirctory = CurDir & "\" & Profile_Folder
            If Dir(Current_Dirctory, vbDirectory) = Empty Then
                MkDir Current_Dirctory
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
                ElseIf LCase(gl_GetInstrumentType_Dic(sPowerPin)) Like "*dcvi*" Then
                    dspw = TheHdw.DCVI.Pins(sPowerPin).Capture.Signals(CapSignalName).DSPWave
                Else
                    Call Print_Error_Message(Error_Warning_Info.Warning_Info, "VBT_LIB_Common", "Plot_Profile", "Use wrong instrument to plot current profile !!")
                End If
            
                For Each site In TheExec.sites
                    If TheExec.enableWord("Print_IProfileValue") = True Then
                        Call Print_IProfileValue(sPowerPin, TestName, dspw, site, percent_control)
                    End If
                    If TheExec.enableWord("DisablePlot_IProfile") = False Then
                        'new instance name ex:SocSaChain_MXXXX_SRMDSSC_MN_CREA0_S_00001
        '                TheHdw.Digital.Patgen.ReadLastStart lastBurstPat, isGrp, lastLabel
                        If LCase(gl_GetInstrumentType_Dic(sPowerPin)) Like "*dcvs*" Then
                            sampleR = CStr(TheHdw.DCVS.Pins(sPowerPin).Capture.SampleRate)
                            sampleSize = CStr(TheHdw.DCVS.Pins(sPowerPin).Capture.sampleSize)
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
                            Dim tempStr As String
                            tempStr = Current_Dirctory & "\"  ''20211005
                            Dim fso As New FileSystemObject
                             
                            If Dir(tempStr, vbDirectory) = Empty Then
                                MkDir tempStr
                            End If
                            dspw.FileExport tempStr & "\" & filename, File_txt
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
                If TheExec.enableWord("DisablePlot_IProfile") = False Then
                    Set fs = CreateObject("Scripting.FileSystemObject")
                    'ex: site  powerPin   original_InstName : Instance_mappingName
                    mapping_txt = CStr(sPowerPin) + "    " + original_InstName + " : " + Instance_mappingName
                    If fs.FileExists(tempStr & "mapping.txt") = False Then
                        'create a mapping file
                        Set mapping_textStream = fs.CreateTextFile(tempStr & "mapping.txt", True)
                        mapping_textStream.WriteLine "ProjectName : " & TheExec.TestProgram.name
                        mapping_textStream.WriteLine "JobName : " & TheExec.CurrentJob
                        mapping_textStream.WriteLine ""
                        mapping_textStream.WriteLine "---------------------------------------------------------"
                        mapping_textStream.WriteLine "| PowerPin | OriginalInstanceName : MappingInstanceName |"
                        mapping_textStream.WriteLine "---------------------------------------------------------"
                        mapping_textStream.WriteLine mapping_txt
                        
                    Else
                        If UBound(Profile_Split) >= 6 Then mapping_textStream.WriteLine mapping_txt
                    End If
                    Set dspw = Nothing
                End If
            End If
        Next p
        If Calculate_ProfileInfo Then Call Calculate_Plot_Profile(Plot_Profile_Label)

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

For Each site In TheExec.sites.Active
    If TheHdw.DCVI.Pins(PinName).Meter.mode = tlDCVIMeterCurrent Then
        Label = "Current Profile for Site: " & site
    Else
        Label = "Voltage Profile for Site: " & site
End If
    
     dspw.Plot Label
    
Next site

Exit Function

errHandler:
        If isDebugMode Then TheExec.AddOutput "Error in the Plot Profile"
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Plot_profile_DCVI") 'Add ErrHandler 2023/08/18
                If AbortTest Then Exit Function Else Resume Next
End Function

Function Start_Profile_AutoResolution(PinName As String, WhatToCapture As String, Optional CapSignalName As String = "Capture_signal", Optional Plottime As Double = 0, Optional ByFlow As Boolean = False, Optional nSampleSize As Long = 1)
'start current or voltage profile capturing

On Error GoTo errHandler
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
    If ByFlow Then
        Profile_Header = UCase(TheExec.Flow.CurrentFlowSheetName)
    Else
        Profile_Header = UCase(TheExec.DataManager.instanceName)
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
    
    If TheExec.enableWord("DownSample_IProfile") = True Then
    
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
        Profile_SampleSize_Hex = nSampleSize
        Call ProfileAutoResolution(glbConstIns_HEXVS, Plottime, Profile_SampleSize_Hex, Profile_SampleRate_Hex, _
                                    Down_SampleRatio_Hex, nSampleSize)
        StartProfile sPin_HEXVS, WhatToCapture, Profile_SampleRate_Hex, Profile_SampleSize_Hex, CapSignalName, glbConstIns_HEXVS
    End If
        
    If sPin_VHDVS <> "" Then
        Profile_SampleSize_UVS = nSampleSize
        Call ProfileAutoResolution(glbConstIns_VHDVS, Plottime, Profile_SampleSize_UVS, Profile_SampleRate_UVS, _
                                    Down_SampleRatio_UVS, nSampleSize)
        StartProfile sPin_VHDVS, WhatToCapture, Profile_SampleRate_UVS, Profile_SampleSize_UVS, CapSignalName, glbConstIns_VHDVS
    End If
    
    If sPin_VSM <> "" Then
        Profile_SampleSize_VSM = nSampleSize
        Call ProfileAutoResolution(glbConstIns_VSM, Plottime, Profile_SampleSize_VSM, Profile_SampleRate_VSM, _
                                    Down_SampleRatio_VSM, nSampleSize)
        StartProfile sPin_VSM, WhatToCapture, Profile_SampleRate_VSM, Profile_SampleSize_VSM, CapSignalName, glbConstIns_VSM
    End If
    
    If sPin_VS5A <> "" Then
        Profile_SampleSize_VS5A = nSampleSize
        Call ProfileAutoResolution(glbConstIns_VS5A, Plottime, Profile_SampleSize_VS5A, Profile_SampleRate_VS5A, _
                                    Down_SampleRatio_VS5A, nSampleSize)
        StartProfile sPin_VS5A, WhatToCapture, Profile_SampleRate_VS5A, Profile_SampleSize_VS5A, CapSignalName, glbConstIns_VS5A
    End If
    
    If sPin_VS800MA <> "" Then
        Profile_SampleSize_VS800mA = nSampleSize
        Call ProfileAutoResolution(glbConstIns_VS800MA, Plottime, Profile_SampleSize_VS800mA, Profile_SampleRate_VS800mA, _
                                    Down_SampleRatio_VS800mA, nSampleSize)
        StartProfile sPin_VS800MA, WhatToCapture, Profile_SampleRate_VS800mA, Profile_SampleSize_VS800mA, CapSignalName, glbConstIns_VS800MA
    End If
    
    If sPin_DC07 <> "" Then
        Call ProfileAutoResolution_DCVI(glbConstIns_DC07, Plottime, Profile_SampleSize_DC07, Profile_SampleRate_DC07, Down_SampleRatio_DC07)    ', UVI80_Mode:=TheHdw.DCDiffMeter.pins(sPin_DC07).MeterMode
        StartProfile_DCVI sPin_DC07, WhatToCapture, Profile_SampleRate_DC07, Profile_SampleSize_DC07, CapSignalName, glbConstIns_DC07
    End If
    
    If sPin_DC30 <> "" Then
        Call ProfileAutoResolution_DCVI(glbConstIns_DC30, Plottime, Profile_SampleSize_DC30, Profile_SampleRate_DC30, Down_SampleRatio_DC30)
        StartProfile_DCVI sPin_DC30, WhatToCapture, Profile_SampleRate_DC30, Profile_SampleSize_DC30, CapSignalName, glbConstIns_DC30
    End If
    If sPin_DC75 <> "" Then
        Call ProfileAutoResolution_DCVI(glbConstIns_DC75, Plottime, Profile_SampleSize_DC75, Profile_SampleRate_DC75, Down_SampleRatio_DC75)
        StartProfile_DCVI sPin_DC75, WhatToCapture, Profile_SampleRate_DC75, Profile_SampleSize_DC75, CapSignalName, glbConstIns_DC75
    End If
    
   
   '==============for  DownSample_IProfile  debug==================================
    TheExec.Datalog.WriteComment "DownSample_IProfile:" & TheExec.enableWord("DownSample_IProfile")
    TheExec.Datalog.WriteComment "Profile_SampleSize_Hex : " & Profile_SampleSize_Hex & "     Profile_SampleRate_Hex : " & Profile_SampleRate_Hex & "    Down_SampleRatio_Hex : " & Down_SampleRatio_Hex
    TheExec.Datalog.WriteComment "Profile_SampleSize_UVS : " & Profile_SampleSize_UVS & "     Profile_SampleRate_UVS : " & Profile_SampleRate_UVS & "    Down_SampleRatio_UVS : " & Down_SampleRatio_UVS
    TheExec.Datalog.WriteComment "Profile_SampleSize_VSM : " & Profile_SampleSize_VSM & "     Profile_SampleRate_VSM : " & Profile_SampleRate_VSM & "    Down_SampleRatio_VSM : " & Down_SampleRatio_VSM
    TheExec.Datalog.WriteComment "Profile_SampleSize_VS5A : " & Profile_SampleSize_VS5A & "     Profile_SampleRate_VS5A : " & Profile_SampleRate_VS5A & "    Down_SampleRatio_VS5A : " & Down_SampleRatio_VS5A
    TheExec.Datalog.WriteComment "Profile_SampleSize_VS800mA : " & Profile_SampleSize_VS800mA & "     Profile_SampleRate_VS800mA : " & Profile_SampleRate_VS800mA & "    Down_SampleRatio_VS800mA : " & Down_SampleRatio_VS800mA
    TheExec.Datalog.WriteComment "====DCVI===="
    TheExec.Datalog.WriteComment "Profile_SampleSize_DC07 : " & Profile_SampleSize_DC07 & "     Profile_SampleRate_DC07 : " & Profile_SampleRate_DC07 & "    Down_SampleRatio_DC07 : " & Down_SampleRatio_DC07
    TheExec.Datalog.WriteComment "Profile_SampleSize_DC30 : " & Profile_SampleSize_DC30 & "     Profile_SampleRate_DC30 : " & Profile_SampleRate_DC30 & "    Down_SampleRatio_DC30 : " & Down_SampleRatio_DC30
    TheExec.Datalog.WriteComment "Profile_SampleSize_DC75 : " & Profile_SampleSize_DC75 & "     Profile_SampleRate_DC75 : " & Profile_SampleRate_DC75 & "    Down_SampleRatio_DC75 : " & Down_SampleRatio_DC75
    TheExec.Datalog.WriteComment "====DCVI===="
   '==============for  DownSample_IProfile  debug==================================
   
    Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Start_Profile_AutoResolution") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function Print_Footer(PrintInfo As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    TheExec.Datalog.WriteComment "******************************"
    TheExec.Datalog.WriteComment "*print: " & PrintInfo & " end*"
    TheExec.Datalog.WriteComment "******************************"
    
    If InStr(PrintInfo, "SOC_CHAIN_") = 0 And InStr(PrintInfo, "SOC_SAA_") = 0 And InStr(PrintInfo, "SOC_TD_") = 0 _
    And InStr(PrintInfo, "CPU_CHAIN_") = 0 And InStr(PrintInfo, "CPU_SAA_") = 0 And InStr(PrintInfo, "CPU_TD_") = 0 _
    And InStr(PrintInfo, "GFX_CHAIN_") = 0 And InStr(PrintInfo, "GFX_SAA_") = 0 And InStr(PrintInfo, "GFX_TD_") = 0 Then
        ''20210504 reset the datalog format
        Init_Datalog_Setup
    End If

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Print_Footer") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
Public Function Print_Header(PrintInfo As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    TheExec.Datalog.WriteComment "********************************"
    TheExec.Datalog.WriteComment "*print: " & PrintInfo & " start*"
    TheExec.Datalog.WriteComment "********************************"

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

    TheExec.Datalog.WriteComment "********************************"
    TheExec.Datalog.WriteComment "*print: Program information    *"
    TheExec.Datalog.WriteComment "********************************"
    
    OS_Info
    IGXL_Info
    RunOptions_Info
    Version_Info
    Print_Version
    Print_DIBInfo
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Print_PgmInfo") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Print_DIBInfo()
    Dim ProCardID As String
    
    ProCardID = RegKeyRead("PROBECARD_ID")
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment "-----------------------------------------------"
    If ProCardID <> "" Then
        TheExec.Datalog.WriteComment "Probe Card ID: " & ProCardID
    Else
        TheExec.Datalog.WriteComment "No Probe Card ID"
    End If
    TheExec.Datalog.WriteComment "-----------------------------------------------"
    TheExec.Datalog.WriteComment ""
    
    Exit Function
errHandler:     'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Print_PgmInfo") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
'*****************************************
'******            Read/Write EPPROM******
'*****************************************
Public Function Write_DIB_EEPROM(Optional DIB_SerialNumber As String) As Long
    On Error GoTo errHandler
    Dim CurrJob As String
    Dim config(2) As Long
    config(0) = 32768
    config(1) = 0
    config(2) = 0
    
    TheHdw.DIB.PIBEEPROM.program (config)
'''    CurrJob = TheExec.CurrentJob    'CP/FT judge
'''    change to LCase to prevent ft1 FT1 issue

    If TheExec.Flow.enableWord("Write_EEPROM_DIBID") = True Then
'''     If (CurrJob Like "cp*") Then
        If LCase(TheExec.CurrentJob) Like "cp*" Then
            If RegKeyRead("PROBECARD_ID") <> "" Then
                DIB_SerialNumber = RegKeyRead("PROBECARD_ID")
            Else
                DIB_SerialNumber = vbNullString
            End If
'''     ElseIf (CurrJob Like "ft*") Then
        ElseIf LCase(TheExec.CurrentJob) Like "ft*" Then
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
        TheExec.Datalog.WriteComment ("print: Write DIB ID " & DIB_SerialNumber)
    End If  'end enable wd block
    'only fuse once
    TheExec.Flow.enableWord("Write_EEPROM_DIBID") = False
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Write_DIB_EEPROM") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
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
    On Error GoTo errHandler
    
    Dim REC() As IDIB_EEPROM_RecordObj
    Dim i As Integer
    Dim DIB_ID As String
    
    If TheHdw.DIB.PIBEEPROM.IsProgrammed Then
        'Debug.Print "The PIB EEPROM is programmed"
        REC = TheHdw.DIB.PIBEEPROM.Record.list
    
'''        For i = 0 To UBound(rec)
'''            Debug.Print "Record " + rec(i).ID + " = " + rec(i).Value
'''        Next i

        DIB_ID = REC(0).value
        TheExec.Datalog.WriteComment ("print: Read DIB ID is " & DIB_ID)
    Else
        'Debug.Print "The PIB EEPROM is not programmed"
        TheExec.Datalog.WriteComment ("print: Read DIB ID is not programmed, read fail")
    End If
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Read_DIB_EEPROM") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ReadProberTemp(Temp_Hilimit As Double, Temp_Lolimit As Double)

Dim Prober_Temp As Double
Dim Prober_Temp_str As String
Dim Handler_Temp As Double
Dim tempHandler_Temp() As String
Dim Handler_Temp_str As String
On Error GoTo errHandler
    If UCase(TheExec.CurrentChanMap) Like "CHANNELMAP_FT*" Then
        TheExec.Datalog.WriteComment ""
        TheExec.Datalog.WriteComment "********************************"
        TheExec.Datalog.WriteComment "*print: Read Handler Temperature*"
        TheExec.Datalog.WriteComment "********************************"
        If CStr(RegKeyRead("Dut_Temperature")) <> "" Then
            Handler_Temp_str = RegKeyRead("Dut_Temperature")
            If InStr(1, Handler_Temp_str, "_") Then
                tempHandler_Temp = Split(Handler_Temp_str, "_")
                If UCase(tempHandler_Temp(1)) = "NULL" Then
                    tempHandler_Temp(1) = "-999"
                End If
                Handler_Temp = CDbl(tempHandler_Temp(1))
            Else
                TheExec.Datalog.WriteComment ("Please Check Handler RegKey Dut_Temperature")
            End If
        Else
            TheExec.Datalog.WriteComment "Registry is empty"
            Handler_Temp = "00000"
        End If
        gL_ProductionTemp = CStr(Handler_Temp) 'upload to global variant
        If TheExec.TesterMode = testModeOffline Then    ''' offline mode
            TheExec.Datalog.WriteComment "Handler_Temp(Registry) : offline_mode_fake_temperature_value_bypass_check" ''offline_mode_bypass_check --> offline_mode_fake_temperature_value_bypass_check
            TheExec.Flow.TestLimit 1, lowVal:=1, hiVal:=1, Tname:="Handler_Temp_test" & CStr(mid(TheExec.DataManager.instanceName, 16, 18))
        Else                                            ''' online mode
            If TheExec.RunMode = runModeDebug Then      ''' debug mode
                TheExec.Datalog.WriteComment "Handler_Temp(Registry) : engineering_mode_fake_temperature_value_bypass_check" ''debug_mode_bypass_check --> engineering_mode_fake_temperature_value_bypass_check
                TheExec.Flow.TestLimit 1, lowVal:=1, hiVal:=1, Tname:="Handler_Temp_test" & CStr(mid(TheExec.DataManager.instanceName, 16, 18))
            Else                                        ''' production mode
                TheExec.Datalog.WriteComment "Handler_Temp(Registry) : " & RegKeyRead("Dut_Temperature")
                TheExec.Flow.TestLimit Handler_Temp, lowVal:=Temp_Lolimit, hiVal:=Temp_Hilimit, Tname:="Handler_Temp_" & CStr(mid(TheExec.DataManager.instanceName, 16, 18))
            End If
        End If
    ElseIf UCase(TheExec.CurrentChanMap) Like "CHANNELMAP_CP*" Or UCase(TheExec.CurrentChanMap) Like "CHANNELMAP_WLFT*" Then
        TheExec.Datalog.WriteComment ""
        'TheExec.Datalog.WriteComment "/********* Read Prober Temperature *********/"
        TheExec.Datalog.WriteComment "********************************"
        TheExec.Datalog.WriteComment "*print: Read Prober Temperature*"
        TheExec.Datalog.WriteComment "********************************"
        
        If IsNumeric(RegKeyRead("Prober_Temp")) = True Then
            Prober_Temp_str = RegKeyRead("Prober_Temp")

        Else
            If RegKeyRead("Prober_Temp") = "" Then
                TheExec.Datalog.WriteComment "Registry is empty"
                Prober_Temp_str = "00000"
            Else
                TheExec.Datalog.WriteComment "Registry is not empty nor number."
                Prober_Temp_str = "99999"
            End If
        End If
        
        If mid(Prober_Temp_str, 1, 1) Like "+" Then
                Prober_Temp = CDbl(mid(Prober_Temp_str, 2, 5))
        Else
                Prober_Temp = CDbl(Prober_Temp_str)
        End If
        
        gL_ProductionTemp = Prober_Temp 'upload to global variant
        
        'offline mode simulation
        If TheExec.TesterMode = testModeOffline Then
            TheExec.Datalog.WriteComment "Prober_Temp(Registry) : offline_mode_fake_temperature_value_bypass_check" ''offline_mode_bypass_check --> offline_mode_fake_temperature_value_bypass_check
            TheExec.Flow.TestLimit 1, lowVal:=1, hiVal:=1, Tname:="Prober_Temp_test" & CStr(mid(TheExec.DataManager.instanceName, 16, 18))
        Else                                            ''' online mode
            If TheExec.RunMode = runModeDebug Then      ''' debug mode
                TheExec.Datalog.WriteComment "Prober_Temp(Registry) : engineering_mode_fake_temperature_value_bypass_check" ''debug_mode_bypass_check --> engineering_mode_fake_temperature_value_bypass_check
                TheExec.Flow.TestLimit 1, lowVal:=1, hiVal:=1, Tname:="Prober_Temp_test" & CStr(mid(TheExec.DataManager.instanceName, 16, 18))
            Else                                        ''' production mode
        TheExec.Datalog.WriteComment "Prober_Temp(Registry) : " & RegKeyRead("Prober_Temp")
        TheExec.Flow.TestLimit Prober_Temp, lowVal:=Temp_Lolimit, hiVal:=Temp_Hilimit, Tname:="Prober_Temp_" & CStr(mid(TheExec.DataManager.instanceName, 16, 18))
            End If
        End If
    Else
         TheExec.Datalog.WriteComment TheExec.DataManager.instanceName & " Can not Judge the Channel Map Name, Please Check the Channel Map Naming "
    End If
       
    Exit Function
errHandler:
    If UCase(TheExec.CurrentChanMap) Like "CHANNELMAP_FT*" Then
        TheExec.Datalog.WriteComment "Read Handler Temp VBT function is error "
        TheExec.Datalog.WriteComment "Registry String :" & RegKeyRead("Dut_Temperature")
        TheExec.Datalog.WriteComment ("Error #: " & str(err.number) & " " & err.Description)
    Else
        TheExec.Datalog.WriteComment "Read Prober Temp VBT function is error "
        TheExec.Datalog.WriteComment "Registry String :" & RegKeyRead("Prober_Temp")
        TheExec.Datalog.WriteComment ("Error #: " & str(err.number) & " " & err.Description)
    End If
        Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "ReadProberTemp") 'Add ErrHandler 2023/08/18
        If AbortTest Then Exit Function Else Resume Next
End Function


Public Function SetupInitialCondition(DisableConnectPins As PinList, DrivePins As PinList, Optional DriveVolt As Double = 0.1) As Long
    If (DisableConnectPins <> "") Then
        TheHdw.Digital.Pins(DisableConnectPins).Disconnect
    End If
    If (DrivePins <> "") Then
        TheHdw.Digital.Pins(DrivePins).Disconnect
        TheHdw.Wait (100 * us)
        
        With TheHdw.PPMU.Pins(DrivePins)
            .Gate = tlOff
            .ForceV DriveVolt, 0.002
            .Connect
            TheHdw.Wait (100 * us)
            .Gate = tlOn
        End With
    End If
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


On Error GoTo errHandler


    
    If Pin_GP1.value = "" Then
    Else
        TheHdw.PPMU.Pins(Pin_GP1).ClampVHi = Pin_GP1_Vch
    End If
    
    If Pin_GP2.value = "" Then
    Else
        TheHdw.PPMU.Pins(Pin_GP2).ClampVHi = Pin_GP2_Vch
    End If
    
    If Pin_GP3.value = "" Then
    Else
        TheHdw.PPMU.Pins(Pin_GP3).ClampVHi = Pin_GP3_Vch
    End If

    If Pin_GP4.value = "" Then
    Else
        TheHdw.PPMU.Pins(Pin_GP4).ClampVHi = Pin_GP4_Vch
    End If
 
    If Pin_GP5.value = "" Then
    Else
        TheHdw.PPMU.Pins(Pin_GP5).ClampVHi = Pin_GP5_Vch
    End If
 
    If Pin_GP6.value = "" Then
    Else
        TheHdw.PPMU.Pins(Pin_GP6).ClampVHi = Pin_GP6_Vch
    End If
 
    If Pin_GP7.value = "" Then
    Else
        TheHdw.PPMU.Pins(Pin_GP7).ClampVHi = Pin_GP7_Vch
    End If
 
    If Pin_GP8.value = "" Then
    Else
        TheHdw.PPMU.Pins(Pin_GP8).ClampVHi = Pin_GP8_Vch
    End If
 
    If Pin_GP9.value = "" Then
    Else
        TheHdw.PPMU.Pins(Pin_GP9).ClampVLo = Pin_GP9_Vcl
    End If
    If Pin_GP10.value = "" Then
    Else
        TheHdw.PPMU.Pins(Pin_GP10).ClampVLo = Pin_GP10_Vcl
    End If
 
    Exit Function
 
errHandler:

                If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Read_Package_ID()

Dim site As Variant

On Error GoTo errHandler

        TheExec.Datalog.WriteComment "********************************"
        
        For Each site In TheExec.sites
            TheExec.Datalog.WriteComment ("Site:" & site & " Device ID : " & RegKeyRead("ManualTestDeviceID"))
        Next site
        
        TheExec.Datalog.WriteComment "********************************"
        
        Dim TestName As String
        TestName = "Device ID"
        Dim ResultVal_DeviceID As Double
        ResultVal_DeviceID = CDbl(RegKeyRead("ManualTestDeviceID"))
        For Each site In TheExec.sites
            
            TheExec.Flow.TestLimit resultVal:=ResultVal_DeviceID, Tname:=TestName, ForceResults:=tlForceNone
        Next site
        
       
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Set_PPMU_Clamp") 'Add ErrHandler 2023/08/18
        If AbortTest Then Exit Function Else Resume Next
End Function


Public Function FreeRunClk_Disable_MultiPort(PortName As String) As Long '''update for multi nWire 20170718
    Dim site As Variant
    Dim i As Integer
    Dim A_PortName() As String
    
    Call Disable_FRC(vbNullString, False)
''    theexec.Datalog.WriteComment "******************  Disable freerunning clock ****************"
''
''    A_PortName = Split(PortName, ",")
''
''    For i = 0 To UBound(A_PortName)
''        ' Disable the nWire engine.
''        'If TheExec.Flow.EnableWord("XI0_nWire") = True Then
''            TheHdw.Protocol.ports(A_PortName(i)).Halt
''            TheHdw.Protocol.ports(A_PortName(i)).Enabled = False
''        'End If
'''        If A_PortName(i) Like "Clock_Port" Then
'''            TheHdw.Digital.Pins("Xi0_PA").InitState = chInitLo
'''        ElseIf PortName Like "RTCLK_Port" Then
'''            TheHdw.Digital.Pins("RT_CLK32768_PA").InitState = chInitLo
'''        End If
'''        TheHdw.Wait 0.001
'''        TheHdw.Digital.Pins(A_PortName(i)).InitState = chInitoff
''
''        'TheHdw.DIB.SupportBoardClock.Stop
''    Next i
''
''    ''upload to global constant
''    FreeRunFreq_debug = 0
''    clock_Vih_debug = 0
''    clock_Vil_debug = 0

End Function


Public Function CheckFlag(Optional B_PrintFlag As Boolean = False)

    Dim F_Sa_Flag As String
    Dim F_SaChain_Flag As String
    Dim F_TD_Flag As String
    Dim F_Bist_Flag As String
    Dim F_Bist_MC_Flag As String
    Dim F_SaHV_Flag As String
    Dim F_SaLV_Flag As String
    Dim F_SaChainHV_Flag As String
    Dim F_SaChainLV_Flag As String
    Dim F_Pass As String
    Dim F_Fail As String
    Dim site As Variant
    Dim DLY_PassCount As Long
    Dim DLY_Top1_PassCount As Long: DLY_Top1_PassCount = 16
    Dim DLY_Top2_PassCount As Long: DLY_Top2_PassCount = 16
    Dim PLY_PassCount As Long
    Dim PLY_Top1_PassCount As Long: PLY_Top1_PassCount = 16
    Dim PLY_Top2_PassCount As Long: PLY_Top2_PassCount = 16
    Dim S_FlagName As String
    Dim S_FlagState As String
    Dim S_Sa_FlagState As String
    Dim S_SaChain_FlagState As String
    Dim S_TD_FlagState As String
    Dim S_Bist_FlagState As String
    Dim S_Bist_MC_FlagState As String
    Dim S_SaHV_FlagState As String
    Dim S_SaLV_FlagState As String
    Dim S_SaChainHV_FlagState As String
    Dim S_SaChainLV_FlagState As String
    Dim B_FirstSite As Boolean
    Dim i As Integer
    Dim A_FlagName() As String
    Dim A_FlagState() As String
    
    B_FirstSite = True
    
    For Each site In TheExec.sites
        S_FlagName = vbNullString
        S_FlagState = vbNullString
        DLY_Top1_PassCount = 16
        DLY_Top2_PassCount = 16
        PLY_Top1_PassCount = 16
        PLY_Top2_PassCount = 16
        
        For i = 0 To 31
            'Right(CStr(i + 100), 2) -> to get 2 character while convert long to string
            F_Sa_Flag = "F_cpusa_B" & right(CStr(i + 100), 2) & "_MHLV"
            F_SaChain_Flag = "F_cpusachain_B" & right(CStr(i + 100), 2) & "_MHLV"
            F_TD_Flag = "F_CpuTD_B" & right(CStr(i + 100), 2) & "_NV"
            F_Bist_Flag = "F_cpubist_B" & right(CStr(i + 100), 2) & "_MHLV"
            F_Bist_MC_Flag = "F_cpubist_B" & right(CStr(i + 100), 2) & "_NV"
            
            F_SaHV_Flag = "F_cpusa_B" & right(CStr(i + 100), 2) & "_HV"
            F_SaLV_Flag = "F_cpusa_B" & right(CStr(i + 100), 2) & "_LV"
            F_SaChainHV_Flag = "F_cpusachain_B" & right(CStr(i + 100), 2) & "_HV"
            F_SaChainLV_Flag = "F_cpusachain_B" & right(CStr(i + 100), 2) & "_LV"
            
            ' set SA MHLV flag
            If TheExec.sites.item(site).FlagState(F_SaHV_Flag) = logicFalse Or TheExec.sites.item(site).FlagState(F_SaLV_Flag) = logicFalse Then
                TheExec.sites.item(site).FlagState(F_Sa_Flag) = logicFalse
            ElseIf TheExec.sites.item(site).FlagState(F_SaHV_Flag) = logicTrue Or TheExec.sites.item(site).FlagState(F_SaLV_Flag) = logicTrue Then
                TheExec.sites.item(site).FlagState(F_Sa_Flag) = logicTrue
            Else
                TheExec.sites.item(site).FlagState(F_Sa_Flag) = logicClear
            End If
            
            ' set SA Chain MHLV flag
            If TheExec.sites.item(site).FlagState(F_SaChainHV_Flag) = logicFalse Or TheExec.sites.item(site).FlagState(F_SaChainLV_Flag) = logicFalse Then
                TheExec.sites.item(site).FlagState(F_SaChain_Flag) = logicFalse
            ElseIf TheExec.sites.item(site).FlagState(F_SaChainHV_Flag) = logicTrue Or TheExec.sites.item(site).FlagState(F_SaChainLV_Flag) = logicTrue Then
                TheExec.sites.item(site).FlagState(F_SaChain_Flag) = logicTrue
            Else
                TheExec.sites.item(site).FlagState(F_SaChain_Flag) = logicClear
            End If
            
            If i < 16 Then
                ' if any of SA, SAChain, Bist_MHLV fail on this Top1 cpu core, then this cpu core fail
                If TheExec.sites.item(site).FlagState(F_Sa_Flag) = logicTrue Or TheExec.sites.item(site).FlagState(F_SaChain_Flag) = logicTrue Or TheExec.sites.item(site).FlagState(F_Bist_Flag) = logicTrue Then
                    DLY_Top1_PassCount = DLY_Top1_PassCount - 1
                End If
                ' if any of SA, SAChain, TD, Bist, Bist_MC fail on this Top1 cpu core, then this cpu core fail
                If TheExec.sites.item(site).FlagState(F_Sa_Flag) = logicTrue Or TheExec.sites.item(site).FlagState(F_SaChain_Flag) = logicTrue Or TheExec.sites.item(site).FlagState(F_TD_Flag) = logicTrue Or TheExec.sites.item(site).FlagState(F_Bist_Flag) = logicTrue Or TheExec.sites.item(site).FlagState(F_Bist_MC_Flag) = logicTrue Then
                    PLY_Top1_PassCount = PLY_Top1_PassCount - 1
                End If
            Else
                ' if any of SA, SAChain, Bist_MHLV fail on this Top2 cpu core, then this cpu core fail
                If TheExec.sites.item(site).FlagState(F_Sa_Flag) = logicTrue Or TheExec.sites.item(site).FlagState(F_SaChain_Flag) = logicTrue Or TheExec.sites.item(site).FlagState(F_Bist_Flag) = logicTrue Then
                    DLY_Top2_PassCount = DLY_Top2_PassCount - 1
                End If
                ' if any of SA, SAChain, TD, Bist, Bist_MC fail on this Top2 cpu core, then this cpu core fail
                If TheExec.sites.item(site).FlagState(F_Sa_Flag) = logicTrue Or TheExec.sites.item(site).FlagState(F_SaChain_Flag) = logicTrue Or TheExec.sites.item(site).FlagState(F_TD_Flag) = logicTrue Or TheExec.sites.item(site).FlagState(F_Bist_Flag) = logicTrue Or TheExec.sites.item(site).FlagState(F_Bist_MC_Flag) = logicTrue Then
                    PLY_Top2_PassCount = PLY_Top2_PassCount - 1
                End If
            End If
            
            ' collect flag state
            If B_PrintFlag = True Then
                If TheExec.sites.item(site).FlagState(F_SaHV_Flag) = logicTrue Then
                    S_SaHV_FlagState = "T"
                ElseIf TheExec.sites.item(site).FlagState(F_SaHV_Flag) = logicFalse Then
                    S_SaHV_FlagState = "F"
                Else
                    S_SaHV_FlagState = "C"
                End If
                If TheExec.sites.item(site).FlagState(F_SaLV_Flag) = logicTrue Then
                    S_SaLV_FlagState = "T"
                ElseIf TheExec.sites.item(site).FlagState(F_SaLV_Flag) = logicFalse Then
                    S_SaLV_FlagState = "F"
                Else
                    S_SaLV_FlagState = "C"
                End If
                If TheExec.sites.item(site).FlagState(F_Sa_Flag) = logicTrue Then
                    S_Sa_FlagState = "T"
                ElseIf TheExec.sites.item(site).FlagState(F_Sa_Flag) = logicFalse Then
                    S_Sa_FlagState = "F"
                Else
                    S_Sa_FlagState = "C"
                End If
                If TheExec.sites.item(site).FlagState(F_SaChainHV_Flag) = logicTrue Then
                    S_SaChainHV_FlagState = "T"
                ElseIf TheExec.sites.item(site).FlagState(F_SaChainHV_Flag) = logicFalse Then
                    S_SaChainHV_FlagState = "F"
                Else
                    S_SaChainHV_FlagState = "C"
                End If
                If TheExec.sites.item(site).FlagState(F_SaChainLV_Flag) = logicTrue Then
                    S_SaChainLV_FlagState = "T"
                ElseIf TheExec.sites.item(site).FlagState(F_SaChainLV_Flag) = logicFalse Then
                    S_SaChainLV_FlagState = "F"
                Else
                    S_SaChainLV_FlagState = "C"
                End If
                If TheExec.sites.item(site).FlagState(F_SaChain_Flag) = logicTrue Then
                    S_SaChain_FlagState = "T"
                ElseIf TheExec.sites.item(site).FlagState(F_SaChain_Flag) = logicFalse Then
                    S_SaChain_FlagState = "F"
                Else
                    S_SaChain_FlagState = "C"
                End If
                If TheExec.sites.item(site).FlagState(F_TD_Flag) = logicTrue Then
                    S_TD_FlagState = "T"
                ElseIf TheExec.sites.item(site).FlagState(F_TD_Flag) = logicFalse Then
                    S_TD_FlagState = "F"
                Else
                    S_TD_FlagState = "C"
                End If
                If TheExec.sites.item(site).FlagState(F_Bist_Flag) = logicTrue Then
                    S_Bist_FlagState = "T"
                ElseIf TheExec.sites.item(site).FlagState(F_Bist_Flag) = logicFalse Then
                    S_Bist_FlagState = "F"
                Else
                    S_Bist_FlagState = "C"
                End If
                If TheExec.sites.item(site).FlagState(F_Bist_MC_Flag) = logicTrue Then
                    S_Bist_MC_FlagState = "T"
                ElseIf TheExec.sites.item(site).FlagState(F_Bist_MC_Flag) = logicFalse Then
                    S_Bist_MC_FlagState = "F"
                Else
                    S_Bist_MC_FlagState = "C"
                End If
                
                S_FlagName = S_FlagName & "," & F_SaHV_Flag & "," & F_SaLV_Flag & "," & F_Sa_Flag & "," & F_SaChainHV_Flag & "," & F_SaChainLV_Flag & "," & F_SaChain_Flag & "," & F_TD_Flag & "," & F_Bist_Flag & "," & F_Bist_MC_Flag
                S_FlagState = S_FlagState & "," & S_SaHV_FlagState & "," & S_SaLV_FlagState & "," & S_Sa_FlagState & "," & S_SaChainHV_FlagState & "," & S_SaChainLV_FlagState & "," & S_SaChain_FlagState & "," & S_TD_FlagState & "," & S_Bist_FlagState & "," & S_Bist_MC_FlagState
            End If
            
        Next i
        
        
        If DLY_Top1_PassCount < DLY_Top2_PassCount Then
            DLY_PassCount = DLY_Top1_PassCount
        Else
            DLY_PassCount = DLY_Top2_PassCount
        End If
        
        If PLY_Top1_PassCount < PLY_Top2_PassCount Then
            PLY_PassCount = PLY_Top1_PassCount
        Else
            PLY_PassCount = PLY_Top2_PassCount
        End If
        
        
        F_Pass = "P_DLY_Cpu_Core_" & DLY_PassCount & "_pass"
        TheExec.sites.item(site).FlagState(F_Pass) = logicTrue
        
        F_Pass = "P_PLY_Cpu_Core_" & PLY_PassCount & "_pass"
        TheExec.sites.item(site).FlagState(F_Pass) = logicTrue
        
        F_Pass = "P_PLY_Top1_Cpu_Core_" & PLY_Top1_PassCount & "_pass"
        TheExec.sites.item(site).FlagState(F_Pass) = logicTrue
        
        F_Pass = "P_PLY_Top2_Cpu_Core_" & PLY_Top2_PassCount & "_pass"
        TheExec.sites.item(site).FlagState(F_Pass) = logicTrue
        
        If B_PrintFlag = True Then
            If S_FlagName <> "" Then S_FlagName = right(S_FlagName, Len(S_FlagName) - 1)
            If S_FlagState <> "" Then S_FlagState = right(S_FlagState, Len(S_FlagState) - 1)
            
            A_FlagName = Split(S_FlagName, ",")
            A_FlagState = Split(S_FlagState, ",")
            
'            If B_FirstSite = True Then
                For i = 0 To UBound(A_FlagName)
                    TheExec.Datalog.WriteComment "Site(" & site & "), " & A_FlagName(i) & ", " & A_FlagState(i)
'                    TheExec.Datalog.WriteComment "FlagState_Site(" & Site & ")    " & S_FlagState
                Next i
'                B_FirstSite = False
'            Else
'                TheExec.Datalog.WriteComment "FlagState_Site(" & Site & ")    " & S_FlagState
'            End If


        If TheExec.sites.item(site).FlagState("F_socmbist_bist_hard_defect") = logicTrue Then
                    TheExec.Datalog.WriteComment "Site(" & site & "), " & "F_socmbist_bist_hard_defect" & ", T"
                ElseIf TheExec.sites.item(site).FlagState("F_socmbist_bist_hard_defect") = logicFalse Then
                    TheExec.Datalog.WriteComment "Site(" & site & "), " & "F_socmbist_bist_hard_defect" & ", F"
                Else
                    TheExec.Datalog.WriteComment "Site(" & site & "), " & "F_socmbist_bist_hard_defect" & ", C"
                End If
                
                If TheExec.sites.item(site).FlagState("F_socmbist_bist_bitcell") = logicTrue Then
                    TheExec.Datalog.WriteComment "Site(" & site & "), " & "F_socmbist_bist_bitcell" & ", T"
                ElseIf TheExec.sites.item(site).FlagState("F_socmbist_bist_bitcell") = logicFalse Then
                    TheExec.Datalog.WriteComment "Site(" & site & "), " & "F_socmbist_bist_bitcell" & ", F"
                Else
                    TheExec.Datalog.WriteComment "Site(" & site & "), " & "F_socmbist_bist_bitcell" & ", C"
                End If
        
                If TheExec.sites.item(site).FlagState("F_soc_fail") = logicTrue Then
                    TheExec.Datalog.WriteComment "Site(" & site & "), " & "F_soc_fail" & ", T"
                ElseIf TheExec.sites.item(site).FlagState("F_soc_fail") = logicFalse Then
                    TheExec.Datalog.WriteComment "Site(" & site & "), " & "F_soc_fail" & ", F"
                Else
                    TheExec.Datalog.WriteComment "Site(" & site & "), " & "F_soc_fail" & ", C"
                End If
                
        End If
        
        
    Next site
    
End Function


Public Function Alarm_binout()


    Dim site As Variant

    'bin out alarm
    For Each site In TheExec.sites

       If TheHdw.Alarms.Check = True Then

                'TheExec.Sites.Item(Site).SortNumber = 908
                'TheExec.Sites.Item(Site).BinNumber = 6
                TheExec.sites.item(site).result = tlResultFail

       End If

    Next site


    Exit Function
    
errHandler:

      If AbortTest Then Exit Function Else Resume Next

End Function
Public Function Set_DCVS_alarm(pin As PinList, AlarmTime As Double, Pin_GP2 As PinList, AlarmTime_GP2 As Double, Optional b_AlarmForceBin As Boolean = False)
        'add boolean to set force_bin or force_fail

On Error GoTo errHandler

    If pin.value = "" And Pin_GP2.value = "" Then
        TheExec.Datalog.WriteComment "*******************************************************************"
        TheExec.Datalog.WriteComment "Error on Set_DCVS_alarm, please fill in pin name to set alarm time."
        TheExec.Datalog.WriteComment "*******************************************************************"
    End If

    With TheHdw.DCVS.Pins(pin)
        .Gate = False
        TheHdw.Wait 1 * ms '20170209 Add to wait gate off to avoid set ifold time out error
        .mode = tlDCVSModeVoltage
        .Voltage.Main = 0
        
        If b_AlarmForceBin = False Then
            .Alarm(tlDCVSAlarmAll) = tlAlarmForceFail
        Else
            .Alarm(tlDCVSAlarmAll) = tlAlarmForceBin
        End If
        If CStr(AlarmTime) = "" Or AlarmTime < 0.0001 Then
            TheExec.Datalog.WriteComment "************************************************************"
            TheExec.Datalog.WriteComment "Error on Set_DCVS_alarm, please put a reasonable alarm time."
            TheExec.Datalog.WriteComment "************************************************************"
        End If
        '.CurrentLimit.Sink.FoldLimit.TimeOut = AlarmTime
        .CurrentLimit.Source.FoldLimit.TimeOut = AlarmTime
    End With
    
    If Pin_GP2.value = "" Then
    Else
        With TheHdw.DCVS.Pins(Pin_GP2)
            .Gate = False
            TheHdw.Wait 1 * ms '20170209 Add to wait gate off to avoid set ifold time out error
            .mode = tlDCVSModeVoltage
            .Voltage.Main = 0
            
            If b_AlarmForceBin = False Then
                .Alarm(tlDCVSAlarmAll) = tlAlarmForceFail
            Else
                .Alarm(tlDCVSAlarmAll) = tlAlarmForceBin
            End If
            If CStr(AlarmTime_GP2) = "" Or AlarmTime_GP2 < 0.0001 Then
                TheExec.Datalog.WriteComment "*******************************************************************"
                TheExec.Datalog.WriteComment "Error on Set_DCVS_alarm, please put a reasonable alarm time on GP2."
                TheExec.Datalog.WriteComment "*******************************************************************"
            End If
            '.CurrentLimit.Sink.FoldLimit.TimeOut = AlarmTime
            .CurrentLimit.Source.FoldLimit.TimeOut = AlarmTime_GP2
        End With
    End If
    
    
    Dim CurrentChans As String
     CurrentChans = TheExec.CurrentChanMap 'obtain FT or CP channel map information
    If CurrentChans Like "*FT*" Then
       'AllPowerPinlist = "All_Power_FT"
       'Utility_list = "FT_Utility"
        'increase the foldlimit time out for Vdd_cpu_sram
'        With thehdw.DCVS.pins("vdd_sram")
'           .CurrentLimit.Source.FoldLimit.TimeOut = 0.02
'
'        End With
    Else
     'increase the foldlimit time out for Vdd_cpu_sram
'        With thehdw.DCVS.pins("vdd_cpu_sram")
'           .CurrentLimit.Source.FoldLimit.TimeOut = 0.02
'
'        End With
      ' AllPowerPinlist = "All_Power_CP"
       'Utility_list = "CP_Utility"
    End If

 
    Exit Function
 
errHandler:

                If AbortTest Then Exit Function Else Resume Next
End Function
Public Function KA_start() As Long
    TheHdw.Digital.Patgen.KeepAlive.Flag = cpuB
    TheHdw.Digital.Patgen.KeepAlive.Enable = True
End Function

Public Function KA_end() As Long
    TheHdw.Digital.Patgen.Halt
    Call TheHdw.Digital.Patgen.Continue(0, cpuB)
    TheHdw.Digital.Patgen.KeepAlive.Enable = False
End Function
Public Function Disable_compare(DisableCompare_Pin As PinList)

    TheHdw.Digital.Pins(DisableCompare_Pin).DisableCompare = True
    
    TheExec.Datalog.WriteComment "*************************************************"
    TheExec.Datalog.WriteComment "*Disable Compare Pin:" & DisableCompare_Pin
    TheExec.Datalog.WriteComment "*************************************************"
    
End Function

Public Function Enble_compare(EnableCompare_Pin As PinList)

    TheHdw.Digital.Pins(EnableCompare_Pin).DisableCompare = False

    TheExec.Datalog.WriteComment "*************************************************"
    TheExec.Datalog.WriteComment "*Enable Compare Pin:" & EnableCompare_Pin
    TheExec.Datalog.WriteComment "*************************************************"
    
End Function
Public Function Plot_Profile_on_disk(PinName As PinList, SampleRate As Double, sampleSize As Long, Optional CapSignalName As String = "Capture_signal")
'Plot profiles

    Dim dspw As New DSPWave
    Dim Label As String
    Dim site As Variant
    Dim Pin_Ary() As String
    Dim Pin_Cnt As Long
    Dim ForCount As Long
    Dim ForCount2 As Long
    Dim p As Variant
    Dim FolderName As String
    Dim filename As String
    Dim FirstActiveSite As Long
    Dim RowString As String
    Dim t1, t2 As Double
    Dim DebugLog As Boolean
    Dim dataArray() As String
    FirstActiveSite = -1
    DebugLog = True
    
    On Error GoTo errHandler

    Do While TheHdw.DCVS.Pins(PinName).Capture.IsRunning = True
    Loop

    ' Get the captured samples from the instrument
    Call TheExec.DataManager.DecomposePinList(PinName, Pin_Ary(), Pin_Cnt)

    ReDim dataArray(sampleSize, Pin_Cnt + 1)

    FolderName = "D:\"
    t1 = Timer
    
    For ForCount = 1 To sampleSize
        If ForCount > 1000 And ForCount > sampleSize - 100 Then Exit For 'avoid less capture
        dataArray(ForCount + 1, 1) = (ForCount - 1) / SampleRate * 1000
    Next ForCount

    t2 = Timer
    If DebugLog Then TheExec.Datalog.WriteComment ("    Print First column time : " & (t2 - t1) * 1000 & " ms.")

    ForCount = 0
    For Each p In Pin_Ary
        If TheExec.DataManager.ChannelType(p) <> "N/C" Then
            dspw = TheHdw.DCVS.Pins(p).Capture.Signals(CapSignalName).DSPWave
            For Each site In TheExec.sites
                If TheHdw.DCVS.Pins(p).Meter.mode = tlDCVSMeterCurrent Then
                    Label = "Current Profile for Site: " & site & " " & p
                    filename = FolderName & site & "_Pin_" & p & ".txt"
                Else
                    Label = "Voltage Profile for Site: " & site & " " & p
                    filename = FolderName & site & "_Pin_" & p & ".txt"
                End If
                
                If False Then dspw.Plot Label   'for pliot
                If False Then dspw.FileExport filename, File_txt  'for file export
                If True Then 'for Save All pin Data into output file
                    If FirstActiveSite = -1 Then
                        FirstActiveSite = site
                    End If
                    If site = FirstActiveSite Then
                        t1 = Timer
                        dataArray(1, ForCount + 2) = p
                        For ForCount2 = 1 To sampleSize
                            If ForCount2 > 1000 And ForCount2 > sampleSize - 100 Then Exit For 'avoid less capture
                            dataArray(ForCount2 + 1, ForCount + 2) = dspw.Element(ForCount2)
                        Next ForCount2
                        t2 = Timer
                        If DebugLog Then TheExec.Datalog.WriteComment ("    Plot Profile Pin: " & p & " is done. Execute Time : " & (t2 - t1) * 1000 & " ms.")
                        ForCount = ForCount + 1
                    End If

                End If
            Next site
        End If
    Next p
    
    t1 = Timer
    filename = FolderName & "VIProfile_site" & FirstActiveSite & "_" & t1 \ 1 & ".txt"
    Open filename For Output As #1
    For ForCount = 1 To sampleSize - 100
        RowString = vbNullString
        For ForCount2 = 1 To Pin_Cnt + 1
            RowString = RowString & dataArray(ForCount, ForCount2) & " "
        Next ForCount2
        Print #1, RowString
    Next ForCount
    t2 = Timer
    If DebugLog Then TheExec.Datalog.WriteComment ("    Plot Profile Pin: " & p & " is done. Execute Time : " & (t2 - t1) * 1000 & " ms.")
    
    Close #1
    
    TheExec.Datalog.WriteComment ("    Plot Profile All Pins are done.")
    
    Exit Function
errHandler:
    ErrorDescription ("Plot_Profile_NEW")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function SiteResultCheck()
On Error GoTo errHandler
    Dim funcName As String:: funcName = "SiteResultCheck"

    Dim site As Variant
    Dim b_SitePass As New SiteLong

    For Each site In TheExec.sites
        If TheExec.sites.item(site).result = tlResultPass Then
            b_SitePass = 1
        ElseIf TheExec.sites.item(site).result = tlResultFail Then
            b_SitePass = 0
        Else
            b_SitePass = 1
        End If
    Next site

    TheExec.Flow.TestLimit b_SitePass, 1, 1

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function FreeRunclk_Enable(PortName As String, Optional freq As Double = 0)

    Dim sFuncName As String:: sFuncName = "FreeRunclk_Enable"
    On Error GoTo errHandler
    
    If TheExec.DataManager.instanceName <> "" Then
        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    End If
    
    ' Add New argument to set nWire Frequency 221013
    If glb_TesterType = "Jaguar" Then
        If freq <> 0 Then TheExec.Datalog.WriteComment "This Argument is not support on UltraFlex"
    ElseIf glb_TesterType = "UltraFLEXplus" Then
        gl_nWireFreq = freq
        If InStr(PortName, "_Port") <> 0 Then
            PortName = Replace(PortName, "_Port", "")
        End If
    End If
    
    Call Enable_FRC(PortName, False)
    TheHdw.Wait 0.001
    Exit Function
    
errHandler:
        Call Print_Error_Message(Error_Info, "VBT_LIB_Common", sFuncName)
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
    TheExec.Datalog.WriteComment "print: Power down start, Power pins: " & PowerPins
    '------------------------------------------------------------------'set vt mode , then set digital pin to 0v , at last disconnect
    With TheHdw.Digital.Pins(DisconnectPinList)
        .Levels.DriverMode = tlDriverModeVt
        '.Levels.value(chVt) = 0
        .Disconnect
    End With
    
    TheExec.Datalog.WriteComment "print: Power down digital disconnect, Digital pins: " & DisconnectPinList
    TheExec.Datalog.WriteComment RepeatChr("*", 120)
    '------------------------------------------------------------------
    
    If Not powerDownEnable Then
        Call PowerUpDownClockSeq(pwrDownSeq, False)
        Call PowerUpDownPinSeq(PowerPins, pwrDownSeq, False)
        powerDownEnable = True
    End If
    
    If debugF Then
        If pwrDownSeq.PowerSeq99Log <> "" Then TheExec.Datalog.WriteComment pwrDownSeq.PowerSeq99Log
        If pwrDownSeq.PowerSeqNCLog <> "" Then TheExec.Datalog.WriteComment pwrDownSeq.PowerSeqNCLog
    End If
    
    PowerUpDown_Process WaitConnectTime, debugF, False
    
    ''' Power Down Flag for restart UCSDM
    For Each site In TheExec.sites.Active
        TheExec.sites.item(site).FlagState("F_PowerDown_Flag") = logicTrue
    Next site
    
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
    
    powerUpPins = PowerPins
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    Call Print_Header("Power up sequence")
    TheExec.Datalog.WriteComment "print: Power up start, Power pins: " & PowerPins
    instSet.Initialize PowerPins
    Call instSet.ApplyPower(0, fVoltage)
    '------------------------------------------------------------------disconnect I/O pins
    TheHdw.Digital.Pins(DisconnectPinList).Connect
    TheHdw.Digital.Pins(DisconnectPinList).initState = chInitoff
    
    TheExec.Datalog.WriteComment "print: Power up digital disconnect, Digital pins: " & DisconnectPinList
    TheExec.Datalog.WriteComment RepeatChr("*", 120)
    '------------------------------------------------------------------
    
    If Not powerUpEnable Then
        Call PowerUpDownClockSeq(pwrUpSeq, True)
        Call PowerUpDownIOSeq(ioPinGrp, True)
        Call PowerUpDownPinSeq(PowerPins, pwrUpSeq, True)
        powerUpEnable = True
        End If
    
    If debugF Then
        If pwrUpSeq.PowerSeq99Log <> "" Then TheExec.Datalog.WriteComment pwrUpSeq.PowerSeq99Log
        If pwrUpSeq.PowerSeqNCLog <> "" Then TheExec.Datalog.WriteComment pwrUpSeq.PowerSeqNCLog
                                        End If
                    
    PowerUpDown_Process WaitConnectTime, debugF, True
    TheHdw.Digital.Pins(DisconnectPinList).Disconnect
    Call Print_Footer("Power up sequence")
    
    ' Record which site has do power up
    For Each v_site In TheExec.sites
        powerUpDone(v_site) = True
    Next v_site
    
    If gl_EnableCurrentProfile Or gl_EnableVoltageProfile Then
        DoAll_save = TheExec.RunOptions.DoAll
        OverRide_FailStop = TheExec.RunOptions.OverrideFailStop
        TheExec.RunOptions.DoAll = True
        TheExec.RunOptions.OverrideFailStop = True
        TheExec.Datalog.WriteComment "Enable Word:CurrentProfile = Enable"
        TheExec.Datalog.WriteComment "DoAll = Enable"
        TheExec.Datalog.WriteComment "OverrideFailStop = Enable"
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
Public Function Set_Power_Alarm(Pin_DCVS As String, SourceFoldAlarmTime_DCVS As Double, SinkFoldLevelRatio_DCVS As Double, SinkFoldAlarmTime_DCVS As Double, Pin_DCVI As String, AlarmTime_DCVI As Double, Pin_DCVS_GP2 As String, SourceFlodAlarmTime_GP2 As Double, SinkFoldLevelRatio_GP2 As Double, SinkFoldAlarmTime_GP2 As Double, Optional b_AlarmForceBin As Boolean = False)        'add boolean to set force_bin or force_fail
        'add boolean to set force_bin or force_fail

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
        TheExec.Datalog.WriteComment "*************************************************************************"
        TheExec.Datalog.WriteComment "Error on Set_Power_pins_alarm, please fill in pin name to set alarm time."
        TheExec.Datalog.WriteComment "*************************************************************************"
    End If
    
    If Pin_DCVS <> "" Then
        If (UCase(Pin_DCVS) Like "*CP*=*" Or UCase(Pin_DCVS) Like "*FT*=*") Then
            s_DCVSPin = Select_MeasPin(Pin_DCVS, UCase(currentJobName))
        Else
            s_DCVSPin = Pin_DCVS
        End If
        TheExec.DataManager.DecomposePinList s_DCVSPin, sa_DCVSPinAry(), n_DCVSPinCnt
    End If
    If Pin_DCVI <> "" Then
        If (UCase(Pin_DCVI) Like "*CP*=*" Or UCase(Pin_DCVI) Like "*FT*=*") Then
            s_DCVIPin = Select_MeasPin(Pin_DCVI, UCase(currentJobName))
        Else
            s_DCVIPin = Pin_DCVI
        End If
        TheExec.DataManager.DecomposePinList s_DCVIPin, sa_DCVIPinAry(), n_DCVIPinCnt
    End If
    If Pin_DCVS_GP2 <> "" Then
        If (UCase(Pin_DCVS_GP2) Like "*CP*=*" Or UCase(Pin_DCVS_GP2) Like "*FT*=*") Then
            s_DCVS_GP2 = Select_MeasPin(Pin_DCVS_GP2, UCase(currentJobName))
        Else
            s_DCVS_GP2 = Pin_DCVS_GP2
        End If
        TheExec.DataManager.DecomposePinList s_DCVS_GP2, sa_PinGP2Ary(), n_PinGP2Cnt
    End If
    If TheExec.Flow.enableWord("CurrentProfile") Or TheExec.Flow.enableWord("VoltageProfile") Then
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
                    .CurrentLimit.Sink.FoldLimit.level = SinkFoldLevelRatio_DCVS * .CurrentLimit.Sink.FoldLimit.level.Max
                    .CurrentLimit.Source.FoldLimit.TimeOut = SourceFoldAlarmTime_DCVS
                    ' 20230428  Update for fixing alarm issue at CurrentProfile
                    
                    ' 20230602 Source/Sink set gate off at SetPowerAlarm
                    .CurrentLimit.Sink.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorGateOff
                    .CurrentLimit.Source.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorGateOff
                    If glb_TesterType = "UltraFLEXplus" Then
                        .Alarm(tlDCVSAlarmSinkFoldCurrentLimitLevel) = tlAlarmOff
                        .Alarm(tlDCVSAlarmSourceFoldCurrentLimitLevel) = tlAlarmOff
                    End If
                End With
            Next i
        End If
        '==================================================================
        TheExec.enableWord("HardIP_Autorange") = False
    Else
    
        '==================================================================
        ''' Donan Set_Power_Alarm TTR @William 230710
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
                    TheExec.Datalog.WriteComment "************************************************************"
                    TheExec.Datalog.WriteComment "Error on Set_DCVS_alarm, please put a reasonable alarm time."
                    TheExec.Datalog.WriteComment "************************************************************"
                End If
                '.CurrentLimit.Sink.FoldLimit.TimeOut = AlarmTime_DCVS
                ' 20220113  modify sink limit for negative current
                .CurrentLimit.Sink.FoldLimit.TimeOut = SinkFoldAlarmTime_DCVS
                ' 20220526  modify sink limit for negative current level
                
'                    .CurrentRange = SinkFoldLevelRatio_DCVS * .CurrentLimit.Sink.FoldLimit.level.max  '20221221 Jayden

                .CurrentLimit.Source.FoldLimit.TimeOut = SourceFoldAlarmTime_DCVS
                
                ' 20230602 Source/Sink set gate off at SetPowerAlarm
                .CurrentLimit.Sink.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorGateOff
                .CurrentLimit.Source.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorGateOff
                If glb_TesterType = "UltraFLEXplus" Then
                    .Alarm(tlDCVSAlarmSinkFoldCurrentLimitLevel) = tlAlarmOff
                    .Alarm(tlDCVSAlarmSourceFoldCurrentLimitLevel) = tlAlarmOff
                End If
                
            End With
        
            For i = 0 To n_DCVSPinCnt - 1
                TheHdw.DCVS.Pins(sa_DCVSPinAry(i)).CurrentLimit.Sink.FoldLimit.level = SinkFoldLevelRatio_DCVS * TheHdw.DCVS.Pins(sa_DCVSPinAry(i)).CurrentLimit.Sink.FoldLimit.level.Max
            Next i
            '''
            
''            For i = 0 To PinCnt - 1
''                With TheHdw.DCVS.Pins(DCVSPinAry(i))
''                    .Gate = False
''                    TheHdw.Wait 1 * ms '20170209 Add to wait gate off to avoid set ifold time out error
''                    .mode = tlDCVSModeVoltage
''                    .Voltage.Main = 0
''
''                    If b_AlarmForceBin = False Then
''                        .Alarm(tlDCVSAlarmAll) = tlAlarmForceFail
''                    Else
''                        .Alarm(tlDCVSAlarmAll) = tlAlarmForceBin
''                    End If
''                    If CStr(SourceFoldAlarmTime_DCVS) = "" Or SourceFoldAlarmTime_DCVS < 0.0001 Then
''                        theexec.Datalog.WriteComment "************************************************************"
''                        theexec.Datalog.WriteComment "Error on Set_DCVS_alarm, please put a reasonable alarm time."
''                        theexec.Datalog.WriteComment "************************************************************"
''                    End If
''                    '.CurrentLimit.Sink.FoldLimit.TimeOut = AlarmTime_DCVS
''                    ' 20220113  modify sink limit for negative current
''                    .CurrentLimit.Sink.FoldLimit.TimeOut = SinkFoldAlarmTime_DCVS
''                    ' 20220526  modify sink limit for negative current level
''
'''                    .CurrentRange = SinkFoldLevelRatio_DCVS * .CurrentLimit.Sink.FoldLimit.level.max  '20221221 Jayden
''
''                    .CurrentLimit.Sink.FoldLimit.level = SinkFoldLevelRatio_DCVS * .CurrentLimit.Sink.FoldLimit.level.max
''                    .CurrentLimit.Source.FoldLimit.TimeOut = SourceFoldAlarmTime_DCVS
''
''                    ' 20230602 Source/Sink set gate off at SetPowerAlarm
''                    .CurrentLimit.Sink.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorGateOff
''                    .CurrentLimit.Source.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorGateOff
''                    If glb_TesterType = "UltraFLEXplus" Then
''                        .Alarm(tlDCVSAlarmSinkFoldCurrentLimitLevel) = tlAlarmOff
''                        .Alarm(tlDCVSAlarmSourceFoldCurrentLimitLevel) = tlAlarmOff
''                    End If
''
''                End With
''            Next i

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
                    TheExec.Datalog.WriteComment "*******************************************************************"
                    TheExec.Datalog.WriteComment "Error on Set_DCVS_alarm, please put a reasonable alarm time on GP2."
                    TheExec.Datalog.WriteComment "*******************************************************************"
                End If
                '.CurrentLimit.Sink.FoldLimit.TimeOut = AlarmTime
                ' 20220113  modify sink limit for negative current
                .CurrentLimit.Sink.FoldLimit.TimeOut = SinkFoldAlarmTime_GP2
                ' 20220526  modify sink limit for negative current level
                .CurrentLimit.Source.FoldLimit.TimeOut = SourceFlodAlarmTime_GP2
                ' 20230602 Source/Sink set gate off at SetPowerAlarm
                .CurrentLimit.Sink.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorGateOff
                .CurrentLimit.Source.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorGateOff
                If glb_TesterType = "UltraFLEXplus" Then
                    .Alarm(tlDCVSAlarmSinkFoldCurrentLimitLevel) = tlAlarmOff
                    .Alarm(tlDCVSAlarmSourceFoldCurrentLimitLevel) = tlAlarmOff
                End If
            End With
            For i = 0 To n_PinGP2Cnt - 1
                TheHdw.DCVS.Pins(sa_PinGP2Ary(i)).CurrentLimit.Sink.FoldLimit.level = SinkFoldLevelRatio_GP2 * TheHdw.DCVS.Pins(sa_PinGP2Ary(i)).CurrentLimit.Sink.FoldLimit.level.Max
            Next i
        End If
        
    End If
    
    ' 20210112 temporary add for open Kelvin Alarm bypass for open socket check
    If TheExec.Flow.enableWord("OpenKelvinAlarmOff") = True Then
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
    TheExec.DataManager.DecomposePinList All_Groups_arr(i), PinAry(), PinCnt
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

If isDebugMode Then TheExec.AddOutput "*********ExportUnExistPins Search End*********"



End Function


Public Function License_Mapping()


Dim version_str As String
Dim tester_path As String
Dim File_path As String
Dim OutputFilePath As String
Dim file_num As Long
Dim tmpstr As String
Dim line_info() As String
ReDim line_info(1)
Dim i As Long
Dim j As Long

Dim Memory_dep_tester_index As Long
Dim Memory_dep_tester_end_index As Long
Dim Speed_tester_index As Long
Dim Speed_tester_end_index As Long


Dim Memory_dep_index As Long
Dim Memory_dep_info() As String

Dim Speed_index As Long

Dim LVM_Fab As Long
Dim Speed_Fab As Long
LVM_Fab = 512
Speed_Fab = 200000000

i = 0
j = 0

If TheExec.Flow.enableWord("License_check") = True Then
        gL_License_check = 1
End If

If (gL_License_check = 0) Then
    gL_License_check = gL_License_check + 1
    Exit Function
ElseIf gL_License_check = 1 Then
    gL_License_check = 2
Else
    Exit Function
End If



version_str = Trim(Split(CStr(TheExec.SoftwareVersion), "(")(0))
tester_path = "C:\Program Files (x86)\Teradyne\IG-XL\" & version_str & "\tester\"
File_path = "C:\Program Files (x86)\Teradyne\IG-XL\" & version_str & "\tester\LicenseMappingFile_HSD.txt"

OutputFilePath = Dir(File_path)
file_num = freefile


Do While OutputFilePath <> ""
    Open File_path For Input As #file_num
    
    Do Until EOF(file_num)
        Line Input #1, tmpstr
        ReDim Preserve line_info(i)
            If InStr(tmpstr, "HSD_MemoryDepth Licenses Assigned:") <> 0 Then
                Memory_dep_tester_index = i
            End If
            
            If InStr(tmpstr, " HSD_MemoryDepth Licenses Assigned:") <> 0 Then
                Memory_dep_tester_index = i + 2
            ElseIf InStr(tmpstr, " HSD_Speed Licenses Assigned:") <> 0 Then
                Speed_tester_index = i + 2
            ElseIf InStr(tmpstr, "VM DEPTHS") <> 0 Then
                Memory_dep_index = i + 2
            ElseIf InStr(tmpstr, "DATARATES") <> 0 Then
                Speed_index = i + 2
            End If
            
            line_info(i) = tmpstr
            i = i + 1
    Loop
    
    For i = Memory_dep_tester_index + 1 To Memory_dep_tester_index + 10
        If line_info(i) = "" Then Memory_dep_tester_end_index = i - 1:: Exit For
    Next i
    
    For i = Speed_tester_index + 1 To Speed_tester_index + 10
        If line_info(i) = "" Then Speed_tester_end_index = i - 1:: Exit For
    Next i
    
    
    
    
    TheExec.Datalog.WriteComment "======== License information start ========" & Chr(10)

    If (Memory_dep_index <> 0 And Speed_index <> 0) Then
        
        TheExec.Datalog.WriteComment "HSD_MemoryDepth Licenses Assigned:"
        For i = Memory_dep_tester_index To Memory_dep_tester_end_index
            TheExec.Datalog.WriteComment "QUANTITY:" & CStr(Trim(Split(line_info(i), Chr(9))(0))) & "   ,ENABLE LEVEL(M license units):" & CStr(Trim(Split(line_info(i), Chr(9))(1)))
        Next i
        TheExec.Datalog.WriteComment ""
        
        
        TheExec.Datalog.WriteComment "HSD_Speed Licenses Assigned:"
        For i = Speed_tester_index To Speed_tester_end_index
            TheExec.Datalog.WriteComment "QUANTITY:" & CStr(Trim(Split(line_info(i), Chr(9))(0))) & "   ,ENABLE LEVEL(Mbps):" & CStr(Trim(Split(line_info(i), Chr(9))(1)))
        Next i
        TheExec.Datalog.WriteComment ""

        TheExec.Flow.TestLimit resultVal:=CDbl(Split(line_info(Memory_dep_index), Chr(9))(1)) + 20, hiVal:=LVM_Fab, unit:=unitNone, Tname:="Program LVM", ForceResults:=tlForceNone
        TheExec.Flow.TestLimit resultVal:=CDbl(Split(line_info(Speed_index), Chr(9))(0)) * 1000000, hiVal:=Speed_Fab, unit:=unitHz, Tname:="Program DataRate", scaletype:=scaleMega, ForceResults:=tlForceNone
    
    Else

    'theexec.AddOutput "Please check 'theexec.Licenses.GenerateLicenseRequirementsFile' is set be true and run on-line "
    TheExec.Datalog.WriteComment "Please check 'theexec.Licenses.GenerateLicenseRequirementsFile' is set be true and run on-line "
    End If
    
    TheExec.Datalog.WriteComment "======== License information end ========" & Chr(10)
    
    
    
'Debug.Print OutputFilePath
    Close #file_num
    OutputFilePath = Dir()
Loop

End Function

''20190604AddFunction
Public Function HIP_Init_Datalog_Setup() ''TER190530
    
    TheExec.Datalog.Setup.DatalogSetup.PartResult = True
    TheExec.Datalog.Setup.DatalogSetup.XYCoordinates = True
    TheExec.Datalog.Setup.DatalogSetup.DisableChannelNumberInPTR = True 'disable channel name to stdf, PE's datalog request -- 131225, chihome
    TheExec.Datalog.Setup.DatalogSetup.OutputWidth = 0

''' 20210709 for datalog format alignment
'''    TheExec.Datalog.Setup.Shared.ascii.Columns.EnableCustomWidths = True
'''    TheExec.Datalog.Setup.Shared.ascii.Columns.Parametric.TestName.Width = 60
'''    TheExec.Datalog.Setup.Shared.ascii.Columns.Functional.TestName.Width = 60
'''    TheExec.Datalog.Setup.Shared.ascii.Columns.Functional.pattern.Width = 70

    TheExec.Datalog.ApplySetup  'must need to apply after datalog setup
    
End Function

Public Function Common_UnitTest()

On Error GoTo errHandler
    
    Dim idx As Variant
    Dim long_ As Long
    Dim dec_result_ As Long
    
    Dim dec_ As Double
    Dim dec_ary_() As Double
    Dim dec_result_double_ As Double
    
    Dim bin_ As String
    Dim bin_ary_() As String
    Dim bin_result_ As String
    
    bin_ = "01001111"
    dec_result_ = 79
    long_ = Bin2Dec(bin_)
    If long_ = dec_result_ Then
        TheExec.Datalog.WriteComment "Binary " & bin_ & " Transform and get the same result " & "Decimal " & long_ & " by Function Bin2Dec"

    Else
        TheExec.Datalog.WriteComment "Binary " & bin_ & " Transform but do not get the same result " & "Decimal " & long_ & " by Function Bin2Dec"
        
    End If
    
    dec_result_ = 242
    long_ = Bin2Dec_rev(bin_)
    If long_ = dec_result_ Then
        TheExec.Datalog.WriteComment "Binary " & bin_ & " Transform and get the same result " & "Decimal " & dec_ & " by Function Bin2Dec_rev"

    Else
        TheExec.Datalog.WriteComment "Binary " & bin_ & " Transform but do not get the same result " & "Decimal " & long_ & " by Function Bin2Dec_rev"
        
    End If
    
    dec_result_ = 242
    dec_ = Bin2Dec_rev_Double(bin_)
    If dec_ = dec_result_ Then
        TheExec.Datalog.WriteComment "Binary " & bin_ & " Transform and get the same result " & "Double " & dec_ & " by Function Bin2Dec_rev_Double"
      
    Else
        TheExec.Datalog.WriteComment "Binary " & bin_ & " Transform but do not get the same result " & "Double " & dec_ & " by Function Bin2Dec_rev_Double"
        
    End If
    TheExec.Datalog.WriteComment "Binary " & bin_ & " Transform to " & "Decimal " & dec_ & " by Function Bin2Dec_rev_Double"
    
    dec_result_double_ = 0.9453125
    dec_ = Bin2Dec_rev_Fractional(bin_)
    If dec_ = dec_result_double_ Then
        TheExec.Datalog.WriteComment "Binary " & bin_ & " Transform and get the same result " & "Fractional " & dec_ & " by Function Bin2Dec_rev_Fractional"

    Else
        TheExec.Datalog.WriteComment "Binary " & bin_ & " Transform but do not get the same result " & "Fractional " & dec_ & " by Function Bin2Dec_rev_Fractional"
        
    End If
    TheExec.Datalog.WriteComment "Binary " & bin_ & " Transform to " & "Decimal " & dec_ & " by Function Bin2Dec_rev_Fractional"
    
    bin_result_ = "011110010"
    bin_ = Dec2BinStr32Bit(8, long_)
    If bin_ = bin_result_ Then
        TheExec.Datalog.WriteComment "Decimal " & long_ & " Transform and get the same result " & "Binary " & bin_ & " by Function Dec2BinStr32Bit"

    Else
        TheExec.Datalog.WriteComment "Decimal " & long_ & " Transform but do not get the same result " & "Binary " & bin_ & " by Function Dec2BinStr32Bit"
        
    End If

    
Exit Function
errHandler:
        If isDebugMode Then TheExec.AddOutput "Error in the VBT Common_UnitTest"
        TheExec.Datalog.WriteComment "Error in the VBT Common_UnitTest"
        If AbortTest Then Exit Function Else Resume Next
End Function

Public Function auto_ReadHandlerOCRData() 'VBT function

    On Error GoTo errHandler
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
        SiteCount = TheExec.sites.Existing.Count
        ReDim OCR_StringBySite_Sort(SiteCount - 1) As String
        For Each site In TheExec.sites
            OCR_StringBySite_Sort(site) = vbNullString
        Next site
    End If
    
    TheExec.Datalog.WriteComment ""
    
    'To write OCR data in STDF file
    For Each site In TheExec.sites
        TheExec.Datalog.WriteComment ("<@OCR_Data=" & site & "|" & OCR_StringBySite_Sort(site) & ">")
    Next site
    
    'To check OCR is available, string or something else but not empty
    For Each site In TheExec.sites
        OCR_available_chk = 0
        If (TheExec.TesterMode = testModeOffline) Then
            OCR_available_chk = 1
        Else
            If IsEmpty(OCR_StringBySite_Sort(site)) Or OCR_StringBySite_Sort(site) = "" Then
                OCR_available_chk = 0
            Else
                OCR_available_chk = 1 '1 means OCR is not ""
            End If
        End If
        TheExec.Flow.TestLimit OCR_available_chk, 1, 1, Tname:="OCR_Checking"
    Next site
    
    
    TheExec.Datalog.WriteComment ""

    Exit Function
    
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "auto_ReadHandlerOCRData") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next

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
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Init_Datalog_Setup"
    
    With TheExec.Datalog
        .WriteComment "----------------------------------------"
        .WriteComment "Reset to default datalog format!!!!!!!!!"
        .WriteComment "----------------------------------------"
    End With
    
    With TheExec.Datalog.Setup.DatalogSetup
        .PartResult = True                      'Default: TRUE
        .XYCoordinates = True                   'Default: TRUE
        .DisableChannelNumberInPTR = True       'Default: FALSE 'disable channel name to stdf, PE's datalog request -- 131225, chihome
        .OutputWidth = 0                        'Default: 0
        .DisableInstanceNameInPTR = False       'Default: FALSE '20210519 add
        .DisablePinNameInPTR = False            'Default: FALSE '20210519 add
        .PTR_InstanceNameIsTINameOnly = True   'Default: FALSE '20210519 add

    End With
    
    With TheExec.Datalog.Setup.Shared.ascii.Columns
        .EnableCustomWidths = True
        .Parametric.TestName.Width = 200
        .Parametric.measured.Width = 16
        .Parametric.number.Width = 12
        .Parametric.pin.Width = 30
        .Functional.TestName.Width = 200
        .Functional.Pattern.Width = 100
        .Functional.number.Width = 12
        .Functional.Cycle.Width = 9
        
    End With
    
    TheExec.Datalog.ApplySetup  'must need to apply after datalog setup
    
    Exit Function
    
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Init_Datalog_Setup") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next

End Function



''------------------------------------------------------------
''20210504
''Customize datalog format
''------------------------------------------------------------
Public Function Customize_Datalog_Setup(BlockName As String, _
                                        Optional P_TestNameWidth As Long = 200, _
                                        Optional P_MeasuredWidth As Long = 16, _
                                        Optional P_NumberWidth As Long = 12, _
                                        Optional P_PinWidth As Long = 30, _
                                        Optional F_TestNameWidth As Long = 200, _
                                        Optional F_PatternWidth As Long = 100, _
                                        Optional F_NumberWidth As Long = 12, _
                                        Optional F_CycleWidth As Long = 9) As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Custom_Datalog_Setup"
    
    With TheExec.Datalog
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

    With TheExec.Datalog.Setup.DatalogSetup
        .PartResult = True
        .XYCoordinates = True
        .DisableChannelNumberInPTR = True 'disable channel name to stdf, PE's datalog request -- 131225, chihome
        .OutputWidth = 0
        .DisableInstanceNameInPTR = False
        .DisablePinNameInPTR = False
        '.PTR_InstanceNameIsTINameOnly = False

    End With
    With TheExec.Datalog.Setup.Shared.ascii.Columns
        .EnableCustomWidths = True
        .Parametric.TestName.Width = P_TestNameWidth
        .Parametric.measured.Width = P_MeasuredWidth
        .Parametric.number.Width = P_NumberWidth
        .Parametric.pin.Width = P_PinWidth
        .Functional.TestName.Width = F_TestNameWidth
        .Functional.Pattern.Width = F_PatternWidth
        .Functional.number.Width = F_NumberWidth
        .Functional.Cycle.Width = F_CycleWidth
    End With
    
    TheExec.Datalog.ApplySetup  'must need to apply after datalog setup
    
    Exit Function
    
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Customize_Datalog_Setup") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function Print_Version()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    TheExec.Datalog.WriteComment "-----------------------------------------------"
    TheExec.Datalog.WriteComment "      DATE : " & CStr(DATE_VER) & "   SVN_VER : " & CStr(SVN_VER) & "     "
    TheExec.Datalog.WriteComment "-----------------------------------------------"

        TheExec.Datalog.WriteComment "CZ            : " & CStr(CZ_VER)
        TheExec.Datalog.WriteComment "COMMON        : " & CStr(COMMON_VER)
        TheExec.Datalog.WriteComment "DC            : " & CStr(DC_VER)
        TheExec.Datalog.WriteComment "DIGITAL       : " & CStr(DIGITAL_VER)
        TheExec.Datalog.WriteComment "EFUSE         : " & CStr(EFUSE_VER)
        TheExec.Datalog.WriteComment "HIP           : " & CStr(HIP_VER)
        TheExec.Datalog.WriteComment "MBIST         : " & CStr(MBIST_VER)
        TheExec.Datalog.WriteComment "SPIROM        : " & CStr(SPIROM_VER)
        TheExec.Datalog.WriteComment "VDDBINNING    : " & CStr(VDDBINNING_VER)
        TheExec.Datalog.WriteComment "RF_FUNC       : " & CStr(RF_FUNC_VER)
        TheExec.Datalog.WriteComment "RF_DVDC       : " & CStr(RF_DVDC_VER)
        TheExec.Datalog.WriteComment "LCD_OTP       : " & CStr(LCD_OTP_VER)
        TheExec.Datalog.WriteComment "LCD_DVDC      : " & CStr(LCD_DVDC_VER)
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Print_Version") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
Public Function GetDownSampleRatio(InstrType As String, ArySize As Integer) As Double
    Dim divide As Double
    Dim PinNumEachGroup As Integer: PinNumEachGroup = 10
    
   
        divide = Ceiling(ArySize / PinNumEachGroup)
        If divide >= 6 Then divide = 6
        
        If UCase(InstrType) = "UVS" Then
            If divide >= 1 Then
                GetDownSampleRatio = 2
            Else
                GetDownSampleRatio = 1
            End If
        Else
            GetDownSampleRatio = 2 ^ divide
        End If
  
End Function

Public Function show_FRC() As Long
    On Error GoTo errHandler
    Dim key As Variant
    For Each key In gl_nWireFreq_Value_Dict.keys
        TheExec.Datalog.WriteComment "********** freerunning clock pin = " & key & ", clock =" & gl_nWireFreq_Value_Dict(key) & " Hz "
    Next
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "show_FRC")
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Sub FlagSetting()
    ENG_SweepPin = False
End Sub

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
        TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD    'clean all cpu flag
        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    End If
    
    Call CRWT.Initialize(TunePins)
    If LoadTuneResult Then Call CRWT.LoadCRWTResultToDict
    If Pattern <> "" Then
        Call TheHdw.Patterns(Pattern).start
        Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)  'Meas during CPUA loop
    
        For i = 0 To UBound(sq_array)
            If sq_array(i) <> "I" Then
                TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD    'clean all cpu flag
                TheHdw.Wait 0.01
            Else
                Call CRWT.FixedVCR_WaitTime(TunePinVoltages, CurrentRanges, AccuracyRatio, NoRunBelowCurrentRange, SmoothCnt, WriteCurrentProfileFile)
                TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD    'clean all cpu flag
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
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
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
        TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD    'clean all cpu flag
        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    End If
    
    Call CRWT.Initialize(TunePins)
    If LoadTuneResult Then Call CRWT.LoadCRWTResultToDict
    If Pattern <> "" Then
        Call TheHdw.Patterns(Pattern).start
        Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)  'Meas during CPUA loop
    
        For i = 0 To UBound(sq_array)
            If sq_array(i) <> "I" Then
                TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD    'clean all cpu flag
                TheHdw.Wait 0.01
            Else
                Call CRWT.LoopVCR_WaitTime(StartVoltage, EndVoltage, StepVoltage, AccuracyRatio, NoRunBelowCurrentRange, SmoothCnt, WriteCurrentProfileFile)
                TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD    'clean all cpu flag
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
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

' RAK and IDS binout used
Public Function onProgramStartedBinOutFunction()
    Call HardIP_RAK_Init
    Call ParseIDSMappingTable(True)
End Function
Public Function Setup_EVS_Output_Format(Optional Apply_EVS_Format As Boolean = True)
Dim funcName As String:: funcName = "Setup_EVS_Output_Format"
On Error GoTo errHandler
    If Apply_EVS_Format Then
        TheExec.Datalog.Setup.Shared.ascii.Columns.EnableCustomWidths = True
        TheExec.Datalog.Setup.Shared.ascii.Columns.Parametric.TestName.Width = 70
        TheExec.Datalog.Setup.Shared.ascii.Columns.Functional.TestName.Width = 75
        TheExec.Datalog.Setup.Shared.ascii.Columns.Functional.Pattern.Width = 75
        TheExec.Datalog.ApplySetup
    Else
        TheExec.Datalog.Setup.Shared.ascii.Columns.EnableCustomWidths = True
        TheExec.Datalog.Setup.Shared.ascii.Columns.Parametric.TestName.Width = 200
        TheExec.Datalog.Setup.Shared.ascii.Columns.Functional.TestName.Width = 200
        TheExec.Datalog.Setup.Shared.ascii.Columns.Functional.Pattern.Width = 100
        TheExec.Datalog.ApplySetup
    End If
    
Exit Function
errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Setup_ATPG_Output_Format(Optional Apply_ATPG_Format As Boolean = True)
Dim funcName As String:: funcName = "Setup_ATPG_Output_Format"
On Error GoTo errHandler
    If Apply_ATPG_Format Then
        TheExec.Datalog.Setup.Shared.ascii.Columns.EnableCustomWidths = True
        TheExec.Datalog.Setup.Shared.ascii.Columns.Functional.TestName.Width = 75
        TheExec.Datalog.Setup.Shared.ascii.Columns.Functional.Pattern.Width = 75
        TheExec.Datalog.Setup.Shared.ascii.Columns.Parametric.TestName.Width = 75
        TheExec.Datalog.ApplySetup
    Else
        Init_Datalog_Setup
    End If
    
Exit Function
errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Print_Footer_TMPS_MON(PrintInfo As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    TheExec.Datalog.WriteComment "<TMPS_TestInsatnce_" & glb_TestInstance & "_End>"
    TheExec.Datalog.WriteComment "******************************"
    TheExec.Datalog.WriteComment "*print: " & PrintInfo & " end*"
    TheExec.Datalog.WriteComment "******************************"
    
    
'    If InStr(PrintInfo, "SOC_CHAIN_") = 0 And InStr(PrintInfo, "SOC_SAA_") = 0 And InStr(PrintInfo, "SOC_TD_") = 0 _
'    And InStr(PrintInfo, "CPU_CHAIN_") = 0 And InStr(PrintInfo, "CPU_SAA_") = 0 And InStr(PrintInfo, "CPU_TD_") = 0 _
'    And InStr(PrintInfo, "GFX_CHAIN_") = 0 And InStr(PrintInfo, "GFX_SAA_") = 0 And InStr(PrintInfo, "GFX_TD_") = 0 Then
'        ''20210504 reset the datalog format
'        Init_Datalog_Setup
'    End If

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Print_Footer") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
Public Function Print_Header_TMPS_MON(PrintInfo As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    TheExec.Datalog.WriteComment "********************************"
    TheExec.Datalog.WriteComment "*print: " & PrintInfo & " start*"
    TheExec.Datalog.WriteComment "********************************"
    
    If glb_TestInstance = "" Then
        glb_TestInstance_TMPS = "dummy"
    Else
        glb_TestInstance_TMPS = glb_TestInstance
    End If
    
    TheExec.Datalog.WriteComment "<TMPS_TestInsatnce_" & glb_TestInstance_TMPS & "_Start>"

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common", "Print_Header") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Init_IO_State()
Dim funcName As String:: funcName = "Init_IO_State"
On Error GoTo errHandler

    TheHdw.Digital.Pins("All_Digital").initState = chInitLo
    
Exit Function
errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
End Function

Public Function TMPS_MON_Test_Control(WaitTime As Double, GLBIndex As String, flagName As String, Optional isApplyPower As Boolean = False)
    Dim EndFlow As Boolean
    EndFlow = True
    For Each site In TheExec.sites
        If (TheExec.sites.item(site).FlagState(flagName) = logicTrue) Then
            EndFlow = False
            If isApplyPower Then TheHdw.PinLevels.ApplyPower
            TheHdw.Wait WaitTime
            TheExec.Datalog.WriteComment "Add Wait time = " & CStr(WaitTime) & " s"
            Exit For
        End If
    Next site

    If EndFlow = True Then
        For Each site In TheExec.sites
            TheExec.sites.item(site).FlagState("F_TMPS_End") = logicTrue
        Next site
        If glb_TMPS_count_int > 0 Then
            TheExec.Datalog.WriteComment "*************************************************"
            TheExec.Datalog.WriteComment "  MTRTSNS excute  " + CStr((glb_TMPS_count_int) * WaitTime) + "  (s) for cooling !!"
            TheExec.Datalog.WriteComment "*************************************************"
            
        End If
    End If
End Function

Function Handler2DIDs(IDString As String)
    On Error GoTo errHandler
    
'    Dim BarcodeStrInfor() As String
'    Dim BarcodeHeader As String
'    Dim i As Long

    
    'Dim s As String
    'Dim n As Integer
    'n = freefile()
    'Open "C:\Ian\2DID.txt" For Append As n
    'IDString = "Barcode:0,1,2,3,4,5,6,7,8,9;0"
    'MsgBox ("Here is the 2DID from driver" & vbCrLf & IDString)
    
    If IDString <> "" Then
        BarcodeString = IDString
        'BarcodeHeader = Split(Replace(IDString, ";", ""), ":")(0)
        'BarcodeStrInfor = Split(Split(Replace(IDString, ";", ""), ":")(1), ",")
'        If BarcodeHeader <> "" Then
'            For i = 0 To UBound(BarcodeStrInfor)
'                If BarcodeStrInfor(i) <> "0" Then
'                    TheExec.Datalog.WriteComment " Site = " & UBound(BarcodeStrInfor) - i & ", Barcode = " & BarcodeStrInfor(i)
'                    'Debug.Print " Site = " & UBound(BarcodeStrInfor) - i & ", Barcode = " & BarcodeStrInfor(i)
'                    'Print #n, " Site = " & UBound(BarcodeStrInfor) - i & ", Barcode = " & BarcodeStrInfor(i)
'                End If
'            Next i
'            TheExec.Datalog.WriteComment "---------Capture 2DID Information Complete -------------"
'
'        Else
'            TheExec.Datalog.WriteComment "Please check the IC has the 2DID information"
'        End If
    Else
        TheExec.Datalog.WriteComment "IDString = Empty"
    End If
    
    'Close #n
    Exit Function
    
errHandler:
     If AbortTest Then Exit Function Else Resume Next
    
    
End Function



Public Function GetBarcode()
    
'    Dim start_time As Double
'    Dim end_time As Double
'    Dim test_time As Double
'    start_time = TheExec.Timer
   
    Dim BarcodeStrInfor() As String
    Dim BarcodeHeader As String
    Dim i As Long
    Dim BarcodeSiteVariant As New SiteVariant
    Dim IEDAInputStr As String
    Dim SiteCount As Integer
    SiteCount = TheExec.sites.Existing.Count
    ReDim IEDAInputStrArr(SiteCount - 1) As String
    'theexec.sites.Active
    If BarcodeString <> "" Then
        BarcodeHeader = Split(Replace(BarcodeString, ";", ""), ":")(0)
        BarcodeStrInfor = Split(Split(Replace(BarcodeString, ";", ""), ":")(1), ",")
        If BarcodeHeader <> "" Then
            
            '20240502 Mason - Clear Handler2DID Registry
            Call IEDA_Initialize(IEDAInputStr)
            
            TheExec.Datalog.WriteComment "---------Capture 2DID Information Start -------------"
            TheExec.Datalog.WriteComment "BarcodeString: " & BarcodeString
            For i = 0 To UBound(BarcodeStrInfor)
                If BarcodeStrInfor(i) <> "0" Then
                    TheExec.Datalog.WriteComment " Site = " & UBound(BarcodeStrInfor) - i & ", Barcode = " & BarcodeStrInfor(i)
                    BarcodeSiteVariant(UBound(BarcodeStrInfor) - i) = BarcodeStrInfor(i)
                    '20240502 Mason - Store Barcode to array by site
                    IEDAInputStrArr(UBound(BarcodeStrInfor) - i) = BarcodeStrInfor(i)
                End If
            Next i
            
            '20240502 Mason - String processing
            For i = 0 To UBound(IEDAInputStrArr)
                If i = UBound(IEDAInputStrArr) Then
                    IEDAInputStr = IEDAInputStr & IEDAInputStrArr(i)
                Else
                    IEDAInputStr = IEDAInputStr & IEDAInputStrArr(i) & ","
                End If
            Next i
            '20240502 Mason - Save value to registry
            Call IEDA_SaveRegistry(IEDAInputStr, "Handler2DID")
            
            '20240502 Mason - Print to text file----------------
'            Dim Current_Dirctory As String
'            Dim Folder_Name As String
'
'            Folder_Name = "BarCodeData" & "_" & right("0" & CStr(Month(Now)), 2) & right("0" & CStr(Day(Now)), 2) & right("0" & CStr(Hour(Now)), 2) & right("0" & CStr(Minute(Now)), 2)
'            Current_Dirctory = CurDir & "\" & Folder_Name
'
'            Dim TempStr As String
'            TempStr = Current_Dirctory & "\"
'            Dim fso As New FileSystemObject
'            If Dir(TempStr, vbDirectory) = Empty Then
'                MkDir TempStr
'            End If
'            fso.CreateTextFile TempStr & "Barcode.txt", True
'            Dim s As String
'            Dim n As Integer
'            n = FreeFile()
'            Open TempStr & "Barcode.txt" For Append As n
'            Print #n, IEDAInputStr
'            Close #n
            '20240502 Mason - Print to text file----------------
            
            TheExec.Datalog.WriteComment "Handler2DID: " & IEDAInputStr
            
            '20231219 Mason - Add test limit to print the barcode to stdf
            For Each site In TheExec.sites
                TheExec.Flow.TestLimit 0, Tname:=BarcodeSiteVariant
            Next
            
            TheExec.Datalog.WriteComment "---------Capture 2DID Information Complete -------------"
            
            BarcodeString = ""
        Else
            TheExec.Datalog.WriteComment "Please check the IC has the 2DID information"
        End If
    Else
        TheExec.Datalog.WriteComment "IDString = Empty"
    End If
    
    
'    end_time = TheExec.Timer
'    test_time = end_time - start_time
'    TheExec.Datalog.WriteComment "Test time=" & test_time
    
 End Function

Public Function EnableWord_Initialization()
Dim funcName As String:: funcName = "EnableWord_Initialization"
On Error GoTo errHandler

    glb_isSFC_Enabled = TheExec.enableWord("Enable_SFC")
    
Exit Function
errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
End Function
