Attribute VB_Name = "LIB_Common_DSSC_HRAM"
Option Explicit
'Revision History:
'V0.0 initial bring up
'*****************************************
'******                         DSSC******
'*****************************************

''''20170920 update for the mutiple Src cases as CFG_RAW+MSP
Public Function DSSC_SetupDigSrcWave(patt As String, DigSrcPin As PinList, SignalName As String, SegmentSize As Long, WaveDefArray() As Long)
On Error GoTo errHandler
    Dim funcName As String:: funcName = "DSSC_SetupDigSrcWave"
    
    'store efuse program bit into a DSP wave
    Dim InWave As New DSPWave
    Dim site As Variant
    Dim WaveDef As String
    
    ''WaveDef = "WaveDef" ''''was
    InWave.data = WaveDefArray
    site = TheExec.sites.siteNumber
    ''''20170920 <NOTICE> if multiple apply this function call/sequence to avoid the following SrcWave to overwrite the previous one
    WaveDef = "WaveDef_" + SignalName + "_" & site
    TheHdw.patterns(patt).Load
    TheExec.WaveDefinitions.CreateWaveDefinition WaveDef, InWave, True
    TheHdw.DSSC.pins(DigSrcPin).Pattern(patt).Source.Signals.Add SignalName
    With TheHdw.DSSC.pins(DigSrcPin).Pattern(patt).Source.Signals(SignalName)
        .WaveDefinitionName = WaveDef
        .SampleSize = SegmentSize
        .Amplitude = 1
        .LoadSamples
        If glb_TesterType = "Jaguar" Then .LoadSettings
    End With
    TheHdw.DSSC.pins(DigSrcPin).Pattern(patt).Source.Signals.DefaultSignal = SignalName
    
Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
End Function

'*****************************************
'******                         HRAM******
'*****************************************
Function DatalogHRAMVecNum(excludePreTrig As Boolean, StvNum As Long, DataOutputPins As PinList)

    Dim maxIdx As Long
    Dim idx As Long
    Dim preTrig As Long
    preTrig = 0
    Dim PinData As New PinListData
    Dim Binary_String As String

    '=== Simulated Data ===
    If (TheExec.TesterMode = testModeOffline) Then
       ' PinData = thehdw.Digital.pins(DataOutputPins).HRAM.PinData(0, 1, StvNum)
        With TheHdw.Digital.hram
            If excludePreTrig = True Then
                preTrig = .PreTrigCycles
            End If
            'TheExec.Datalog.WriteComment "  Pattern: " + CStr(.PatGenInfo(idx, pgPattern))
    
            Binary_String = vbNullString
            maxIdx = StvNum - 1
            For idx = preTrig To maxIdx
                TheExec.Datalog.WriteComment "      Hram index:" + CStr(idx) + _
                " Vector number: " + CStr(idx) + "   DUT state : " + "0" '+ PinData.pins(DataOutputPins).Value(0)(idx)
                Binary_String = Binary_String + "0"
            Next
        End With
    Else

  
        PinData = TheHdw.Digital.pins(DataOutputPins).hram.PinData(0, 1, StvNum)
        With TheHdw.Digital.hram
            If excludePreTrig = True Then
                preTrig = .PreTrigCycles
            End If
            TheExec.Datalog.WriteComment "  Pattern: " + CStr(.PatGenInfo(idx, pgPattern))
    
            Binary_String = vbNullString
            maxIdx = .CapturedCycles - 1
            For idx = preTrig To maxIdx
                TheExec.Datalog.WriteComment "      Hram index:" + CStr(idx) + _
                " Vector number: " + CStr(.PatGenInfo(idx, pgVector)) + "   DUT state : " + PinData.pins(DataOutputPins).value(0)(idx)
                If LCase(PinData.pins("tdo").value(0)(idx)) = "l" Then
                    Binary_String = Binary_String + "0"
                Else
                    Binary_String = Binary_String + "1"
                End If
            Next
        End With
    

    End If
    
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment vbTab + "Binary Code = " + Binary_String

End Function
Public Function Hram_Trig_Setup(TrigType As TrigType, CaptType As CaptType)
    
    Dim WaitForEvent As Boolean
    Dim preTrigCnt As Integer
    Dim stopFull As Boolean

    WaitForEvent = False
    preTrigCnt = 0
    stopFull = True
    With TheHdw.Digital.hram
        .SetTrigger TrigType, WaitForEvent, preTrigCnt, stopFull
        .CaptureType = CaptType  'the vector to be captured only at stv micro-code
        .size = 0
    End With

End Function


Public Function DSSC_PreSetupDigSrcWave(patt As String, DigSrcPin As PinList, SignalName As String)
On Error GoTo errHandler
 ''' 20211214 Digital Source Pin Command "Start With Lebel"
    Dim funcName As String:: funcName = "DSSC_PreSetupDigSrcWave"
    Dim tempVarArray As Variant
    tempVarArray = TheHdw.DSSC.pins(DigSrcPin).Pattern(patt).Source.Labels.list
    If tempVarArray(0) = "" Then
        SignalName = "Meas_Src"
    Else
        SignalName = tempVarArray(0)
    End If
    
    TheHdw.patterns(patt).Load
    TheHdw.DSSC.pins(DigSrcPin).Pattern(patt).Source.Signals.Add SignalName

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_DSSC_HRAM", "DSSC_PreSetupDigSrcWave") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function DSSC_PostSetupDigSrcWave(patt As String, DigSrcPin As PinList, SignalName As String)
On Error GoTo errHandler
    
    Dim funcName As String:: funcName = "DSSC_PostSetupDigSrcWave"
    
    With TheHdw.DSSC.pins(DigSrcPin).Pattern(patt).Source.Signals(SignalName)
        .Amplitude = 1
        .LoadSamples
        If LCase(glb_TesterType) = "jaguar" Then
            .LoadSettings
        End If
    End With
                                                                                                                                                               
    TheHdw.DSSC.pins(DigSrcPin).Pattern(patt).Source.Signals.DefaultSignal = SignalName

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_DSSC_HRAM", "DSSC_PostSetupDigSrcWave") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function DSSC_BodySetupDigSrcWave(patt As String, DigSrcPin As PinList, SignalName As String, SegmentSize As Long, WaveDefArray() As Long)
On Error GoTo errHandler
    Dim funcName As String:: funcName = "DSSC_BodySetupDigSrcWave"
    
    'store efuse program bit into a DSP wave
    Dim InWave As New DSPWave
    Dim site As Variant
    Dim WaveDef As String
    
    ''WaveDef = "WaveDef" ''''was
    InWave.data = WaveDefArray
    
    site = TheExec.sites.siteNumber
    
    ''''20170920 <NOTICE> if multiple apply this function call/sequence to avoid the following SrcWave to overwrite the previous one
    WaveDef = "WaveDef_" + SignalName + "_" & site
    
    TheExec.WaveDefinitions.CreateWaveDefinition WaveDef, InWave, True
    TheHdw.DSSC.pins(DigSrcPin).Pattern(patt).Source.Signals.Add SignalName
    With TheHdw.DSSC.pins(DigSrcPin).Pattern(patt).Source.Signals(SignalName)
        .WaveDefinitionName = WaveDef
        .SampleSize = SegmentSize
    End With

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_DSSC_HRAM", "DSSC_BodySetupDigSrcWave") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

