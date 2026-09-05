Attribute VB_Name = "LIB_Common_DSSC_HRAM"
Option Explicit
'Revision History:
'V0.0 initial bring up
'*****************************************
'******                         DSSC******
'*****************************************

Public Function DSSC_PreSetupDigSrcWave(patt As String, digSrcPin As PinList, SignalName As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
 ''' 20211214 Digital Source Pin Command "Start With Lebel"
    Dim funcName As String:: funcName = "DSSC_PreSetupDigSrcWave"
    Dim tempVarArray As Variant
    tempVarArray = TheHdw.DSSC.Pins(digSrcPin).Pattern(patt).Source.Labels.list
    If tempVarArray(0) = "" Then
        SignalName = "Meas_Src"
    Else
        SignalName = tempVarArray(0)
    End If
    
    TheHdw.Patterns(patt).Load
    TheHdw.DSSC.Pins(digSrcPin).Pattern(patt).Source.Signals.Add SignalName

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_DSSC_HRAM", "DSSC_PreSetupDigSrcWave") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function DSSC_PostSetupDigSrcWave(patt As String, digSrcPin As PinList, SignalName As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    
    Dim funcName As String:: funcName = "DSSC_PostSetupDigSrcWave"
    
    With TheHdw.DSSC.Pins(digSrcPin).Pattern(patt).Source.Signals(SignalName)
        .Amplitude = 1
        .LoadSamples
        If LCase(glb_TesterType) = "jaguar" Then
            .LoadSettings
        End If
    End With
                                                                                                                                                               
    TheHdw.DSSC.Pins(digSrcPin).Pattern(patt).Source.Signals.DefaultSignal = SignalName

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_DSSC_HRAM", "DSSC_PostSetupDigSrcWave") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function DSSC_BodySetupDigSrcWave(patt As String, digSrcPin As PinList, SignalName As String, SegmentSize As Long, WaveDefArray() As Long)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim funcName As String:: funcName = "DSSC_BodySetupDigSrcWave"
    
    'store efuse program bit into a DSP wave
    Dim InWave As New DSPWave
    Dim site As Variant
    Dim WaveDef As String
    
    ''WaveDef = "WaveDef" ''''was
    InWave.data = WaveDefArray
    
    site = Theexec.sites.siteNumber
    
    ''''20170920 <NOTICE> if multiple apply this function call/sequence to avoid the following SrcWave to overwrite the previous one
    WaveDef = "WaveDef_" + SignalName + "_" & site
    
    Theexec.WaveDefinitions.CreateWaveDefinition WaveDef, InWave, True
    
    With TheHdw.DSSC.Pins(digSrcPin).Pattern(patt).Source.Signals(SignalName)
        .WaveDefinitionName = WaveDef
        .SampleSize = SegmentSize
    End With

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_DSSC_HRAM", "DSSC_BodySetupDigSrcWave") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
