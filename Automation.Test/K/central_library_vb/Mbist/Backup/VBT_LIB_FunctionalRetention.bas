Attribute VB_Name = "VBT_LIB_FunctionalRetention"
Public Function Functional_DataRetention(PatSet_First As Pattern, PatSet_Second As Pattern, _
    Optional DurationTime As Double = 0.1, Optional DisableConnectPins As PinList) As Long

On Error GoTo errHandler

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    If (DisableConnectPins <> "") Then TheHdw.Digital.Pins(DisableConnectPins).Disconnect

    TheHdw.Wait (100 * us)
    '' Run the first pattern to do write
    Call TheHdw.Patterns(PatSet_First).test(pfAlways, 0)
        
    '' Setup duration time after write pattern
    TheHdw.Wait (DurationTime)
    
    '' Run the first pattern to do write
    Call TheHdw.Patterns(PatSet_Second).test(pfAlways, 0)
    
    If (DisableConnectPins <> "") Then TheHdw.Digital.Pins(DisableConnectPins).Connect

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in Functional_DataRetention"
    If AbortTest Then Exit Function Else Resume Next
End Function



''Public Function VerifyInitialCondition(DisableConnectPins As PinList, DriveLowPins As PinList) As Long
''    Dim Pins As String
''    NumberPins As Long
''    Call TheExec.DataManager.DecomposePinList(DisableConnectPins, Pins(), NumberPins)
''
''    Dim i As Long
''    For i = 0 To NumberPins - 1
''    TheHdw.Digital.Pins(DriveLowPins).Disconnect
''    Next i
''
''    Call TheExec.DataManager.DecomposePinList(DriveLowPins, Pins(), NumberPins)
''
''    For i = 0 To NumberPins - 1
''
''    Next i
''End Function
