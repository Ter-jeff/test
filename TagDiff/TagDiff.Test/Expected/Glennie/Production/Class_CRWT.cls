VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_CRWT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Private variable
' Pins
Private pins() As String ' The array of Tune Pins
Private Powers As String ' Tune Pins
Private pinCnt As Long ' Tune Pins Number
Private DCVS_Pins As String
Private DCVS_pin() As String
Private DCVS_Cnt As Long
Private DCVI_Pins As String
Private DCVI_pin() As String
Private DCVI_Cnt As Long


' For resotre used
Private Init_PinName() As String
Private Init_CurrentRange() As Double
Private Init_Source_FoldLimit() As Double
Private Init_Sink_FoldLimit() As Double
Private Init_Voltage() As Double

' Sheet
Private sheetName As String ' Output sheet Name of Tune Result
Private sheetFind As Boolean ' The boolean of sheet if exists is Tune Result exists

' Tune used
Private TuneCnt As Long ' The Count of Tune Result
Private CurrentRangeNow() As Double ' For loop tune used
Public Dict_TuneTime As New Dictionary ' Sotre TuneResult Dict
Private meterSamples As Double ' meter sample size
Private meterSampleRate As Double ' meter sample rate
Private B_PassSTCR As Boolean ' Fixed Current Range Wait Time used
Private DSPW As New DSPWave

Public Function Initialize(TunePins As String)
On Error GoTo errHandler
    Dim funcName As String:: funcName = "Class_CRWT.Initialize"
    Dim i As Integer
    
    Powers = TunePins
    DCVS_Cnt = 0
    DCVI_Cnt = 0
    If sheetName = "" Then sheetName = "TuneCRWTResult_" & UCase(currentJobName)
    sheetFind = FindSheet(sheetName, DMGR_SHEET_TYPE_USER)
    
    ' Split Pins and ReDim Array
    TheExec.DataManager.DecomposePinList Powers, pins(), pinCnt
    ReDim CurrentRangeNow(pinCnt - 1) As Double
    ReDim Init_PinName(pinCnt - 1) As String
    ReDim Init_CurrentRange(pinCnt - 1) As Double
    ReDim Init_Source_FoldLimit(pinCnt - 1) As Double
    ReDim Init_Sink_FoldLimit(pinCnt - 1) As Double
    ReDim Init_Voltage(pinCnt - 1) As Double
    
    TuneCnt = 0
    If meterSamples = 0 Then meterSamples = 2048
    If meterSampleRate = 0 Then meterSampleRate = 2048
    
    For i = 0 To pinCnt - 1
        ' Init CurrentRange and store origin CurrentRange/IFold
        If UCase(TheExec.DataManager.ChannelType(pins(i))) Like "*DCVS*" Then
            With TheHdw.DCVS.pins(pins(i))
                Init_PinName(i) = pins(i)
                Init_CurrentRange(i) = .CurrentRange.value
                Init_Source_FoldLimit(i) = .CurrentLimit.Source.FoldLimit.level.value
                'Init_Sink_FoldLimit(i) = .CurrentLimit.Sink.FoldLimit.level.value
                Init_Voltage(i) = .Voltage.value
                .SetCurrentRanges .CurrentRange.Max, .CurrentRange.Max
            End With
            If DCVS_Pins = "" Then
                DCVS_Pins = pins(i)
            Else
                DCVS_Pins = DCVS_Pins & "," & pins(i)
            End If
            ReDim Preserve DCVS_pin(DCVS_Cnt) As String
            DCVS_Cnt = DCVS_Cnt + 1
        ElseIf UCase(TheExec.DataManager.ChannelType(pins(i))) Like "*DCVI*" Then
            With TheHdw.DCVI.pins(pins(i))
                Init_PinName(i) = pins(i)
                Init_CurrentRange(i) = .CurrentRange.value
                Init_Voltage(i) = .Voltage.value
                .SetCurrentAndRange .CurrentRange.Max, .CurrentRange.Max
            End With
            If DCVI_Pins = "" Then
                DCVI_Pins = pins(i)
            Else
                DCVI_Pins = DCVI_Pins & "," & pins(i)
            End If
            ReDim Preserve DCVI_pin(DCVI_Cnt) As String
            DCVI_Cnt = DCVI_Cnt + 1
        End If
    Next i
    
    ' Turn off SourceTimeout and Overrage Alarm
    If DCVS_Pins <> "" Then
        TheHdw.DCVS.pins(DCVS_Pins).Gate = False
        TheHdw.DCVS.pins(DCVS_Pins).Disconnect
        TheHdw.DCVS.pins(DCVS_Pins).Alarm(tlDCVSAlarmOverrange) = tlAlarmOff
        TheHdw.DCVS.pins(DCVS_Pins).mode = tlDCVSModeVoltage
        TheHdw.DCVS.pins(DCVS_Pins).Meter.mode = tlDCVSMeterCurrent
        TheHdw.DCVS.pins(DCVS_Pins).Gate = True
        TheHdw.DCVS.pins(DCVS_Pins).Connect
    End If
    
    If DCVI_Pins <> "" Then
        TheHdw.DCVI.pins(DCVI_Pins).Gate = False
        TheHdw.DCVI.pins(DCVI_Pins).Disconnect
        TheHdw.DCVI.pins(DCVI_Pins).FoldCurrentLimit.Behavior = tlDCVIFoldCurrentLimitBehaviorDoNotGateOff
        TheHdw.DCVI.pins(DCVI_Pins).Alarm(tlDCVIAlarmFoldCurrentLimit) = tlAlarmOff
        TheHdw.DCVI.pins(DCVI_Pins).Alarm(tlDCVIAlarmOverRange) = tlAlarmOff
        TheHdw.DCVI.pins(DCVI_Pins).mode = tlDCVIModeVoltage
        TheHdw.DCVI.pins(DCVI_Pins).Meter.mode = tlDCVIMeterCurrent
        TheHdw.DCVI.pins(DCVI_Pins).Gate = True
        TheHdw.DCVI.pins(DCVI_Pins).Connect
    End If
    
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

' Reset Global sheetName
Public Function setShetName(ShtName As String)
    sheetName = ShtName
End Function

' Change B_PassSTCR boolean status
Public Function SetStatus_BPassSTCR(F_enable As Boolean)
    B_PassSTCR = F_enable
End Function

' Change meter Sample Size and Rate
Public Function ChangeSampleRateAndSizes(targetSamplesSize As Long, targetSampleRate As Long)
    If targetSamplesSize < 1 Then
        TheExec.Datalog.WriteComment "Your targetSamplesSize: " & targetSamplesSize & " < 1"
        Exit Function
    ElseIf targetSamplesSize > 2048 Then
        TheExec.Datalog.WriteComment "Your targetSamplesSize: " & targetSamplesSize & " > 2048, meter samples maximum is 2048"
        Exit Function
    End If
    If targetSampleRate < 1 Then
        TheExec.Datalog.WriteComment "Your targetSampleRate: " & targetSampleRate & " < 1"
        Exit Function
    ElseIf targetSampleRate > 250 * kHz Then
        TheExec.Datalog.WriteComment "Your targetSampleRate: " & targetSampleRate & " > 250k Hz, meter samples rate maximum is 250k Hz"
        Exit Function
    End If
    meterSamples = targetSamplesSize
    meterSampleRate = targetSampleRate
End Function

' Class close
Public Function TuneClose()
On Error GoTo errHandler
    Dim funcName As String:: funcName = "Class_CRWT.TuneClose"
    
    Dim i As Integer
    ' Restore CurrentRange and IFold Value
    For i = 0 To pinCnt - 1
        If UCase(TheExec.DataManager.ChannelType(pins(i))) Like "*DCVS*" Then
            With TheHdw.DCVS.pins(pins(i))
                .SetCurrentRanges Init_CurrentRange(i), Init_CurrentRange(i)
                .CurrentLimit.Source.FoldLimit.level.value = Init_Source_FoldLimit(i)
                '.CurrentLimit.Sink.FoldLimit.level.value = Init_Sink_FoldLimit(i)
                .Voltage.value = Init_Voltage(i)
            End With
        ElseIf UCase(TheExec.DataManager.ChannelType(pins(i))) Like "*DCVI*" Then
            With TheHdw.DCVI.pins(pins(i))
                .SetCurrentAndRange Init_CurrentRange(i), Init_CurrentRange(i)
                .Voltage.value = Init_Voltage(i)
            End With
        End If
    Next i
    
    If DCVS_Pins <> "" Then
        ' Open SourceTimeout and Overrange Alarm
        TheHdw.DCVS.pins(DCVS_Pins).Gate = False
        TheHdw.DCVS.pins(DCVS_Pins).Disconnect
        TheHdw.DCVS.pins(DCVS_Pins).Alarm(tlDCVSAlarmOverrange) = tlAlarmForceFail
        TheHdw.DCVS.pins(DCVS_Pins).Gate = True
        TheHdw.DCVS.pins(DCVS_Pins).Connect
    End If
    If DCVI_Pins <> "" Then
        TheHdw.DCVI.pins(DCVI_Pins).Gate = False
        TheHdw.DCVI.pins(DCVI_Pins).Disconnect
        TheHdw.DCVI.pins(DCVI_Pins).Alarm(tlDCVIAlarmFoldCurrentLimit) = tlAlarmOff
        TheHdw.DCVI.pins(DCVI_Pins).Alarm(tlDCVIAlarmOverRange) = tlAlarmOff
        TheHdw.DCVI.pins(DCVI_Pins).Gate = True
        TheHdw.DCVI.pins(DCVI_Pins).Connect
    End If
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

' Set ApplyPins voltage to PowerV and setting Gate and Connect Status
Public Function DCVS_ApplyVoltage(ApplyPins As String, Optional PowerV As Double = 0, Optional gateStatus As gateStatus = gKeep, _
                                Optional ConnectedStatus As ConnectedStatus = KeepStatus)
On Error GoTo errHandler
    Dim funcName As String:: funcName = "Class_CRWT.DCVS_ApplyVoltage"
    
    If PowerV <> 0 Then
        TheHdw.DCVS.pins(ApplyPins).Voltage.value = PowerV
    End If
    If gateStatus = gOn Then
        TheHdw.DCVS.pins(ApplyPins).Gate = True
    ElseIf gateStatus = gOff Then
        TheHdw.DCVS.pins(ApplyPins).Gate = False
    End If
    If ConnectedStatus = Connect Then
        TheHdw.DCVS.pins(ApplyPins).Connect
    ElseIf ConnectedStatus = Disconnect Then
        TheHdw.DCVS.pins(ApplyPins).Disconnect
    End If
    
    ' Wait Voltage stable
    TheHdw.Wait 10 * ms
    
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
' Set ApplyPins voltage to PowerV and setting Gate and Connect Status
Public Function DCVI_ApplyVoltage(ApplyPins As String, Optional PowerV As Double = 0, Optional gateStatus As gateStatus = gKeep, _
                                Optional ConnectedStatus As ConnectedStatus = KeepStatus)
On Error GoTo errHandler
    Dim funcName As String:: funcName = "Class_CRWT.DCVI_ApplyVoltage"
    
    If PowerV <> 0 Then
        TheHdw.DCVI.pins(ApplyPins).Voltage.value = PowerV
    End If
    If gateStatus = gOn Then
        TheHdw.DCVI.pins(ApplyPins).Gate = True
    ElseIf gateStatus = gOff Then
        TheHdw.DCVI.pins(ApplyPins).Gate = False
    End If
    If ConnectedStatus = Connect Then
        TheHdw.DCVI.pins(ApplyPins).Connect
    ElseIf ConnectedStatus = Disconnect Then
        TheHdw.DCVI.pins(ApplyPins).Disconnect
    End If
    
    ' Wait Voltage stable
    TheHdw.Wait 10 * ms
    
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
' Finxed Voltage and Current Range Used
Public Function FixedVCR_WaitTime(ForceVoltages As String, CurrentRanges As String, AccuracyRatio As Double, NoRunBelowCurrentRange As Double, _
                                                    SmoothCnt As Integer, WriteCurrentProfileFile As Boolean)
On Error GoTo errHandler
    Dim funcName As String:: funcName = "Class_CRWT.FixedVCR_WaitTime"
    Dim TuneTime As Double
    Dim ForceVoltage() As String
    Dim CurrentRange() As String
    Dim i As Integer
    Dim j As Integer
    Dim FoundCurrent As Boolean
    Dim TuneAccuracyRatio As Double
    Dim pininfo As AutoRange_Info
    
    If ForceVoltages <> "" Then
        ForceVoltage = Split(ForceVoltages, ",")
    Else
        ReDim ForceVoltage(UBound(pins)) As String
    End If
    If CurrentRanges <> "" Then
        CurrentRange = Split(CurrentRanges, ",")
    Else
        ReDim CurrentRange(UBound(pins)) As String
    End If
    TuneAccuracyRatio = AccuracyRatio
    B_PassSTCR = True
    For i = 0 To pinCnt - 1
        If UCase(TheExec.DataManager.ChannelType(pins(i))) <> "N/C" Then
            pininfo = PowerPin_range_ary(gl_dicPowerPinIndex(LCase(pins(i))))
            If ForceVoltage(i) <> "" Then
                If pininfo.MergeType Like "DCVS*" Then Call DCVS_ApplyVoltage(pins(i), CDbl(ForceVoltage(i)))
                If pininfo.MergeType Like "DCVI*" Then Call DCVI_ApplyVoltage(pins(i), CDbl(ForceVoltage(i)))
            Else
                If pininfo.MergeType Like "DCVS*" Then ForceVoltage(i) = CStr(Format(TheHdw.DCVS.pins(pininfo.PinName).Voltage.value, "0.000"))
                If pininfo.MergeType Like "DCVI*" Then ForceVoltage(i) = CStr(Format(TheHdw.DCVI.pins(pininfo.PinName).Voltage.value, "0.000"))
            End If
            If CurrentRange(i) = "" Then CurrentRange(i) = "0"
            For j = 0 To UBound(pininfo.Range_List)
                If pininfo.Range_List(j) >= CDbl(CurrentRange(i)) Then
                    TheExec.Datalog.WriteComment "Pin " & pins(i) & " forcing " & ForceVoltage(i) & " CurrentRange is " & pininfo.Range_List(j)
                    TuneTime = TuneWaitTimeWithMeter(pininfo, TuneAccuracyRatio, CDbl(ForceVoltage(i)), j, WriteCurrentProfileFile, SmoothCnt)
                    If TuneTime <> 0 Then ' Get Tune Result Time
                        Call StoreTuneTimeToDict(pins(i), TuneTime, pininfo.WaitTime_List(j), pininfo.Range_List(j), TuneAccuracyRatio)
                        FoundCurrent = True
                        Exit For
                    End If
                Else
                    FoundCurrent = False
                End If
            Next j
            If Not FoundCurrent Then
                TheExec.Datalog.WriteComment "Pin " & pins(i) & ", CurrentRange " & CurrentRange(i) & " doesn't exist!"
            End If
            If pininfo.MergeType Like "DCVS*" Then Call DCVS_ApplyVoltage(pins(i), Init_Voltage(i))
            If pininfo.MergeType Like "DCVI*" Then Call DCVI_ApplyVoltage(pins(i), Init_Voltage(i))
        End If
        TuneAccuracyRatio = AccuracyRatio
    Next i
    
    If TuneCnt > 0 Then Call WriteTuneResult(sheetName, Dict_TuneTime, sheetFind)
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
' Finxed Voltage and Current Range Used
Public Function FixedVCR_WaitTime_parallel(ForceVoltages As String, CurrentRanges As String, AccuracyRatio As Double, NoRunBelowCurrentRange As Double, _
                                                    SmoothCnt As Integer, WriteCurrentProfileFile As Boolean)
On Error GoTo errHandler
    Dim funcName As String:: funcName = "Class_CRWT.FixedVCR_WaitTime"
    Dim TuneTime As Double
    Dim ForceVoltage() As String
    Dim CurrentRange() As String
    Dim i As Integer
    Dim j As Integer
    Dim FoundCurrent As Boolean
    Dim TuneAccuracyRatio As Double
    Dim pininfo As AutoRange_Info
    
    If ForceVoltages <> "" Then
        ForceVoltage = Split(ForceVoltages, ",")
    Else
        ReDim ForceVoltage(UBound(pins)) As String
    End If
    If CurrentRanges <> "" Then
        CurrentRange = Split(CurrentRanges, ",")
    Else
        ReDim CurrentRange(UBound(pins)) As String
    End If
    TuneAccuracyRatio = AccuracyRatio
    B_PassSTCR = True
    For i = 0 To pinCnt - 1
        If UCase(TheExec.DataManager.ChannelType(pins(i))) <> "N/C" Then
            pininfo = PowerPin_range_ary(gl_dicPowerPinIndex(LCase(pins(i))))
            If ForceVoltage(i) <> "" Then
                If pininfo.MergeType Like "DCVS*" Then Call DCVS_ApplyVoltage(pins(i), CDbl(ForceVoltage(i)))
                If pininfo.MergeType Like "DCVI*" Then Call DCVI_ApplyVoltage(pins(i), CDbl(ForceVoltage(i)))
            Else
                If pininfo.MergeType Like "DCVS*" Then ForceVoltage(i) = CStr(Format(TheHdw.DCVS.pins(pininfo.PinName).Voltage.value, "0.000"))
                If pininfo.MergeType Like "DCVI*" Then ForceVoltage(i) = CStr(Format(TheHdw.DCVI.pins(pininfo.PinName).Voltage.value, "0.000"))
            End If
            If CurrentRange(i) = "" Then CurrentRange(i) = "0"
            For j = 0 To UBound(pininfo.Range_List)
                If pininfo.Range_List(j) >= CDbl(CurrentRange(i)) Then
                    TheExec.Datalog.WriteComment "Pin " & pins(i) & " forcing " & ForceVoltage(i) & " CurrentRange is " & pininfo.Range_List(j)
                    TuneTime = TuneWaitTimeWithMeter(pininfo, TuneAccuracyRatio, CDbl(ForceVoltage(i)), j, WriteCurrentProfileFile, SmoothCnt)
                    If TuneTime <> 0 Then ' Get Tune Result Time
                        Call StoreTuneTimeToDict(pins(i), TuneTime, pininfo.WaitTime_List(j), pininfo.Range_List(j), TuneAccuracyRatio)
                        FoundCurrent = True
                        Exit For
                    End If
                Else
                    FoundCurrent = False
                End If
            Next j
            If Not FoundCurrent Then
                TheExec.Datalog.WriteComment "Pin " & pins(i) & ", CurrentRange " & CurrentRange(i) & " doesn't exist!"
            End If
            If pininfo.MergeType Like "DCVS*" Then Call DCVS_ApplyVoltage(pins(i), Init_Voltage(i))
            If pininfo.MergeType Like "DCVI*" Then Call DCVI_ApplyVoltage(pins(i), Init_Voltage(i))
        End If
        TuneAccuracyRatio = AccuracyRatio
    Next i
    For i = 0 To pinCnt - 1
        If UCase(TheExec.DataManager.ChannelType(pins(i))) <> "N/C" Then
            pininfo = PowerPin_range_ary(gl_dicPowerPinIndex(LCase(pins(i))))
            If ForceVoltage(i) <> "" Then
                If pininfo.MergeType Like "DCVS*" Then Call DCVS_ApplyVoltage(pins(i), CDbl(ForceVoltage(i)))
                If pininfo.MergeType Like "DCVI*" Then Call DCVI_ApplyVoltage(pins(i), CDbl(ForceVoltage(i)))
            Else
                If pininfo.MergeType Like "DCVS*" Then ForceVoltage(i) = CStr(Format(TheHdw.DCVS.pins(pininfo.PinName).Voltage.value, "0.000"))
                If pininfo.MergeType Like "DCVI*" Then ForceVoltage(i) = CStr(Format(TheHdw.DCVI.pins(pininfo.PinName).Voltage.value, "0.000"))
            End If
            If CurrentRange(i) = "" Then CurrentRange(i) = "0"
            For j = 0 To UBound(pininfo.Range_List)
                If pininfo.Range_List(j) >= CDbl(CurrentRange(i)) Then
                    TheExec.Datalog.WriteComment "Pin " & pins(i) & " forcing " & ForceVoltage(i) & " CurrentRange is " & pininfo.Range_List(j)
                    TuneTime = TuneWaitTimeWithMeter(pininfo, TuneAccuracyRatio, CDbl(ForceVoltage(i)), j, WriteCurrentProfileFile, SmoothCnt)
                    If TuneTime <> 0 Then ' Get Tune Result Time
                        Call StoreTuneTimeToDict(pins(i), TuneTime, pininfo.WaitTime_List(j), pininfo.Range_List(j), TuneAccuracyRatio)
                        FoundCurrent = True
                        Exit For
                    End If
                Else
                    FoundCurrent = False
                End If
            Next j
            If Not FoundCurrent Then
                TheExec.Datalog.WriteComment "Pin " & pins(i) & ", CurrentRange " & CurrentRange(i) & " doesn't exist!"
            End If
            If pininfo.MergeType Like "DCVS*" Then Call DCVS_ApplyVoltage(pins(i), Init_Voltage(i))
            If pininfo.MergeType Like "DCVI*" Then Call DCVI_ApplyVoltage(pins(i), Init_Voltage(i))
        End If
        TuneAccuracyRatio = AccuracyRatio
    Next i
    If TuneCnt > 0 Then Call WriteTuneResult(sheetName, Dict_TuneTime, sheetFind)
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
' Loop Voltage and Auto Change Current Range
Public Function LoopVCR_WaitTime(StartVoltage As Double, EndVoltage As Double, StepVoltage As Double, AccuracyRatio As Double, NoRunBelowCurrentRange As Double, _
                                                SmoothCnt As Integer, WriteCurrentProfileFile As Boolean)
On Error GoTo errHandler
    Dim funcName As String:: funcName = "Class_CRWT.LoopVCR_WaitTime"

    Dim i As Integer
    Dim j As Integer
    Dim k As Long
    Dim ForceVLoop As Double
    Dim TuneTime As Double
    Dim TuneAccuracyRatio As Double
    Dim pininfo As AutoRange_Info
    
    TuneAccuracyRatio = AccuracyRatio
    
    ' Voltage Loop
    For ForceVLoop = StartVoltage To EndVoltage + 0.01 Step StepVoltage
        TheExec.Datalog.WriteComment "===============================================ForceV: " & ForceVLoop & "==============================================="
        If DCVS_Pins <> "" Then Call DCVS_ApplyVoltage(DCVS_Pins, ForceVLoop)
        If DCVI_Pins <> "" Then Call DCVI_ApplyVoltage(DCVI_Pins, ForceVLoop)
        For i = 0 To pinCnt - 1
            If UCase(TheExec.DataManager.ChannelType(pins(i))) <> "N/C" Then
                pininfo = PowerPin_range_ary(gl_dicPowerPinIndex(LCase(pins(i))))
                For j = 0 To UBound(pininfo.Range_List)
                    If pininfo.Range_List(j) > NoRunBelowCurrentRange And pininfo.Range_List(j) > CurrentRangeNow(i) Then
                        TheExec.Datalog.WriteComment "Pin " & pins(i) & ", TuneCurrentRange: " & CStr(pininfo.Range_List(j)) & " > AlreadyTuneCurrentRange: " & CStr(CurrentRangeNow(i))
                        TuneTime = TuneWaitTimeWithMeter(pininfo, TuneAccuracyRatio, ForceVLoop, j, WriteCurrentProfileFile, SmoothCnt)
                        If TuneTime = -1 Then ' Meter Result is small than Next CurrentRange
                            Exit For
                        ElseIf TuneTime <> 0 Then ' Get Tune Result Time
                            Call StoreTuneTimeToDict(pins(i), TuneTime, pininfo.WaitTime_List(j), pininfo.Range_List(j), TuneAccuracyRatio)
                            CurrentRangeNow(i) = pininfo.Range_List(j)
                            Exit For
                        End If
                    End If
                Next j
            End If
            TuneAccuracyRatio = AccuracyRatio
        Next i
    Next ForceVLoop
    
    If TuneCnt > 0 Then Call WriteTuneResult(sheetName, Dict_TuneTime, sheetFind)
    
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

' Load Sheet Data to DC structure
Public Function LoadCRWTToPowerPinAry(Optional ShtName As String = vbNullString, Optional OriStoreName As String = vbNullString) As Boolean
On Error GoTo errHandler
    Dim funcName As String:: funcName = "Class_CRWT.LoadCRWTToPowerPinAry"
    Dim sheetOrgFind As Boolean
    Dim TuneCnt As Long
    Dim i As Integer, j As Integer
    Dim all_key() As Variant
    Dim tmpstr() As String
    Dim Dict_OrgTime As New Dictionary
    Dim PowerPin_range As AutoRange_Info
    
    ' If shtName change, re-find Sheet
    If ShtName <> "" Then
        sheetName = ShtName
        sheetFind = FindSheet(sheetName, DMGR_SHEET_TYPE_USER)
    End If
    If OriStoreName <> "" Then sheetOrgFind = FindSheet(OriStoreName, DMGR_SHEET_TYPE_USER)

    If sheetFind Then
        ' Load Result to Dictionary
        Call LoadCRWTResultToDict
        all_key = Dict_TuneTime.Keys
        For i = 0 To UBound(all_key)
            tmpstr = Split(all_key(i), "-")
            If UCase(TheExec.DataManager.ChannelType(tmpstr(0))) <> "N/C" Then
'                With PowerPin_range_ary(gl_dicPowerPinIndex(LCase(tmpStr(0))))
'                    For j = 0 To UBound(.Range_List)
'                        If CStr(.Range_List(j)) = tmpStr(1) Then
'                            Dict_OrgTime.Add all_key(i), .WaitTime_List(j)
'                            .WaitTime_List(j) = Dict_TuneTime(all_key(i))
'                            Exit For
'                        End If
'                        TheExec.Datalog.WriteComment "Doesn't find Current Range: " & CStr(.Range_List(j)) & " , Pin: " & tmpStr(0)
'                    Next j
'                End With

                PowerPin_range = PowerPin_range_ary_CRWT(gl_dicPowerPinIndex(LCase(tmpstr(0))))
                
                For j = 0 To UBound(PowerPin_range.Range_List)
                    If CStr(PowerPin_range.Range_List(j)) = tmpstr(1) Then
                        'PowerPin_range.WaitTime_List(j) = Dict_TuneTime(all_key(i))
                        PowerPin_range_ary_CRWT(gl_dicPowerPinIndex(LCase(tmpstr(0)))).WaitTime_List(j) = Dict_TuneTime(all_key(i))
                        'Exit For
                    Else
                        If (LCase(PowerPin_range.ChanMapType) = "hexvs") Or (LCase(PowerPin_range.ChanMapType) = "vsm") Or (LCase(PowerPin_range.ChanMapType) = "vs-5a") Or (LCase(PowerPin_range.ChanMapType) = "vs-800ma") Then  '''''UFP_Corr
                            If PowerPin_range.Range_List(j) = 0.01 Then
                                PowerPin_range_ary_CRWT(gl_dicPowerPinIndex(LCase(tmpstr(0)))).WaitTime_List(j) = 100 * ms
                            Else
                                PowerPin_range_ary_CRWT(gl_dicPowerPinIndex(LCase(tmpstr(0)))).WaitTime_List(j) = 30 * ms
                            End If
                        Else
                            PowerPin_range_ary_CRWT(gl_dicPowerPinIndex(LCase(tmpstr(0)))).WaitTime_List(j) = 5 * ms + 0.01
                            If (UCase(PowerPin_range.PinName) = "VDD_SRAM_SOC") Or (UCase(PowerPin_range.PinName) = "VDDIO12_AOP") Or (UCase(PowerPin_range.PinName) = "VDDIO12_GRP") Then
                                PowerPin_range_ary_CRWT(gl_dicPowerPinIndex(LCase(tmpstr(0)))).WaitTime_List(j) = PowerPin_range_ary_CRWT(gl_dicPowerPinIndex(LCase(tmpstr(0)))).WaitTime_List(j) + 0.01
                            End If
                        End If
                    TheExec.Datalog.WriteComment "Doesn't find Current Range: " & CStr(PowerPin_range.Range_List(j)) & " , Pin: " & tmpstr(0)
                    End If
                Next j
            Else
                TheExec.Datalog.WriteComment "Pin: " & tmpstr(0) & " is N/C"
            End If
        Next i
        
        If OriStoreName <> "" Then Call WriteTuneResult(OriStoreName, Dict_OrgTime, sheetOrgFind)
        
        TheExec.Datalog.WriteComment "Load CurrentRange WaitTime Finished!!"
        LoadCRWTToPowerPinAry = True
    Else
        TheExec.Datalog.WriteComment "sheetName " & sheetName & " Not Found!!!"
        LoadCRWTToPowerPinAry = False
    End If
    
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

' Load Sheet to Global Dictionary
Public Function LoadCRWTResultToDict()
On Error GoTo errHandler
    Dim funcName As String:: funcName = "Class_CRWT.LoadCRWTResultToDict"
    Dim MySheet As Worksheet:: Set MySheet = Nothing
    Dim MaxColumn As Long, MaxRow As Long
    Dim TuneCurrListArr As Variant
    Dim i, j As Long
    Dim temp_str As String
    
    If sheetFind Then
        Set MySheet = ActiveWorkbook.Sheets(sheetName)
        MySheet.Select
        MaxRow = MySheet.UsedRange.Rows.Count
        MaxColumn = MySheet.UsedRange.Columns.Count
        If MaxRow = 0 Then Exit Function
        
        Dict_TuneTime.RemoveAll
        TuneCurrListArr = MySheet.range(Cells(1, 1), Cells(MaxRow, MaxColumn)).value
        TuneCnt = MaxRow
        
        For i = 1 To MaxRow
            temp_str = TuneCurrListArr(i, 1) & "-" & CStr(TuneCurrListArr(i, 2))
            Dict_TuneTime.Add temp_str, TuneCurrListArr(i, 3)
        Next i
    End If
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function InstrumentSetSampleRateAndSamples(pininfo As AutoRange_Info, CurrentRange As Double, OutputCurrent As Double)
    If LCase(pininfo.ChanMapType) = "vs-5a" Then
        TheHdw.DCVS.pins(pininfo.PinName).CurrentLimit.Source.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorDoNotGateOff
        TheHdw.DCVS.pins(pininfo.PinName).CurrentLimit.Sink.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorDoNotGateOff
        TheHdw.DCVS.pins(pininfo.PinName).Alarm(tlDCVSAlarmSinkFoldCurrentLimitTimeout) = tlAlarmOff
        TheHdw.DCVS.pins(pininfo.PinName).Alarm(tlDCVSAlarmSourceFoldCurrentLimitTimeout) = tlAlarmOff
        OutputCurrent = CurrentRange
    ElseIf LCase(pininfo.ChanMapType) = "vs-800ma" Then
        TheHdw.DCVS.pins(pininfo.PinName).CurrentLimit.Source.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorDoNotGateOff
        TheHdw.DCVS.pins(pininfo.PinName).CurrentLimit.Sink.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorDoNotGateOff
        TheHdw.DCVS.pins(pininfo.PinName).Alarm(tlDCVSAlarmSinkFoldCurrentLimitTimeout) = tlAlarmOff
        TheHdw.DCVS.pins(pininfo.PinName).Alarm(tlDCVSAlarmSourceFoldCurrentLimitTimeout) = tlAlarmOff
        OutputCurrent = CurrentRange
    ElseIf LCase(pininfo.ChanMapType) = "hexvs" Then
        TheHdw.DCVS.pins(pininfo.PinName).Gate = False
        TheHdw.DCVS.pins(pininfo.PinName).Disconnect
        TheHdw.DCVS.pins(pininfo.PinName).CurrentLimit.Source.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorDoNotGateOff
        TheHdw.DCVS.pins(pininfo.PinName).CurrentLimit.Sink.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorDoNotGateOff
        TheHdw.DCVS.pins(pininfo.PinName).Alarm(tlDCVSAlarmSinkFoldCurrentLimitTimeout) = tlAlarmOff
        TheHdw.DCVS.pins(pininfo.PinName).Alarm(tlDCVSAlarmSourceFoldCurrentLimitTimeout) = tlAlarmOff
        If CurrentRange >= 15 Then
            OutputCurrent = CurrentRange
        Else
            OutputCurrent = 15
        End If
        TheHdw.DCVS.pins(pininfo.PinName).Gate = True
        TheHdw.DCVS.pins(pininfo.PinName).Connect
    ElseIf LCase(pininfo.ChanMapType) = "vhdvs" Then
        TheHdw.DCVS.pins(pininfo.PinName).Gate = False
        TheHdw.DCVS.pins(pininfo.PinName).Disconnect
        TheHdw.DCVS.pins(pininfo.PinName).Alarm(tlDCVSAlarmFoldCurrentLimitTimeout) = tlAlarmOff
        If CurrentRange >= pininfo.Range_List(UBound(pininfo.Range_List)) Then
            OutputCurrent = CurrentRange
        Else
            OutputCurrent = pininfo.Range_List(UBound(pininfo.Range_List))
        End If
        TheHdw.DCVS.pins(pininfo.PinName).Gate = True
        TheHdw.DCVS.pins(pininfo.PinName).Connect
    ElseIf LCase(pininfo.ChanMapType) = "dc-30" Then
        OutputCurrent = CurrentRange
    ElseIf LCase(pininfo.ChanMapType) = "dc-70" Then
        OutputCurrent = CurrentRange
    End If
End Function
Private Function GetDSPDataToPinData(PinName As String, CapSigName As String, MeterData As PinData)
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Class_CRWT.GetDSPDataToPinListData"
    Dim site As Variant
    Dim i As Integer
    Do While TheHdw.DCVS.pins(PinName).Capture.IsRunning = True
    Loop
    
    DSPW = TheHdw.DCVS.pins(PinName).Capture.Signals(CapSigName).DSPWave
    MeterData.Add PinName
    For Each site In TheExec.sites
        MeterData.value = DSPW.data
    Next site
    
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
Private Function MeterWithDSPStart(PinName As String, CapSigName As String)
    On Error GoTo errHandler
    Dim i As Long
    Dim pinCnt As Long
    Dim pinAry() As String
    Dim funcName As String:: funcName = "Class_CRWT.MeterWithDSPStart"
    
    ' Wait if another capture is running
    Do While TheHdw.DCVS.pins(PinName).Capture.IsRunning = True
    Loop
    
    ' Clear capture memory
    TheHdw.DCVS.pins(PinName).ClearCaptureMemory
    
    ' Create a SIGNAL to set up instrument
    TheHdw.DCVS.pins(PinName).Capture.Signals.Add CapSigName

    ' Set this as the default signal
    TheHdw.DCVS.pins(PinName).Capture.Signals.DefaultSignal = CapSigName
    
    TheExec.DataManager.DecomposePinList PinName, pinAry(), pinCnt
    
    ' Define the signal used for the capture
    For i = 0 To pinCnt - 1
        With TheHdw.DCVS.pins(pinAry(i)).Capture.Signals.item(CapSigName)
            .Reinitialize
            .mode = tlDCVSMeterCurrent
            .range = TheHdw.DCVS.pins(pinAry(i)).Meter.CurrentRange.Max
            .SampleRate = meterSampleRate
            .SampleSize = meterSamples
        End With
    Next i
    ' Setup the hardware by loading the signal
    TheHdw.DCVS.pins(PinName).Capture.Signals.item(CapSigName).LoadSettings

    ' Start the capture
    TheHdw.DCVS.pins(PinName).Capture.Signals.item(CapSigName).Trigger
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function SetCurrentRangeAndMeter(pininfo As AutoRange_Info, MR As PinData, MC As SiteDouble, CurrentRange As Double, CurrentRangeMax As Double)
On Error GoTo errHandler
    Dim funcName As String:: funcName = "Class_CRWT.SetCurrentRangeAndMeter"
    Dim OutputCurrent As Double
    Dim MeterResult() As Double
    Dim sigName As String
    Dim site As Variant 'Carter, 20240304
    sigName = "Capture1"
    InstrumentSetSampleRateAndSamples pininfo, CurrentRange, OutputCurrent
    If pininfo.MergeType Like "DCVS*" Then
        With TheHdw.DCVS.pins(pininfo.PinName)
            CurrentRangeMax = .CurrentRange.Max
            'MeterWithDSPStart PinInfo.PinName, SigName
            Call ProfileAutoResolution(UCase(pininfo.ChanMapType), 1, meterSamples, meterSampleRate, 1)
            StartProfile pininfo.PinName, "I", meterSampleRate, meterSamples, sigName, UCase(pininfo.ChanMapType)
            If LCase(pininfo.ChanMapType) = "vhdvs" Then
                .SetCurrentRanges CurrentRange, OutputCurrent
            Else
                .SetCurrentRanges OutputCurrent, CurrentRange
            End If
            'MR = .Meter.Read(tlStrobe, meterSamples, meterSampleRate, tlDCVSMeterReadingFormatArray)
            GetDSPDataToPinData pininfo.PinName, sigName, MR
            For Each site In TheExec.sites
                meterSamples = UBound(MR.value, 1)
                Exit For
            Next site
            MC = .Meter.Read(tlStrobe, 10, meterSampleRate, tlDCVSMeterReadingFormatAverage)
        End With
    ElseIf pininfo.MergeType Like "DCVI*" Then
        With TheHdw.DCVI.pins(pininfo.PinName)
            'CurrentRangeMax = .CurrentRange.max
            '.SetCurrentAndRange OutputCurrent, CurrentRange
            meterSamples = 1023
            MR = .Meter.Read(tlStrobe, meterSamples, meterSampleRate, tlDCVSMeterReadingFormatArray)
            MC = .Meter.Read(tlStrobe, 10, meterSampleRate, tlDCVSMeterReadingFormatAverage)
            'MR = .Meter.Read(tlStrobe, 1024, meterSampleRate, tlDCVSMeterReadingFormatArray)
            'MC = .Meter.Read(tlStrobe, 10, meterSampleRate, tlDCVSMeterReadingFormatAverage)
        End With
    End If
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function SetCurrentRange(pininfo As AutoRange_Info, CurrentRange As Double, OutputCurrent As Double)
    If pininfo.MergeType Like "DCVS*" Then
        TheHdw.DCVS.pins(pininfo.PinName).SetCurrentRanges OutputCurrent, CurrentRange
    ElseIf pininfo.MergeType Like "DCVI*" Then
        'TheHdw.DCVI.Pins(PinInfo.PinName).SetCurrentAndRange OutputCurrent, CurrentRange
    End If
End Function

' Find CRWT with Meter
Private Function TuneWaitTimeWithMeter(pininfo As AutoRange_Info, CurrentRangeAccuracy As Double, _
                                        CurrentVoltage As Double, CurrentRangeindex As Integer, CR_Profile As Boolean, SmoothCnt As Integer) As Double
On Error GoTo errHandler
    Dim funcName As String:: funcName = "Class_CRWT.TuneWaitTimeWithMeter"
    Dim MeterResult As New PinData
    Dim MeterCompare As New SiteDouble
    Dim i As Long
    Dim site As Variant
    Dim DropRngSite As New SiteBoolean
    Dim CntPoint As New SiteLong
    Dim CurrentRangeMax As Double

    SetCurrentRangeAndMeter pininfo, MeterResult, MeterCompare, pininfo.Range_List(CurrentRangeindex), CurrentRangeMax
    ' When current is large than CurrentRange, it is not correct setting, need to change CurrentRange
    DropRngSite = MeterCompare.compare(GreaterThan, pininfo.Range_List(CurrentRangeindex))
    If DropRngSite.Any(True) Then
        TuneWaitTimeWithMeter = 0
        SetCurrentRange pininfo, CurrentRangeMax, CurrentRangeMax
        Exit Function
    End If
    
    ' When current is small than next CurrentRange, it is not correct setting, need to change CurrentRange
    If CurrentRangeindex > 0 And Not B_PassSTCR Then
        DropRngSite = MeterCompare.compare(LessThan, pininfo.Range_List(CurrentRangeindex - 1))
        If DropRngSite.Any(True) Then
            TuneWaitTimeWithMeter = -1
            SetCurrentRange pininfo, CurrentRangeMax, CurrentRangeMax
            Exit Function
        End If
    End If
    ' Write CR Porfile if CR_Profile enable
    If CR_Profile Then
        Call WriteCurrentInTxt(MeterResult, pininfo.PinName, CurrentVoltage, pininfo.Range_List(CurrentRangeindex))
    End If
    
    'Find stable current value and calculate stable time
Retry:
    For i = 0 To CLng(meterSamples) - 1
        For Each site In TheExec.sites
            If Abs(MeterCompare(site) - MeterResult.value(site)(i)) < _
                Abs((MeterCompare(site) + pininfo.Range_List(CurrentRangeindex)) * CurrentRangeAccuracy) Then
                CntPoint(site) = CntPoint(site) + 1
            Else
                CntPoint(site) = 0
            End If
            If CntPoint(site) > SmoothCnt Then DropRngSite(site) = True ' Waitting wave to smooth
        Next site
        
        If DropRngSite.All(True) Then
            TuneWaitTimeWithMeter = i / meterSampleRate
            Exit For
        End If
    Next i
    
    ' If any site doesn't find result, increase Accuracy and re-find
    If DropRngSite.Any(False) Then
        CurrentRangeAccuracy = CurrentRangeAccuracy * 2
        GoTo Retry
    End If
        
    SetCurrentRange pininfo, CurrentRangeMax, CurrentRangeMax
    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


Private Function StoreTuneTimeToDict(pin As String, TuneTime As Double, WaitTime As Double, CurrentRange As Double, TuneAccuracyRatio As Double)
On Error GoTo errHandler
    Dim funcName As String:: funcName = "Class_CRWT.StoreTuneTimeToDict"
    Dim temp_str As String
    If TuneTime <> WaitTime Then
        temp_str = pin & "-" & CStr(CurrentRange)
        If Dict_TuneTime.Exists(temp_str) Then
            If Dict_TuneTime(temp_str) < TuneTime Then
                Dict_TuneTime.Remove temp_str
                Dict_TuneTime.Add temp_str, TuneTime
                TheExec.Datalog.WriteComment "Pin " & pin & " CurrentRange " & CurrentRange & _
                ", Aurracy Ratio is " & TuneAccuracyRatio & ", Tune Wait Time Result is " & TuneTime & " sec."
            End If
        Else
            Dict_TuneTime.Add temp_str, TuneTime
            TuneCnt = TuneCnt + 1
            TheExec.Datalog.WriteComment "Pin " & pin & " CurrentRange " & CurrentRange & _
            ", Aurracy Ratio is " & TuneAccuracyRatio & ", Tune Wait Time Result is " & TuneTime & " sec."
        End If
    End If
    
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function WriteCurrentInTxt(MeterResult As PinData, pin As String, CurrentVoltage As Double, CurrentRange As Double)
On Error GoTo errHandler
    Dim funcName As String:: funcName = "Class_CRWT.WriteCurrentInTxt"
    Dim Current_Dirctory As String
    Dim day_code As String
    Dim txtFileName As String
    Dim FileTxt As TextStream
    Dim site As Variant
    Dim i As Long
    
    day_code = right("0" & CStr(Month(Now)), 2) & right("0" & CStr(Day(Now)), 2)
    day_code = day_code & right("0" & CStr(Hour(Now)), 2) & right("0" & CStr(Minute(Now)), 2) & right("0" & CStr(Second(Now)), 2)
    Current_Dirctory = CurDir & "\CRWT-" & TestProgram_Day_Code & "\"
    If Dir(Current_Dirctory, vbDirectory) = Empty Then
        MkDir Current_Dirctory
    End If
    For Each site In TheExec.sites
        txtFileName = Current_Dirctory & "CurrentProfile-Site" & site & "-" & pin & "-" & meterSampleRate & "-" & meterSamples & "-" & _
                        CurrentVoltage & "V-" & CurrentRange & "A-" & day_code & ".txt"
        'CreateTxtFile txtFileName, FileTxt
        DSPW.FileExport txtFileName, File_txt
'        For i = 0 To CLng(meterSamples)
'            FileTxt.WriteLine CStr(MeterResult.value(site)(i))
'        Next i
    Next site
    'FileTxt.Close
    
    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function FindSheet(sheetName As String, SheetType As String) As Boolean
    Dim sheetnames() As String
    Dim idxSheet As Integer
    sheetnames = TheExec.job.GetSheetNamesOfType(SheetType)

    For idxSheet = 0 To UBound(sheetnames)
        If LCase(sheetnames(idxSheet)) Like LCase(sheetName) Then
            FindSheet = True
            Exit Function
        End If
    Next idxSheet
    FindSheet = False
End Function

Private Function WriteTuneResult(ShtName As String, Dict_TuneTime As Dictionary, sheetFind As Boolean)
On Error GoTo errHandler
    Dim funcName As String:: funcName = "Class_CRWT.WriteTuneResult"
    Dim MySheet As Worksheet:: Set MySheet = Nothing
    Dim i As Long
    Dim all_key() As Variant
    If Not sheetFind Then Sheets.Add.name = ShtName
    Set MySheet = ActiveWorkbook.Sheets(ShtName)
    MySheet.Select
    
    all_key = Dict_TuneTime.Keys
    For i = 0 To UBound(all_key)
        Cells(i + 1, 1) = Split(all_key(i), "-")(0)
        Cells(i + 1, 2) = Split(all_key(i), "-")(1)
        Cells(i + 1, 3) = Dict_TuneTime(all_key(i))
    Next i

    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function CreateTxtFile(filename As String, FileToCreate As TextStream)
    Dim fso As New FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set FileToCreate = fso.CreateTextFile(filename)
End Function

