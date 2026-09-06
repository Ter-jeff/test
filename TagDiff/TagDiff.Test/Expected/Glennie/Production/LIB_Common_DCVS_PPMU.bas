Attribute VB_Name = "LIB_Common_DCVS_PPMU"
Option Explicit
'Revision History:
'V0.0 initial bring up
Public Const Version_Lib_Common_DC = "0.1" 'library version
Private Const moduleName = "LIB_Common_DCVS_PPMU"
Private functionName As String

'enums
Enum dcvs_type
    DCVS_HexVs = 1
    DCVS_HDVS = 2
    DCVS_UVS256 = 3
End Enum

Enum pinType
    SingleEnd = 0
    Differential = 1
End Enum


Public Function DCVS_MeterRead(dcvs_type As dcvs_type, powerPin As String, Sample_Size As Long, ByRef MeasCurr As PinListData)
'strobe measure values with DCVS instrument
    On Error GoTo errHandler
    Dim site As Variant
    Dim p As Variant
    Dim ResetI As Boolean
    Dim Default_CurrentRange As Double
    Dim s_ErrorMsg As String
    ResetI = False
    Select Case dcvs_type
    
        Case DCVS_HexVs, DCVS_HDVS:
            MeasCurr = TheHdw.DCVS.pins(powerPin).Meter.Read(tlStrobe, Sample_Size, 10000, tlDCVSMeterReadingFormatAverage)
             
        Case DCVS_UVS256:
            TheHdw.DCVS.pins(powerPin).Meter.Filter.value = TheHdw.DCVS.pins(powerPin).Meter.Filter.Max / Sample_Size ' UVS average 10 samples
            MeasCurr = TheHdw.DCVS.pins(powerPin).Meter.Read(tlStrobe, 1)  ' UVS only allow one sample
            
        Case Else:
            TheExec.flow.TestLimit resultVal:=1, lowVal:=-1, hiVal:=-1, unit:=unitNone, Tname:=TheExec.DataManager.instancename
            s_ErrorMsg = "Instrument type : " & CStr(dcvs_type) & "not define in function please check!!"
            Call Print_Error_Message(Error_Warning_Info.Warning_Info, "LIB_Common_DCVS_PPMU", "DCVS_MeterRead", s_ErrorMsg)
            Exit Function
    End Select
    
    Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_DCVS_PPMU", "DCVS_MeterRead") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function DCVS_PowerOn_I_Meter(pin As String, v As Double, i_rng As Double, wait_before_gate As Double, wait_after_gate As Double, Steps As Integer, RiseTime As Double, Optional debugPrintEnable As Boolean = False)
'set Force voltage and Current/Meter Range
''===============================================
''Description: __                __
''                __|
''            __|
''        __|
''    __| >|   |<--stepT  __        v
''__|                   __ stepV   __
''|<-- steps -->
''|<-FallTime ->
''===============================================
    On Error GoTo errHandler
    
    Dim i_meter_rng As Double
    Dim setV As Double
    Dim StepV As Double
    Dim stepT As Double
    
    Dim i As Integer
    
    i_meter_rng = i_rng
    StepV = v / Steps
    stepT = RiseTime / Steps

    With TheHdw.DCVS.pins(pin)
        .Connect
        .mode = tlDCVSModeVoltage
        .Voltage.Main = 0
        .Meter.mode = tlDCVSMeterCurrent
        
        If i_rng <> -99 Then    'bypass range setup if does not need to
            .SetCurrentRanges i_rng, i_meter_rng
            .CurrentRange.value = i_rng
            .CurrentLimit.Source.FoldLimit.level.value = i_rng
            If LCase(glb_TesterType) = "jaguar" Then
                .Meter.CurrentRange = i_rng
            End If
        End If
        
        TheHdw.Wait wait_before_gate   'wait for relay connect
        
        .Gate = True
    End With
    
''Pwr On Ramp up slew-rate control============================
    For i = 1 To Steps
        setV = i * StepV
        TheHdw.DCVS.pins(pin).Voltage.Main = setV
        
        If debugPrintEnable = True Then
            TheExec.Datalog.WriteComment "  Curr_" & pin & " Pwr Up Voltage (" & CStr(i) & ") : " & CStr(setV) & " V"
        End If
        
        TheHdw.Wait stepT
    Next i
''============================================================

    TheHdw.Wait wait_after_gate
      
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_DCVS_PPMU", "DCVS_PowerOn_I_Meter") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function DCVS_PowerOff_I_Meter_Parallel(pin As String, wait_before_gate As Double, wait_after_gate As Double, PowerSequence As Long, Optional debugPrintEnable As Boolean = False)
'set Force voltage and Current/Meter Range
''===============================================
''Description
''__                              __
''  |__
''     |_>|  |<--stepT             v
''        |__          __
''           |__       __ stepV   __
''|<-- steps -->
''|<-FallTime ->
''===============================================
    On Error GoTo errHandler
    
    Dim i_meter_rng As Double   'meter range
    Dim setV As Double          'current voltage
    Dim StepV As Double         'step voltage
    Dim stepT As Double         'step time
    Dim PowerPins() As String
    Dim pinCnt As Long
    Dim powerPin As Variant
    Dim TempString As String
    Dim VMain() As Double
    Dim Irange As Double
    Dim step As Integer
    Dim PreStep As Integer:: PreStep = 0
    Dim RiseTime As Double
    
    Dim i As Integer:: i = 1
    Dim j As Integer:: j = 1

    TheExec.DataManager.DecomposePinList pin, PowerPins(), pinCnt
    
    ReDim VMain(pinCnt - 1) As Double
    For Each powerPin In PowerPins
        'If TheExec.DataManager.ChannelType(PowerPin) <> "N/C" Then 'check CP for FT form NC pins
        TempString = powerPin & "_GLB"
        'Vmain(i) = TheExec.specs.Globals(TempString).ContextValue
        VMain(i) = TheHdw.DCVS.pins(powerPin).Voltage.Main.value
        'get Ifold limit spec value
        TempString = powerPin & "_Ifold_GLB"
        Irange = TheExec.Specs.Globals(TempString).ContextValue
        
        'auto calculate steps
        step = VMain(i) / 0.1 '0.1v per step
        If step = 0 Then step = 10 'default value
        If step > PreStep Then PreStep = step
        
        RiseTime = step * ms
        i_meter_rng = Irange
        
        With TheHdw.DCVS.pins(powerPin)
            .mode = tlDCVSModeVoltage
            .SetCurrentRanges Irange, i_meter_rng
            '.Meter.mode = tlDCVSMeterCurrent
            .CurrentRange.value = Irange
            .CurrentLimit.Source.FoldLimit.level.value = Irange
            ''202107
            If LCase(glb_TesterType) = "jaguar" Then
                .Meter.CurrentRange = Irange
            End If

        End With
        
        If debugPrintEnable = True Then    'debugprint
            TheExec.Datalog.WriteComment "print: Pin " & FormatNumericDatalog(powerPin, 30, False) & ", Vmain " & Format(VMain(i), "0.000") & " V, Irange " & FormatNumericDatalog(Format(Irange, "0.000"), 7, True) & " A, Step " & FormatNumericDatalog(step, 2, True) & ", FallTime " & FormatNumericDatalog(RiseTime * 1000, 2, True) & " ms" & ", PowerSequence " & FormatNumericDatalog(PowerSequence, 3, True)
        End If
        
        i = i + 1
'        Else
'            If DebugPrintEnable = True Then    'debugprint
'                TheExec.Datalog.WriteComment "print: Pin " & PowerPin & " not turn on by 'NC pin', PowerSequence " & PowerSequence & " ,Warning!!!"
'            End If
'        End If
    Next powerPin
    
    step = PreStep
    RiseTime = step * ms
    stepT = RiseTime / step
    
    With TheHdw.DCVS.pins(pin)
        .Connect
        TheHdw.Wait wait_before_gate
        .Gate = True
    End With
    
    ''Pwr On Ramp Down slew-rate control============================
    For j = 1 To step
        i = 1
        For Each powerPin In PowerPins
''            If TheExec.DataManager.ChannelType(PowerPin) <> "N/C" Then 'check CP for FT form NC pins  'no need to double check NC pin
            setV = VMain(i) - (j * VMain(i) / step)
            TheHdw.DCVS.pins(powerPin).Voltage.Main = setV
                
''                If DebugPrintEnable = True Then
''                    TheExec.Datalog.WriteComment "  Curr_" & PowerPin & " Pwr Down Voltage (" & CStr(i) & ") : " & Format(setV, "0.000") & " V"
''                End If
            i = i + 1
''            End If
        Next powerPin
        TheHdw.Wait stepT   'wait step time
    Next j
    
    setV = 0    'final step, return to 0V anyway
    TheHdw.DCVS.pins(pin).Voltage.Main = setV
    
''    If DebugPrintEnable = True Then
''        TheExec.Datalog.WriteComment "  Curr_" & PowerPin & " Pwr Down Voltage (" & CStr(i) & ") : " & CStr(setV) & " V"
''    End If
    ''==============================================================
    
    TheHdw.Wait wait_after_gate
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_DCVS_PPMU", "DCVS_PowerOff_I_Meter_Parallel") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function DCVS_PowerOn_I_Meter_Parallel(pin As String, wait_before_gate As Double, wait_after_gate As Double, PowerSequence As Long, Optional debugPrintEnable As Boolean = False)
'set Force voltage and Current/Meter Range
''===============================================
''Description: __                __
''                __|
''            __|
''        __|
''    __| >|   |<--stepT  __        v
''__|                   __ stepV   __
''|<-- steps -->
''|<-FallTime ->
''===============================================
    On Error GoTo errHandler
    
    Dim i_meter_rng As Double
    Dim setV As Double
    Dim StepV As Double
    Dim stepT As Double
    Dim PowerPins() As String
    Dim pinCnt As Long
    Dim powerPin As Variant
    Dim TempString As String
    Dim VMain() As Double
    Dim Irange As Double
    Dim step As Integer
    Dim PreStep As Integer:: PreStep = 0
    Dim RiseTime As Double
    
    Dim i As Integer:: i = 1
    Dim j As Integer:: j = 1
    Dim temp_str As String
    
    TheExec.DataManager.DecomposePinList pin, PowerPins(), pinCnt

    If TheExec.Specs.DC.Contains(PowerPins(0) & "_VAR_H") Then 'Carter, 20190502
        temp_str = "_VAR_H"
    Else
        temp_str = "_VAR"
    End If

    ReDim VMain(pinCnt - 1) As Double
    For Each powerPin In PowerPins
        TempString = powerPin & temp_str
        VMain(i) = TheExec.Specs.DC.item(TempString).ContextValue
        
        'get Ifold limit spec value
        TempString = powerPin & "_Ifold_GLB"
        Irange = TheExec.Specs.Globals(TempString).ContextValue
        
        'auto calculate steps
        step = VMain(i) / 0.1 '0.1v per step
        If step = 0 Then step = 10 'default value
        If step > PreStep Then PreStep = step   'calculate largest ramp up steps from all powers in the same sequence
        
        RiseTime = step * ms
        i_meter_rng = Irange
    
        With TheHdw.DCVS.pins(powerPin)
            .mode = tlDCVSModeVoltage
            .SetCurrentRanges Irange, i_meter_rng
            '.Meter.mode = tlDCVSMeterCurrent
            .CurrentRange.value = Irange
            .CurrentLimit.Source.FoldLimit.level.value = Irange
            ''202107
            If LCase(glb_TesterType) = "jaguar" Then
                .Meter.CurrentRange = Irange
            End If
        End With
        
        If debugPrintEnable = True Then    'debugprint
            TheExec.Datalog.WriteComment "print: Pin " & FormatNumericDatalog(powerPin, 30, False) & ", Vmain " & Format(VMain(i), "0.000") & " V, Irange " & FormatNumericDatalog(Format(Irange, "0.000"), 7, True) & " A, Step " & FormatNumericDatalog(step, 2, True) & ", RiseTime " & FormatNumericDatalog(RiseTime * 1000, 2, True) & " ms" & ", PowerSequence " & FormatNumericDatalog(PowerSequence, 3, True)
        End If
        
        i = i + 1
    Next powerPin
    
    step = PreStep
    RiseTime = step * ms
    stepT = RiseTime / step
    
    With TheHdw.DCVS.pins(pin)
        .Connect
        .Voltage.Main = 0
        TheHdw.Wait wait_before_gate
        .Gate = True
    End With
    

''Pwr On Ramp up slew-rate control============================
    For j = 1 To step
        i = 1
        For Each powerPin In PowerPins
            setV = j * VMain(i) / step
            TheHdw.DCVS.pins(powerPin).Voltage.Main = setV
''            If DebugPrintEnable = True Then
''                TheExec.Datalog.WriteComment "  Curr_" & PowerPin & " Pwr Up Voltage (" & CStr(i) & ") : " & Format(setV, "0.000") & " V"
''            End If
            i = i + 1
        Next powerPin
        
        TheHdw.Wait stepT
        
    Next j
''============================================================

    TheHdw.Wait wait_after_gate
  
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_DCVS_PPMU", "DCVS_PowerOn_I_Meter_Parallel") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function DCVS_PowerOff_I_Meter(pin As String, v As Double, i_rng As Double, wait_before_gate As Double, wait_after_gate As Double, Steps As Integer, FallTime As Double, Optional debugPrintEnable As Boolean = False)
'set Force voltage and Current/Meter Range
''===============================================
''Description
''__                              __
''  |__
''     |_>|  |<--stepT             v
''        |__          __
''           |__       __ stepV   __
''|<-- steps -->
''|<-FallTime ->
''===============================================
    On Error GoTo errHandler
    
    Dim i_meter_rng As Double   'meter range
    Dim setV As Double          'current voltage
    Dim StepV As Double         'step voltage
    Dim stepT As Double         'step time
    
    Dim i As Integer
    Dim stepsm As Integer
    
    i_meter_rng = i_rng
    StepV = v / Steps
    stepT = FallTime / Steps
    
    With TheHdw.DCVS.pins(pin)
        .Connect
        .mode = tlDCVSModeVoltage
        .Voltage.Main = v
        .SetCurrentRanges i_rng, i_meter_rng
        .Meter.mode = tlDCVSMeterCurrent
        .CurrentRange.value = i_rng
        .CurrentLimit.Source.FoldLimit.level.value = i_rng
        ''202107
        If LCase(glb_TesterType) = "jaguar" Then
            .Meter.CurrentRange = i_rng
        End If
        
        TheHdw.Wait wait_before_gate   'wait for relay connect
        
        .Gate = True
    End With
        
    ''Pwr On Ramp Down slew-rate control============================
    stepsm = Steps - 1
    For i = 0 To stepsm
        setV = v - (i * StepV)
        TheHdw.DCVS.pins(pin).Voltage.Main = setV
        
        If debugPrintEnable = True Then
            TheExec.Datalog.WriteComment "  Curr_" & pin & " Pwr Down Voltage (" & CStr(i) & ") : " & CStr(setV) & " V"
        End If
        
        TheHdw.Wait stepT   'wait step time
    Next i
    
    setV = 0    'final step, return to 0V anyway
    TheHdw.DCVS.pins(pin).Voltage.Main = setV
    
    If debugPrintEnable = True Then
        TheExec.Datalog.WriteComment "  Curr_" & pin & " Pwr Down Voltage (" & CStr(i) & ") : " & CStr(setV) & " V"
    End If
    ''==============================================================
    
    TheHdw.Wait wait_after_gate

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_DCVS_PPMU", "DCVS_PowerOff_I_Meter") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' [20240311][T-All][Clyde] Modularize Set current range for DCVI
Public Function SetCurrentRangeDCVI(pin As String, Irange As Double)
On Error GoTo errHandler

    functionName = "SetCurrentRangeDCVI"
    
    With TheHdw.DCVI.pins(pin)
        .mode = tlDCVIModeVoltage
        .SetCurrentAndRange Irange, Irange
        '.CurrentRange.value = Irange(i)       'T-Col TTR approve by Si -- 230413
        .Current = Irange
        .Meter.CurrentRange = Irange
    End With
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, moduleName, functionName)
    If AbortTest Then Exit Function Else Resume Next
End Function

' [20240311][T-All][Clyde] Modularize Set current range for DCVS
Public Function SetCurrentRangeDCVS(pin As String, Irange As Double)
On Error GoTo errHandler

    functionName = "SetCurrentRangeDCVS"
    
    With TheHdw.DCVS.pins(pin)
        .mode = tlDCVSModeVoltage
        .SetCurrentRanges Irange, Irange
        '.Meter.mode = tlDCVSMeterCurrent
        '.CurrentRange.value = Irange(i)       'T-Col TTR approve by Si -- 230413
        .CurrentLimit.Source.FoldLimit.level.value = Irange
        If LCase(glb_TesterType) = "jaguar" Then
            .Meter.CurrentRange = Irange
        End If
    End With
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, moduleName, functionName)
    If AbortTest Then Exit Function Else Resume Next
End Function

' [20240311][T-All][Clyde] RampingApplyPower function for Power up/down function
Public Function RampingApplyPower(pin() As String, baseVoltage() As Double, TargetVoltage() As Double, step() As Integer, WaitTime As Double, _
                    Optional outputDCVS As DCVSoutput = dKeep)
On Error GoTo errHandler
    Dim i As Integer
    Dim j As Integer
    Dim stepGrp() As Double
    Dim targetVoltageGrp() As Double
    Dim baseVoltageGrp() As Double
    Dim pinGrp() As String
    Dim grpCnt As Integer
    Dim maxStep As Integer
    Dim ForceVoltage As Double
    Dim haveRamping As Boolean
    Dim InstPins() As New instrumentUtility
    
    functionName = "RampingApplyPower"
    grpCnt = 0
    maxStep = 0
    
    If UBound(pin) <> UBound(baseVoltage) And UBound(pin) <> UBound(TargetVoltage) And UBound(pin) <> UBound(step) Then
        Call Print_Error_Message(Error_Info, moduleName, functionName, "Array not equal, please check")
        TheExec.flow.TestLimit -1, 0, 0
        Exit Function
    Else
        ReDim stepGrp(0) As Double
        ReDim targetVoltageGrp(0) As Double
        ReDim baseVoltageGrp(0) As Double
        ReDim pinGrp(0) As String
        For i = 0 To UBound(pin)
            If grpCnt = 0 Then
                stepGrp(grpCnt) = step(i)
                maxStep = step(i)
                targetVoltageGrp(grpCnt) = TargetVoltage(i)
                baseVoltageGrp(grpCnt) = baseVoltage(i)
                pinGrp(grpCnt) = pin(i)
                grpCnt = grpCnt + 1
            Else
                For j = 0 To UBound(pinGrp)
                    If stepGrp(j) = step(i) And targetVoltageGrp(j) = TargetVoltage(i) And baseVoltageGrp(j) = baseVoltage(i) Then
                        haveRamping = True
                        Exit For
                    End If
                Next j
                If haveRamping Then
                    pinGrp(j) = CombineStringList(pinGrp(j), pin(i))
                    haveRamping = False
                Else
                    
                    ReDim Preserve stepGrp(grpCnt) As Double
                    ReDim Preserve targetVoltageGrp(grpCnt) As Double
                    ReDim Preserve baseVoltageGrp(grpCnt) As Double
                    ReDim Preserve pinGrp(grpCnt) As String
                    stepGrp(grpCnt) = step(i)
                    targetVoltageGrp(grpCnt) = TargetVoltage(i)
                    baseVoltageGrp(grpCnt) = baseVoltage(i)
                    pinGrp(grpCnt) = pin(i)
                    grpCnt = grpCnt + 1
                    If step(i) > maxStep Then maxStep = step(i)
                End If
            End If
        Next i
    End If
    
    ReDim InstPins(grpCnt - 1) As New instrumentUtility
    For i = 0 To grpCnt - 1
        InstPins(i).Initialize pinGrp(i)
    Next i
    
    For i = 0 To maxStep - 1
        For j = 0 To UBound(pinGrp)
            If i < stepGrp(j) Then
                ForceVoltage = (i + 1) * (targetVoltageGrp(j) - baseVoltageGrp(j)) / stepGrp(j) + baseVoltageGrp(j)
                InstPins(j).ApplyPower ForceVoltage, fVoltage, gKeep, KeepStatus, outputDCVS
            End If
        Next j
        TheHdw.Wait WaitTime
    Next i
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, moduleName, functionName)
    If AbortTest Then Exit Function Else Resume Next
End Function

''20230223: Modidfied to process pin classify by Host and Dot .
' [20230524][All][Si] add TTR reduce .CurrentRange.value
' [20240311][T-All][Clyde] TTR
Public Function PowerOnOff_I_Meter_Parallel(pins As String, waitBeforeGate As Double, waitAfterGate As Double, pwrSequence As Integer, isUp As Boolean, Optional debugPrintEnable As Boolean = False)
On Error GoTo errHandler
    Dim PwrPin() As String
    Dim pinCnt As Long
    Dim tmp As String
    Dim i As Integer
    Dim TargetVoltage() As Double
    Dim baseVoltage() As Double
    Dim Irange As Double
    Dim step() As Integer
    Dim RiseTime As Double
    Dim InstPins As New instrumentUtility
    
    functionName = "PowerOnOff_I_Meter_Parallel"
    InstPins.Initialize pins
    TheExec.DataManager.DecomposePinList pins, PwrPin(), pinCnt
    
    ReDim TargetVoltage(pinCnt - 1) As Double
    ReDim step(pinCnt - 1) As Integer
    ReDim baseVoltage(pinCnt - 1) As Double
    
    For i = 0 To pinCnt - 1
        'get Ifold limit spec value
        tmp = PwrPin(i) & "_Ifold_GLB"
        tmp = Replace(tmp, "_DCVI", "")
        tmp = Replace(tmp, "_DCVS", "")
        Irange = TheExec.Specs.Globals(tmp).ContextValue
        
        'get Target voltage value
        tmp = PwrPin(i) & "_VAR"
        tmp = Replace(tmp, "_DCVI", "")
        tmp = Replace(tmp, "_DCVS", "")
                
        ' set current range
        If gl_GetInstrumentType_Dic(LCase(PwrPin(i))) Like "*DCVS*" Then
            SetCurrentRangeDCVS PwrPin(i), Irange
            If Not isUp Then
                TargetVoltage(i) = 0
                baseVoltage(i) = TheHdw.DCVS.pins(PwrPin(i)).Voltage.Main.value
            End If
            step(i) = (TargetVoltage(i) - baseVoltage(i)) / 0.1 '0.1v per step
        ElseIf gl_GetInstrumentType_Dic(LCase(PwrPin(i))) Like "*DCVI*" Then
            SetCurrentRangeDCVI PwrPin(i), Irange
            If Not isUp Then
                TargetVoltage(i) = 0
                baseVoltage(i) = TheHdw.DCVI.pins(PwrPin(i)).Voltage.value
            End If
        End If
        
        If isUp Then
            TargetVoltage(i) = TheExec.Specs.DC(tmp).ContextValue
            baseVoltage(i) = 0
        End If
        
        step(i) = Abs(TargetVoltage(i) - baseVoltage(i)) / 0.1 '0.1v per step
        RiseTime = step(i) * ms
        
        If debugPrintEnable Then
            If isUp Then
                TheExec.Datalog.WriteComment PowerUpDownPrintDatalogFormat(PwrPin(i), TargetVoltage(i), Irange, step(i), RiseTime, pwrSequence, True)
            Else
                TheExec.Datalog.WriteComment PowerUpDownPrintDatalogFormat(PwrPin(i), baseVoltage(i), Irange, step(i), RiseTime, pwrSequence, False)
            End If
        End If
    Next i
    
    If isUp Then
        ' Apply 0 voltage
        InstPins.ApplyPower 0, fVoltage, gOn, Connect, dVmain, False
        'Call ApplyPower(Pins, 0, fVoltage, GateOn, Connect, dVmain, False)
    Else
        InstPins.ApplyPower 0, fNone, gOn, Connect, dVmain, False
        'Call ApplyPower(Pins, 0, fNone, GateOn, Connect, dVmain, False)
    End If
    
    TheHdw.Wait waitBeforeGate
    Call RampingApplyPower(PwrPin, baseVoltage, TargetVoltage, step, 0.001, dVmain)
    TheHdw.Wait waitAfterGate
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, moduleName, functionName)
    If AbortTest Then Exit Function Else Resume Next
End Function
