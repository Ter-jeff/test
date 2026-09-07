Attribute VB_Name = "LIB_Digital_Debug"
#Const isUFP = True
Option Explicit

Public Const LVCC_boundary_Switch = 4 '1~10 means only get fail log at LVCC boundary with how many times
                                      '0 means get fail log with full search range
Public Const Shmoo_faillog_test_number = 100

Enum Shmoo_direction_enum
    High_to_Low = 1
    Low_to_High = 2
End Enum

Public Function FailingBoundaryDatalog_Func_Multi_Power(Power_Search_String As String, _
                    Shmoo_LotID As String, Shmoo_wafer As String, Shmoo_X As String, Shmoo_Y As String, Shmoo_Pattern As String, _
                                     Shmoo_status As String, Direction As Shmoo_direction_enum, CurrSite As Variant) As Long
On Error GoTo errHandler
'Shmoo with faillog capture version 1.0 careated by JT 2014/02/20.

    Dim PinCnt As Long
    Dim PinAry() As String
    Dim i As Integer
    Dim k As Integer
    Dim j As Integer
    Dim Fail_log_cnt As Integer
    Dim patternArray() As String
    Dim PowerV As Double
    Dim p As Integer
    Dim Org_Test_Number As Long
    Dim current_site As Integer
    Dim Timelist As String
    Dim TimeGroup() As String
    Dim CurrTiming As Variant
    Dim TimeDomainlist As String
    Dim TimeDomaingroup() As String
    Dim CurrTimeDomain As Variant
    Dim TimeDomainIn As String
    
    TheExec.Datalog.WriteComment "***************** Shmoo fail log capture start *****************"
    TheExec.Datalog.WriteComment vbNullString
    TheExec.Datalog.WriteComment "Lot : " & Shmoo_LotID
    TheExec.Datalog.WriteComment "Wafer :" & Shmoo_wafer
    TheExec.Datalog.WriteComment "Die X :" & Shmoo_X
    TheExec.Datalog.WriteComment "Die Y :" & Shmoo_Y
    TheExec.Datalog.WriteComment "Pattern :" & Shmoo_Pattern
    TheExec.Datalog.WriteComment "Shmoo status :" & Shmoo_status
    TheExec.Datalog.WriteComment vbNullString

'list time ing and frerunning clock
    TimeDomainlist = thehdw.Digital.Timing.TimeDomainlist
    
    TimeDomaingroup = Split(TimeDomainlist, ",")
    
    For Each CurrTimeDomain In TimeDomaingroup
        
        If CStr(CurrTimeDomain) = "All" Then
            TimeDomainIn = ""
        Else
            TimeDomainIn = CStr(CurrTimeDomain)
        End If
        
        Timelist = thehdw.Digital.TimeDomains(TimeDomainIn).Timing.TimeSetNameList
        'TimeGroup
        TimeGroup = Split(Timelist, ",")
        For Each CurrTiming In TimeGroup
            If CurrTiming = "" Then Exit For
            TheExec.Datalog.WriteComment "Time Doamin : " & CurrTimeDomain & ", TimeSet : " & CStr(CurrTiming) & " = " & (1 / thehdw.Digital.TimeDomains(TimeDomainIn).Timing.period(CStr(CurrTiming))) / 1000000 & " Mhz"

        Next CurrTiming
    Next CurrTimeDomain

    'Add for XI0 free running clk
    'TheExec.Datalog.WriteComment "  FreeRunFreq : " & thehdw.DIB.SupportBoardClock.Frequency / 1000000 & " Mhz , clock_Vih: " & thehdw.DIB.SupportBoardClock.Vih & " v , clock_Vil: " & thehdw.DIB.SupportBoardClock.vil & " v"
    'TheExec.Datalog.WriteComment "FreeRunFreq : " & TheHdw.DIB.SupportBoardClock.Frequency / 1000000 & " Mhz" ', clock_Vih: " & clock_Vih_debug & " v , clock_Vil: " & clock_Vil_debug & " v"
    TheExec.Datalog.WriteComment vbNullString


    Dim power_list() As String
    Dim Power_number As Integer
    Dim Power_pins(20) As String
    Dim Power_RangeA(20) As Double
    Dim Power_RangeB(20) As Double
    Dim Power_StepSize(20) As Double
    Dim power_Temp() As String
    Dim Power_range_temp() As String
    Dim Shmoo_steps As Double
    Dim axis_type As tlDevCharShmooAxis
    Dim SetupName As String
    Dim VmainOrValt As String
        
    power_list = Split(Power_Search_String, ",")
    Power_number = UBound(power_list)
        
    For i = 0 To Power_number
        power_Temp() = Split(power_list(i), "=")
        Power_range_temp() = Split(power_Temp(1), ":")
        Power_pins(i) = power_Temp(0)
        Power_RangeA(i) = CDbl(Power_range_temp(0))
        Power_RangeB(i) = CDbl(Power_range_temp(1))
        Power_StepSize(i) = CDbl(Power_range_temp(2))
        Shmoo_steps = Abs(Power_RangeA(i) - Power_RangeB(i)) / Abs(Power_StepSize(i))
        'Power_setting()
    Next i

    k = Shmoo_steps
    Fail_log_cnt = 1

    SetupName = TheExec.DevChar.Setups.ActiveSetupName
    VmainOrValt = LCase(TheExec.DevChar.Setups.item(SetupName).Shmoo.Axes.item(axis_type).Parameter.name)

    For j = 0 To k
        'loop power by step
        If Direction = Low_to_High Then
            For i = 0 To Power_number
                If Power_RangeA(i) < Power_RangeB(i) Then
                    If VmainOrValt = "vmain" Then
                        thehdw.DCVS.Pins(Power_pins(i)).Voltage.Main.value = Power_RangeA(i) + Abs(Power_StepSize(i)) * j
                    ElseIf VmainOrValt = "valt" Then
                        thehdw.DCVS.Pins(Power_pins(i)).Voltage.Alt.value = Power_RangeA(i) + Abs(Power_StepSize(i)) * j
                    End If
                Else
                    If VmainOrValt = "vmain" Then
                        thehdw.DCVS.Pins(Power_pins(i)).Voltage.Main.value = Power_RangeB(i) + Abs(Power_StepSize(i)) * j
                    ElseIf VmainOrValt = "valt" Then
                        thehdw.DCVS.Pins(Power_pins(i)).Voltage.Alt.value = Power_RangeB(i) + Abs(Power_StepSize(i)) * j
                    End If
                End If
            Next i
        Else
            For i = 0 To Power_number
                If Power_RangeA(i) > Power_RangeB(i) Then
                    If VmainOrValt = "vmain" Then
                        thehdw.DCVS.Pins(Power_pins(i)).Voltage.Main.value = Power_RangeA(i) - Abs(Power_StepSize(i)) * j
                    ElseIf VmainOrValt = "valt" Then
                        thehdw.DCVS.Pins(Power_pins(i)).Voltage.Alt.value = Power_RangeA(i) - Abs(Power_StepSize(i)) * j
                    End If
                Else
                    If VmainOrValt = "vmain" Then
                        thehdw.DCVS.Pins(Power_pins(i)).Voltage.Main.value = Power_RangeB(i) - Abs(Power_StepSize(i)) * j
                    ElseIf VmainOrValt = "valt" Then
                        thehdw.DCVS.Pins(Power_pins(i)).Voltage.Alt.value = Power_RangeB(i) - Abs(Power_StepSize(i)) * j
                    End If
                End If
            Next i
        End If
        
        For i = 0 To Power_number
            If VmainOrValt = "vmain" Then
                TheExec.Datalog.WriteComment Power_pins(i) + "= " & CStr(Format(thehdw.DCVS.Pins(Power_pins(i)).Voltage.Main.value, "0.000"))
                TheExec.Datalog.WriteComment ("VOLTAGE=") + CStr(Format(thehdw.DCVS.Pins(Power_pins(i)).Voltage.Main.value, "0.000")) + ("V")
            ElseIf VmainOrValt = "valt" Then
                TheExec.Datalog.WriteComment Power_pins(i) + "= " & CStr(Format(thehdw.DCVS.Pins(Power_pins(i)).Voltage.Alt.value, "0.000"))
                TheExec.Datalog.WriteComment ("VOLTAGE=") + CStr(Format(thehdw.DCVS.Pins(Power_pins(i)).Voltage.Alt.value, "0.000")) + ("V")
            End If
        Next i
        
        thehdw.Wait 0.002
        thehdw.Patterns(Shmoo_Pattern).test pfAlways, 0
        
        For i = 0 To Power_number
            If VmainOrValt = "vmain" Then
                TheExec.Datalog.WriteComment Power_pins(i) + "= " & CStr(Format(thehdw.DCVS.Pins(Power_pins(i)).Voltage.Main.value, "0.000"))
                TheExec.Datalog.WriteComment ("VOLTAGE=") + CStr(Format(thehdw.DCVS.Pins(Power_pins(i)).Voltage.Main.value, "0.000")) + ("V")
            ElseIf VmainOrValt = "valt" Then
                TheExec.Datalog.WriteComment Power_pins(i) + "= " & CStr(Format(thehdw.DCVS.Pins(Power_pins(i)).Voltage.Alt.value, "0.000"))
                TheExec.Datalog.WriteComment ("VOLTAGE=") + CStr(Format(thehdw.DCVS.Pins(Power_pins(i)).Voltage.Alt.value, "0.000")) + ("V")
            End If
        Next i
        
        If thehdw.Digital.Patgen.PatternBurstPassed(CurrSite) = False And LVCC_boundary_Switch = Fail_log_cnt Then Exit For
        If thehdw.Digital.Patgen.PatternBurstPassed(CurrSite) = False Then Fail_log_cnt = Fail_log_cnt + 1

    Next j

    TheExec.Datalog.WriteComment "***************** Shmoo fail log capture end *****************"

    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Debug", "FailingBoundaryDatalog_Func_Multi_Power")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function FailingDatalog_Lvcc_Boundary(Power_Search_String As String, _
                    Shmoo_LotID As String, Shmoo_wafer As String, Shmoo_X As String, Shmoo_Y As String, Shmoo_Pattern As String, _
                                     Shmoo_status As String, Direction As Shmoo_direction_enum, CurrSite As Variant, Optional RangeFrom As Double, Optional RangeTo As Double, Optional RangeSteps As Long, Optional RangeStepSize As Double) As Long
On Error GoTo errHandler
'Shmoo with faillog capture version 1.0 careated by JT 2014/02/20.

    Dim PinCnt As Long
    Dim PinAry() As String
    Dim i As Integer
    Dim k As Integer
    Dim j As Integer
    Dim Fail_log_cnt As Integer
    Dim patternArray() As String
    Dim PowerV As Double
    Dim p As Integer
    Dim Org_Test_Number As Long
    Dim current_site As Integer
    Dim Timelist As String
    Dim TimeGroup() As String
    Dim CurrTiming As Variant
    Dim TimeDomainlist As String
    Dim TimeDomaingroup() As String
    Dim CurrTimeDomain As Variant
    Dim TimeDomainIn As String
    Dim ShmooPatternSplit() As String
    Dim TestNumber As Long
    Dim inst_name As String
    Dim shmoopowerpin As String
    Dim Failed_Pins() As String
    Dim AllFailPins As String
    Dim OutputString As String
    
    
    ShmooPatternSplit() = Split(Shmoo_Pattern, ",")
    inst_name = UCase(TheExec.DataManager.instancename)
    
    Dim Context As String: Context = vbNullString
    Dim TimeSet_Str As String: TimeSet_Str = vbNullString
    
    Context = TheExec.Contexts.ActiveSelection
    TimeSet_Str = TheExec.Contexts(Context).Sheets.Timesets
    
    If TheExec.Flow.enableWord("FailPinsOnly") = False Then
        TheExec.Datalog.WriteComment "***************** Shmoo fail log capture start *****************"
        TheExec.Datalog.WriteComment vbNullString
        TheExec.Datalog.WriteComment "Lot : " & Shmoo_LotID
        TheExec.Datalog.WriteComment "Wafer :" & Shmoo_wafer
        TheExec.Datalog.WriteComment "Die X :" & Shmoo_X
        TheExec.Datalog.WriteComment "Die Y :" & Shmoo_Y
        TheExec.Datalog.WriteComment "Pattern :" & Shmoo_Pattern
        TheExec.Datalog.WriteComment "Shmoo status :" & Shmoo_status
        TheExec.Datalog.WriteComment vbNullString
        TheExec.Datalog.WriteComment "Activity Timeset Sheet :" & TimeSet_Str

'list time ing and frerunning clock
        TimeDomainlist = thehdw.Digital.Timing.TimeDomainlist
        
        TimeDomaingroup = Split(TimeDomainlist, ",")
        
        For Each CurrTimeDomain In TimeDomaingroup
             
            If CStr(CurrTimeDomain) = "All" Then
                TimeDomainIn = ""
            Else
                TimeDomainIn = CStr(CurrTimeDomain)
            End If
             
            Timelist = thehdw.Digital.TimeDomains(TimeDomainIn).Timing.TimeSetNameList
             'TimeGroup
            TimeGroup = Split(Timelist, ",")
            For Each CurrTiming In TimeGroup
                If CurrTiming = "" Then Exit For
                TheExec.Datalog.WriteComment "Time Domain : " & CurrTimeDomain & ", TimeSet : " & CStr(CurrTiming) & " = " & (1 / thehdw.Digital.TimeDomains(TimeDomainIn).Timing.period(CStr(CurrTiming))) / 1000000 & " Mhz"
            Next CurrTiming
        Next CurrTimeDomain
        
         '' add for XI0 free running clk
        'TheExec.Datalog.WriteComment "  FreeRunFreq : " & thehdw.DIB.SupportBoardClock.Frequency / 1000000 & " Mhz , clock_Vih: " & thehdw.DIB.SupportBoardClock.Vih & " v , clock_Vil: " & thehdw.DIB.SupportBoardClock.vil & " v"
        'TheExec.Datalog.WriteComment "FreeRunFreq : " & TheHdw.DIB.SupportBoardClock.Frequency / 1000000 & " Mhz" ', clock_Vih: " & clock_Vih_debug & " v , clock_Vil: " & clock_Vil_debug & " v"
        TheExec.Datalog.WriteComment vbNullString
    End If

    Dim power_list() As String
    Dim Power_number As Integer

    Dim Power_pins() As String
    Dim Power_RangeA(20) As Double
    Dim Power_RangeB(20) As Double
    Dim Power_StepSize(20) As Double
    Dim power_Temp() As String
    Dim Power_range_temp() As String
    Dim Shmoo_steps As Double
    Dim StepValue As Double

    power_list = Split(Power_Search_String, ",")
    Power_number = UBound(power_list)
    ReDim Power_pins(Power_number)
    
    For i = 0 To Power_number
         power_Temp() = Split(power_list(i), "=")
         Power_range_temp() = Split(power_Temp(1), ":")
         Power_pins(i) = power_Temp(0)
    Next i
    
    If TheExec.Flow.enableWord("Find_shmoo_hole_low_power_twice") = True Then
       Shmoo_steps = 3
    Else
	    If RangeFrom > RangeTo Then
	       Shmoo_steps = ((Shmoo_Vcc_Min(CurrSite) - RangeTo) / RangeStepSize)
	    ElseIf RangeTo > RangeFrom Then
	       Shmoo_steps = (Shmoo_Vcc_Min(CurrSite) - RangeFrom) / RangeStepSize
	    Else
	       Shmoo_steps = 3
	    End If
    End if   
    If TheExec.Flow.enableWord("FailPinsOnly") = False Then
       k = Shmoo_steps
    Else
       k = 0
    End If
    Fail_log_cnt = 1

    For j = 0 To k
        
            'loop power by step
            
        For i = 0 To Power_number
            If TheExec.Flow.enableWord("FailPinsOnly") = True Then
                thehdw.DCVS.Pins(Power_pins(i)).Voltage.Main.value = Shmoo_Vcc_Min(CurrSite) - 0.005
            Else
                thehdw.DCVS.Pins(Power_pins(i)).Voltage.Main.value = Shmoo_Vcc_Min(CurrSite) - RangeStepSize * (j + 1)
            End If
        Next i
        If TheExec.Flow.enableWord("FailPinsOnly") = False Then
            StepValue = Fail_log_cnt * 3.125
            
            TheExec.Datalog.WriteComment "Power setup (Vmin- " & StepValue & "mV) "
            For i = 0 To Power_number
                TheExec.Datalog.WriteComment Power_pins(i) + "= " & CStr(Format(thehdw.DCVS.Pins(Power_pins(i)).Voltage.Main.value, "0.00000"))
                TheExec.Datalog.WriteComment ("VOLTAGE=") + CStr(Format(thehdw.DCVS.Pins(Power_pins(i)).Voltage.Main.value, "0.00000")) + ("V")
            Next i
        End If
            
        thehdw.Wait 0.002
            
        If TheExec.Flow.enableWord("FailPinsOnly") = True Then
            For i = 0 To UBound(ShmooPatternSplit)
                Call thehdw.Patterns(ShmooPatternSplit(i)).Load
                Call thehdw.Patterns(ShmooPatternSplit(i)).start
                thehdw.Digital.Patgen.HaltWait
                If thehdw.Digital.Patgen.PatternBurstPassed(CurrSite) = False Then
                    shmoopowerpin = Join(Power_pins, ",")
                    Failed_Pins() = thehdw.Digital.FailedPins(CurrSite)
                    AllFailPins = Join(Failed_Pins, ",")
                    OutputString = "[" & "FailPins" & "," & Shmoo_LotID & "-" & Shmoo_wafer & "," & Shmoo_X & "," & Shmoo_Y & "," & "Site" & CStr(CurrSite)
                    OutputString = OutputString & "," & inst_name & "," & "Pattern./" & ShmooPatternSplit(i) & "," & "ShmooPowerPin:" & shmoopowerpin & "," & "ApplyVoltage(Vmin-Guardband 5mV)" & "=" & CStr(Format((Shmoo_Vcc_Min(CurrSite) - 0.005), "0.00000"))
                    OutputString = OutputString & "," & "FailPins = " & UCase(AllFailPins)
                    TheExec.Datalog.WriteComment OutputString & "]"
                End If
            Next i
        Else
        
            Dim Temp_patary() As String
            Dim TempPat As Variant
            Temp_patary() = Split(Shmoo_Pattern, ",")
            For Each TempPat In Temp_patary
                TestNumber = TheExec.sites.item(CurrSite).TestNumber
                thehdw.Patterns(TempPat).test pfNever, 0
                If (thehdw.Digital.Patgen.PatternBurstPassed(CurrSite) = True) Then
                    Call TheExec.Datalog.WriteFunctionalResult(CurrSite, TestNumber, logTestPass)
                Else
                    Call TheExec.Datalog.WriteFunctionalResult(CurrSite, TestNumber, logTestFail)
                End If
                TestNumber = TestNumber + 1
                TheExec.sites.item(CurrSite).TestNumber = TestNumber
            Next TempPat
            
            TheExec.Datalog.WriteComment "                                                "
  
            If thehdw.Digital.Patgen.PatternBurstPassed(CurrSite) = False And LVCC_boundary_Switch = Fail_log_cnt Then GoTo Endfor
            If thehdw.Digital.Patgen.PatternBurstPassed(CurrSite) = False Then Fail_log_cnt = Fail_log_cnt + 1
         End If
    Next j
Endfor:

If TheExec.Flow.enableWord("FailPinsOnly") = False Then
TheExec.Datalog.WriteComment "***************** Shmoo fail log/Pins capture end *****************"
End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Debug", "FailingDatalog_Lvcc_Boundary")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function FailingDatalog_Hvcc_Boundary(Power_Search_String As String, _
                    Shmoo_LotID As String, Shmoo_wafer As String, Shmoo_X As String, Shmoo_Y As String, Shmoo_Pattern As String, _
                                     Shmoo_status As String, Direction As Shmoo_direction_enum, CurrSite As Variant, Optional RangeFrom As Double, Optional RangeTo As Double, Optional RangeSteps As Long, Optional RangeStepSize As Double) As Long
On Error GoTo errHandler
'Shmoo with faillog capture version 1.0 careated by JT 2014/02/20.

    Dim PinCnt As Long
    Dim PinAry() As String
    Dim i As Integer
    Dim k As Integer
    Dim j As Integer
    Dim Fail_log_cnt As Integer
    Dim patternArray() As String
    Dim PowerV As Double
    Dim p As Integer
    Dim Org_Test_Number As Long
    Dim current_site As Integer
    Dim Timelist As String
    Dim TimeGroup() As String
    Dim CurrTiming As Variant
    Dim TimeDomainlist As String
    Dim TimeDomaingroup() As String
    Dim CurrTimeDomain As Variant
    Dim TimeDomainIn As String
    Dim ShmooPatternSplit() As String
    Dim TestNumber As Long
    Dim inst_name As String
    Dim shmoopowerpin As String
    Dim Failed_Pins() As String
    Dim AllFailPins As String
    Dim OutputString As String
    
    
    ShmooPatternSplit() = Split(Shmoo_Pattern, ",")
    inst_name = UCase(TheExec.DataManager.instancename)
    
    Dim Context As String: Context = vbNullString
    Dim TimeSet_Str As String: TimeSet_Str = vbNullString
    Context = TheExec.Contexts.ActiveSelection
    TimeSet_Str = TheExec.Contexts(Context).Sheets.Timesets
    
    If TheExec.Flow.enableWord("FailPinsOnly") = False Then
        TheExec.Datalog.WriteComment "***************** Shmoo fail log capture start *****************"
        TheExec.Datalog.WriteComment vbNullString
        TheExec.Datalog.WriteComment "Lot : " & Shmoo_LotID
        TheExec.Datalog.WriteComment "Wafer :" & Shmoo_wafer
        TheExec.Datalog.WriteComment "Die X :" & Shmoo_X
        TheExec.Datalog.WriteComment "Die Y :" & Shmoo_Y
        TheExec.Datalog.WriteComment "Pattern :" & Shmoo_Pattern
        TheExec.Datalog.WriteComment "Shmoo status :" & Shmoo_status
        TheExec.Datalog.WriteComment vbNullString
        TheExec.Datalog.WriteComment "Activity Timeset Sheet :" & TimeSet_Str
'list time ing and frerunning clock
        TimeDomainlist = thehdw.Digital.Timing.TimeDomainlist
        
        TimeDomaingroup = Split(TimeDomainlist, ",")
        
        For Each CurrTimeDomain In TimeDomaingroup
            
            If CStr(CurrTimeDomain) = "All" Then
                TimeDomainIn = ""
            Else
                TimeDomainIn = CStr(CurrTimeDomain)
            End If
            Timelist = thehdw.Digital.TimeDomains(TimeDomainIn).Timing.TimeSetNameList
            'TimeGroup
            TimeGroup = Split(Timelist, ",")
            For Each CurrTiming In TimeGroup
                If CurrTiming = vbNullString Then Exit For
                TheExec.Datalog.WriteComment "Time Domain : " & CurrTimeDomain & ", TimeSet : " & CStr(CurrTiming) & " = " & (1 / thehdw.Digital.TimeDomains(TimeDomainIn).Timing.period(CStr(CurrTiming))) / 1000000 & " Mhz"
            Next CurrTiming
        Next CurrTimeDomain

        'Add for XI0 free running clk
        'TheExec.Datalog.WriteComment "  FreeRunFreq : " & thehdw.DIB.SupportBoardClock.Frequency / 1000000 & " Mhz , clock_Vih: " & thehdw.DIB.SupportBoardClock.Vih & " v , clock_Vil: " & thehdw.DIB.SupportBoardClock.vil & " v"
        'TheExec.Datalog.WriteComment "FreeRunFreq : " & TheHdw.DIB.SupportBoardClock.Frequency / 1000000 & " Mhz" ', clock_Vih: " & clock_Vih_debug & " v , clock_Vil: " & clock_Vil_debug & " v"
        TheExec.Datalog.WriteComment vbNullString
    End If

    Dim power_list() As String
    Dim Power_number As Integer

    Dim Power_pins() As String
    Dim Power_RangeA(20) As Double
    Dim Power_RangeB(20) As Double
    Dim Power_StepSize(20) As Double
    Dim power_Temp() As String
    Dim Power_range_temp() As String
    Dim Shmoo_steps As Double
    Dim StepValue As Double

    power_list = Split(Power_Search_String, ",")
    Power_number = UBound(power_list)
    ReDim Power_pins(Power_number)
        
    For i = 0 To Power_number
         power_Temp() = Split(power_list(i), "=")
         Power_range_temp() = Split(power_Temp(1), ":")
         Power_pins(i) = power_Temp(0)
    Next i
        
        
    If RangeFrom > RangeTo Then
       Shmoo_steps = ((RangeFrom - Shmoo_Vcc_Max(CurrSite)) / RangeStepSize)
    ElseIf RangeTo > RangeFrom Then
       Shmoo_steps = (RangeTo - Shmoo_Vcc_Max(CurrSite)) / RangeStepSize
    Else
       Shmoo_steps = 3
    End If
       
    If TheExec.Flow.enableWord("FailPinsOnly") = False Then
       k = Shmoo_steps
    Else
       k = 0
    End If
    Fail_log_cnt = 1

    For j = 0 To k
        
            'loop power by step
    
        For i = 0 To Power_number
            If TheExec.Flow.enableWord("FailPinsOnly") = True Then
                thehdw.DCVS.Pins(Power_pins(i)).Voltage.Main.value = Shmoo_Vcc_Max(CurrSite) + 0.005
            Else
                thehdw.DCVS.Pins(Power_pins(i)).Voltage.Main.value = Shmoo_Vcc_Max(CurrSite) + RangeStepSize * (j + 1)
            End If
        Next i
                
        If TheExec.Flow.enableWord("FailPinsOnly") = False Then

            StepValue = Fail_log_cnt * 3.125
            
            TheExec.Datalog.WriteComment "Power setup (Vmax+ " & StepValue & "mV) "
            For i = 0 To Power_number
                TheExec.Datalog.WriteComment Power_pins(i) + "= " & CStr(Format(thehdw.DCVS.Pins(Power_pins(i)).Voltage.Main.value, "0.00000"))
                TheExec.Datalog.WriteComment ("VOLTAGE=") + CStr(Format(thehdw.DCVS.Pins(Power_pins(i)).Voltage.Main.value, "0.00000")) + ("V")
            Next i
        End If
            
        thehdw.Wait 0.002
            
        If TheExec.Flow.enableWord("FailPinsOnly") = True Then
            For i = 0 To UBound(ShmooPatternSplit)
                Call thehdw.Patterns(ShmooPatternSplit(i)).Load
                Call thehdw.Patterns(ShmooPatternSplit(i)).start
                thehdw.Digital.Patgen.HaltWait
                If thehdw.Digital.Patgen.PatternBurstPassed(CurrSite) = False Then
                    shmoopowerpin = Join(Power_pins, ",")
                    Failed_Pins() = thehdw.Digital.FailedPins(CurrSite)
                    AllFailPins = Join(Failed_Pins, ",")
                    OutputString = "[" & "FailPins" & "," & Shmoo_LotID & "-" & Shmoo_wafer & "," & Shmoo_X & "," & Shmoo_Y & "," & "Site" & CStr(CurrSite)
                    OutputString = OutputString & "," & inst_name & "," & "Pattern./" & ShmooPatternSplit(i) & "," & "ShmooPowerPin:" & shmoopowerpin & "," & "ApplyVoltage(Vmax+Guardband 5mV)" & "=" & CStr(Format((Shmoo_Vcc_Max(CurrSite) + 0.005), "0.00000"))
                    OutputString = OutputString & "," & "FailPins = " & UCase(AllFailPins)
                    TheExec.Datalog.WriteComment OutputString & "]"
                End If
           Next i
        Else
            Dim Temp_patary() As String
            Dim TempPat As Variant
            Temp_patary() = Split(Shmoo_Pattern, ",")
            For Each TempPat In Temp_patary
                TestNumber = TheExec.sites.item(CurrSite).TestNumber
                thehdw.Patterns(TempPat).test pfNever, 0
                If (thehdw.Digital.Patgen.PatternBurstPassed(CurrSite) = True) Then
                    Call TheExec.Datalog.WriteFunctionalResult(CurrSite, TestNumber, logTestPass)
                Else
                    Call TheExec.Datalog.WriteFunctionalResult(CurrSite, TestNumber, logTestFail)
                End If
                TestNumber = TestNumber + 1
                TheExec.sites.item(CurrSite).TestNumber = TestNumber
            Next TempPat
            
            TheExec.Datalog.WriteComment "                                                "
            
            If thehdw.Digital.Patgen.PatternBurstPassed(CurrSite) = False And LVCC_boundary_Switch = Fail_log_cnt Then GoTo Endfor
            If thehdw.Digital.Patgen.PatternBurstPassed(CurrSite) = False Then Fail_log_cnt = Fail_log_cnt + 1
        End If
    Next j
Endfor:

If TheExec.Flow.enableWord("FailPinsOnly") = False Then
TheExec.Datalog.WriteComment "***************** Shmoo fail log/Pins capture end *****************"
End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Debug", "FailingDatalog_Hvcc_Boundary")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function FailingDatalog_HLvcc_Boundary_SELSRM(Power_Search_String As String, _
                    Shmoo_LotID As String, Shmoo_wafer As String, Shmoo_X As String, Shmoo_Y As String, Shmoo_Pattern As String, _
                                     Shmoo_status As String, Direction As Shmoo_direction_enum, CurrSite As Variant, Optional RangeFrom As Double, Optional RangeTo As Double, Optional RangeSteps As Long, Optional RangeStepSize As Double, Optional dssc_pat As String) As Long
On Error GoTo errHandler
'Shmoo SELSRM HLVCC boundary failure log by Cebu 201807
    
    Dim PinCnt As Long
    Dim PinAry() As String
    Dim i As Integer
    Dim k As Integer
    Dim j As Integer
    Dim m As Integer
        Dim L As Integer
    Dim Fail_log_cnt As Integer
    Dim patternArray() As String
    Dim PowerV As Double
    Dim p As Integer
    Dim Org_Test_Number As Long
    Dim current_site As Integer
    Dim Timelist As String
    Dim TimeGroup() As String
    Dim CurrTiming As Variant
    Dim TimeDomainlist As String
    Dim TimeDomaingroup() As String
    Dim CurrTimeDomain As Variant
    Dim TimeDomainIn As String
    Dim ShmooPatternSplit() As String
    Dim TestNumber As Long
    Dim inst_name As String
    Dim shmoopowerpin As String
    Dim Failed_Pins() As String
    Dim AllFailPins As String
    Dim OutputString As String
    
    Dim nWire_port_ary() As String
    nWire_port_ary = Split(nWire_Ports_GLB, ",")
    Dim pat_count As Long, pat_array_count As Long
    Dim pat_array() As String, pat_array_nu() As String
    Dim power_list_dic() As String
    Dim power_count_dic As Long
    
    Dim power_list() As String
    Dim Power_number As Integer
    Dim Power_pins() As String
    Dim Power_RangeA(20) As Double
    Dim Power_RangeB(20) As Double
    Dim Power_StepSize(20) As Double
    Dim power_Temp() As String
    Dim Power_range_temp() As String
    Dim Shmoo_steps As Double
    Dim StepValue As Double
    
    Dim pat_list_array() As String
    Dim pat_list_count As Long
    Dim confirm_pat_load As Boolean
    ''''''''''''''''''''''''''''''''''''''
    Dim DC_Spec_Level As New PinListData   '''from Vmain or Valt
    Dim SELSRM_Rails As String
    Dim Shmoo_value As New PinListData
    Dim DigSrc_wav As New DSPWave
    Dim powerPin As String
    Dim pin_name() As String
    Dim n As Integer
    Dim DSSC_string As String
    Dim pat_str_ary() As String
    ''''''''''''''''''''''''''''''''''''''
    Dim patCnt As Long
    Dim Lvcc As Boolean:: Lvcc = True
    Dim dssc_pat_exist As Boolean:: dssc_pat = False
    Dim RET_InfoAry() As String
    Dim sub_RET_InfoAry() As String
    Dim RET_PatAry() As String
    Dim RET_WaitAry() As Double
    Dim RET_PLD As New PinListData
    Dim RET_PinAry() As String
    Dim RET_PinCnt As Long
    Dim Retention_Fail_Cycle As Boolean
    Dim RET_Power_StoreIDx As Long
    '20240531shmoo hole fail log
    Dim step_10, step_2 As Integer
    Dim Store_orginal_voltage As Double
        
    Retention_Fail_Cycle = False
    If Shmoo_status <> "" Then
       If UCase(Shmoo_status) Like "SHMOO HVCC" Then Lvcc = False
    End If
   
    Set RET_PLD = Nothing
    RET_Power_StoreIDx = 0
    Set DigSrc_wav = Nothing
    If InStr(Shmoo_Pattern, ",") > 0 Then
        ShmooPatternSplit() = Split(Shmoo_Pattern, ",")
    Else
        ShmooPatternSplit() = TheExec.DataManager.Raw.GetPatternsInSet(Shmoo_Pattern, patCnt)
    End If
    inst_name = UCase(TheExec.DataManager.instancename)
    
    If UCase(g_PR_Scenario) = "INIT_NV_PL_NV" And g_PR_Scenario <> "" Then
        RET_InfoAry = Split(g_Retention_Info, ",")
        ReDim RET_PatAry(UBound(RET_InfoAry)) As String
        ReDim RET_WaitAry(UBound(RET_InfoAry)) As Double
        For i = 0 To UBound(RET_InfoAry)
            sub_RET_InfoAry = Split(RET_InfoAry(i), "|")
            RET_PatAry(i) = sub_RET_InfoAry(0)
            RET_WaitAry(i) = sub_RET_InfoAry(1)
        Next i
        Retention_Fail_Cycle = True
    End If
    
    If digSrc_EQ_GB <> "" Then
        DigSrc_wav.CreateConstant 0, Len(digSrc_EQ_GB), DspLong
        Decide_DC_Level DC_Spec_Level, g_ApplyLevelTimingValt, g_ApplyLevelTimingVmain, BlockType_GB, TestType_GB
    End If
    'pin_name() = Split(g_VDDForce, ",")
    pin_name() = Split(g_ForceCond_VDD, ",")
'----------------------------------------------------------------------
    TheExec.Datalog.WriteComment "***************** Shmoo fail log capture start *****************"
    TheExec.Datalog.WriteComment vbNullString
    TheExec.Datalog.WriteComment "Lot : " & Shmoo_LotID
    TheExec.Datalog.WriteComment "Wafer :" & Shmoo_wafer
    TheExec.Datalog.WriteComment "Die X :" & Shmoo_X
    TheExec.Datalog.WriteComment "Die Y :" & Shmoo_Y
    TheExec.Datalog.WriteComment "Pattern :" & Shmoo_Pattern
    TheExec.Datalog.WriteComment "Shmoo status :" & Shmoo_status
    TheExec.Datalog.WriteComment vbNullString

    TimeDomainlist = thehdw.Digital.Timing.TimeDomainlist
    TimeDomaingroup = Split(TimeDomainlist, ",")

    For Each CurrTimeDomain In TimeDomaingroup
        If CStr(CurrTimeDomain) = "All" Then
            TimeDomainIn = ""
        Else
            TimeDomainIn = CStr(CurrTimeDomain)
        End If
        
        Timelist = thehdw.Digital.TimeDomains(TimeDomainIn).Timing.TimeSetNameList
        'TimeGroup
        TimeGroup = Split(Timelist, ",")
        For Each CurrTiming In TimeGroup
            If CurrTiming = "" Then Exit For
            If TimeDomainIn = "" Then
                TheExec.Datalog.WriteComment "Time Doamin : TimeDomainIn = N/A"
            Else
                TheExec.Datalog.WriteComment "Time Doamin : " & CurrTimeDomain & ", TimeSet : " & CStr(CurrTiming) & " = " & (1 / thehdw.Digital.TimeDomains(TimeDomainIn).Timing.period(CStr(CurrTiming))) / 1000000 & " Mhz"
            End If
        Next CurrTiming
    Next CurrTimeDomain
        TheExec.Datalog.WriteComment vbNullString

    power_list = Split(Power_Search_String, ",")
    Power_number = UBound(power_list)
    ReDim Power_pins(Power_number)

    For i = 0 To Power_number
        power_Temp() = Split(power_list(i), "=")
        Power_range_temp() = Split(power_Temp(1), ":")
        Power_pins(i) = power_Temp(0)
        powerPin = powerPin + "," + Power_pins(i)
        Shmoo_value.AddPin Power_pins(i)
    Next i
    powerPin = mid(powerPin, 2, Len(powerPin))

    If TheExec.TesterMode = testModeOffline Then
        Shmoo_Vcc_Min(CurrSite) = RangeTo + 0.3
        Shmoo_Vcc_Max(CurrSite) = RangeTo - 0.3
    End If

    If Lvcc = True Then
        If RangeFrom > RangeTo Then
            Shmoo_steps = Fix((Shmoo_Vcc_Min(CurrSite) - RangeTo) / RangeStepSize)
        ElseIf RangeTo > RangeFrom Then
            Shmoo_steps = Fix((Shmoo_Vcc_Min(CurrSite) - RangeFrom) / RangeStepSize)
        Else
            Shmoo_steps = 3
        End If
    Else
        If RangeFrom > RangeTo Then
            Shmoo_steps = Fix((RangeFrom - Shmoo_Vcc_Max(CurrSite)) / RangeStepSize)
        ElseIf RangeTo > RangeFrom Then
            Shmoo_steps = Fix((RangeTo - Shmoo_Vcc_Max(CurrSite)) / RangeStepSize)
        Else
            Shmoo_steps = 3
        End If
    End If

    If Not Shmoo_steps < 1 Then
        k = Shmoo_steps
    Else
        k = 0
    End If

    Fail_log_cnt = 1
    
    '220411 Modify for DFC 2 diff LVCC Collect Range
    '220420 Workable General DFC diff LVCC Collect Range
    '220420 HVCC add
    '220428 format change to SSizeXSNumY
    '220428 Diff DFC Size
    '220429 Combine 2 format in 1
    Dim DFC_Range() As String
    Dim inst_info() As String
    Dim log_const As Integer
    Dim DFC_Step As Double
    Dim const_Size As Boolean
    Dim infoCnt As Long
    Dim isDFCCond As Boolean
    inst_info() = Split(inst_name, "_")
'    ReDim Preserve Inst_info(11)
    isDFCCond = False
    For infoCnt = UBound(inst_info) To 0 Step -1
        If inst_info(infoCnt) Like "DFC*" Then
            Dim DFC_Temp_Num() As String
            Dim DFC_Temp_Size() As String
    
            If Lvcc = True Then
                If inst_info(infoCnt) Like "DFCM*" Then
                    DFC_Range = Split(inst_info(infoCnt), "M")
                    log_const = UBound(DFC_Range) - LBound(DFC_Range)
                    const_Size = False
                    isDFCCond = True
                ElseIf inst_info(infoCnt) Like "DFCS*" Then
                    DFC_Range() = Split(inst_info(infoCnt), "S")
                    DFC_Temp_Size = Split(DFC_Range(2), "IZE")
                    DFC_Temp_Num = Split(DFC_Range(3), "NUM")
                    log_const = 1
                    const_Size = True
                    isDFCCond = True
                Else
                    TheExec.Datalog.WriteComment "DFC Format mismatch."
                End If
        Else
            If inst_info(infoCnt) Like "DFCP*" Then
                DFC_Range = Split(inst_info(infoCnt), "P")
                log_const = UBound(DFC_Range) - LBound(DFC_Range)
                const_Size = False
                isDFCCond = True
            ElseIf inst_info(infoCnt) Like "DFCS*" Then
                DFC_Range() = Split(inst_info(infoCnt), "S")
                DFC_Temp_Size = Split(DFC_Range(2), "IZE")
                DFC_Temp_Num = Split(DFC_Range(3), "NUM")
                log_const = 1
                const_Size = True
                isDFCCond = True
            Else
                TheExec.Datalog.WriteComment "DFC Format mismatch."
            End If
        End If
'     Else
'         DFC_Step = RangeStepSize
'         log_const = 1
'         const_Size = True
        End If
        If isDFCCond = True Then Exit For
    Next infoCnt
    
    If isDFCCond = False Then
        DFC_Step = RangeStepSize
        log_const = 1
        const_Size = True
    End If

    Dim payload_len As Long
    Dim Fail_log_cnt_flag As Boolean
    Dim hramInfos As LVCC_VminBoundary
    Set hramInfos = New LVCC_VminBoundary
    'Chris added 220418
    'William edit 220419 general purpose
    'hramInfos.LvccVoltage = 0.005
    'hramInfos.LvccCount = 3
    '220428 Diff DFC Size
    Dim log_cnt As Integer
    If Not log_const = 0 Then
        For log_cnt = 0 To log_const
            ''20240531shmoo hole fail log
            'If isDFCCond = True Then
            '    If const_Size = True Then
            '        hramInfos.LvccVoltage = DFC_Temp_Size(1) / 1000
            '        hramInfos.LvccCount = DFC_Temp_Num(1) + 1
            '    Else
            '        hramInfos.LvccVoltage = DFC_Range(log_cnt + 1) / 1000
            '        hramInfos.LvccCount = 2
            '    End If
            'Else
            '    hramInfos.LvccVoltage = RangeStepSize
            '    hramInfos.LvccCount = 2
            'End If
            'If Lvcc = True Then
            '    hramInfos.LvccValue = Shmoo_Vcc_Min(CurrSite)
            'Else
            '    hramInfos.LvccValue = Shmoo_Vcc_Max(CurrSite)
            'End If
            'If Not k = 0 Then
            If (Not k = 0) Or Find_shmoo_hole_low_power_twice_hvcc_voltage(0)(CurrSite) <> 0 Then '20240531shmoo hole fail log
        '=============================================================================================================================
                For j = 0 To k                      '''''step loop
                    For m = 0 To UBound(ShmooPatternSplit)      '''''pat loop
                    '----------------------------------------------------------------------confirm init or pl
                        GetPatFromPatternSet CStr(ShmooPatternSplit(m)), pat_list_array, pat_list_count
                        pat_str_ary = Split(pat_list_array(0), ":")
                        If pat_list_array(0) Like "*:*" Then
                            pat_str_ary = Split(pat_str_ary(1), "_")
                        Else
                            pat_str_ary = Split(pat_str_ary(0), "\")
                            pat_str_ary = Split(pat_str_ary(UBound(pat_str_ary)), "_")
                        End If
    
                        If LCase(pat_str_ary(3)) Like LCase("pl*") Or LCase(pat_str_ary(3)) Like LCase("fu*") Then
                            confirm_pat_load = True
                            If LCase(pat_str_ary(3)) Like LCase("pllp") Or LCase(pat_str_ary(3)) Like LCase("fulp") Then
                               confirm_pat_load = True
                            Else
                                For payload_len = 3 To Len(pat_str_ary(3))
                                    If mid(pat_str_ary(3), payload_len, 1) >= "0" And mid(pat_str_ary(3), payload_len, 1) <= "9" Then
                                        If confirm_pat_load <> False Then confirm_pat_load = True
                                    Else
                                        confirm_pat_load = False
                                    End If
                                Next payload_len
                            End If
                        Else
                            confirm_pat_load = False
                        End If
						'20240531shmoo hole fail log
						'initial do shmoo flag
                        do_shmoo_hole_flag = False
                        Set hramInfos = New LVCC_VminBoundary
                        If inst_info(11) Like "DFC*" Then
                            If const_Size = True Then
                                hramInfos.LvccVoltage = DFC_Temp_Size(1) / 1000
                                hramInfos.LvccCount = DFC_Temp_Num(1) 'DFC_Temp_Num(1) + 1
                            Else
                                hramInfos.LvccVoltage = DFC_Range(log_cnt + 1) / 1000
                                hramInfos.LvccCount = 1 '2
                            End If
                        End If
                        If Lvcc = True Then
                            hramInfos.LvccValue = Shmoo_Vcc_Min(CurrSite)
                            hramInfos.LvccCount = 1 '2
                        Else
                            hramInfos.LvccValue = Shmoo_Vcc_Max(CurrSite)
                            hramInfos.LvccCount = 1 '2
                        End If
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''init
                        If confirm_pat_load <> True Then
                            '-----------------------------------------------------------------assign power voltage
                            If m = 0 Then
                                thehdw.DCVS.Pins("All_Power").Voltage.Output = tlDCVSVoltageMain
                                thehdw.Wait 0.0001
                                For i = 0 To g_ApplyLevelTimingVmain.Pins.Count - 1
                                    'DFC INT Failure for voltage application'TheHdw.DCVS.Pins(g_ApplyLevelTimingVmain.Pins(i).name).Voltage.value = g_ApplyLevelTimingVmain.Pins(i).value
                                    thehdw.DCVS.Pins(g_ApplyLevelTimingVmain.Pins(i).name).Voltage.Main.value = g_ApplyLevelTimingVmain.Pins(i).value
                                    thehdw.DCVS.Pins(g_ApplyLevelTimingVmain.Pins(i).name).Voltage.Alt.value = g_ApplyLevelTimingValt.Pins(i).value
                                Next i
                            End If
                            '''''''''''''''''''''''''''''''''confirm whether DSSC source pat
                            If InStr(dssc_pat_init_GB, ShmooPatternSplit(m)) > 0 Then
                              dssc_pat_exist = True
                                For i = 0 To Power_number
                                    If Lvcc = True Then
                                        If Not hramInfos.IsEnableDFTLHFC Then
                                            Shmoo_value.Pins(Power_pins(i)).value = Shmoo_Vcc_Min(CurrSite) - RangeStepSize * (j + 1)
                                        Else
                                            Shmoo_value.Pins(Power_pins(i)).value = Shmoo_Vcc_Min(CurrSite) - hramInfos.LvccVoltage * (j + 1)
                                        End If
                                    Else
                                        If Not hramInfos.IsEnableDFTLHFC Then
                                            Shmoo_value.Pins(Power_pins(i)).value = Shmoo_Vcc_Max(CurrSite) + RangeStepSize * (j + 1)
                                        Else
                                            Shmoo_value.Pins(Power_pins(i)).value = Shmoo_Vcc_Max(CurrSite) + hramInfos.LvccVoltage * (j + 1)
                                        End If
                                    End If
                                Next i
                                
                                DSSC_string = vbNullString
                                Dim SEL_print As New SiteVariant
                                                                
                                Dim sSrcSigName As String
                                Dim tempVarArray As Variant
                                tempVarArray = thehdw.DSSC.Pins(DigSrc_pin_GB).pattern(dssc_pat_init_GB).Source.Labels.list
                                sSrcSigName = tempVarArray(0)
                                If sSrcSigName = "" Then
                                        sSrcSigName = "FUNC_SRC"
                                End If
                                With g_CharPattInfoAry(m)
                                    DigSrc_wav.CreateConstant 0, Len(.DynamicSourceBit)
                                    Decide_DC_Level_Mod DC_Spec_Level, ""
                                    DSSC_string = Decide_Switching_Bit_Debug_LVCC(digSrc_EQ_GB, DigSrc_wav, DC_Spec_Level, "", SELSRM_Rails, powerPin, Shmoo_value, g_ForceCond_VDD, .DictApplyVol, ShmooPatternSplit(m), SEL_print, True)
                                    Call SetupDigSrcDspWave(dssc_pat_init_GB, DigSrc_pin_GB, sSrcSigName, CLng(.DigSrc_BitSize), DigSrc_wav)
                                End With
                            End If
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            If m = 0 Then
                                If Not hramInfos.IsEnableDFTLHFC Then
                                    StepValue = (j + 1) * RangeStepSize
                                Else
                                    StepValue = (j + 1) * hramInfos.LvccVoltage
                                End If
								
								'20240531shmoo hole fail log
								If LVCC = True And TheExec.enableWord("Find_shmoo_hole_low_power_twice") = False Then
                                'If Lvcc = True Then
                                    TheExec.Datalog.WriteComment "Power setup (Vmin - " & StepValue * 1000 & "mV) "
                                Else
                                    TheExec.Datalog.WriteComment "Power setup (Vmax + " & StepValue * 1000 & "mV) "
                                End If
                            End If
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            'If judgment add by Chris 220311 for MemFP in Char
                            'Add judgment for Mbist pattern only by William_Jr & Berton 220318
                            If (hramInfos.IsEnableDFTLHFC And (TheExec.enableWord("Mbist_FingerPrint") = True Or TheExec.enableWord("Mbist_FingerPrint_Vector") = True)) And Shmoo_Pattern Like "*_BI_*" Then
                                Call Finger_print(ShmooPatternSplit(m), True, "", False)
                            Else
                                thehdw.Patterns(ShmooPatternSplit(m)).start
                                thehdw.Digital.Patgen.HaltWait
                                HardIP_WriteFuncResult
                            End If
                              ''' store hram data
                            If hramInfos.IsEnableDFTLHFC And Not hramInfos.IsAnySiteFail Then hramInfos.storeHRAMData Join(pat_str_ary, "_"), Power_pins, CurrSite
                            
                            If Retention_Fail_Cycle = True Then
                                For L = 0 To UBound(RET_PatAry)
                                    If ShmooPatternSplit(m) Like RET_PatAry(L) Then
'                                    Stop
                                        Set RET_PLD = Nothing
                                        For i = 0 To Power_number
                                            TheExec.DataManager.DecomposePinList Power_pins(i), RET_PinAry, RET_PinCnt
                                            For p = 0 To RET_PinCnt - 1
                                            
                                                RET_PLD.AddPin RET_PinAry(p)
                                                If Lvcc = True Then
                                                    RET_PLD.Pins(RET_PinAry(p)).value = Shmoo_Vcc_Min(CurrSite) - RangeStepSize * (j + 1)
                                                Else
                                                    RET_PLD.Pins(RET_PinAry(p)).value = Shmoo_Vcc_Max(CurrSite) + RangeStepSize * (j + 1)
                                                End If
                                            Next p
                                        Next i
                                        
                                        Retention_RampdownUp Shmoo_Apply_Pin, "DOWN", True, RET_PLD
                                        thehdw.Wait RET_WaitAry(L)
                                        Retention_RampdownUp Shmoo_Apply_Pin, "UP", True, RET_PLD
                                        Exit For
                                    End If
                                Next L
                            End If
                                                        
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''pl
                        ElseIf confirm_pat_load = True Then
                            '---------------------------------------------for non shmoo pins
							If TheExec.enableWord("Find_shmoo_hole_low_power_twice") = True Then '20231018
								If Not Find_shmoo_hole_low_power_twice_hvcc_voltage(0)(CurrSite) = 0 Then
									For step_10 = 0 To 9
										TheExec.Datalog.WriteComment "Do Shmoo Hole Loop ," + "(" & step_10 + 1 & " time)"
											For step_2 = 0 To UBound(Find_shmoo_hole_low_power_twice_hvcc_voltage)
												'/RESTORE
												Store_orginal_voltage = Thehdw.DCVS.Pins(powerPin).Voltage.Alt.value
												TheExec.Datalog.WriteComment "[ Shmoo Hole Voltage ] " + powerPin + " = " & CStr(Format(Find_shmoo_hole_low_power_twice_hvcc_voltage(step_2)(CurrSite), "0.000")) & " V"
												Thehdw.DCVS.Pins(powerPin).Voltage.Alt.value = Find_shmoo_hole_low_power_twice_hvcc_voltage(step_2)
												Thehdw.DCVS.Pins(powerPin).Voltage.Output = tlDCVSVoltageAlt
												Thehdw.Wait 0.0001

												Thehdw.Patterns(ShmooPatternSplit(m)).start
												Thehdw.Digital.Patgen.HaltWait
												HardIP_WriteFuncResult
												If hramInfos.IsEnableDFTLHFC And Not hramInfos.IsAnySiteFail Then
													hramInfos.storeHRAMData Join(pat_str_ary, "_"), Power_pins, CurrSite
													do_shmoo_hole_flag = True
												End If
												Thehdw.DCVS.Pins(powerPin).Voltage.Alt.value = Store_orginal_voltage

											Next step_2
										If do_shmoo_hole_flag = True Then Exit For '20231018
											'"/ judge PF
									Next step_10
								End If	
                            Else
                            'If g_CharInputString_Voltage_Dict.Count <> 0 Then
                                For i = 0 To UBound(pin_name)       '--------apply force_condition(Dictionary)
                                   'If Not g_CharInputString_Voltage_Dict.Exists(pin_name(i)) Then
                                    'thehdw.DCVS.Pins(pin_name(i)).Voltage.Alt.value = g_CharInputString_Voltage_Dict(pin_name(i))
                                   'End If
                                   thehdw.DCVS.Pins(pin_name(i)).Voltage.Alt.value = g_CharPattInfoAry(m).ForceVoltage.Pins(UCase(pin_name(i))).value
                                Next i
                            'End If
                            
								If Retention_Fail_Cycle = True Then
									For L = 0 To UBound(RET_PatAry)
										If ShmooPatternSplit(m) Like RET_PatAry(L) Then
											Set RET_PLD = Nothing
											For i = 0 To Power_number
												RET_PLD.AddPin Power_pins(i)
												If Lvcc = True Then
													RET_PLD.Pins(Power_pins(i)).value = Shmoo_Vcc_Min(CurrSite) - RangeStepSize * (j + 1)
												Else
													RET_PLD.Pins(Power_pins(i)).value = Shmoo_Vcc_Max(CurrSite) + RangeStepSize * (j + 1)
												End If
											Next i
											Exit For
										End If
									Next L
								Else
								'---------------------------------------------for shmoo pins
									For i = 0 To Power_number
										If Lvcc = True Then
											If Not hramInfos.IsEnableDFTLHFC Then
												TheHdw.DCVS.Pins(Power_pins(i)).Voltage.Alt.value = Shmoo_Vcc_Min(CurrSite) - RangeStepSize * (j + 1)
											Else
												TheHdw.DCVS.Pins(Power_pins(i)).Voltage.Alt.value = Shmoo_Vcc_Min(CurrSite) - hramInfos.LvccVoltage * (j + 1)
											End If
										Else
											If Not hramInfos.IsEnableDFTLHFC Then
												TheHdw.DCVS.Pins(Power_pins(i)).Voltage.Alt.value = Shmoo_Vcc_Max(CurrSite) + RangeStepSize * (j + 1)
											Else
												TheHdw.DCVS.Pins(Power_pins(i)).Voltage.Alt.value = Shmoo_Vcc_Max(CurrSite) + hramInfos.LvccVoltage * (j + 1)
											End If
										End If
									Next i
								End If
								'---------------------------------------------
								'---------------------------------------------
								TheHdw.DCVS.Pins("All_Power").Voltage.Output = tlDCVSVoltageAlt
								TheHdw.Wait 0.0001
								'---------------------------------------------
								'If judgment add by Chris 220311 for MemFP in Char
								'Add judgment for Mbist pattern only by William_Jr & Berton 220318
								If (hramInfos.IsEnableDFTLHFC And (TheExec.enableWord("Mbist_FingerPrint") = True Or TheExec.enableWord("Mbist_FingerPrint_Vector") = True)) And (Shmoo_Pattern Like "*_BI_*") Then
									Call Finger_print(ShmooPatternSplit(m), True, "", False)
								Else
									TheHdw.Patterns(ShmooPatternSplit(m)).start
									TheHdw.Digital.Patgen.HaltWait
									HardIP_WriteFuncResult
								End If
								
								If Retention_Fail_Cycle = True Then
									For L = 0 To UBound(RET_PatAry)
										If ShmooPatternSplit(m) Like RET_PatAry(L) Then
											Retention_RampdownUp Shmoo_Apply_Pin, "DOWN", True, RET_PLD
											TheHdw.Wait RET_WaitAry(L)
											Retention_RampdownUp Shmoo_Apply_Pin, "UP", True, RET_PLD
											Retention_Fail_Cycle = False
										Else
										End If
									Next L
								End If
													   
								TheExec.Datalog.WriteComment "                                                "
								If TheHdw.Digital.Patgen.PatternBurstPassed(CurrSite) = False And LVCC_boundary_Switch = Fail_log_cnt Then
									For i = 0 To Power_number
										TheExec.Datalog.WriteComment Power_pins(i) + "= " & CStr(Format(TheHdw.DCVS.Pins(Power_pins(i)).Voltage.value, "0.000"))
										TheExec.Datalog.WriteComment ("VOLTAGE=") + CStr(Format(TheHdw.DCVS.Pins(Power_pins(i)).Voltage.value, "0.000")) + ("V")
									Next i
									GoTo Endfor
								End If
								If TheHdw.Digital.Patgen.PatternBurstPassed(CurrSite) = False Then Fail_log_cnt_flag = True
								'=============================================================================================================print log
								For i = 0 To Power_number
									TheExec.Datalog.WriteComment Power_pins(i) + "= " & CStr(Format(TheHdw.DCVS.Pins(Power_pins(i)).Voltage.value, "0.000"))
									TheExec.Datalog.WriteComment ("VOLTAGE=") + CStr(Format(TheHdw.DCVS.Pins(Power_pins(i)).Voltage.value, "0.000")) + ("V")
								Next i
								If dssc_pat_exist = True Then TheExec.Datalog.WriteComment "DSSC Sorce = " & DSSC_string & ";  SELSRM_Rails=" & SELSRM_Rails
								'=============================================================================================================
							''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''error
								''' store hram data
								If hramInfos.IsEnableDFTLHFC And Not hramInfos.IsAnySiteFail Then hramInfos.storeHRAMData Join(pat_str_ary, "_"), Power_pins, CurrSite
							End If
						End If
						''' print out ffc fail log and initial for next run (each pattern print fail cycle/vector/pins)
                        If hramInfos.IsEnableDFTLHFC And hramInfos.IsStoreAnyHRAM Then hramInfos.printFailInfo: hramInfos.initialFFC
                        Set hramInfos = Nothing
					Next m
                    ''' print out ffc fail log and initial for next run
                    'If hramInfos.IsEnableDFTLHFC And hramInfos.IsStoreAnyHRAM Then hramInfos.printFailInfo: hramInfos.initialFFC
                        
                    If Fail_log_cnt_flag = True Then
                        
                        If hramInfos.IsEnableDFTLHFC Or hramInfos.IsEnableFAILLOG Then
                            'If hramInfos.LvccCount - 1 = Fail_log_cnt - 1 Then GoTo Endfor
                            If hramInfos.LvccCount = Fail_log_cnt Then Exit For
                        Else
                            'If LVCC_boundary_Switch - 1 = Fail_log_cnt - 1 Then GoTo Endfor
                            If LVCC_boundary_Switch = Fail_log_cnt Then Exit For
                        End If
                        Fail_log_cnt = Fail_log_cnt + 1
                    End If
                    Fail_log_cnt_flag = False
                    If (hramInfos.IsEnableDFTLHFC Or hramInfos.IsEnableFAILLOG) And hramInfos.IsOnlyOnePoint Then GoTo Endfor
                Next j
            '=============================================================================================================================
            Else
                TheExec.Datalog.WriteComment "Shmoo HVCC all pass"
            End If
            Fail_log_cnt = 1
            If log_cnt + 1 = log_const Then GoTo Endfor
        Next log_cnt
    Else
        TheExec.Datalog.WriteComment "No DFC Step Value."
    End If
    Set hramInfos = Nothing
Endfor:
    Set hramInfos = Nothing
    TheExec.Datalog.WriteComment "***************** Shmoo fail log/Pins capture end *****************"
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Digital_Debug", "FailingDatalog_HLvcc_Boundary_SELSRM")
    If AbortTest Then Exit Function Else Resume Next
    End Function
