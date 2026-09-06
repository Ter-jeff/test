Attribute VB_Name = "VBT_LIB_DC_Leak"
Option Explicit
'Revision History:
'V0.0 initial bring up
' This module should be used for VBT Tests.  All functions in this module
' will be available to be used from the Test Instance sheet.
' Additional modules may be added as needed (all starting with "VBT_").
'
' The required signature for a VBT Test is:
'
' Public Function FuncName(<arglist>) As Long
'   where <arglist> is any list of arguments supported by VBT Tests.
'
' See online help for supported argument types in VBT Tests.
'
'
' It is highly suggested to use error handlers in VBT Tests.  A sample
' VBT Test with a suggeseted error handler is shown below:
'
' Function FuncName() As Long
'     On Error GoTo errHandler
'
'     Exit Function
' errHandler:
'     If AbortTest Then Exit Function Else Resume Next
' End Function



Public Function HiZ_Leakage_Parallel(patset As Pattern, ForceV_IiH As Double, ForceV_IiL As Double, I_Meas_Range As Double, leakage_pins As PinList)

Dim site As Variant
'Dim SeqLeakPins As String
Dim PinArr() As String, PinCount As Long, i As Long
Dim MeasVal As New PinListData
Dim TestNum As Long
On Error GoTo errHandler


TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

TheHdw.patterns(patset).Load
Call TheHdw.patterns(patset).start
TheHdw.Digital.Patgen.HaltWait
  

Call TheHdw.Digital.ApplyLevelsTiming(True, True, False, tlPowered, leakage_pins, , leakage_pins)

TheHdw.Digital.pins(leakage_pins).initState = chInitoff


TheHdw.PPMU.pins(leakage_pins).ForceV (0)
TheHdw.PPMU.pins(leakage_pins).Gate = tlOff
TheHdw.PPMU.pins(leakage_pins).Disconnect


    TestNum = TheExec.sites.item(0).TestNumber
    'If TheExec.DataManager.ChannelType(PinArr(i)) <> "N/C" Then
    
    TheHdw.Digital.pins(leakage_pins).Disconnect
        
        With TheHdw.PPMU(leakage_pins)
            .Connect
            .Gate = tlOn
            .ForceV ForceV_IiH, I_Meas_Range
             TheHdw.Wait 0.001
             DebugPrintFunc_PPMU leakage_pins.value
             MeasVal = .Read(tlPPMUReadMeasurements)
            
            .ForceV (0)
            .Gate = tlOff
            .Disconnect
        End With
    TheHdw.Digital.pins(leakage_pins).Connect 'connect the tested pin back to the PE
              
    TheExec.flow.TestLimit resultVal:=MeasVal, unit:=unitAmp, ForceVal:=ForceV_IiH, ForceUnit:=unitVolt, lowVal:=-0.000006, hiVal:=0.000006, Tname:="JtagTap_HiZ_HV"
    
    'If CurrentJobName Like "*char*" Then Call Char_show(MeasVal, -0.000006, 0.000006, TestNum, "A", CStr(i))
    If TheExec.sites.Active.Count = 0 Then Exit Function 'chihome
        


'Low

TestNum = TheExec.sites.item(0).TestNumber
    'If TheExec.DataManager.ChannelType(PinArr(i)) <> "N/C" Then
       TheHdw.Digital.pins(leakage_pins).Disconnect
        With TheHdw.PPMU(leakage_pins)
            .Connect
            .Gate = tlOn
            .ForceV ForceV_IiL, I_Meas_Range
             TheHdw.Wait 0.001
             DebugPrintFunc_PPMU leakage_pins.value
             MeasVal = .Read(tlPPMUReadMeasurements)
            
            .ForceV (0)
            .Gate = tlOff
            .Disconnect
        End With
             TheHdw.Digital.pins(leakage_pins).Connect 'connect the tested pin back to the PE


       TheExec.flow.TestLimit resultVal:=MeasVal, unit:=unitAmp, ForceVal:=ForceV_IiL, ForceUnit:=unitVolt, lowVal:=-0.000006, hiVal:=0.000006, Tname:="JtagTap_HiZ_0V"
       
       'If CurrentJobName Like "*char*" Then Call Char_show(MeasVal, -0.000006, 0.000006, TestNum, "A", CStr(i))
       If TheExec.sites.Active.Count = 0 Then Exit Function 'chihome



DebugPrintFunc patset.value


Exit Function
errHandler:
        TheExec.AddOutput "Error in the Parallel Leakage Test"
        If AbortTest Then Exit Function Else Resume Next

End Function


Public Function HiZ_Leakage(patset As Pattern, ForceV_IiH As Double, ForceV_IiL As Double, I_Meas_Range As Double, leakage_pins As PinList, Hi_Limit As Double, Lo_Limit As Double, _
                            Port_name As String, CPUA_Flag_In_Pat As Boolean, Optional DisableClock As Boolean = False)
                

Dim site As Variant
'Dim SeqLeakPins As String
Dim PinArr() As String, PinCount As Long, i As Long
Dim p As Variant
Dim MeasVal As New PinListData
Dim TestNum As Long
Dim Tname As String
Dim AllSitePass As Boolean
Dim BurstResult As New SiteLong

On Error GoTo errHandler

Call TheHdw.Digital.ApplyLevelsTiming(True, True, False, tlPowered, leakage_pins, , leakage_pins)

TheHdw.patterns(patset).Load
''Call TheHdw.Patterns(patset).Start
''TheHdw.digital.Patgen.HaltWait

'Check if pattern passed
For i = 0 To 1
    Call TheHdw.patterns(patset).start
    TheHdw.Digital.Patgen.HaltWait
    AllSitePass = True
    For Each site In TheExec.sites
        BurstResult(site) = 1
    Next site
    
    For Each site In TheExec.sites
        If (TheHdw.Digital.Patgen.PatternBurstPassed(site) = False) Then
            TheExec.Datalog.WriteComment vbCrLf & patset & "_" & vbTab & "Run " & i & " : Fail."
            BurstResult(site) = 0
            AllSitePass = False
        End If
    Next site
    If AllSitePass = True Then Exit For
Next i
TheExec.flow.TestLimit BurstResult, 1, 1, tlSignGreaterEqual, tlSignLessEqual, Tname:="Burst_Result" 'BurstResult=1:Pass
TheExec.Datalog.WriteComment ""

TheHdw.Digital.pins(leakage_pins).initState = chInitoff

TheHdw.PPMU.pins(leakage_pins).ForceV (0)
TheHdw.PPMU.pins(leakage_pins).Gate = tlOff
TheHdw.PPMU.pins(leakage_pins).Disconnect

    ''''''use the "theexec.DataManager.DecomposePinList" to serialize the pins to be tested sequentially'''''
    TheExec.DataManager.DecomposePinList leakage_pins, PinArr(), PinCount
'    DCVS_Trim_NC_Pin PinArr(), PinCount
    
    ''High
    
    'If TheExec.DataManager.ChannelType(PinArr(i)) <> "N/C" Then
    TheHdw.Digital.pins(leakage_pins).Disconnect
    
    With TheHdw.PPMU(leakage_pins)
        .Connect
        .Gate = tlOn
        .ForceV ForceV_IiH, I_Meas_Range
         TheHdw.Wait 0.001
         DebugPrintFunc_PPMU leakage_pins.value
         MeasVal = .Read(tlPPMUReadMeasurements)
         
        .ForceV (0)
        .Gate = tlOff
        .Disconnect
    End With
        
    'TheHdw.digital.Pins(leakage_pins).Connect 'connect the tested pin back to the PE
    
    'offline mode simulation
    If TheExec.TesterMode = testModeOffline Then

        For Each site In TheExec.sites
            For Each p In PinArr()
                If TheExec.DataManager.ChannelType(p) <> "N/C" Then MeasVal.pins(p).value = 5 * uA + Rnd() * 0.1 * uA
            Next p
        Next site
    End If
  
  
    TheExec.flow.TestLimit resultVal:=MeasVal, lowVal:=Lo_Limit, hiVal:=Hi_Limit, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:="Leak_Hi", ForceVal:=ForceV_IiH, ForceUnit:=unitVolt, ForceResults:=tlForceNone
    'glb_TestInstance = theexec.DataManager.instancename
    'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
    'TheExec.Flow.TestLimit resultval:=PPMUMeasure.Pins(DUTPin), lowval:=LowLimit, hival:=HiLimit, ScaleType:=scaleNone, unit:=unitVolt, formatstr:="%.3f", tname:=tname, forceVal:=force_i, forceunit:=unitAmp, forceResults:=tlForceNone


    '' Low
    
    ''If TheExec.DataManager.ChannelType(PinArr(i)) <> "N/C" Then
    TheHdw.Digital.pins(leakage_pins).Disconnect

    With TheHdw.PPMU(leakage_pins)
        .Connect
        .Gate = tlOn
        .ForceV ForceV_IiL, I_Meas_Range
         TheHdw.Wait 0.001
         DebugPrintFunc_PPMU leakage_pins.value
         MeasVal = .Read(tlPPMUReadMeasurements)
         
        .ForceV (0)
        .Gate = tlOff
        .Disconnect
    End With
    


    'offline mode simulation
    If TheExec.TesterMode = testModeOffline Then
''        TheExec.DataManager.DecomposePinList leakage_pins, PinArr(), PinCount
''        DCVS_Trim_NC_Pin PinArr(), PinCount
        For Each site In TheExec.sites
            For Each p In PinArr()
                If TheExec.DataManager.ChannelType(p) <> "N/C" Then MeasVal.pins(p).value = -(5 * uA + Rnd() * 0.1 * uA)
            Next p
        Next site
    End If

    TheExec.flow.TestLimit resultVal:=MeasVal, lowVal:=Lo_Limit, hiVal:=Hi_Limit, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:="Leak_Lo", ForceVal:=ForceV_IiL, ForceUnit:=unitVolt, ForceResults:=tlForceNone
    'glb_TestInstance = theexec.DataManager.instancename
    'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
    TheHdw.Digital.pins(leakage_pins).Connect 'connect the tested pin back to the PE
        
    
    DebugPrintFunc patset.value


Exit Function
errHandler:
    TheExec.AddOutput "Error in the Seq Leakage Test"
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function HiZ_Leakage_MeasPower(patset As Pattern, ForceV_IiH As Double, ForceV_IiL As Double, I_Meas_Range As Double, leakage_pins As PinList, Hi_Limit As Double, Lo_Limit As Double, _
                            MeasureI_pin As PinList, I_LowLimit As Double, I_HighLimit As Double, Irange As String, Optional Flag_Low As Boolean = False, Optional CharInputString As String)
                

    Dim site As Variant
    'Dim SeqLeakPins As String
    Dim PinArr() As String, PinCount As Long, i As Long
    Dim p As Variant
    Dim MeasVal As New PinListData
    Dim TestNum As Long
    Dim Tname As String
    Dim AllSitePass As Boolean
    Dim BurstResult As New SiteLong
    
    Dim MeasCurr As New PinListData
    Dim TestSeqNum As Integer
    Dim PowerVoltage As Double
    Dim patt_ary() As String
    Dim pat_count As Long
    Dim pat As String
    Dim patt As Variant
    

    Dim MaxSpec As Double
    Dim ppmuRange As Double
    
    On Error GoTo errHandler

    Call TheHdw.Digital.ApplyLevelsTiming(True, True, False, tlPowered, , , leakage_pins)
    
    
    '' 20150625 - Apply Char setup
    If UCase(TheExec.CurrentJob) Like "*CHAR*" Then
        If CharInputString <> "" Then
            Call SetForceCondition(CharInputString)
        End If
    End If
    
    TheHdw.patterns(patset).Load
    Call PATT_GetPatListFromPatternSet(patset.value, patt_ary, pat_count)

    For Each patt In patt_ary
        pat = CStr(patt)
            
        Call TheHdw.patterns(pat).start
        TheHdw.Digital.Patgen.HaltWait

        Call DC_Func_WriteFuncResult
        
        'digital, ppmu initialized
        TheHdw.Digital.pins(leakage_pins).initState = chInitoff
        
        TheHdw.PPMU.pins(leakage_pins).ForceV 0, I_Meas_Range
        TheHdw.PPMU.pins(leakage_pins).Gate = tlOff
        TheHdw.PPMU.pins(leakage_pins).Disconnect

        ppmuRange = TheHdw.PPMU.pins(leakage_pins).MeasureCurrentRange
        'TheExec.Datalog.WriteComment ("Force Voltatage  = " & ForceV & ", Range = " & ppmuRange)
        MaxSpec = Max(Abs(Hi_Limit), Abs(Lo_Limit))
        
        TheExec.flow.TestLimit ppmuRange, MaxSpec, , tlSignGreaterEqual, tlSignLessEqual, Tname:="Abs(Spec) vs Range" 'PPMU Range check PASS

        ''High
        TheHdw.Digital.pins(leakage_pins).Disconnect
    
        If LCase(currentJobName) Like "*cp*" Then
            ForceV_IiH = ForceV_IiH * 1
        ElseIf LCase(currentJobName) Like "*ft*" Then
            ForceV_IiH = ForceV_IiH * 1
        End If
       
    
        With TheHdw.PPMU(leakage_pins)
            .Connect
            .Gate = tlOn
            .ForceV ForceV_IiH, I_Meas_Range
             TheHdw.Wait 0.001
             DebugPrintFunc_PPMU leakage_pins.value
             MeasVal = .Read(tlPPMUReadMeasurements)
             
            .ForceV (0)
            .Gate = tlOff
            .Disconnect
        End With
        
        TheHdw.Digital.pins(leakage_pins).Connect
        
        'offline mode simulation
        If TheExec.TesterMode = testModeOffline Then
            TheExec.DataManager.DecomposePinList leakage_pins, PinArr(), PinCount  'decompose for serial tests
            For Each site In TheExec.sites
                For Each p In PinArr()
                    If TheExec.DataManager.ChannelType(p) <> "N/C" Then MeasVal.pins(p).value = 0.5 * uA + Rnd() * 0.1 * uA
                Next p
            Next site
        End If
            If LCase(TheExec.CurrentJob) Like "*char*" And UCase(TheExec.DataManager.instancename) Like "*GPIO_PP_CMN_S_FULP_IO_BSCN_BSD_JIO_UNS_ALLFV_SI_HIGHZ*" Then
                TheExec.flow.TestLimit resultVal:=MeasVal, lowVal:=Lo_Limit, hiVal:=Hi_Limit, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", ForceVal:=ForceV_IiH, ForceUnit:=unitVolt, ForceResults:=tlForceFlow
            Else
                TheExec.flow.TestLimit resultVal:=MeasVal, lowVal:=Lo_Limit, hiVal:=Hi_Limit, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", ForceVal:=ForceV_IiH, ForceUnit:=unitVolt, ForceResults:=tlForceNone
                'glb_TestInstance = theexec.DataManager.instancename
                'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
            End If
''''        'DCVS meas
''''        Call DCVS_Set_Meter_Range(MeasureI_pin, Irange)
''''        DCVS_MeterRead DCVS_UVS256, CStr(MeasureI_pin), 10, MeasCurr
''''        PowerVoltage = Format(TheHdw.DCVS.Pins(MeasureI_pin).Voltage.Main.Value, "0.00")
''''
''''        'offline mode simulation
''''        If TheExec.TesterMode = testModeOffline Then
''''            TheExec.DataManager.DecomposePinList MeasureI_pin, PinArr(), PinCount  'decompose for serial tests
''''            For Each Site In TheExec.Sites
''''                For Each p In PinArr()
''''                    If TheExec.DataManager.ChannelType(p) <> "N/C" Then MeasCurr.Pins(p).Value = (5 * mA + Rnd() * 0.1 * mA)
''''                Next p
''''            Next Site
''''        End If
''''
''''        TheExec.Flow.TestLimit resultVal:=MeasCurr, lowVal:=I_LowLimit, hiVal:=I_HighLimit, ScaleType:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:="Curr_meas_0_" + "@COND:PATTERN=" + PATT_ExculdePath(Pat), forceVal:=PowerVoltage, forceunit:=unitVolt, ForceResults:=tlForceNone

        '' Low
        If Flag_Low = True Then
            TheHdw.Digital.pins(leakage_pins).Disconnect
        
            With TheHdw.PPMU(leakage_pins)
                .Connect
                .Gate = tlOn
                .ForceV ForceV_IiL, I_Meas_Range
                 TheHdw.Wait 0.001
                 DebugPrintFunc_PPMU leakage_pins.value
                 MeasVal = .Read(tlPPMUReadMeasurements)
                 
                .ForceV (0)
                .Gate = tlOff
                .Disconnect
            End With
            
            'offline mode simulation
            If TheExec.TesterMode = testModeOffline Then
                TheExec.DataManager.DecomposePinList leakage_pins, PinArr(), PinCount  'decompose for serial tests
                For Each site In TheExec.sites
                    For Each p In PinArr()
                        If TheExec.DataManager.ChannelType(p) <> "N/C" Then MeasVal.pins(p).value = -(0.5 * uA + Rnd() * 0.1 * uA)
                    Next p
                Next site
            End If
            
            TheHdw.Digital.pins(leakage_pins).Connect
    
            TheExec.flow.TestLimit resultVal:=MeasVal, lowVal:=Lo_Limit, hiVal:=Hi_Limit, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:="Leak_Lo", ForceVal:=ForceV_IiL, ForceUnit:=unitVolt, ForceResults:=tlForceNone
            'glb_TestInstance = theexec.DataManager.instancename
            'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
''''            Call DCVS_Set_Meter_Range(MeasureI_pin, Irange)
''''            DCVS_MeterRead DCVS_UVS256, CStr(MeasureI_pin), 10, MeasCurr
''''            PowerVoltage = Format(TheHdw.DCVS.Pins(MeasureI_pin).Voltage.Main.Value, "0.00")
''''
''''            'offline mode simulation
''''            If TheExec.TesterMode = testModeOffline Then
''''                TheExec.DataManager.DecomposePinList MeasureI_pin, PinArr(), PinCount  'decompose for serial tests
''''                For Each Site In TheExec.Sites
''''                    For Each p In PinArr()
''''                        If TheExec.DataManager.ChannelType(p) <> "N/C" Then MeasCurr.Pins(p).Value = (5 * mA + Rnd() * 0.1 * mA)
''''                    Next p
''''                Next Site
''''            End If
''''
''''            TheExec.Flow.TestLimit resultVal:=MeasCurr, lowVal:=I_LowLimit, hiVal:=I_HighLimit, ScaleType:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:="Curr_meas_1_" + "@COND:PATTERN=" + PATT_ExculdePath(Pat), forceVal:=PowerVoltage, forceunit:=unitVolt, ForceResults:=tlForceNone

       End If
    Next patt
   
    DebugPrintFunc patset.value


Exit Function
errHandler:
    'TheExec.AddOutput "Error in the Seq Leakage Test"
    ErrorDescription ("HiZ_Leakage_MeasPower")
    If AbortTest Then Exit Function Else Resume Next

End Function



Public Function HiZ_Leakage_Parallel_GPIO(patset As Pattern, ForceV_IiH As Double, ForceV_IiL As Double, I_Meas_Range As Double, _
        leakage_pins As PinList, Low_limit As Double, High_limit As Double, Optional power_gate_off As Boolean = False)

Dim site As Variant
'Dim SeqLeakPins As String
Dim PinArr() As String, PinCount As Long, i As Long
Dim MeasVal As New PinListData
Dim TestNum As Long
On Error GoTo errHandler


TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

TheHdw.patterns(patset).Load
Call TheHdw.patterns(patset).start
TheHdw.Digital.Patgen.HaltWait
  

Call TheHdw.Digital.ApplyLevelsTiming(True, True, False, tlPowered, leakage_pins, , leakage_pins)

TheHdw.Digital.pins(leakage_pins).initState = chInitoff

If power_gate_off = True Then
TheHdw.DCVS.pins("vddio18_grp1,vddio18_grp2").Gate = False
TheExec.Datalog.WriteComment "VDDIO Gate power off"
End If

TheHdw.PPMU.pins(leakage_pins).ForceV (0)
TheHdw.PPMU.pins(leakage_pins).Gate = tlOff
TheHdw.PPMU.pins(leakage_pins).Disconnect


    TestNum = TheExec.sites.item(0).TestNumber
    'If TheExec.DataManager.ChannelType(PinArr(i)) <> "N/C" Then
    
    TheHdw.Digital.pins(leakage_pins).Disconnect
        
        With TheHdw.PPMU(leakage_pins)
            .Connect
            .Gate = tlOn
            .ForceV ForceV_IiH, I_Meas_Range
             TheHdw.Wait 0.001
             DebugPrintFunc_PPMU leakage_pins.value
             MeasVal = .Read(tlPPMUReadMeasurements)
            
            .ForceV (0)
            .Gate = tlOff
            .Disconnect
        End With
    TheHdw.Digital.pins(leakage_pins).Connect 'connect the tested pin back to the PE
              
    TheExec.flow.TestLimit resultVal:=MeasVal, unit:=unitAmp, ForceVal:=ForceV_IiH, ForceUnit:=unitVolt, lowVal:=Low_limit, hiVal:=High_limit, Tname:="GPIO_Iih"
    
    'If CurrentJobName Like "*char*" Then Call Char_show(MeasVal, -0.000006, 0.000006, TestNum, "A", CStr(i))
    If TheExec.sites.Active.Count = 0 Then Exit Function 'chihome
        


'Low

TestNum = TheExec.sites.item(0).TestNumber
    'If TheExec.DataManager.ChannelType(PinArr(i)) <> "N/C" Then
       TheHdw.Digital.pins(leakage_pins).Disconnect
        With TheHdw.PPMU(leakage_pins)
            .Connect
            .Gate = tlOn
            .ForceV ForceV_IiL, I_Meas_Range
             TheHdw.Wait 0.001
             DebugPrintFunc_PPMU leakage_pins.value
             MeasVal = .Read(tlPPMUReadMeasurements)
            
            .ForceV (0)
            .Gate = tlOff
            .Disconnect
        End With
             TheHdw.Digital.pins(leakage_pins).Connect 'connect the tested pin back to the PE


       TheExec.flow.TestLimit resultVal:=MeasVal, unit:=unitAmp, ForceVal:=ForceV_IiL, ForceUnit:=unitVolt, lowVal:=Low_limit, hiVal:=High_limit, Tname:="GPIO_Iil"
       
       'If CurrentJobName Like "*char*" Then Call Char_show(MeasVal, -0.000006, 0.000006, TestNum, "A", CStr(i))
       If TheExec.sites.Active.Count = 0 Then Exit Function 'chihome



DebugPrintFunc patset.value


Exit Function
errHandler:
        TheExec.AddOutput "Error in the Parallel Leakage Test"
        If AbortTest Then Exit Function Else Resume Next

End Function






