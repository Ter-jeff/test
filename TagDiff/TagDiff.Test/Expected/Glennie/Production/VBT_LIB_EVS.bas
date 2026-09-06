Attribute VB_Name = "VBT_LIB_EVS"
'====================================================
'=   Define the variables for EVS  =
'====================================================
Public GL_TestName As String
Public GL_EVS_Ramp_Steps As Long
Public GL_Measure_Current_EVS As New PinListData
Public GL_NV_Voltage_Pins As New PinListData
Public GL_EVS_Voltage_Pins As New PinListData
Public GL_Force_Voltage_Pins As New PinListData
Public GL_EVS_Pin_IFold_Max As New PinListData
Public GL_All_Power_Pins As New PinList
Public GL_DCVS_Pin_Output As New PinListData
Public GL_Gate_Check  As New SiteBoolean
Public GL_Flag_Checkboard As String
Public GL_Flag_Reverse_Checkboard As String
Public GL_inst_name As String


Public Function auto_Dummy_Item(patterns As Pattern) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_Dummy_Category"

    Dim m_TestInstName As String
    m_TestInstName = Trim(TheExec.DataManager.instancename)
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    '''//// For cooling ////'''
'''    TheHdw.DCVS.Pins("VDD_AVE,VDD_CPU_SRAM,VDD_DCS_DDR,VDD_DISP,VDD_ECPU,VDD_FIXED,VDD_GPU,VDD_LOW,VDD_PCPU,VDD_SOC,VDD_SRAM_GPU,VDD_SRAM_SOC").Voltage.Value = 0.5
'''    TheHdw.Digital.Pins("All_Digital").InitState = chInitLo
'''    TheHdw.Wait 0.5
    '''//// For cooling ////'''
    
    TheExec.flow.TestLimit 1, 1, 1, , , , , , Tname:="Dummy"
    If (False) Then
        TheExec.Datalog.WriteComment m_TestInstName + "...Load ApplyLevelsTiming..."
    End If
    
    Call SetupDatalogFormat(TestNameW:=115, PatternW:=100)
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function


'Public Function auto_Checkboard_EVS_Probe_Location()
'On Error GoTo errHandler
'    Dim funcName As String:: funcName = "auto_Check_Probe_Location"
'    Dim site As Variant
'
'    For Each site In TheExec.sites
'
'        Dim Sum_XY As New SiteLong
'        GL_Flag_Checkboard = "Checkboard"
'        GL_Flag_Reverse_Checkboard = "Reverse_Checkboard"
'        Sum_XY(site) = XCoord(site) + YCoord(site)
'        If Sum_XY Mod 2 = 0 Then
'            TheExec.sites(site).FlagState(GL_Flag_Checkboard) = logicTrue
'            TheExec.sites(site).FlagState(GL_Flag_Reverse_Checkboard) = logicFalse
'        Else
'            TheExec.sites(site).FlagState(GL_Flag_Checkboard) = logicFalse
'            TheExec.sites(site).FlagState(GL_Flag_Reverse_Checkboard) = logicTrue
'        End If
'
'    Next site
'
'Exit Function
'
'errHandler:
'     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
'     If AbortTest Then Exit Function Else Resume Next
'End Function


Public Function IFold_Max_Cal(pins() As String, site As Long)

Dim i As Integer
Dim Pin_CH(), Pin_Slot(), Pin_CH_Gate(), ChannelType(), Channel_Str() As String
Dim Ifold_Unit As Double
Dim Channel_Count As Long
Dim Pin_Str As Variant

ReDim ChannelType(UBound(pins))
ReDim Channel_Str(UBound(pins))
ReDim Pin_CH(UBound(pins))
ReDim Pin_Slot(UBound(pins))
ReDim Pin_CH_Gate(UBound(pins))

If GL_EVS_Pin_IFold_Max.pins.Count = 0 Then
   For Each Pin_Str In pins
       GL_EVS_Pin_IFold_Max.AddPin(Pin_Str).value = 2
   Next
End If
    
For i = 0 To UBound(pins)
    Call TheExec.DataManager.GetChannelTypes(pins(i), Channel_Count, Channel_Str)
    ChannelType(i) = Channel_Str(0)
    Pin_CH(i) = TheHdw.pins(pins(i)).ChanFromSite(site)
    
    
    Select Case mid(Pin_CH(i), 1, InStr(Pin_CH(i), ".") - 1)
        Case 0, 1, 11, 12, 13, 20, 21
            Pin_Slot(i) = " HEXVS": Ifold_Unit = 15
        Case Else
            Pin_Slot(i) = " UVS256"
            If UCase(Pin_CH(i)) Like UCase("*HC*") Then
                Ifold_Unit = 0.7
            Else
                Ifold_Unit = 0.2
            End If
    End Select
    
    If UCase(ChannelType(i)) Like UCase("*Merged*") Then Pin_CH_Gate(i) = right(ChannelType(i), 1)
    If Pin_CH_Gate(i) = "" Then Pin_CH_Gate(i) = CStr(1)
    GL_EVS_Pin_IFold_Max.pins(i).value = CLng(Pin_CH_Gate(i)) * Ifold_Unit
    TheHdw.DCVS.pins(pins(i)).CurrentRange.value = TheHdw.DCVS.pins(pins(i)).CurrentRange.Max
    TheHdw.DCVS.pins(pins(i)).CurrentLimit.Source.FoldLimit.level.value = GL_EVS_Pin_IFold_Max.pins(i).value
Next i

End Function


Public Function PrintCorePowerGate(GL_All_Power_Pins_gate() As String, inst_name As String)
Dim EVS_core_power_gate As New SiteBoolean
Dim i As Integer

For i = 0 To UBound(GL_All_Power_Pins_gate)
    For Each site In TheExec.sites.Active

        EVS_core_power_gate(site) = TheHdw.DCVS.pins(GL_All_Power_Pins_gate(i)).Gate
        If EVS_core_power_gate(site) = False Then GL_Gate_Check(site) = True
        
    Next site
    
    TheExec.flow.TestLimit EVS_core_power_gate, -1, -1, , , , , , Tname:=inst_name & "  Gate_" & GL_All_Power_Pins_gate(i), PinName:=GL_All_Power_Pins_gate(i)
    
Next i

End Function


Public Function CheckCorePowerGate(GL_All_Power_Pins_gate() As String)
Dim EVS_core_power_gate As New SiteBoolean
Dim i As Integer

For i = 0 To UBound(GL_All_Power_Pins_gate)

    For Each site In TheExec.sites.Active

        EVS_core_power_gate(site) = TheHdw.DCVS.pins(GL_All_Power_Pins_gate(i)).Gate
        If EVS_core_power_gate(site) = False Then GL_Gate_Check(site) = True
  
    Next site
        
Next i

End Function
Public Function EVS_Check_AlarmTimeOut(pins As String)

If TheHdw.DCVS.pins(pins).CurrentLimit.Source.FoldLimit.TimeOut > 0.025 Then
    MsgBox ("<!!!WARNING!!!> There has risk of burning needles! <!!!WARNING!!!>")
    For Each site In TheExec.sites
        TheExec.sites.item(site).FlagState("F_TimeOut_alarm_check") = logicTrue
    Next site
End If

End Function

Public Function EVS_Power_Ramp(Direction As String, EVS_Power_Pin As String, EVS_Voltage As Double, S_WaitTime As Double, Block As String, Optional Ramp_Step_Function As Boolean = True, Optional Ramp_Max_Step As Long = 20, Optional Step_Voltage As Double = 0.05, Optional Rising_Delay_time As Double = 0#, Optional IFold_Max As Boolean = False, Optional Looping_Contorl As Boolean = False, Optional Looping_Index_Name As String = "", Optional Looping_Max_Steps_Name As String = "", Optional Looping_Volt_Start_Value As Double = 1.3, Optional Open_LatchUp_Measure As Boolean = False, Optional LatchUp_Volt_End_Value As Double = 1.6, Optional Multi_Function As Boolean = False, Optional Multi_EVS_Index_Name As String = "", Optional Multi_EVS_Index As Long = 1, Optional Ramp_Rate_Function As Boolean = False, Optional Ramp_Rate As Long = 50) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "EVS_Power_Ramp"

    Dim EVS_Voltage_Sub As Double

    Dim VBT_LIB_Common_ProfileMark_2249 As Long: VBT_LIB_Common_ProfileMark_2249 = ProfileMarkEnter(2, "EVS_RAMP_MAIN_SETTING")

    '''////=======================================================================================////'''
    EVS_Voltage_Sub = 1
    GL_All_Power_Pins = ("VDD_AVE,VDD_CPU_SRAM,VDD_DCS_DDR,VDD_DISP,VDD_ECPU,VDD_FIXED,VDD_GPU,VDD_LOW,VDD_PCPU,VDD_SOC,VDD_SRAM_GPU,VDD_SRAM_SOC,VDDIO12_GRP,VDDIO18_GRP1")
    '''////=======================================================================================////'''
    
     '// ----Multi EVS ----
    If Open_LatchUp_Measure Then Multi_Function = False
    If Multi_Function = False Then
        Multi_EVS_Index = 1
    Else
        Multi_EVS_Index = Multi_EVS_Index
    End If
    S_WaitTime = S_WaitTime / Multi_EVS_Index
    
    '''////setting current/voltage profile samplerate/samplesize////'''
    '''////Define sample rate and sample size for profile base on pattern execution time////'''
    Dim step As Long
    Dim pin As Variant
    Dim i, j As Long
    Dim DomainStr As String
    Dim LoopCnt As Long
    Dim Force_Voltage As Double
    Dim Force_EVS_Pin As String
    Dim All_Power_Pins_Str() As String
    Dim EVS_Power_Pins_Str() As String
    Dim PinsCnt As Long
    Dim pins As Variant
    Dim PowerVolt1 As Double
    Dim Match_Pins() As Boolean
    Dim site As Variant
    Dim Dict_EVS_Pins As New Dictionary
    
    If Open_LatchUp_Measure And Looping_Contorl Then GoTo errHandler

    
    ''' the looping flow is for vtrig collection, if that is production flow, adjust the looping value to max value, it will jump out the looping flow'''
    Dim EVS_Loop_Var_Max As New SiteLong
    For Each site In TheExec.sites.Active
        GL_Gate_Check(site) = False
        If Looping_Contorl = False And Open_LatchUp_Measure = False Then TheExec.sites(site).SiteVariableValue(Looping_Index_Name) = TheExec.flow.var(Looping_Max_Steps_Name).value
        If Multi_Function = False Then TheExec.sites(site).SiteVariableValue(Multi_EVS_Index_Name) = TheExec.flow.var(Looping_Max_Steps_Name).value + 100
    Next site

    GL_inst_name = TheExec.DataManager.instancename
    
    DomainStr = LCase(Block)
    EVS_Power_Pins_Str = Split(EVS_Power_Pin, ",")
    
    Call TheExec.DataManager.DecomposePinList(GL_All_Power_Pins, All_Power_Pins_Str, PinsCnt)
        
    
    For Each pin In EVS_Power_Pins_Str
        Dict_EVS_Pins.Add pin, pin
    Next
    For Each pin In EVS_Power_Pins_Str
        If Not Dict_EVS_Pins.Exists(pin) Then
            Dict_EVS_Pins.Add pin, pin
        End If
    Next
    
    If GL_EVS_Voltage_Pins.pins.Count = 0 Then
        For Each pins In All_Power_Pins_Str
            GL_EVS_Voltage_Pins.AddPin(pins).value = EVS_Voltage_Sub
''''''            GL_DCVS_Pin_Output.AddPin(Pins).Value = 0
        Next
    End If
    
    
    '''////check the EVS core pin string fill in argument if correct////'''
    i = 0: ReDim Match_Pins(99)
    For Each pins In All_Power_Pins_Str
        For Each pin In EVS_Power_Pins_Str
            If pin = pins Then
                Match_Pins(i) = True
                i = i + 1
            End If
        Next
    Next
    For i = 0 To UBound(EVS_Power_Pins_Str) - 1
        If Match_Pins(i) = False Then GoTo errHandler
    Next
    
    
    '''////from each starting, force voltage of all core power to 1v / VDDIO12_GRP to 1.2V / VDDIO18_GRP to 1.8V/////'''
    If UCase(Direction) = "UP" Then
        For i = 0 To UBound(EVS_Power_Pins_Str)
            GL_EVS_Voltage_Pins.pins.item(i).value = EVS_Voltage_Sub
''''''            GL_DCVS_Pin_Output.Pins.Item(i).Value = TheHdw.DCVS.Pins(All_Power_Pins_Str(i)).Voltage.output             ''''Read Vmain/Valt reley
''''''            If GL_DCVS_Pin_Output.Pins.Item(i).Value = 0 Then GoTo errHandler
            If GL_EVS_Voltage_Pins.pins.item(i) = "VDDIO12_GRP" Then GL_EVS_Voltage_Pins.pins.item(EVS_Pins_Index(All_Power_Pins_Str, "VDDIO12_GRP")).value = TheHdw.DCVS.pins("VDDIO12_GRP").Voltage.ValuePerSite   ''' Sicily
            If GL_EVS_Voltage_Pins.pins.item(i) = "VDDIO18_GRP1" Then GL_EVS_Voltage_Pins.pins.item(EVS_Pins_Index(All_Power_Pins_Str, "VDDIO18_GRP1")).value = TheHdw.DCVS.pins("VDDIO18_GRP1").Voltage.ValuePerSite    ''' Sicily
        Next
    End If
    
            
    '''////set EVS voltage value on each core power pins////'''
    If Looping_Contorl Then
        For Each site In TheExec.sites.Active
            EVS_Voltage = Looping_Volt_Start_Value + (TheExec.sites(site).SiteVariableValue(Looping_Index_Name) * 0.02)
        Next site
        S_WaitTime = 0.5
    ElseIf Open_LatchUp_Measure Then
        Ramp_Rate_Function = False
        For Each site In TheExec.sites.Active
            TheExec.sites(site).SiteVariableValue(Looping_Index_Name) = TheExec.flow.var(Looping_Max_Steps_Name).value + 100
        Next site
        EVS_Voltage = LatchUp_Volt_End_Value
        S_WaitTime = 0
    End If
    
    
    For i = 0 To UBound(EVS_Power_Pins_Str)
        GL_EVS_Voltage_Pins.pins.item(EVS_Pins_Index(All_Power_Pins_Str, (EVS_Power_Pins_Str(i)))).value = EVS_Voltage
'''////special pins need special voltage setting////'''
       If UCase(GL_inst_name) Like UCase("*SocMbistEVS*") And EVS_Power_Pins_Str(i) = "VDD_LOW" Then
            If GL_EVS_Voltage_Pins.pins.item(EVS_Pins_Index(All_Power_Pins_Str, "VDD_LOW")).value > 1.43 Then GL_EVS_Voltage_Pins.pins.item(EVS_Pins_Index(All_Power_Pins_Str, "VDD_LOW")).value = 1.43
            If UCase(TheExec.CurrentJob) Like "CP2" And GL_EVS_Voltage_Pins.pins.item(EVS_Pins_Index(All_Power_Pins_Str, "VDD_LOW")).value > 1.3 Then GL_EVS_Voltage_Pins.pins.item(EVS_Pins_Index(All_Power_Pins_Str, "VDD_LOW")).value = 1.3
       End If
       If UCase(GL_inst_name) Like UCase("*GfxSaChainEVS*") And EVS_Power_Pins_Str(i) = "VDD_DISP" Then
                If GL_EVS_Voltage_Pins.pins.item(EVS_Pins_Index(All_Power_Pins_Str, "VDD_DISP")).value > 1.4 Then GL_EVS_Voltage_Pins.pins.item(EVS_Pins_Index(All_Power_Pins_Str, "VDD_DISP")).value = 1.4
       End If
       If UCase(GL_inst_name) Like UCase("*GfxSaChainEVS*") And EVS_Power_Pins_Str(i) = "VDD_SOC" Then
                If GL_EVS_Voltage_Pins.pins.item(EVS_Pins_Index(All_Power_Pins_Str, "VDD_SOC")).value > 1.48 Then GL_EVS_Voltage_Pins.pins.item(EVS_Pins_Index(All_Power_Pins_Str, "VDD_SOC")).value = 1.48
       End If
       If UCase(GL_inst_name) Like UCase("*CpuSaChainEVS*") And UCase(GL_inst_name) Like UCase("*SRVP*") And EVS_Power_Pins_Str(i) = "VDDIO12_GRP" Then GL_EVS_Voltage_Pins.pins.item(EVS_Pins_Index(All_Power_Pins_Str, "VDDIO12_GRP")).value = 1.6
       If UCase(GL_inst_name) Like UCase("*CpuSaChainEVS*") And UCase(GL_inst_name) Like UCase("*SRVP*") And EVS_Power_Pins_Str(i) = "VDDIO18_GRP1" Then GL_EVS_Voltage_Pins.pins.item(EVS_Pins_Index(All_Power_Pins_Str, "VDDIO18_GRP1")).value = 2.4
'''////special pins need special voltage setting////'''
    Next


'''''''''    TheHdw.Alarms.CloseAlarmWindow
'''''''''    TheHdw.Alarms.StartMonitoringAlarms


    TheExec.Datalog.Setup.Shared.ascii.Columns.EnableCustomWidths = True
    TheExec.Datalog.Setup.Shared.ascii.Columns.Parametric.TestName.Width = 110
    TheExec.Datalog.Setup.Shared.ascii.Columns.Functional.TestName.Width = 70
    TheExec.Datalog.Setup.Shared.ascii.Columns.Functional.Pattern.Width = 80
    TheExec.Datalog.ApplySetup
    
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment "-- RAMP " + UCase(Direction) + " Cap On --"
    TheExec.Datalog.WriteComment "InstName = " + GL_inst_name
    TheExec.Datalog.WriteComment "EVS_Volt = " + CStr(EVS_Voltage) + "V"
    TheExec.Datalog.WriteComment ""
    
    GL_inst_name = GL_inst_name + "_" + CStr(EVS_Voltage) + "V"
   
    ProfileMarkLeave VBT_LIB_Common_ProfileMark_2249
 
    
    
    If UCase(Direction) = "UP" Then
    
    
        Set GL_NV_Voltage_Pins = New PinListData
        Set GL_Force_Voltage_Pins = New PinListData
    
    
        Dim VBT_LIB_Common_ProfileMark_2259 As Long: VBT_LIB_Common_ProfileMark_2259 = ProfileMarkEnter(2, "EVS_RAMP_UP_PRE_SETTING")    ' Profile Mark
            
        '''////add pins////'''
        If GL_NV_Voltage_Pins.pins.Count = 0 Then
            For Each pins In EVS_Power_Pins_Str
                GL_NV_Voltage_Pins.AddPin(pins).value = TheHdw.DCVS.pins(pins).Voltage.ValuePerSite
                GL_Force_Voltage_Pins.AddPin(pins).value = TheHdw.DCVS.pins(pins).Voltage.ValuePerSite
            Next
        End If
        
        
        '''////save core power conditions after the last pattern////'''
        For Each pin In Dict_EVS_Pins.Items
            GL_NV_Voltage_Pins.pins(pin).value = TheHdw.DCVS.pins(pin).Voltage.ValuePerSite
            GL_Force_Voltage_Pins.pins(pin).value = GL_NV_Voltage_Pins.pins(pin).value
        Next
        
        
''''''            '''////print all core power values in Datalog before EVS power ramp////'''
''''''            theexec.Datalog.WriteComment "******************Print_Core_EVS_Power_Pins_Voltage_Before_RampUP***********************"
''''''            For i = 0 To UBound(All_Power_Pins_Str)
''''''                    theexec.Datalog.WriteComment All_Power_Pins_Str(i) & ": " & Format(GL_NV_Voltage_Pins.Pins.Item(i).Value, "0.000")
''''''            Next
''''''            theexec.Datalog.WriteComment "**************************************** end ***************************************"

        If UCase(TheExec.CurrentChanMap) Like UCase("*V2*") Then
            '''////For Sicily new P/C, change iFold of VDD_SOC from 15A to 28A////'''
            TheExec.Datalog.WriteComment "**********************Update VDD_SOC IFold to 28A for EVS**************************"
            TheExec.Datalog.WriteComment "Original VDD_SOC IFOLD= " & Format(TheHdw.DCVS.pins("VDD_SOC").CurrentLimit.Source.FoldLimit.level.value, "0.000")
            TheHdw.DCVS.pins("VDD_SOC").CurrentRange.value = TheHdw.DCVS.pins("VDD_SOC").CurrentRange.Max
            TheHdw.DCVS.pins("VDD_SOC").CurrentLimit.Source.FoldLimit.level.value = 28
            TheExec.Datalog.WriteComment "New VDD_SOC IFOLD= " & Format(TheHdw.DCVS.pins("VDD_SOC").CurrentLimit.Source.FoldLimit.level.value, "0.000")
            TheExec.Datalog.WriteComment "**************************************** end ***************************************"
            '''////For Sicily new P/C, change iFold of VDD_SOC from 15A to 28A////'''
        End If

        '''////relax each pin with max H/W limits////'''
        If IFold_Max Then Call IFold_Max_Cal(All_Power_Pins_Str, 0)
'''        DebugPrintFunc ""
       
       
       
        If Ramp_Step_Function Then
            GL_EVS_Ramp_Steps = Ramp_Max_Step
        Else
                '''////calculate how many steps from NV to EVS voltage////'''
                For i = 0 To UBound(EVS_Power_Pins_Str)
                    step = (GL_EVS_Voltage_Pins.pins.item(i).value - GL_NV_Voltage_Pins.pins.item(i).value) / Step_Voltage
                    If step > GL_EVS_Ramp_Steps Then GL_EVS_Ramp_Steps = step
                Next
                If GL_EVS_Ramp_Steps < 3 Then GL_EVS_Ramp_Steps = 3                  '''////default set to 3 steps////'''
        End If
               
        TheExec.Datalog.WriteComment "  -= Start_EVS_RAMP_UP =-"
        TheExec.Datalog.WriteComment "Ramp_Steps: " & GL_EVS_Ramp_Steps
        
        ProfileMarkLeave VBT_LIB_Common_ProfileMark_2259
                
         
         
        Dim VBT_LIB_Common_ProfileMark_2269 As Long: VBT_LIB_Common_ProfileMark_2269 = ProfileMarkEnter(2, "EVS_RAMP_UP")    ' Profile Mark
        
        ''' force EVS core power voltage by steps '''
        For i = 1 To GL_EVS_Ramp_Steps
            For Each pin In Dict_EVS_Pins.Items
                If Ramp_Rate_Function = False Then
                        TheHdw.DCVS.pins(pin).Voltage.ValuePerSite = GL_NV_Voltage_Pins.pins(pin).value + i * (GL_EVS_Voltage_Pins.pins(pin).value - GL_NV_Voltage_Pins.pins(pin).value) / GL_EVS_Ramp_Steps
                        GL_Force_Voltage_Pins.pins(pin).value = GL_NV_Voltage_Pins.pins(pin).value + i * (GL_EVS_Voltage_Pins.pins(pin).value - GL_NV_Voltage_Pins.pins(pin).value) / GL_EVS_Ramp_Steps
                Else
                        If i = GL_EVS_Ramp_Steps Then
                            TheHdw.DCVS.pins(pin).Voltage.ValuePerSite = GL_EVS_Voltage_Pins.pins(pin).value
                        Else
                            TheHdw.DCVS.pins(pin).Voltage.ValuePerSite = TheHdw.DCVS.pins(pin).Voltage.ValuePerSite + (GL_EVS_Voltage_Pins.pins(pin).value - TheHdw.DCVS.pins(pin).Voltage.ValuePerSite) * Ramp_Rate * 0.01
                        End If
                End If
            Next
            
            TheHdw.Wait Rising_Delay_time
            
            If Open_LatchUp_Measure = True Then
                For Each pins In EVS_Power_Pins_Str
                    TheHdw.DCVS.pins(pins).CurrentRange = TheHdw.DCVS.pins(pins).CurrentRange.Max
                Next
                TheHdw.Wait 0.001
            
                Do While TheHdw.DCVS.pins(GL_All_Power_Pins).Capture.IsRunning = True
                Loop
                GL_Measure_Current_EVS = TheHdw.DCVS.pins(GL_All_Power_Pins).Meter.Read(tlStrobe, 1)
                    
                For Each site In TheExec.sites
                    For Each pins In EVS_Power_Pins_Str
                        Force_Voltage = GL_Force_Voltage_Pins.pins(pins).value
                        TheExec.Datalog.WriteComment UCase(GL_inst_name) & ":RampUP:X:" & XCoord & ":Y:" & YCoord & ":Site:" & CStr(site) & ": " & pins & Space(13 - Len(pins)) & ": ForceV: " & Format(Force_Voltage, "0.0000") & " : MeasureI: " & Format(GL_Measure_Current_EVS.pins(pins).value(site), "0.0000")
                    Next pins
                Next site
            End If
        Next i
        DebugPrintFunc ""
            
        ProfileMarkLeave VBT_LIB_Common_ProfileMark_2269
           
           
           
        Dim VBT_LIB_Common_ProfileMark_2279 As Long: VBT_LIB_Common_ProfileMark_2279 = ProfileMarkEnter(2, "EVS_RAMP_STRESS")    ' Profile Mark
           
        '''////==================== EVS stress ==================////'''
        TheHdw.Wait S_WaitTime
            
        ProfileMarkLeave VBT_LIB_Common_ProfileMark_2279
        
        
        
        Dim VBT_LIB_Common_ProfileMark_2289 As Long: VBT_LIB_Common_ProfileMark_2289 = ProfileMarkEnter(2, "EVS_RAMP_UP_POST_SETTING")    ' Profile Mark
        
        TheExec.Datalog.WriteComment "  -= End_EVS_RAMP_UP =-"
        
        TheExec.Datalog.WriteComment ""
        TheExec.Datalog.WriteComment "-- RAMP " + UCase(Direction) + " Cap Off --"
        TheExec.Datalog.WriteComment ""
        TheExec.flow.TestLimit S_WaitTime, , , tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNoScaling, unit:=unitCustom, Tname:=GL_inst_name & "  EVS_Time", customUnit:="sec", formatStr:="%.6f"
        
        For i = 0 To UBound(EVS_Power_Pins_Str)
            PowerVolt1 = Format(TheHdw.DCVS.pins(EVS_Power_Pins_Str(i)).Voltage.ValuePerSite, "0.000")
            TheExec.flow.TestLimit PowerVolt1, , , tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, unit:=unitCustom, Tname:=GL_inst_name & "  EVS_Volt", PinName:=EVS_Power_Pins_Str(EVS_Pins_Index(EVS_Power_Pins_Str, (EVS_Power_Pins_Str(i)))), customUnit:="V"
        Next i
        
''''''        Call PrintCorePowerGate(All_Power_Pins_Str, GL_inst_name) ''' Myst have                  '''''For TTR mask
        
        ProfileMarkLeave VBT_LIB_Common_ProfileMark_2289
            
            
            
    ElseIf UCase(Direction) = "DOWN" Then
        
        If Open_LatchUp_Measure = False Then
                
                
                
        Dim VBT_LIB_Common_ProfileMark_2298 As Long: VBT_LIB_Common_ProfileMark_2298 = ProfileMarkEnter(2, "EVS_RAMP_DOWN_PRE_SETTING")
        
        
            If Ramp_Step_Function Then
                GL_EVS_Ramp_Steps = Ramp_Max_Step
            Else
                    '''////calculate how many steps from NV to EVS voltage////'''
                    For i = 0 To UBound(EVS_Power_Pins_Str)
                        step = (GL_EVS_Voltage_Pins.pins.item(i).value - GL_NV_Voltage_Pins.pins.item(i).value) / Step_Voltage
                        If step > GL_EVS_Ramp_Steps Then GL_EVS_Ramp_Steps = step
                    Next
                    If GL_EVS_Ramp_Steps < 3 Then GL_EVS_Ramp_Steps = 3                  '''////default set to 3 steps////'''
            End If
    
                               
            TheExec.Datalog.WriteComment "  -= Start_EVS_RAMP_DOWN =-"
    
            TheExec.Datalog.WriteComment "Ramp_Steps: " & GL_EVS_Ramp_Steps         '''////GL_EVS_Ramp_Steps for ENG////'''
         
        ProfileMarkLeave VBT_LIB_Common_ProfileMark_2298
         
         
         
        Dim VBT_LIB_Common_ProfileMark_2299 As Long: VBT_LIB_Common_ProfileMark_2299 = ProfileMarkEnter(2, "EVS_RAMP_DOWN")

            For i = 1 To GL_EVS_Ramp_Steps
                For Each pin In Dict_EVS_Pins.Items
                    If Ramp_Rate_Function = False Then
                        TheHdw.DCVS.pins(pin).Voltage.ValuePerSite = GL_EVS_Voltage_Pins.pins(pin).value - i * (GL_EVS_Voltage_Pins.pins(pin).value - GL_NV_Voltage_Pins.pins(pin).value) / GL_EVS_Ramp_Steps
                        GL_Force_Voltage_Pins.pins(pin).value = GL_EVS_Voltage_Pins.pins(pin).value - i * (GL_EVS_Voltage_Pins.pins(pin).value - GL_NV_Voltage_Pins.pins(pin).value) / GL_EVS_Ramp_Steps
                    Else
                        If i = GL_EVS_Ramp_Steps Then
                            TheHdw.DCVS.pins(pin).Voltage.ValuePerSite = GL_NV_Voltage_Pins.pins(pin).value
                        Else
                            TheHdw.DCVS.pins(pin).Voltage.ValuePerSite = TheHdw.DCVS.pins(pin).Voltage.ValuePerSite - (TheHdw.DCVS.pins(pin).Voltage.ValuePerSite - GL_NV_Voltage_Pins.pins(pin).value) * Ramp_Rate * 0.01
                        End If
                    End If
                Next pin
                
                TheHdw.Wait Rising_Delay_time
                
            Next i
'''            DebugPrintFunc ""
        
        ProfileMarkLeave VBT_LIB_Common_ProfileMark_2299
        
        Else
        
             TheExec.Datalog.WriteComment "  -= Start_EVS_RAMP_DOWN =-"

            GL_EVS_Ramp_Steps = 10
            TheExec.Datalog.WriteComment "Ramp_Steps: " & GL_EVS_Ramp_Steps         '''////GL_EVS_Ramp_Steps for ENG////'''
            
            
           
            For i = 1 To GL_EVS_Ramp_Steps
                For Each pin In Dict_EVS_Pins.Items
                    TheHdw.DCVS.pins(pin).Voltage.ValuePerSite = GL_EVS_Voltage_Pins.pins(pin).value - i * (GL_EVS_Voltage_Pins.pins(pin).value - GL_NV_Voltage_Pins.pins(pin).value) / GL_EVS_Ramp_Steps
                    GL_Force_Voltage_Pins.pins(pin).value = GL_EVS_Voltage_Pins.pins(pin).value - i * (GL_EVS_Voltage_Pins.pins(pin).value - GL_NV_Voltage_Pins.pins(pin).value) / GL_EVS_Ramp_Steps
                Next
                
                TheHdw.Wait Rising_Delay_time
                
                For Each pins In EVS_Power_Pins_Str
                    TheHdw.DCVS.pins(pins).CurrentRange = TheHdw.DCVS.pins(pins).CurrentRange.Max
                Next
                TheHdw.Wait 0.001
    
                Do While TheHdw.DCVS.pins(GL_All_Power_Pins).Capture.IsRunning = True
                Loop
                GL_Measure_Current_EVS = TheHdw.DCVS.pins(GL_All_Power_Pins).Meter.Read(tlStrobe, 1)
    
                For Each site In TheExec.sites
                    For Each pins In EVS_Power_Pins_Str
                        Force_Voltage = GL_Force_Voltage_Pins.pins(pins).value
                        TheExec.Datalog.WriteComment UCase(GL_inst_name) & ":RampDown:X:" & XCoord & ":Y:" & YCoord & ":Site:" & CStr(site) & ": " & pins & Space(13 - Len(pins)) & ": ForceV: " & Format(Force_Voltage, "0.0000") & " : MeasureI: " & Format(GL_Measure_Current_EVS.pins(pins).value(site), "0.0000")
                    Next
                Next site
            Next i
            
        End If
        
        
        
        Dim VBT_LIB_Common_ProfileMark_2300 As Long: VBT_LIB_Common_ProfileMark_2300 = ProfileMarkEnter(2, "EVS_RAMP_DOWN_POST_SETTING")
                
        TheExec.Datalog.WriteComment "  -= End_EVS_RAMP_DOWN =-"
        
        TheExec.Datalog.WriteComment "EVS Ramp Down : " & GL_inst_name
                        
        ProfileMarkLeave VBT_LIB_Common_ProfileMark_2300
    
    
    
        Dim VBT_LIB_Common_ProfileMark_2309 As Long: VBT_LIB_Common_ProfileMark_2309 = ProfileMarkEnter(2, "EVS_RAMP_ALARM_CHECK")
        
        
        Call CheckCorePowerGate(EVS_Power_Pins_Str) ''' Myst have
        
        Dim Alarm_Result_Site As New SiteLong
        
        For Each site In TheExec.sites.Active
        
            If TheHdw.Alarms.Check = True Or GL_Gate_Check(site) Then
                Alarm_Result_Site(site) = 0
                TheExec.sites.item(site).FlagState("F_" & DomainStr & "_alarm_check") = logicTrue
                If Open_LatchUp_Measure Then
                Else
                        TheExec.sites(site).SiteVariableValue(Looping_Index_Name) = TheExec.flow.var(Looping_Max_Steps_Name).value + 100
                        TheExec.sites(site).SiteVariableValue(Multi_EVS_Index_Name) = TheExec.flow.var(Looping_Max_Steps_Name).value + 100
                End If
                
            ElseIf TheHdw.Alarms.Check = False Then
            
                Alarm_Result_Site(site) = 1
                
            End If
                
        Next site
    
        TheExec.flow.TestLimit Alarm_Result_Site, 1, 1, Tname:=GL_inst_name
        
        ProfileMarkLeave VBT_LIB_Common_ProfileMark_2309
       
       
       
        Dim VBT_LIB_Common_ProfileMark_2319 As Long: VBT_LIB_Common_ProfileMark_2319 = ProfileMarkEnter(2, "EVS_RAMP_PRINT_GATE_STATUS")
      
        TheExec.Datalog.WriteComment ""
        TheExec.Datalog.WriteComment "-- RAMP " + UCase(Direction) + " Cap Off --"
        TheExec.Datalog.WriteComment ""
        
        Call PrintCorePowerGate(EVS_Power_Pins_Str, GL_inst_name) ''' Myst have
        GL_EVS_Ramp_Steps = 0
        
        ProfileMarkLeave VBT_LIB_Common_ProfileMark_2319
        
        DebugPrintFunc ""
    
    
    End If

Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
        
End Function


Public Function EVS_Pins_Index(pins() As String, pin As String) As Integer

    Dim i As Integer
    For i = 0 To UBound(pins)
        If pins(i) = pin Then EVS_Pins_Index = i
    Next

End Function


Public Function DVS_Power_Ramp(Direction As String, EVS_core_pin As String, EVS_Voltage As Double, S_WaitTime As Double, Block As String, FuncPat As Pattern, PatLoopCount As Long, result_mode As tlResultMode) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "EVS_Power_Ramp"
    Dim EVS_Voltage_Sub, SlewRate_Step, Step_WaitTime, Looping_Volt_Start_Value, LatchUp_Volt_End_Value As Double
    Dim Ins_Type As Long:: Ins_Type = 0
    Dim Plot_Profile, Profile_Volt, Plot_IV, Plot_Vtrig, ENG_Ifold_Relax As Boolean
    
Dim VBT_LIB_Common_ProfileMark_2249 As Long: VBT_LIB_Common_ProfileMark_2249 = ProfileMarkEnter(2, "EVS_RAMP_PRE")

    '''======================================================================================= '''
    ''' Setting Other_Power_V: Slew_Rate_V: Step_Time_S '''
    EVS_Voltage_Sub = 1: SlewRate_Step = 0.05: Step_WaitTime = 0.001
    ''' Sorting Power Pins '''
    GL_All_Power_Pins = ("VDD_CPU,VDD_CPU1,VDD_CPU2,VDD_CPU3,VDD_CPU4,VDD_DISP,VDD_LOW,VDD_SOC,VDD_SRAM,VDD_SRAM_CPU,VDDIO12_GRP")
    '''======================================================================================= '''
    '''=================== EVS debug setting ============================ '''
    Plot_Profile = False                      ''' Collect current/voltage profile flag
    Profile_Volt = False                      ''' Enable voltage profile collection
    Plot_IV = False                             ''' Collect IV-curve flag
    Plot_Vtrig = False                         ''' Collect voltage trigger flag
    ENG_Ifold_Relax = False              ''' ENG Ifold relax to H/W max
    '''=================== EVS debug setting ============================ '''
    
    
    ''' setting current/voltage profile samplerate/samplesize '''
    ''' ========Define sample rate and sample size for profile base on pattern execution time ======================
    
    Dim SampleRateT As Double
    Dim SampleSizeT As Long
    Dim inst_name As String
    Dim step As Long
    Dim pin As Variant
    Dim i, j As Long
    Dim DomainStr As String
    Dim LoopCnt As Long
    Dim Force_Voltage As Double
    Dim Force_EVS_Pin As String
    Dim All_Power_Pins_Str() As String
    Dim EVS_core_pin_Str() As String
    Dim PinsCnt As Long
    Dim pins As Variant
    Dim PowerVolt1 As Double
    Dim Match_Pins() As Boolean
    
    If Ins_Type = 0 Then
            SampleRateT = 50000: SampleSizeT = 256000
            Plot_Profile_Pins = "VDD_SOC"                       '''"VDD_CPU,VDD_CPU1,VDD_CPU4"
    ElseIf Ins_Type = 1 Then
            SampleRateT = 3125: SampleSizeT = 16000
            Plot_Profile_Pins = "VDD_DISP,VDD_LOW,VDD_SRAM"     '''"VDD_CPU2,VDD_CPU3,VDD_SRAM_CPU"
    End If
    
    If Plot_IV And Plot_Vtrig Then GoTo errHandler
        
    
    ''' the looping flow is for vtrig collection, if that is production flow, adjust the looping value to max value, it will jump out the looping flow'''
    Dim EVS_Loop_Var_Max As New SiteLong
    For Each site In TheExec.sites.Active
        EVS_Loop_Var_Max(site) = 100
        If UCase(Direction) = "UP" Then GL_Gate_Check(site) = False
        If Plot_Vtrig = False And Plot_IV = False Then TheExec.sites(site).SiteVariableValue(Looping_Index_Name) = EVS_Loop_Var_Max(site)
    Next site
    

    Dim Pre_InstanceName As String
    
    Pre_InstanceName = m_InstanceName
    m_InstanceName = LCase(TheExec.DataManager.instancename)
    If UCase(m_InstanceName) Like UCase("*GPIODVS*") Then
    Else
        m_InstanceName = Pre_InstanceName
    End If
    
    inst_name = TheExec.DataManager.instancename '''+ "_" + UCase(m_InstanceName)
    DomainStr = LCase(Block)
    
    ''' current loop to whitch blocks '''
    If UCase(gB_block_type) Like UCase("*MBIST*") Then
            If UCase(gB_block_type) Like UCase("*B*") Or UCase(gB_block_type) Like UCase("*SOC*") Then inst_name = inst_name & "_" & gB_block_type & Format(currentBlock_loopCnt, "00")
    End If
    
     
    If UCase(inst_name) Like UCase("GPIODVS*") Then
        inst_name = TheExec.DataManager.instancename
    End If
    
    
    If UCase(Direction) = "UP" Then Call TheHdw.Digital.ApplyLevelsTiming(True, True, True, tlPowered)
    

    EVS_core_pin_Str = Split(EVS_core_pin, ",")
    
    Call TheExec.DataManager.DecomposePinList(GL_All_Power_Pins, All_Power_Pins_Str, PinsCnt)
        
    
    If GL_EVS_Voltage_Pins.pins.Count = 0 Then
        For Each pins In All_Power_Pins_Str
            GL_EVS_Voltage_Pins.AddPin(pins).value = EVS_Voltage_Sub
            GL_DCVS_Pin_Output.AddPin(pins).value = 0
        Next
    End If
    
    
    ''' check the EVS core pin string fill in argument if correct '''
    i = 0: ReDim Match_Pins(99)
    For Each pins In All_Power_Pins_Str
        For Each pin In EVS_core_pin_Str
            If pin = pins Then
                Match_Pins(i) = True
                i = i + 1
            End If
        Next
    Next
    For i = 0 To UBound(EVS_core_pin_Str) - 1
        If Match_Pins(i) = False Then GoTo errHandler
    Next
    
    
    ''' from each starting, force voltage of all core power to 1v and VDDIO18_GRP to 1.8V '''
    If UCase(Direction) = "UP" Then
        For i = 0 To UBound(All_Power_Pins_Str)
            GL_EVS_Voltage_Pins.pins.item(i).value = EVS_Voltage_Sub
            GL_DCVS_Pin_Output.pins.item(i).value = TheHdw.DCVS.pins(All_Power_Pins_Str(i)).Voltage.Output
            If GL_DCVS_Pin_Output.pins.item(i).value = 0 Then GoTo errHandler
            If GL_EVS_Voltage_Pins.pins.item(i) = "VDDIO12_GRP" Then GL_EVS_Voltage_Pins.pins.item(EVS_Pins_Index(All_Power_Pins_Str, "VDDIO12_GRP")).value = 1.2    ''' Skua
        Next
    End If
    
    
    ''' debug ''' debug ''' debug ''' debug '''
    Select Case True
        Case UCase(inst_name) Like UCase("*CpuSaChainDVS*")
            Looping_Volt_Start_Value = 1.4: LatchUp_Volt_End_Value = 2
        Case UCase(inst_name) Like UCase("*CpuMbistDVS*")
            Looping_Volt_Start_Value = 1.4: LatchUp_Volt_End_Value = 2
        Case UCase(inst_name) Like UCase("*SocSaChainDVS*")
            Looping_Volt_Start_Value = 1.4: LatchUp_Volt_End_Value = 2
        Case UCase(inst_name) Like UCase("*SocMbistDVS*")
            Looping_Volt_Start_Value = 1.4: LatchUp_Volt_End_Value = 2
        Case UCase(inst_name) Like UCase("*GPIODVS*")
            Looping_Volt_Start_Value = 1.6: LatchUp_Volt_End_Value = 1.6
    End Select
                
            
    ''' set EVS voltage value on each core power pins '''
    If Plot_Vtrig Then
        For Each site In TheExec.sites.Active
            EVS_Voltage = Looping_Volt_Start_Value + (TheExec.sites(site).SiteVariableValue(Looping_Index_Name) * 0.02)
        Next site
    ElseIf Plot_IV Then
        For Each site In TheExec.sites.Active
            LoopCnt = CLng(TheExec.sites(site).SiteVariableValue(Looping_Index_Name))
        Next site
        EVS_Voltage = LatchUp_Volt_End_Value
        S_WaitTime = 0
    End If
    
    For i = 0 To UBound(EVS_core_pin_Str)
        GL_EVS_Voltage_Pins.pins.item(EVS_Pins_Index(All_Power_Pins_Str, (EVS_core_pin_Str(i)))).value = EVS_Voltage
'''        If UCase(inst_name) Like UCase("*SocMbistEVS*") And EVS_core_pin_Str(i) = "VDD_CPU_SRAM" Then GL_EVS_Voltage_Pins.Pins.item(EVS_Pins_Index(All_Power_Pins_Str, "VDD_CPU_SRAM")).Value = 1.3     '''Cyprus
    Next


'''''''''    TheHdw.Alarms.CloseAlarmWindow
'''''''''    TheHdw.Alarms.StartMonitoringAlarms


    TheExec.Datalog.Setup.Shared.ascii.Columns.EnableCustomWidths = True
    TheExec.Datalog.Setup.Shared.ascii.Columns.Parametric.TestName.Width = 110
    TheExec.Datalog.Setup.Shared.ascii.Columns.Functional.TestName.Width = 70
    TheExec.Datalog.Setup.Shared.ascii.Columns.Functional.Pattern.Width = 80
    TheExec.Datalog.ApplySetup
    
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment "-- RAMP " + Direction + " Cap On --"
    TheExec.Datalog.WriteComment "InstName = " + inst_name
    TheExec.Datalog.WriteComment "EVS_Volt = " + CStr(EVS_Voltage) + "V"
    TheExec.Datalog.WriteComment ""
    
    inst_name = inst_name + "_" + CStr(EVS_Voltage) + "V"
   
   
ProfileMarkLeave VBT_LIB_Common_ProfileMark_2249
 
 
    If UCase(Direction) = "UP" Then


Dim VBT_LIB_Common_ProfileMark_2259 As Long: VBT_LIB_Common_ProfileMark_2259 = ProfileMarkEnter(2, "EVS_RAMP_UP_SETUP")    ' Profile Mark


    
        ''' add pins '''
        If GL_NV_Voltage_Pins.pins.Count = 0 Then
            For Each pins In All_Power_Pins_Str
                GL_NV_Voltage_Pins.AddPin(pins).value = TheHdw.DCVS.pins(pins).Voltage.Main.ValuePerSite
                GL_Force_Voltage_Pins.AddPin(pins).value = TheHdw.DCVS.pins(pins).Voltage.Main.ValuePerSite
            Next
        End If
        
        
        ''' save core power conditions after the last pattern '''
        For i = 0 To UBound(All_Power_Pins_Str)
            If GL_DCVS_Pin_Output.pins.item(i).value = 1 Then
                GL_NV_Voltage_Pins.pins.item(i).value = TheHdw.DCVS.pins(All_Power_Pins_Str(i)).Voltage.Main.ValuePerSite
            Else
                GL_NV_Voltage_Pins.pins.item(i).value = TheHdw.DCVS.pins(All_Power_Pins_Str(i)).Voltage.Alt.ValuePerSite
            End If
            GL_Force_Voltage_Pins.pin(i).value = GL_NV_Voltage_Pins.pins.item(i).value
        Next
        
        
        ''' print all core power values in Datalog before EVS power ramp '''
        TheExec.Datalog.WriteComment "******************Print_Core_Power_Pins_Voltage_Apply_Default***********************"
        For i = 0 To UBound(All_Power_Pins_Str)
                TheExec.Datalog.WriteComment All_Power_Pins_Str(i) & ": " & Format(GL_NV_Voltage_Pins.pins.item(i).value, "0.000")
        Next
        TheExec.Datalog.WriteComment "**************************************** end ***************************************"
        
        
'''        ''' for Cyprus, change iFold of VDD_FIXED from 0.9A to 2A '''
'''        TheExec.DataLog.WriteComment "**********************Update VDD_FIXED IFold to 2A for EVS**************************"
'''        TheExec.DataLog.WriteComment "Original VDD_FIXED IFOLD= " & Format(Thehdw.DCVS.Pins("VDD_FIXED").CurrentLimit.Source.FoldLimit.Level.Value, "0.000")
'''        Thehdw.DCVS.Pins("VDD_FIXED").CurrentLimit.Source.FoldLimit.Level.Value = 2
'''        TheExec.DataLog.WriteComment "New VDD_FIXED IFOLD= " & Format(Thehdw.DCVS.Pins("VDD_FIXED").CurrentLimit.Source.FoldLimit.Level.Value, "0.000")
'''        TheExec.DataLog.WriteComment "**************************************** end ***************************************"


        ''' relax each pin with max H/W limits for vtrig and hysteresis '''
        If Plot_IV Or Plot_Vtrig Or ENG_Ifold_Relax Then Call IFold_Max_Cal(All_Power_Pins_Str, 0)
        DebugPrintFunc ""
       
       
        ''' calculate how many steps from NV to EVS voltage '''
        For i = 0 To UBound(All_Power_Pins_Str)
            step = (GL_EVS_Voltage_Pins.pins.item(i).value - GL_NV_Voltage_Pins.pins.item(i).value) / SlewRate_Step
            If step > GL_EVS_Ramp_Steps Then GL_EVS_Ramp_Steps = step
        Next
       
        If GL_EVS_Ramp_Steps < 5 Then GL_EVS_Ramp_Steps = 5         ''' default set to 5 steps
        If Plot_IV Then GL_EVS_Ramp_Steps = 50                  ''' more setups for plot curve
        
        TheExec.Datalog.WriteComment "  -= Start_EVS_RAMP_UP =-"
       
        TheExec.Datalog.WriteComment "Ramp_Steps: " & GL_EVS_Ramp_Steps
        
        
ProfileMarkLeave VBT_LIB_Common_ProfileMark_2259
        
Dim VBT_LIB_Common_ProfileMark_2269 As Long: VBT_LIB_Common_ProfileMark_2269 = ProfileMarkEnter(2, "EVS_RAMP_UP_RUN")    ' Profile Mark
 
        
        ''' force EVS core power voltage and stress by steps '''
        For i = 1 To GL_EVS_Ramp_Steps
            For j = 0 To UBound(All_Power_Pins_Str)
                If GL_DCVS_Pin_Output.pins.item(j).value = 1 Then
                    TheHdw.DCVS.pins(All_Power_Pins_Str(j)).Voltage.Main.ValuePerSite = GL_NV_Voltage_Pins.pins.item(j).value + i * (GL_EVS_Voltage_Pins.pins.item(j).value - GL_NV_Voltage_Pins.pins.item(j).value) / GL_EVS_Ramp_Steps
                Else
                    TheHdw.DCVS.pins(All_Power_Pins_Str(j)).Voltage.Alt.ValuePerSite = GL_NV_Voltage_Pins.pins.item(j).value + i * (GL_EVS_Voltage_Pins.pins.item(j).value - GL_NV_Voltage_Pins.pins.item(j).value) / GL_EVS_Ramp_Steps
                End If
                GL_Force_Voltage_Pins.pins(j).value = GL_NV_Voltage_Pins.pins.item(j).value + i * (GL_EVS_Voltage_Pins.pins.item(j).value - GL_NV_Voltage_Pins.pins.item(j).value) / GL_EVS_Ramp_Steps
            Next
            
            TheHdw.Wait Step_WaitTime
            
            If Plot_IV = True Then
                For Each pins In All_Power_Pins_Str
                    TheHdw.DCVS.pins(pins).CurrentRange = TheHdw.DCVS.pins(pins).CurrentRange.Max
                Next
                TheHdw.Wait 0.001
            
                Do While TheHdw.DCVS.pins(GL_All_Power_Pins).Capture.IsRunning = True
                Loop
                GL_Measure_Current_EVS = TheHdw.DCVS.pins(GL_All_Power_Pins).Meter.Read(tlStrobe, 1)
                    
                For Each site In TheExec.sites
                
'''                  For Each Pin In GL_Measure_Current_EVS.Pins
                    
                    pin = UCase(EVS_core_pin_Str(LoopCnt))
                    For Each pins In All_Power_Pins_Str
                        If pin = pins Then Force_Voltage = GL_Force_Voltage_Pins.pins(EVS_Pins_Index(All_Power_Pins_Str, CStr(pin))).value
                    Next
                    
                    TheExec.Datalog.WriteComment UCase(inst_name) & ":RampUP:X:" & XCoord & ":Y:" & YCoord & ":Site:" & CStr(site) & ": " & pin & ": ForceV: " & Format(Force_Voltage, "0.0000") & " : MeasureI: " & Format(GL_Measure_Current_EVS.pins(pin).value(site), "0.0000")
'''                    TheExec.Flow.TestLimit GL_Measure_Current_EVS.Pins(Pin).Value(Site), -1, 5, tlSignGreaterEqual, tlSignLessEqual, ScaleType:=scaleNoScaling, unit:=unitAmp, _
'''                    Tname:=m_InstanceName & ":RampUp:X:" & Xcoord & ":Y:" & Ycoord & ":Site:" & CStr(Site) & ":" & Pin & ":ForceV:" & Force_Voltage, customUnit:="A"  'BurstResult=1:Pass

'''                  Next Pin

                Next site
            End If
        Next i
        DebugPrintFunc ""
    
ProfileMarkLeave VBT_LIB_Common_ProfileMark_2269
   
Dim VBT_LIB_Common_ProfileMark_2279 As Long: VBT_LIB_Common_ProfileMark_2279 = ProfileMarkEnter(2, "EVS_RAMP_UP_STRESS")    ' Profile Mark
 
   
        '==================== EVS stress =================='
        TheHdw.Wait S_WaitTime
    
ProfileMarkLeave VBT_LIB_Common_ProfileMark_2279

Dim VBT_LIB_Common_ProfileMark_2289 As Long: VBT_LIB_Common_ProfileMark_2289 = ProfileMarkEnter(2, "EVS_RAMP_UP_POST")    ' Profile Mark

        TheExec.Datalog.WriteComment "  -= End_EVS_RAMP_UP =-"
        
        TheExec.Datalog.WriteComment ""
        TheExec.Datalog.WriteComment "-- RAMP " + Direction + " Cap Off --"
        TheExec.Datalog.WriteComment ""
        TheExec.flow.TestLimit S_WaitTime, , , tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNoScaling, unit:=unitCustom, Tname:=TheExec.DataManager.instancename & "  EVS_Time", customUnit:="sec", formatStr:="%.6f"
        
        
        For i = 0 To UBound(EVS_core_pin_Str)
            If GL_DCVS_Pin_Output.pins.item(EVS_Pins_Index(All_Power_Pins_Str, EVS_core_pin_Str(i))).value = 1 Then
                PowerVolt1 = Format(TheHdw.DCVS.pins(EVS_core_pin_Str(i)).Voltage.Main.ValuePerSite, "0.000")
            Else
                PowerVolt1 = Format(TheHdw.DCVS.pins(EVS_core_pin_Str(i)).Voltage.Alt.ValuePerSite, "0.000")
            End If
            TheExec.flow.TestLimit PowerVolt1, , , tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, unit:=unitCustom, Tname:=inst_name & "  EVS_Volt", PinName:=All_Power_Pins_Str(EVS_Pins_Index(All_Power_Pins_Str, (EVS_core_pin_Str(i)))), customUnit:="V"
        Next
'''        TheExec.DataLog.WriteComment "                " + inst_name + "_EVS_voltage: " + CStr(EVS_voltage) + "V"
'''        TheExec.DataLog.WriteComment "                " + inst_name + "_Stress_Time: " + CStr(S_WaitTime) + "S"
        
        
        Call PrintCorePowerGate(All_Power_Pins_Str, inst_name) ''' Myst have
    
    
        '''================= DVS pattern running ====================='''
        If FuncPat.value <> "" Then
                Call PrLoadPattern(FuncPat.value)
                For i = 1 To PatLoopCount
                    Call TheHdw.patterns(FuncPat).test(pfAlways, 0, tlResultModeDomain)
                Next
        End If
        '''================= DVS pattern running ====================='''
    
'''    TheExec.DataLog.WriteComment ""
'''    TheExec.DataLog.WriteComment "**************Print_Core_Power_Pins_Voltage*********************"
'''    For Each Pins In All_Power_Pins_Str
'''        PowerVolt1 = Format(thehdw.DCVS.Pins(Pins).Voltage.Main.Value, "0.000")
'''        TheExec.DataLog.WriteComment Pins & "= " & PowerVolt1
'''    Next
'''    TheExec.DataLog.WriteComment "******************************** end ***************************"

ProfileMarkLeave VBT_LIB_Common_ProfileMark_2289

ElseIf UCase(Direction) = "DOWN" Then
        
Dim VBT_LIB_Common_ProfileMark_2299 As Long: VBT_LIB_Common_ProfileMark_2299 = ProfileMarkEnter(2, "EVS_RAMP_DOWN_RUN")


    TheExec.Datalog.WriteComment "  -= Start_EVS_RAMP_DOWN =-"


    If Plot_IV = False Then
        
'''        TheExec.Datalog.WriteComment "Ramp_Steps: " & 1         ''' GL_EVS_Ramp_Steps = 1 for TTR
'''
'''        ''' ============ ramp down for 1 step (TTR by 20180402 Leslie) ========== '''
'''        For i = 0 To UBound(All_Power_Pins_Str)
'''            If GL_DCVS_Pin_Output.Pins.item(i).Value = 1 Then
'''                TheHdw.DCVS.Pins(All_Power_Pins_Str(i)).Voltage.Main.Value = GL_NV_Voltage_Pins.Pins.item(i).Value      ''' normal flow running
'''            Else
'''                TheHdw.DCVS.Pins(All_Power_Pins_Str(i)).Voltage.Alt.Value = GL_NV_Voltage_Pins.Pins.item(i).Value       ''' normal flow running
'''            End If
'''        Next
'''        DebugPrintFunc ""
        
        TheExec.Datalog.WriteComment "Ramp_Steps: " & GL_EVS_Ramp_Steps         '''GL_EVS_Ramp_Steps for ENG
        
        For i = 1 To GL_EVS_Ramp_Steps
            For j = 0 To UBound(All_Power_Pins_Str)
                If GL_DCVS_Pin_Output.pins.item(j).value = 1 Then
                    TheHdw.DCVS.pins(All_Power_Pins_Str(j)).Voltage.Main.ValuePerSite = GL_EVS_Voltage_Pins.pins.item(j).value - i * (GL_EVS_Voltage_Pins.pins.item(j).value - GL_NV_Voltage_Pins.pins.item(j).value) / GL_EVS_Ramp_Steps
                Else
                    TheHdw.DCVS.pins(All_Power_Pins_Str(j)).Voltage.Alt.ValuePerSite = GL_EVS_Voltage_Pins.pins.item(j).value - i * (GL_EVS_Voltage_Pins.pins.item(j).value - GL_NV_Voltage_Pins.pins.item(j).value) / GL_EVS_Ramp_Steps
                End If
                GL_Force_Voltage_Pins.pins(j).value = GL_EVS_Voltage_Pins.pin(j).value - i * (GL_EVS_Voltage_Pins.pins(j).value - GL_NV_Voltage_Pins.pins.item(j).value) / GL_EVS_Ramp_Steps
            Next
            
            TheHdw.Wait Step_WaitTime
        Next
        DebugPrintFunc ""
        
    Else
        
        GL_EVS_Ramp_Steps = 25
        TheExec.Datalog.WriteComment "Ramp_Steps: " & GL_EVS_Ramp_Steps         '''GL_EVS_Ramp_Steps for ENG
        
        For i = 1 To GL_EVS_Ramp_Steps
            For j = 0 To UBound(All_Power_Pins_Str)
                If GL_DCVS_Pin_Output.pins.item(j).value = 1 Then
                    TheHdw.DCVS.pins(All_Power_Pins_Str(j)).Voltage.Main.ValuePerSite = GL_EVS_Voltage_Pins.pins.item(j).value - i * (GL_EVS_Voltage_Pins.pins.item(j).value - GL_NV_Voltage_Pins.pins.item(j).value) / GL_EVS_Ramp_Steps
                Else
                    TheHdw.DCVS.pins(All_Power_Pins_Str(j)).Voltage.Alt.ValuePerSite = GL_EVS_Voltage_Pins.pins.item(j).value - i * (GL_EVS_Voltage_Pins.pins.item(j).value - GL_NV_Voltage_Pins.pins.item(j).value) / GL_EVS_Ramp_Steps
                End If
                GL_Force_Voltage_Pins.pins(j).value = GL_EVS_Voltage_Pins.pin(j).value - i * (GL_EVS_Voltage_Pins.pins(j).value - GL_NV_Voltage_Pins.pins.item(j).value) / GL_EVS_Ramp_Steps
            Next
            
            TheHdw.Wait Step_WaitTime
            
            For Each pins In All_Power_Pins_Str
                TheHdw.DCVS.pins(pins).CurrentRange = TheHdw.DCVS.pins(pins).CurrentRange.Max
            Next
            TheHdw.Wait 0.001

            Do While TheHdw.DCVS.pins(GL_All_Power_Pins).Capture.IsRunning = True
            Loop
            GL_Measure_Current_EVS = TheHdw.DCVS.pins(GL_All_Power_Pins).Meter.Read(tlStrobe, 1)

            For Each site In TheExec.sites

                pin = UCase(EVS_core_pin_Str(LoopCnt))
                For Each pins In All_Power_Pins_Str
                    If pin = pins Then Force_Voltage = GL_Force_Voltage_Pins.pins(EVS_Pins_Index(All_Power_Pins_Str, CStr(pin))).value
                Next

                TheExec.Datalog.WriteComment UCase(inst_name) & ":RampDown:X:" & XCoord & ":Y:" & YCoord & ":Site:" & CStr(site) & ": " & pin & ": ForceV: " & Format(Force_Voltage, "0.0000") & " : MeasureI: " & Format(GL_Measure_Current_EVS.pins(pin).value(site), "0.0000")
'''                       TheExec.Flow.TestLimit GL_Measure_Current_EVS.Pins(Pin).Value(Site), -1, 5, tlSignGreaterEqual, tlSignLessEqual, ScaleType:=scaleNoScaling, unit:=unitAmp, _
'''                       Tname:=m_InstanceName & ":RampDown:X:" & Xcoord & ":Y:" & Ycoord & ":Site:" & CStr(Site) & ":" & Pin & ":ForceV:" & Force_Voltage, customUnit:="A"  'BurstResult=1:Pass

            Next site
        Next i
    End If
        
    TheExec.Datalog.WriteComment "  -= End_EVS_RAMP_DOWN =-"
    
    TheExec.Datalog.WriteComment "EVS Ramp Down : " & inst_name


ProfileMarkLeave VBT_LIB_Common_ProfileMark_2299
    
Dim VBT_LIB_Common_ProfileMark_2309 As Long: VBT_LIB_Common_ProfileMark_2309 = ProfileMarkEnter(2, "EVS_ALARM_CHECK")

    Call CheckCorePowerGate(All_Power_Pins_Str)  ''' Myst have

    Dim Alarm_Result_Site As New SiteLong
    
    For Each site In TheExec.sites.Active
    
        If TheHdw.Alarms.Check Or GL_Gate_Check(site) Then
            Alarm_Result_Site(site) = 0
            TheExec.sites.item(site).FlagState("F_" & DomainStr & "_alarm_check") = logicTrue
            If Plot_IV Then
            Else
                    TheExec.sites(site).SiteVariableValue(Looping_Index_Name) = EVS_Loop_Var_Max(site)
            End If
            
            
'''            '''Vtrig twice in succession than fail out
'''            If Plot_IV = True Then
'''                TheExec.sites(Site).SiteVariableValue(Looping_Index_Name) = EVS_Loop_Var_Max(Site)
'''            Else
'''                If TheExec.sites(Site).SiteVariableValue("EVS_Loop_Var_Post") = 0 Then
'''                    TheExec.sites(Site).SiteVariableValue(Looping_Index_Name) = EVS_Loop_Var_Max(Site)
'''                    TheExec.sites(Site).SiteVariableValue("EVS_Loop_Var_Post") = 1
'''                Else
'''                    TheExec.sites(Site).SiteVariableValue("EVS_Loop_Var_Post") = TheExec.sites(Site).SiteVariableValue("EVS_Loop_Var_Post") - 1
'''                End If
'''            End If
            
'''           TheExec.Flow.TestLimit 0, 1, 1, Tname:=inst_name

        ElseIf TheHdw.Alarms.Check = False Then
        
            Alarm_Result_Site(site) = 1
            
'''            TheExec.Flow.TestLimit 1, 1, 1, Tname:=inst_name

        End If
            
    Next site
    
    TheExec.flow.TestLimit Alarm_Result_Site, 1, 1, Tname:=inst_name
    
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment "-- RAMP " + Direction + " Cap Off --"
    TheExec.Datalog.WriteComment ""
    
    Call PrintCorePowerGate(All_Power_Pins_Str, inst_name) ''' Myst have
    GL_EVS_Ramp_Steps = 0
    
    
    ''' exit for loop measure pins for plot IV curve '''
    For Each site In TheExec.sites.Active
        If Plot_IV Then
                If UCase(inst_name) Like UCase("*SA*") And LoopCnt = CLng(UBound(EVS_core_pin_Str)) Then TheExec.sites(site).SiteVariableValue(Looping_Index_Name) = EVS_Loop_Var_Max(site)
                If UCase(inst_name) Like UCase("*WR1*") And LoopCnt = CLng(UBound(EVS_core_pin_Str)) Then TheExec.sites(site).SiteVariableValue(Looping_Index_Name) = EVS_Loop_Var_Max(site)
                If UCase(inst_name) Like UCase("*WRB*") And LoopCnt = CLng(UBound(EVS_core_pin_Str)) Then TheExec.sites(site).SiteVariableValue(Looping_Index_Name) = EVS_Loop_Var_Max(site)
        End If
    Next site
    
    
ProfileMarkLeave VBT_LIB_Common_ProfileMark_2309
    
End If

Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
        
End Function


Function Start_Profile_EVS(PinName As PinList, WhatToCapture As String, SampleRate As Double, SampleSize As Long, Optional CapSignalName As String = "Capture_signal")
'start current or voltage profile capturing

On Error GoTo errHandler
' Wait if another capture is running
    Do While TheHdw.DCVS.pins(PinName).Capture.IsRunning = True
    Loop
    
    'Create a SIGNAL to set up instrument
    TheHdw.DCVS.pins(PinName).Capture.Signals.Add CapSignalName
    
    'Set this as the default signal
    TheHdw.DCVS.pins(PinName).Capture.Signals.DefaultSignal = CapSignalName
    
    'Define the signal used for the capture
    With TheHdw.DCVS.pins(PinName).Capture.Signals.item(CapSignalName)
        .Reinitialize
        If (WhatToCapture = "I") Then
            .mode = tlDCVSMeterCurrent
            .range = TheHdw.DCVS.pins(PinName).CurrentRange.Max '2
        Else
            .mode = tlDCVSMeterVoltage
            .range = 10
        End If
        .SampleRate = SampleRate
        .SampleSize = SampleSize
    
    End With
    
    ' Setup the hardware by loading the signal
    TheHdw.DCVS.pins(PinName).Capture.Signals.item(CapSignalName).LoadSettings
    
    ' Start the capture
    TheHdw.DCVS.pins(PinName).Capture.Signals.item(CapSignalName).Trigger

    Exit Function
errHandler:
    ErrorDescription ("Start_Profile")
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function Plot_Profile_EVS(PinName As PinList, Optional CapSignalName As String = "Capture_signal", Optional SampleRate As String, Optional instancename As String, Optional Find_IMax_Imin As Boolean, Optional Bolck_Name As String)
'Plot profiles

    Dim DSPW As New DSPWave
    Dim Label As String
    Dim site As Variant
    Dim Pin_Ary() As String
    Dim Pin_Cnt As Long
    Dim p As Variant
    Dim filename As String
    Dim CheckPath As String
    
    CheckPath = "D:\Profile"
    If Dir(CheckPath, vbDirectory) = "" Then
        MkDir (CheckPath)
    End If
                
    On Error GoTo errHandler

    Do While TheHdw.DCVS.pins(PinName).Capture.IsRunning = True
    Loop

    Dim lastBurstPat As New SiteVariant
    Dim isGrp As New SiteBoolean
    Dim lastLabel As New SiteVariant

    TheHdw.Digital.Patgen.ReadLastStart lastBurstPat, isGrp, lastLabel

    ' Get the captured samples from the instrument
    Call TheExec.DataManager.DecomposePinList(PinName, Pin_Ary(), Pin_Cnt)
    
    For Each p In Pin_Ary
        If TheExec.DataManager.ChannelType(p) <> "N/C" Then
            DSPW = TheHdw.DCVS.pins(p).Capture.Signals(CapSignalName).DSPWave
            For Each site In TheExec.sites
                If TheHdw.DCVS.pins(p).Meter.mode = tlDCVSMeterCurrent Then
                    Label = "Current Profile for Site: " & site & " " & p
                    filename = "CurrentProfile_Site" & site & "_Pin_" & p & ".txt"
                Else
                    Label = "Voltage Profile for Site: " & site & " " & p
                    filename = "VoltageProfile_Site" & site & "_Pin_" & p & ".txt"
                End If
                Dim Plot_Name As String
                
                
                Dim MaximumMagnitude As Double
                Dim MinimumMagnitude As Double
                If Find_IMax_Imin = True Then                  ' Add by CS 20170908 current profile Imax Imin
            
                   MaximumMagnitude = DSPW.CalcMaximumMagnitude
                   MinimumMagnitude = DSPW.CalcMinimumMagnitude
                   TheExec.Datalog.WriteComment "Site" & site & ":" & lastBurstPat(site) & ":" & CapSignalName & ":" & "currentMax" & ":" & MaximumMagnitude & ":" & "currentMin" & ":" & MinimumMagnitude
                End If
                
                ''' create a new folder
                If Dir(CheckPath & "\" & mid(filename, 1, 14), vbDirectory) = "" Then
                    MkDir (CheckPath & "\" & mid(filename, 1, 14))
                End If
                

                Plot_Name = CheckPath & "\" & mid(filename, 1, 14) & "\" & mid(filename, 1, 14) & "-Site" & site & "-" & p & "-" & SampleRate & "-" & Bolck_Name & ".txt"
                
                
                If True Then DSPW.FileExport Plot_Name, File_txt   'for pliot
                If False Then DSPW.Plot Label  'for file export
            Next site
        End If
    Next p
    
    Exit Function
errHandler:
    ErrorDescription ("Plot_Profile")
    If AbortTest Then Exit Function Else Resume Next
End Function

