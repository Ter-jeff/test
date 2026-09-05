Attribute VB_Name = "VBT_LIB_RTOS"

Option Explicit

Private m_TimeSetSheet As String
Private m_LevelsSheet As String
Private m_InterposeFunctionsSet As Boolean
Public gSB_RTOSBootPatResult As New SiteBoolean
Public gSB_RTOSBistPatResult As New SiteBoolean
Public gDSPData_UART As New PinListData
Public G_cmd1 As String
Public G_cmd2 As String
Public G_cmd3 As String
Public GlobalMergeAry() As String
Public TNTEMP As String
Public PinTemp As String
Public force_val As Double

Public g_RTOS_FirstSetp As Boolean
Public g_LastRTOSPoint As Boolean
Public g_RTOSNwireChar As Boolean
Public g_RTOS2DFirstPoint As Boolean
Public g_RTOSRampStep As Integer
Public g_RTOS_SceVoltage As New PinListData
Public g_MasterCMDOnly As Boolean
Public g_HardwareSiteCount As Long
Public g_DeviceMergedCount As Long
Public g_Reboot_Flag As Boolean
Public RTOSTest_Inst As String
Public BootUpInstanceName As String

Public Function RTOS_Capture_Data(InWave As DSPWave, CapAry() As String, CapNumPerSegAry() As Long)
On Error GoTo errHandler

    Dim site As Variant
    Dim j As Long
    Dim CapIdx As Long
    Dim CaptureData As String
    Dim SegCapData As String
    Dim StartBit As Long

    For Each site In TheExec.sites
        CaptureData = vbNullString
        For j = 0 To InWave.SampleSize - 1
            CaptureData = CaptureData & CStr(InWave(site).Element(j))
        Next j
        StartBit = 1
        TheExec.Datalog.WriteComment "Total Capture Bits: " & InWave.SampleSize
        For CapIdx = 0 To UBound(CapAry)
            SegCapData = mid(CaptureData, StartBit, CapNumPerSegAry(CapIdx))
            TheExec.Datalog.WriteComment "Site: " & site & " " & CapAry(CapIdx) & ", Capture Data: " & SegCapData
            StartBit = StartBit + CapNumPerSegAry(CapIdx)
        Next CapIdx
        'TheExec.Datalog.WriteComment "Site: " & site & ", Capture Data: " & CaptureData
    Next site

    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_RTOS", "RTOS_Capture_Data")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function RTOS_Command(Optional Cmd1 As String, Optional Cmd2 As String, Optional Cmd3 As String, Optional Cmd4 As String, _
            Optional Cmd5 As String, Optional Cmd1TimeOut As Double = 0#, Optional Cmd2TimeOut As Double = 0#, Optional Cmd3TimeOut As Double = 0#, _
            Optional Cmd4TimeOut As Double = 0#, Optional Cmd5TimeOut As Double = 0#) As Long
On Error GoTo errHandler

    ReDim GlobalMergeAry(TheExec.sites.Existing.Count - 1) 'for txt data collection
      
    Dim CmdList As Variant 'String
    Dim CmdListStatus As New SiteLong
    Dim powerPin As String
    Dim instancename As String: instancename = TheExec.DataManager.instancename
    Dim CMDTotalTT As Double

      
    If Cmd1 <> "" Then CmdList = Cmd1
    If Cmd2 <> "" Then CmdList = CmdList + Cmd2
    If Cmd3 <> "" Then CmdList = CmdList + Cmd3
    If Cmd4 <> "" Then CmdList = CmdList + Cmd4
    If Cmd5 <> "" Then CmdList = CmdList + Cmd5
    CMDTotalTT = Cmd1TimeOut + Cmd2TimeOut + Cmd3TimeOut + Cmd4TimeOut + Cmd5TimeOut

    'Scenario Run Conditions
    CmdListStatus = 0
    
    TheExec.Datalog.DatalogSuspended = False
    
    If CmdList <> "" Then Set CmdListStatus = SendCmd(CmdList, CMDTotalTT, False)

    TheExec.flow.TestLimit CmdListStatus, 1, 1       ', , , , , , TestName
    
    RTOS_UART_Print instancename, CmdListStatus
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_RTOS", "RTOS_Command")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function RTOS_eFuse_Read(FuseType As String, m_catename As String, Dict_Store_Code_Name As String, dspwavesize As Long, Optional Efuse_Read_Dec_Flag As Boolean = False, Optional Dict_Store_Dec_Name As String = vbNullString, _
                                Optional Calc_code As String = vbNullString) As Long
On Error GoTo errHandler

    ' Parameter : eFuse Block , eFuse Variable , data , Data Width
    ' Create dictionary , if exist then remove and re-create
    ' MUST :  if necessary , we can set limit if read out value = 0 then bin out .
    
    Dim site As Variant
    Dim Read_Code As New DSPWave
    Dim Read_Value As New DSPWave
    Dim Efuse_Value As New SiteLong
    Dim TempVal As Long
    Dim Efuse_Value_Chk As New SiteVariant
    Dim i As Long
        
    
    Read_Code.CreateConstant 0, dspwavesize

    If Efuse_Read_Dec_Flag = True Then
        Read_Value.CreateConstant 0, 1
    End If
    
    ''====20201230 add for efuse new code====
    Efuse_Value = GetEfuseHipValue(FuseType, m_catename)

    For Each site In TheExec.sites

    ''====20201230 remove for efuse old code====
        'Efuse_Value(site) = auto_eFuse_GetReadDecimal(FuseType, m_catename, True)
'''''        Efuse_Value(Site) = CLng(Site) + 8
'''----------cal get fused code
        If Calc_code <> "" Then
        'Calc_code = "minus,100"
            If Split(Calc_code, ",")(0) = "minus" Then
                Efuse_Value = Efuse_Value.Subtract(Split(Calc_code, ",")(1))
            End If
        End If
'''----------cal get fused code
        If Efuse_Read_Dec_Flag = True Then
            Read_Value.Element(0) = Efuse_Value(site)
        End If

        TempVal = Efuse_Value(site)
        For i = 0 To dspwavesize - 1
            Read_Code.Element(i) = TempVal Mod 2
            TempVal = TempVal \ 2
        Next i
        'If Read out value = 0 then bin out
        If Efuse_Value(site) = 0 Then
            Efuse_Value_Chk(site) = 0
        Else
            Efuse_Value_Chk(site) = 1
        End If
        
    Next site
        
    TheExec.flow.TestLimit resultVal:=Efuse_Value_Chk, lowVal:=1, hiVal:=1, Tname:="NonZero_Val_Chk", ForceResults:=tlForceNone
        
    Call AddStoredCaptureData(Dict_Store_Code_Name, Read_Code)
    
    If Efuse_Read_Dec_Flag = True Then
        Call AddStoredCaptureData(Dict_Store_Dec_Name, Read_Value)
    End If
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_RTOS", "RTOS_eFuse_Read")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function RTOS_eFuse_Write(FuseType As String, m_catename As String, Dict_Store_Code_Name As String, Flag_Name As String, Optional Efuse_Binary_Write_Flag As Boolean = False, _
                                Optional Calc_code As String) As Long
On Error GoTo errHandler

    ' Parameter : eFuse Block , eFuse Variable , data
    Dim site As Variant
    Dim RTOS_eFuseData_Dict As New SiteVariant
    Dim Data_Temp As String
    Dim m_value As New SiteVariant
    Dim j As Integer
    Dim Pass_Fail_Flag As New SiteBoolean
    
    If m_catename = "mtr_fused_t2" Then
        For Each site In TheExec.sites
            m_value = 3
        Next site
    Else
        m_value = GetStoredData(Dict_Store_Code_Name)
    End If

    For Each site In TheExec.sites
        If TheExec.flow.SiteFlag(site, Flag_Name) = 1 Then
            Pass_Fail_Flag(site) = False
        ElseIf TheExec.flow.SiteFlag(site, Flag_Name) = 0 Then
            Pass_Fail_Flag(site) = True
        Else
            Pass_Fail_Flag(site) = False
            TheExec.Datalog.WriteComment ("Error! " & Flag_Name & "(" & site & ")" & " status is Clear !")
        End If

    Next site
    
    Dim opbank As eFuseBdfBank
    Dim field As eFuseBdfField
    Set opbank = GetBdfBank(FuseType)
    Set field = opbank.Fields(m_catename)
    opbank.SetEfuse field.name, m_value, Pass_Fail_Flag, , , , True
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_RTOS", "RTOS_eFuse_Write")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function RTOS_IDS(Cmd As String, Cmdwait As Double) As SiteLong
On Error GoTo errHandler

    
    ReDim GlobalMergeAry(TheExec.sites.Existing.Count - 1) ''for txt data collection
    
    Dim dspData As New PinListData
    Dim LowPins As String
    Dim HighPins As String
    
    Dim LowToHigh As String
    Dim HighToLow As String

    Dim BootResult As New DSPWave
    Dim i As Long, p As Long
    Dim TResult As New SiteLong
    Dim RTOS_IDS_inst As String
    Dim instancename As String
    RTOS_IDS_inst = TheExec.DataManager.instancename
    If Cmdwait < 0.0001 Then    '#16_ELSE_CASE_CHK
        Cmdwait = 0.1
    Else
    End If
    
    TResult = SendCmd(Cmd, Cmdwait)
    
    TheExec.flow.TestLimit TResult, 1, 1, , , , , , "RTOS_IDS_SC94_Cmd"
    
    RTOS_UART_Print instancename, TResult
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_RTOS", "RTOS_IDS")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function RTOS_Boot(BootUsingPattern As Boolean, Optional BootPattern As Pattern, Optional UseJTAG As Boolean = True, Optional shmooing As Boolean, Optional ramp As Boolean = False, _
                        Optional LowPins As String = "RTOS_Boot_Low", Optional HighPins As String = "RTOS_Boot_High", Optional LowToHigh As String = "RTOS_Boot_L2H", Optional HighToLow As String = vbNullString, _
                        Optional BootRelay As String, Optional UART_Log_Capture As Boolean = True, Optional MeasClock As Boolean = False, Optional ClockRelay As String, Optional ClockPin As String = "SPI1_SCLK", _
                        Optional digcap_flag As Boolean = False, Optional CaptureInfo As String = "sgmt0_32", Optional IO_State_Check As Boolean = False, Optional IO_Check_Pins As String = vbNullString, _
                        Optional Enable_JTAG_Relay As Boolean, Optional Realy_On_JTAG As PinList, Optional Relay_Off_JTAG As PinList, _
                        Optional OutputPin As String, Optional OutputPin_Vt As Double, Optional OutputPin_iLoad As Double, Optional Boot_CaptureTime As Double, Optional FRC_Pins As String = vbNullString) As Long

On Error GoTo errHandler


    TheExec.Datalog.Setup.Shared.ascii.Columns.EnableCustomWidths = True
    TheExec.Datalog.Setup.Shared.ascii.Columns.Parametric.TestName.Width = 70
    TheExec.Datalog.ApplySetup
    If Not (shmooing) Then
        ReDim GlobalMergeAry(TheExec.sites.Existing.Count - 1)
    End If
    
    Dim site As Variant
    Dim BootRelayAry() As String
    Dim Relay_On_Boot As New PinList
    Dim Relay_Off_Boot As New PinList
    Dim ClockRelayAry() As String
    Dim Relay_On_Colck As New PinList
    Dim Relay_Off_Clock As New PinList
    Dim dspData As New PinListData
    Dim BootDSP As New DSPWave
    Dim i As Long, j As Long
    Dim TResult As New SiteLong
    Dim instancename As String
    Dim DigCapPin As String
    Dim DFU_DSPWave As New DSPWave
    Dim DSSCCapSize As Long
    Dim CaptureData As String
    Dim IO_Check_PinsAry() As String
    
    Dim MasterSite As New SiteBoolean
    Dim SlaveSite As New SiteBoolean
    Dim RestoreSite As New SiteBoolean
    Dim MasterCMDOnly As Boolean
    Dim reftime As Double
    Dim Elapsetime As Double
    MasterCMDOnly = g_MasterCMDOnly
    RestoreSite = TheExec.sites.Active
        If BootUpInstanceName = "" Then BootUpInstanceName = TheExec.DataManager.instancename

    If CurrentChannelMap = LCase("ChannelMap_FT_4_site_2C") Then    '#16_ELSE_CASE_CHK
        g_HardwareSiteCount = 2
    ElseIf CurrentChannelMap = LCase("ChannelMap_WLFT_2_site") Then
        g_HardwareSiteCount = 1
    Else
    End If
    
    'SPIROM Gating
    If Not write_spirom.Any(False) = True Then  '#16_ELSE_CASE_CHK
        TheExec.AddOutput "<Warnning> Not  test SPI-ROM !! Please turn on the Flow_Table_Write_SPIROM_main! "
        TheExec.Datalog.WriteComment "<Warnning>  Not test SPI-ROM !!  Please turn on the Flow_Table_Write_SPIROM_main! "
        TheExec.Datalog.WriteComment " "
        For Each site In TheExec.sites
            TheExec.sites(site).FlagState("F_Others_Error") = logicTrue
        Next site
        Exit Function
    Else
    End If
    
    For Each site In TheExec.sites
        If site < g_HardwareSiteCount Then
            MasterSite(site) = RestoreSite(site)
        Else
            SlaveSite(site) = RestoreSite(site)
        End If
    Next site
    instancename = TheExec.DataManager.instancename
    
    If InStr(BootRelay, "|") <> 0 Then
        BootRelayAry = Split(BootRelay, "|")
        Relay_On_Boot.value = BootRelayAry(0)
        Relay_Off_Boot.value = BootRelayAry(1)
    Else
        TheExec.Datalog.WriteComment "Please Check the Boot Relay Setting !!!"
    End If
    
    TheHdw.Digital.pins("All_Digital").initState = chInitLo
    TheExec.Datalog.WriteComment "Set All_digital pin InitLo"

    'Process DigCapInfo
    Dim CapAry() As String
    Dim tmpInfoAry() As String
    Dim CapSegAry() As String
    Dim CapNumPerSegAry() As Long
    Dim TotalCapNum As Long
    Dim CapIdx As Long
    Dim SegDataAry() As String
    Dim CapCnt As Long
    
    'Process DigCapInfo
    If CaptureInfo <> "" Then   '#16_ELSE_CASE_CHK
        TotalCapNum = 0
        CapAry = Split(CaptureInfo, "+")
        ReDim CapSegAry(UBound(CapAry))
        ReDim CapNumPerSegAry(UBound(CapAry))
        ReDim SegDataAry(UBound(CapAry))
        For CapIdx = 0 To UBound(CapAry)
            tmpInfoAry = Split(CapAry(CapIdx), "_")
            CapSegAry(CapIdx) = tmpInfoAry(0)
            CapNumPerSegAry(CapIdx) = CLng(tmpInfoAry(1))
            TotalCapNum = TotalCapNum + CLng(tmpInfoAry(1))
        Next CapIdx
    Else
    End If
    
    If ramp Then
       RTOS_Voltage_RampUp
    Else
        'Print shmoo voltage information''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If TheExec.DevChar.Setups.IsRunning = True Then
            Dim active_setup As String, curr_axis As Variant, curr_track As Variant, apply_Pin As String, apply_Pin_arry() As String, pin_count As Long
            active_setup = TheExec.DevChar.Setups.ActiveSetupName
            TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
            For Each curr_axis In TheExec.DevChar.Setups(active_setup).Shmoo.axes.list
                ''exit for if any axis is not power pin -by SY
                If TheExec.DevChar.Setups(active_setup).Shmoo.axes(curr_axis).ApplyTo.pins <> "" Then
                    apply_Pin = TheExec.DevChar.Setups(active_setup).Shmoo.axes(curr_axis).ApplyTo.pins
            '        Add for store shmoo global spec to avoid direct to apply Vmain used for Vbump function
                    Call TheExec.DataManager.DecomposePinList(apply_Pin, apply_Pin_arry, pin_count)
                    For i = 0 To pin_count - 1
                                            If gl_GetInstrument_Dic.Exists(LCase(apply_Pin_arry(i))) Then _
                        TheExec.Datalog.WriteComment "Boot voltage, " & apply_Pin_arry(i) & "                       " & TheHdw.DCVS.pins(apply_Pin_arry(i)).Voltage.value & "V"
                    Next i
                    For Each curr_track In TheExec.DevChar.Setups(active_setup).Shmoo.axes(curr_axis).TrackingParameters.list
                        apply_Pin = TheExec.DevChar.Setups(active_setup).Shmoo.axes(curr_axis).TrackingParameters.item(curr_track).ApplyTo.pins
                        Call TheExec.DataManager.DecomposePinList(apply_Pin, apply_Pin_arry, pin_count)
                        For i = 0 To pin_count - 1
                                                        If gl_GetInstrument_Dic.Exists(LCase(apply_Pin_arry(i))) Then _
                            TheExec.Datalog.WriteComment "Boot voltage, " & apply_Pin_arry(i) & "                       " & TheHdw.DCVS.pins(apply_Pin_arry(i)).Voltage.value & "V"
                        Next i
                    Next curr_track
                End If
            Next curr_axis
        Else
            TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End If

    Relay_Control Relay_On_Boot, Relay_Off_Boot
    ' Re-cycle SPI-ROM power
    TheHdw.DCVS.pins("QSPI_PWR_1P2").Connect
    TheHdw.DCVS.pins("QSPI_PWR_1P2").Voltage.Main.value = 0#
    TheHdw.DCVS.pins("QSPI_PWR_1P2").Gate = True '' add 240607
    TheHdw.Wait 0.001
    TheHdw.DCVS.pins("QSPI_PWR_1P2").Voltage.Main.value = 1.2
    TheHdw.Wait 0.001
    
   

    '''//Follow relay switch //'''
    'TheHdw.Utility.Pins(Relay_Device).State = tlUtilBitOn
    'TheHdw.Wait 0.05
    
    TheHdw.Digital.pins("UART_TX").Levels.value(chVt) = 0
    TheHdw.Digital.pins("UART_TX").Levels.value(chIoh) = -0.0000005
    TheHdw.Digital.pins("UART_TX").Levels.value(chIol) = 0.0000005
    
    If OutputPin <> "" Then '#16_ELSE_CASE_CHK
        TheHdw.Digital.pins(OutputPin).Levels.value(chVt) = OutputPin_Vt
        TheHdw.Digital.pins(OutputPin).Levels.value(chIoh) = 0 - OutputPin_iLoad
        TheHdw.Digital.pins(OutputPin).Levels.value(chIol) = 0 + OutputPin_iLoad
    Else
    End If
    
    
    ''' Tahiti: These 2 Pins have pull up ensure the capability of Boot by Allen.
    
    
    TheHdw.Digital.pins("QSPI0_DQ2").initState = chInitHi
    TheHdw.Digital.pins("QSPI0_DQ3").initState = chInitHi
    
    
    With TheHdw.Protocol.ports("UART_TX")
        .TimeOut.Enabled = True
        .TimeOut.value = 25
        .Enabled = True
        .NWire.MaxHoldUntilTimeout.value = 0.003     '//3msec*1500=4.5sec
        .NWire.CMEM.MoveMode = tlNWireCMEMMoveMode_Databus
    End With
    
    Set dspData = TheHdw.Protocol.ports("UART_TX").NWire.CMEM.DSPWave
    
    'For SPIROM Clock Pin Measurement Relay Setting
    If MeasClock = True Then    '#16_ELSE_CASE_CHK
        Dim tmptimer As Double
        If InStr(ClockRelay, "|") <> 0 Then '#16_ELSE_CASE_CHK
            ClockRelayAry = Split(ClockRelay, "|")
            Relay_On_Colck.value = ClockRelayAry(0)
            Relay_Off_Clock.value = ClockRelayAry(1)
        Else
        End If
        Relay_Control Relay_On_Colck, Relay_Off_Clock
        'Define the Reference Timer
        tmptimer = TheExec.Timer(0)
    Else
    End If
    
    If UART_Log_Capture = True Then '#16_ELSE_CASE_CHK
        If MasterCMDOnly = True Then
            TheExec.sites.Selected = MasterSite
            TheHdw.Protocol.ports("UART_TX").Modules("UART_boot").start
            TheExec.sites.Selected = RestoreSite
        Else
            reftime = TheExec.Timer
            TheHdw.Protocol.ports("UART_TX").Modules("UART_boot").start
        End If
    Else
    End If
    
    If BootUsingPattern Then
       Dim PattAry() As String
       Dim PattCnt As Long
       PATT_GetPatListFromPatternSet BootPattern.value, PattAry, PattCnt
       
      'TheHdw.Digital.Pins("DFU_STATUS").DisableCompare = True
        For i = 0 To PattCnt - 1
            TheHdw.patterns(PattAry(i)).Load
            If UseJTAG Then
            ' Add DigCap function
                If digcap_flag = True Then
                    DFU_DSPWave.Clear
                    DigCapPin = "JTAG_TDO"
                    DSSCCapSize = TotalCapNum
                    
                    RTOS_DigCap_Setting PattAry(i), DigCapPin, "DSSC_Cap_All", DSSCCapSize
        
                    TheHdw.Wait 0.05
                    TheHdw.patterns(PattAry(i)).start
                    TheHdw.Digital.Patgen.HaltWait
                             
                    DFU_DSPWave = TheHdw.DSSC.pins(DigCapPin).Pattern(PattAry(i)).Capture.Signals("DSSC_Cap_All").DSPWave
                    
                    RTOS_Capture_Data DFU_DSPWave, CapAry, CapNumPerSegAry
                    
                    HardIP_WriteFuncResult  ' , , , , PattAry(i)
                Else
                    TheHdw.patterns(PattAry(i)).test pfAlways, 0
                    'TheHdw.Patterns(PattAry(i)).start
                        TheHdw.Digital.Patgen.HaltWait
                    'HardIP_WriteFuncResult
                End If
                DebugPrintFunc "RTOS_BOOT"
                    
                If MeasClock = True Then    '#16_ELSE_CASE_CHK
                    Call RTOS_Freq_Measurement(ClockPin, tmptimer, 5)
                Else
                End If
                
                If Enable_JTAG_Relay = True Then    '#16_ELSE_CASE_CHK
                    Relay_Control Realy_On_JTAG, Relay_Off_JTAG, 0.003
                Else
                End If
            Else
                TheHdw.Wait 0.005
                'thehdw.Digital.Pins("SPI0_MISO").initState = chInitHi ''001 -> 101
                'thehdw.Digital.Pins("SPI0_MOSI").initState = chInitLo ''001 -> 101
                'thehdw.Digital.Pins("SPI0_SCLK").initState = chInitHi ''001 -> 101
                TheHdw.Digital.Patgen.Continue 0, cpuA
                TheHdw.patterns(PattAry(i)).start
                'TheHdw.Digital.Patgen.FlagWait cpuA, 0
                TheHdw.Digital.Patgen.Continue 0, cpuA
                TheHdw.Digital.Patgen.HaltWait
                    
                If Enable_JTAG_Relay = True Then
                    Relay_Control Realy_On_JTAG, Relay_Off_JTAG, 0.003
                End If
            End If
        Next i
    Else
        ' ===== defined in RTOS test plan, to avoid long boot time =====
        TheHdw.Digital.pins("KIS_TO_PMU_REQUEST_DFU").initState = chInitHi
        TheHdw.Digital.pins("KIS_TO_PMU_RESET").initState = chInitHi
        '===============================================================
        TheHdw.Wait 0.01
        TheHdw.Digital.pins(LowPins).initState = chInitLo 'cold H
        TheHdw.Wait 0.01
        TheHdw.Digital.pins(HighPins).initState = chInitHi 'cold H
        TheHdw.Wait 0.01
     ''   thehdw.Digital.Pins("HOLD_RESET").initState = chInitHi 'cold H
        TheHdw.Digital.pins("QSPI0_DQ2").initState = chInitHi ''001 -> 101
        TheHdw.Digital.pins("QSPI0_DQ3").initState = chInitHi ''001 -> 101
        
        TheHdw.Digital.pins(LowToHigh).initState = chInitLo 'cold L
        TheHdw.Wait 0.01
        TheHdw.Digital.pins(LowToHigh).initState = chInitHi 'cold H
        TheHdw.Wait 0.01

        
        ''''''''''''''''''''''''''Boot Config'''''''''''''''''''''''''''''''''
        'QSPI->40MHz for Tahiti
        'TheHdw.Digital.Pins("SPI0_MISO").InitState = chInitHi ''001 -> 111
        'TheHdw.Digital.Pins("SPI0_MOSI").InitState = chInitHi ''001 -> 111
        'TheHdw.Digital.Pins("SPI0_SCLK").InitState = chInitLo ''001 -> 111
        ''''''''''''''''''''''''''Boot Config'''''''''''''''''''''''''''''''''
        
         ''''''''''''''''''''''''''Boot Config'''''''''''''''''''''''''''''''''
        'QSPI->24MHz for Tahiti
        TheHdw.Digital.pins("QSPI0_DQ1").initState = chInitLo ''001 -> 111
        TheHdw.Digital.pins("QSPI0_DQ0").initState = chInitHi ''001 -> 111
        TheHdw.Digital.pins("QSPI0_CK").initState = chInitLo ''001 -> 111
        ''''''''''''''''''''''''''Boot Config'''''''''''''''''''''''''''''''''
        
        
        
        If IO_State_Check = True And IO_Check_Pins <> "" Then   '#16_ELSE_CASE_CHK
            RTOS_IOState_Checker IO_Check_Pins
        Else
        End If
        'Frequency Measurment Within Monitior Time
        If MeasClock = True Then    '#16_ELSE_CASE_CHK
            Call RTOS_Freq_Measurement(ClockPin, tmptimer, 5)
        Else
        End If
    End If
    
    If UART_Log_Capture = True Then
'        TheHdw.Protocol.ports("UART_TX").IdleWait
        'Need to adjust wait time by project
        If Boot_CaptureTime = 0 Then    '#16_ELSE_CASE_CHK
            Boot_CaptureTime = 3
        Else
        End If
      TheExec.flow.Wait Boot_CaptureTime
    ''    TheExec.Flow.Wait 10
        TheHdw.Protocol.ports("UART_TX").Enabled = False
        Elapsetime = TheExec.Timer(reftime)
        
    Else
        Exit Function
    End If
    'Add Offline Simulated Data
    If TheExec.TesterMode = testModeOffline Then
        Dim CompareArrayOffline(649) As Long
        Dim dspdata_offline As New DSPWave
        For i = 0 To 649 Step 4
            If i < 650 Then CompareArrayOffline(i) = 65         'A
            If i + 1 < 650 Then CompareArrayOffline(i + 1) = 84 'T
            If i + 2 < 650 Then CompareArrayOffline(i + 2) = 69 'E
            If i + 3 < 650 Then CompareArrayOffline(i + 3) = 62 '>
        Next i
        dspdata_offline.data = CompareArrayOffline
        rundsp.CheckBootStatus dspdata_offline, TResult
        
        For Each site In TheExec.sites
            BootDSP(site) = dspdata_offline(site).COPY
        Next site
    Else
        If MasterCMDOnly = True Then
            TheExec.sites.Selected = MasterSite
            rundsp.CheckBootStatus dspData, TResult 'Check DSP wave status to determine TResult
            BootDSP = dspData.COPY
            Call LogDUTResponse(BootDSP, TResult)    'Copy DSP wave into an output log
            'TResult_coreup = SendCmd("Core up acc; pmgr mode", 0.2)
            TheExec.sites.Selected = RestoreSite
        Else
            rundsp.CheckBootStatus dspData, TResult 'Check DSP wave status to determine TResult
            BootDSP = dspData.COPY
            Call LogDUTResponse(BootDSP, TResult)    'Copy DSP wave into an output log
            TheExec.Datalog.WriteComment "RTOS Boot time : " + str(Elapsetime)
            TheExec.Datalog.WriteComment "RTOS Boot Wait time : " + str(Boot_CaptureTime)
        End If
    End If
    If MeasClock = True And FRC_Pins <> "" Then
        RTOS_Meas_FRC FRC_Pins
    End If
   
    TheHdw.Protocol.ports("UART_TX").TimeOut.value = 30
    
    If MasterCMDOnly = True Then    '#16_ELSE_CASE_CHK
        For Each site In TheExec.sites
            If site < g_HardwareSiteCount Then  '#16_ELSE_CASE_CHK
                TResult(site + g_HardwareSiteCount) = TResult(site)
            Else
                'TResult_coreup(site + g_HardwareSiteCount) = TResult_coreup(site)
            End If
        Next site
    Else
    End If
    
    If TheExec.DevChar.Setups.IsRunning = False Then    '#16_ELSE_CASE_CHK
        If Not (shmooing) Then  '#16_ELSE_CASE_CHK
            RTOS_UART_Print instancename, TResult
        Else
        End If
        If Not (shmooing) Then  '#16_ELSE_CASE_CHK
            TheExec.flow.TestLimit TResult, 1, 1, , , , , , "Boot Status"
        Else
        End If
        'If MasterCMDOnly = True Then
            'If Not (shmooing) Then theexec.Flow.TestLimit TResult_coreup, 1, 1, , , , , , "Core up acc"
        'End If
    Else
    End If
    
    DebugPrintFunc "RTOS_BOOT"
    
    'SendCmd "core up acc;fs mount SPI1;", 0.05, True
    If Enable_JTAG_Relay = True Then    '#16_ELSE_CASE_CHK
        Relay_Control Realy_On_JTAG, Relay_Off_JTAG, 0.003
    Else
    End If
    
        '20221122 Sbin950 avoid, return to 0
    TheHdw.Digital.pins("UART_TX").Levels.value(chIoh) = 0
    TheHdw.Digital.pins("UART_TX").Levels.value(chIol) = 0
        
        g_Reboot_Flag = False
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_RTOS", "RTOS_Boot")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function RTOS_Boot_CZ(argc As Long, argv() As String) As Long
On Error GoTo errHandler


    Dim dspData As New PinListData
    Dim LowPins As String
    Dim HighPins As String
    
    Dim LowToHigh As String
    Dim HighToLow As String

    Dim BootResult As New DSPWave
    Dim i As Long, p As Long
    Dim TResult As New SiteLong
    
    g_RTOS_FirstSetp = True
    g_LastRTOSPoint = False
    g_RTOS2DFirstPoint = True
    RTOS_Boot False, , , True
         
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_RTOS", "RTOS_Boot_CZ")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function RTOS_IOState_Checker(IO_Check_Pins As String)
On Error GoTo errHandler

    Dim site As Variant
    Dim i As Long
    Dim RTOS_pinInitState As ChInitState
    Dim IO_Check_PinsAry() As String
    
    IO_Check_PinsAry = Split(IO_Check_Pins, ",")
    
    For Each site In TheExec.sites
            For i = 0 To UBound(IO_Check_PinsAry)
                RTOS_pinInitState = TheHdw.Digital.pins(IO_Check_PinsAry(i)).initState
                If RTOS_pinInitState = "0" Then
                    TheExec.Datalog.WriteComment "site" & site & " : " & IO_Check_PinsAry(i) & " = " & RTOS_pinInitState & "   (State = Hi)"
                ElseIf RTOS_pinInitState = "1" Then
                    TheExec.Datalog.WriteComment "site" & site & " : " & IO_Check_PinsAry(i) & " = " & RTOS_pinInitState & "   (State = Lo)"
                Else
                    TheExec.Datalog.WriteComment IO_Check_PinsAry(i) & " state have issue!!!"
                End If
            Next
    Next site
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_RTOS", "RTOS_IOState_Checker")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function RTOS_Meas_FRC(MeasPins As String)
On Error GoTo errHandler
    Dim TestInstName As String
    Dim FRC_PLD As New PinListData
    Dim CounterValue As New PinListData
    Dim FRC_Ary() As String
    Dim FreqInterval As Double
    Dim i As Long
    'MeasPins = "XI0_Diff"
    FreqInterval = 0.01
    
    TestInstName = "FreeRunClock"
    FRC_Ary = Split(MeasPins, ",")
    For i = 0 To UBound(FRC_Ary)
        If TheExec.DataManager.pinType(FRC_Ary(i)) Like "Differential" Then
            TheHdw.Digital.pins(FRC_Ary(i)).DifferentialLevels.value(chVod) = XI0_ref_VOH / 2
        Else
            TheHdw.Digital.pins(FRC_Ary(i)).Levels.value(chVoh) = XI0_ref_VOH / 2
        End If

        With TheHdw.Digital.pins(FRC_Ary(i)).FreqCtr
            .EventSource = VOH
            .EventSlope = Positive
            .Interval = 0.01
            .Enable = IntervalEnable
            .Clear
            TheHdw.Wait FreqInterval
            .start
            CounterValue = .Read()
        End With
        FRC_PLD = CounterValue.Math.divide(FreqInterval)
        TheExec.flow.TestLimit FRC_PLD, , , , , , unitHz, , TestInstName
    Next i
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_RTOS", "RTOS_Meas_FRC")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function RTOS_RunMetrology(SensorName As String, VddSenseFreq As String, VddSenseHeat As String, VddSensePreheat As String, _
                                      CmdROT As String, CmdROV As String, CmdTMP As String, CmdFinal As String, CmdROT_Timeout As Double, CmdROV_Timeout As Double, _
                                      CmdTMP_Timeout As Double, CmdFinal_Timeout As Double, Optional SELSRAM_DSSC As String)
On Error GoTo errHandler

    Dim CmdList As String
    Dim CmdListStatus As New SiteLong
     
    Dim powerPin As String
    Dim SupplyVoltage As Long
    Dim LogTimes As Boolean
    Dim RTOS_SELSRM_STR As String
    Dim site As Variant
    Dim MasterCMDOnly As Boolean
    Dim MasterSite As New SiteBoolean
    Dim SlaveSite As New SiteBoolean
    Dim RestoreSite As New SiteBoolean
    MasterCMDOnly = g_MasterCMDOnly
    TheHdw.PinLevels.ApplyPower
    RestoreSite = TheExec.sites.Active
    For Each site In TheExec.sites
        If site < g_HardwareSiteCount Then
            MasterSite(site) = RestoreSite(site)
        Else
            SlaveSite(site) = RestoreSite(site)
        End If
    Next site
    
    Shmoo_Save_core_power_per_site_for_Vbump
    TheHdw.DCVS.pins("All_Power").Voltage.Output = tlDCVSVoltageMain
    
    If MasterCMDOnly = True Then    '#16_ELSE_CASE_CHK
        TheExec.sites.Selected = MasterSite
    Else
    End If

    ''''////===For Ts5/Ta0/Ta1====////'''
    'If UCase(SensorName) Like UCase("*ta*") Then
    '    SendCmd "pmgr domain enable ane0", 0.1
    'ElseIf UCase(SensorName) Like UCase("*ts*") Then
    '    SendCmd "pmgr power-up isp", 0.1
    'End If
    ''''////===For Ts5/Ta0/Ta1====////'''
    If MasterCMDOnly = True Then    '#16_ELSE_CASE_CHK
        TheExec.sites.Selected = RestoreSite
    Else
    End If
    
    'Select Sram Start
    TheHdw.PinLevels.ApplyPower
    CmdListStatus = 0

    Shmoo_Save_core_power_per_site_for_Vbump
    For Each site In TheExec.sites.Active
        RTOS_SELSRM_STR = Decide_Switching_Bit_RTOS(SELSRAM_DSSC, g_ApplyLevelTimingValt, "RTOS")
        Exit For
    Next site
    '''''''''''''''''''''''''''''''''''''
    'Metrology Selsrm
'    If g_MasterCMDOnly = True Then
'        theexec.sites.Selected = MasterSite
'    End If
'    SendCmd RTOS_SELSRM_STR, 0.1
'    TheHdw.DCVS.Pins("CorePower").Voltage.output = tlDCVSVoltageAlt ' only change the corepower to alt
'    TheHdw.Wait 0.005
'    If MasterCMDOnly = True Then
'        theexec.sites.Selected = RestoreSite
'    End If
    '''''''''''''''''''''''''''''''''''''
    
    Dim OffsetRecord As New SiteLong
    'Delete start, 190628, Leonli
    Dim MTRString_TMPS() As String 'TMPS Per site
    Dim MTRString_ROT() As String 'ROT Per site
    Dim MTRString_ROV() As String 'ROV Per site

    ReDim MTRString_TMPS(TheExec.sites.Existing.Count - 1) 'Reset String Array
    ReDim MTRString_ROT(TheExec.sites.Existing.Count - 1) 'Reset String Array
    ReDim MTRString_ROV(TheExec.sites.Existing.Count - 1) 'Reset String Array

    Dim MTRTmpWave As New DSPWave
    Dim BeforeMTRTmpWave As New DSPWave
    Dim MTRDSPWave_ROT As New DSPWave
    Dim MTRDSPWave_ROV As New DSPWave
    Dim i As Integer
    Dim CmdCheck As String
    Dim SweepCondition_Split() As String: SweepCondition_Split = Split(VddSenseFreq, "+")
    
    For i = 0 To UBound(SweepCondition_Split)
        If i = 0 Then
            'MTRDSPWave_ROT = GetStoredCaptureData(SweepCondition_Split(i) & "-Freq-" & SensorName & "-sensor-ROT")
            MTRDSPWave_ROT = GetStoredCaptureData(SweepCondition_Split(i) & "-Freq-" & SensorName & "--ROT")
            MTRDSPWave_ROV = GetStoredCaptureData(SweepCondition_Split(i) & "-Freq-" & SensorName & "--ROV")
        Else
            For Each site In TheExec.sites
                MTRDSPWave_ROT = MTRDSPWave_ROT.Concatenate(GetStoredCaptureData(SweepCondition_Split(i) & "-Freq-" & SensorName & "--ROT"))
                MTRDSPWave_ROV = MTRDSPWave_ROV.Concatenate(GetStoredCaptureData(SweepCondition_Split(i) & "-Freq-" & SensorName & "--ROV"))
            Next site
        End If
    Next i
    For Each site In TheExec.sites
        MTRDSPWave_ROT = MTRDSPWave_ROT.divide(1000000000#)
        MTRDSPWave_ROV = MTRDSPWave_ROV.divide(1000000000#)
    Next site
    
    'For VddSense average
    Dim VddSenseHeatAry() As String
    Dim VddSensePreheatAry() As String
    Dim SenseCnt As Long
    
    VddSenseHeatAry = Split(VddSenseHeat, ",")
    VddSensePreheatAry = Split(VddSensePreheat, ",")
    If UBound(VddSenseHeatAry) = UBound(VddSensePreheatAry) Then    '#16_ELSE_CASE_CHK
        For SenseCnt = 0 To UBound(VddSenseHeatAry)
            If SenseCnt = 0 Then
                MTRTmpWave = GetStoredCaptureData(VddSenseHeatAry(SenseCnt))
                BeforeMTRTmpWave = GetStoredCaptureData(VddSensePreheatAry(SenseCnt))
            Else
                For Each site In TheExec.sites
                    MTRTmpWave = MTRTmpWave.Concatenate(GetStoredCaptureData(VddSenseHeatAry(SenseCnt))).COPY
                    BeforeMTRTmpWave = BeforeMTRTmpWave.Concatenate(GetStoredCaptureData(VddSensePreheatAry(SenseCnt))).COPY
                Next site
            End If
        Next
    Else
    End If
'    MTRTmpWave = GetStoredCaptureData(VddSenseHeat)
'    BeforeMTRTmpWave = GetStoredCaptureData(VddSensePreheat)

            Dim MasterCMD_ROT() As String: ReDim MasterCMD_ROT(TheExec.sites.Existing.Count - 1)
            Dim MasterCMD_ROV() As String: ReDim MasterCMD_ROV(TheExec.sites.Existing.Count - 1)
            Dim MasterCMD_TMPS() As String: ReDim MasterCMD_TMPS(TheExec.sites.Existing.Count - 1)
            Dim SlaveCMD_ROT() As String: ReDim SlaveCMD_ROT(TheExec.sites.Existing.Count - 1)
            Dim SlaveCMD_ROV() As String: ReDim SlaveCMD_ROV(TheExec.sites.Existing.Count - 1)
            Dim SlaveCMD_TMPS() As String: ReDim SlaveCMD_TMPS(TheExec.sites.Existing.Count - 1)

    For Each site In TheExec.sites.Active
        For i = 0 To UBound(SweepCondition_Split)
            MTRString_ROT(site) = MTRString_ROT(site) + CStr(FormatNumber(MTRDSPWave_ROT.Element(i), 8))
            MTRString_ROV(site) = MTRString_ROV(site) + CStr(FormatNumber(MTRDSPWave_ROV.Element(i), 8))
            MTRString_ROT(site) = MTRString_ROT(site) + " "
            MTRString_ROV(site) = MTRString_ROV(site) + " "
        Next i
        
        If MasterCMDOnly = True Then
            
            If site < g_HardwareSiteCount Then
                MasterCMD_ROT(site) = CmdROT + " " + SensorName + ".0" + " " + MTRString_ROT(site)
                MasterCMD_ROV(site) = CmdROV + " " + SensorName + ".0" + " " + MTRString_ROV(site)
                MasterCMD_TMPS(site) = CmdTMP + " " + SensorName + ".0" + " " + CStr(FormatNumber((BeforeMTRTmpWave.CalcMean / 8), 7)) + " " + CStr(FormatNumber((MTRTmpWave.CalcMean / 8), 7))
            Else
                SlaveCMD_ROT(site - g_HardwareSiteCount) = CmdROT + " " + SensorName + ".1" + " " + MTRString_ROT(site)
                SlaveCMD_ROV(site - g_HardwareSiteCount) = CmdROV + " " + SensorName + ".1" + " " + MTRString_ROV(site)
                SlaveCMD_TMPS(site - g_HardwareSiteCount) = CmdTMP + " " + SensorName + ".1" + " " + CStr(FormatNumber((BeforeMTRTmpWave.CalcMean / 8), 7)) + " " + CStr(FormatNumber((MTRTmpWave.CalcMean / 8), 7))
            End If
        Else
            MTRString_ROT(site) = CmdROT + " " + SensorName + " " + MTRString_ROT(site)
            MTRString_ROV(site) = CmdROV + " " + SensorName + " " + MTRString_ROV(site)
            MTRString_TMPS(site) = CmdTMP + " " + SensorName + " " + CStr(FormatNumber((BeforeMTRTmpWave.CalcMean / 8), 7)) + " " + CStr(FormatNumber((MTRTmpWave.CalcMean / 8), 7))
        End If
    Next site
        

    If MasterCMDOnly = True Then
        Dim Bool_MasterCMD As Boolean: Bool_MasterCMD = False
        TheExec.sites.Selected = MasterSite
        Bool_MasterCMD = True
        SendCmd MasterCMD_ROT, CmdROT_Timeout
        SendCmd MasterCMD_ROV, CmdROV_Timeout
        SendCmd MasterCMD_TMPS, CmdTMP_Timeout
        
        CmdListStatus = 0 'Reset Command Result Status
        CmdCheck = CmdFinal + " " + SensorName + ".0" + " -ts 10 40"
        Set CmdListStatus = SendCmd(CmdCheck, CmdFinal_Timeout, False)
        TheExec.flow.TestLimit CmdListStatus, 1, 1, , , , , , "MasterMTR"
        
        GoTo UART_MTR_Parser
        
SlaveCMDStart:
        ReDim GlobalMergeAry(TheExec.sites.Existing.Count - 1)
        Bool_MasterCMD = False
        
        SendCmd SlaveCMD_ROT, CmdROT_Timeout
        SendCmd SlaveCMD_ROV, CmdROV_Timeout
        SendCmd SlaveCMD_TMPS, CmdTMP_Timeout
        
        CmdListStatus = 0 'Reset Command Result Status
        CmdCheck = CmdFinal + " " + SensorName + ".1" + " -ts 10 40"
        Set CmdListStatus = SendCmd(CmdCheck, CmdFinal_Timeout, False)
        TheExec.flow.TestLimit CmdListStatus, 1, 1, , , , , , "SlaveMTR"
        
        GoTo UART_SlaveMTR_Parser
    Else
        SendCmd MTRString_ROT, CmdROT_Timeout
        SendCmd MTRString_ROV, CmdROV_Timeout
        SendCmd MTRString_TMPS, CmdTMP_Timeout     ' Modify, 190628, Leon Li
    
        CmdListStatus = 0 'Reset Command Result Status
        CmdCheck = CmdFinal + " " + SensorName + " -ts 10 40"                    ''''////remove for FT2 60C testing by 20190715 Leslie commend
        Set CmdListStatus = SendCmd(CmdCheck, CmdFinal_Timeout, False)
        TheExec.flow.TestLimit CmdListStatus, 1, 1 ', , , , , , TestName
    End If
    
    
UART_MTR_Parser:
    '============Metrology UART Parser================================
    TheExec.Datalog.WriteComment vbNullString
    TheExec.Datalog.WriteComment "****** RTOS MTR to EFUSE Hex2Decimal *******"
    
            Dim MTRFieldCount As Integer
            Dim rtos_mtr_fuse_name() As String     'Modify, Leon Li, 20190628
            Dim mm As Integer
            Dim rtos_mtr_fuse_value() As SiteLong     'Modify, Leon Li, 20190628
            Dim MTRFuseName As String

            MTRFieldCount = 15
            ReDim rtos_mtr_fuse_name(MTRFieldCount)

            Dim strLen As Long
            Dim fuse_name_len As Long
            Dim fuse_code_idx As Long
            Dim fuse_code_value_hex() As New SiteVariant
            Dim fuse_code_value() As New SiteDouble
            Dim fuse_start_lo As Long
            Dim fuse_name_in_cate() As String
            
            ReDim fuse_code_value_hex(MTRFieldCount)
            ReDim fuse_code_value(MTRFieldCount)
            ReDim fuse_name_in_cate(MTRFieldCount)
            
            For mm = 0 To MTRFieldCount
                rtos_mtr_fuse_name(mm) = LCase("mtr_" & SensorName & "_c" & Trim(str(mm)) & "=")
                If mm = MTRFieldCount Then rtos_mtr_fuse_name(mm) = LCase("mtr_" & SensorName & "_ss=")
            Next mm
UART_SlaveMTR_Parser:
            For Each site In TheExec.sites.Selected
              If CmdListStatus(site) = 1 Then
                    strLen = Len(GlobalMergeAry(site))
                    For mm = 0 To MTRFieldCount
                        If InStr(1, LCase(GlobalMergeAry(site)), rtos_mtr_fuse_name(mm)) <> 0 Then
                           
                            fuse_name_len = Len(rtos_mtr_fuse_name(mm))
                            fuse_code_idx = InStr(1, LCase(GlobalMergeAry(site)), rtos_mtr_fuse_name(mm))
                            fuse_start_lo = fuse_code_idx + fuse_name_len
                            fuse_name_in_cate(mm) = Replace(rtos_mtr_fuse_name(mm), "=", vbNullString)
                            If MasterCMDOnly = True Then
'                            fuse_code_value_hex(mm) = Mid(LCase(GlobalMergeAry(site)), fuse_start_lo, 10)
                                If Bool_MasterCMD = True Then
                                    fuse_code_value_hex(mm)(site) = mid(LCase(GlobalMergeAry(site)), fuse_start_lo, 10)

                                    fuse_code_value(mm)(site) = auto_HexStr2Value(fuse_code_value_hex(mm)(site))    ' hex to dec
                                    TheExec.Datalog.WriteComment vbTab & "Site(" + CStr(site) + ") " + " MTR Hex EFUSE from UART                       " + Replace(rtos_mtr_fuse_name(mm), "=", " = ") + fuse_code_value_hex(mm)

                                Else
                                    fuse_code_value_hex(mm)(site + g_HardwareSiteCount) = mid(LCase(GlobalMergeAry(site)), fuse_start_lo, 10)

                                    fuse_code_value(mm)(site + g_HardwareSiteCount) = auto_HexStr2Value(fuse_code_value_hex(mm)(site + g_HardwareSiteCount)) ' hex to dec
                                    TheExec.Datalog.WriteComment vbTab & "Site(" + CStr(site + g_HardwareSiteCount) + ") " + " MTR Hex EFUSE from UART                       " + Replace(rtos_mtr_fuse_name(mm), "=", " = ") + fuse_code_value_hex(mm)(site + g_HardwareSiteCount)

                                End If
                            Else
                                fuse_code_value_hex(mm) = mid(LCase(GlobalMergeAry(site)), fuse_start_lo, 10)
                                fuse_code_value(mm) = auto_HexStr2Value(fuse_code_value_hex(mm))     ' hex to dec
                                TheExec.Datalog.WriteComment vbTab & "Site(" + CStr(site) + ") " + " MTR Hex EFUSE from UART                       " + Replace(rtos_mtr_fuse_name(mm), "=", " = ") + fuse_code_value_hex(mm)

                            End If
'                            TheExec.Datalog.WriteComment vbTab & "Site(" + CStr(site) + ") " + " MTR Hex EFUSE from UART                       " + Replace(rtos_mtr_fuse_name(mm), "=", " = ") + fuse_code_value_hex(mm)
                        Else
                            If InStr(1, LCase(GlobalMergeAry(site)), rtos_mtr_fuse_name(MTRFieldCount)) <> 0 Then         '''//// search "_ss" ////'''
                                    fuse_name_len = Len(rtos_mtr_fuse_name(MTRFieldCount))
                                    fuse_code_idx = InStr(1, LCase(GlobalMergeAry(site)), rtos_mtr_fuse_name(MTRFieldCount))
                                    fuse_start_lo = fuse_code_idx + fuse_name_len
                                    fuse_name_in_cate(MTRFieldCount) = Replace(rtos_mtr_fuse_name(MTRFieldCount), "=", vbNullString)
                                If MasterCMDOnly = True Then
                                    If Bool_MasterCMD = True Then
                                        fuse_code_value_hex(MTRFieldCount)(site) = mid(LCase(GlobalMergeAry(site)), fuse_start_lo, 10)
                                        fuse_code_value(MTRFieldCount)(site) = auto_HexStr2Value(fuse_code_value_hex(MTRFieldCount)(site)) ' hex to dec
                                        TheExec.Datalog.WriteComment vbTab & "Site(" + CStr(site) + ") " + " MTR Hex EFUSE from UART                       " + Replace(rtos_mtr_fuse_name(MTRFieldCount), "=", " = ") + fuse_code_value_hex(MTRFieldCount)
                                    Else
                                        fuse_code_value_hex(MTRFieldCount)(site + g_HardwareSiteCount) = mid(LCase(GlobalMergeAry(site)), fuse_start_lo, 10)
                                        fuse_code_value(MTRFieldCount)(site + g_HardwareSiteCount) = auto_HexStr2Value(fuse_code_value_hex(MTRFieldCount)(site + g_HardwareSiteCount)) ' hex to dec
                                        TheExec.Datalog.WriteComment vbTab & "Site(" + CStr(site + g_HardwareSiteCount) + ") " + " MTR Hex EFUSE from UART                       " + Replace(rtos_mtr_fuse_name(MTRFieldCount), "=", " = ") + fuse_code_value_hex(MTRFieldCount)
                                    End If
                                Else
                                    fuse_code_value_hex(MTRFieldCount) = mid(LCase(GlobalMergeAry(site)), fuse_start_lo, 10)
                                    fuse_code_value(MTRFieldCount) = auto_HexStr2Value(fuse_code_value_hex(MTRFieldCount)) ' hex to dec
                                    TheExec.Datalog.WriteComment vbTab & "Site(" + CStr(site) + ") " + " MTR Hex EFUSE from UART                       " + Replace(rtos_mtr_fuse_name(MTRFieldCount), "=", " = ") + fuse_code_value_hex(MTRFieldCount)
                                End If
                            End If
                       
                            If InStr(1, UCase(GlobalMergeAry(site)), UCase("ERROR")) <> 0 Then
                                    fuse_name_in_cate(mm) = "mtr_" & SensorName & "_c" & Trim(str(mm))
                                    If mm = MTRFieldCount Then fuse_name_in_cate(mm) = "mtr_" & SensorName & "_ss"
                                    fuse_code_value(mm) = 0
                                    TheExec.sites(site).FlagState("F_Rtos_Metrology") = logicTrue
                                    GoTo NextForLoop:
                            End If
                            GoTo ExitForLoop:
                        End If
NextForLoop:
                    Next mm
               Else
                '''//// For RTOS MTR fuse, Once MTRSNS fail, that will assign value into "0" ////'''
                    For mm = 0 To MTRFieldCount
                        fuse_name_in_cate(mm) = "mtr_" & SensorName & "_c" & Trim(str(mm))
                        If mm = MTRFieldCount Then  '#16_ELSE_CASE_CHK
                            fuse_name_in_cate(mm) = "mtr_" & SensorName & "_ss"
                        Else
                        End If
                        If MasterCMDOnly = True Then
                            If Bool_MasterCMD = True Then
                                fuse_code_value(mm)(site) = 0
                            Else
                                fuse_code_value(mm)(site + g_HardwareSiteCount) = 0
                            End If
                                
                        Else
                            fuse_code_value(mm) = 0
                        End If
                    Next mm
                '''//// For RTOS MTR fuse, Once MTRSNS fail, that will assign value into "0" ////'''
               End If
ExitForLoop:
               TheExec.Datalog.WriteComment vbNullString
            Next site
            
    If MasterCMDOnly = True Then    '#16_ELSE_CASE_CHK
        If Bool_MasterCMD = True Then
            GoTo SlaveCMDStart
        Else
            TheExec.sites.Selected = RestoreSite
        End If
    Else
    End If
            For mm = 0 To MTRFieldCount
                Call AddStoredData(fuse_name_in_cate(mm), fuse_code_value(mm))
            Next mm
    '============Metrology UART Parser End===========================
    Dim instancename As String
    instancename = TheExec.DataManager.instancename
    
    RTOS_UART_Print instancename, CmdListStatus
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_RTOS", "RTOS_RunMetrology")
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function RTOS_Shmoo_Reboot(argc As Long, argv() As String) As Long
On Error GoTo errHandler

    Dim LowPins As String
    Dim HighPins As String
    Dim HighToLow As String
    Dim LowToHigh As String
    
    Dim MyR As Variant
    Dim TResult As New SiteLong
    Dim dspData As New PinListData
    
    Dim site As Variant
    Dim LastPointFailed As Boolean


    LastPointFailed = False

    For Each site In TheExec.sites.Selected
    
        '''======= For Bora Speceil case SC19 and SC25 cmd "fs mount SPI1" ======'''
        '''If theexec.DataManager.instanceName Like UCase("*SC19*") Or theexec.DataManager.instanceName Like UCase("*SC25*") Then LastPointFailed = True
        '''======= For Bora Speceil case SC19 and SC25 cmd "fs mount SPI1" ======'''

        If Not (LastPointFailed) Then
            If (TheExec.DevChar.results(argv(0)).Shmoo.CurrentPoint.ExecutionResult = tlDevCharResult_Fail) Then
                LastPointFailed = True
            End If
        End If
    Next site

    If UCase(TheExec.DataManager.instancename) Like "*S094*" Then
       If TheExec.enableWord("RTOSRamp") = True Then
           RTOS_Boot False, , , True, True
       Else
           RTOS_Boot False, , , True, False
       End If
    Else
        If (LastPointFailed) Then
    
            g_RTOS_FirstSetp = True
            If TheExec.enableWord("RTOSRamp") = True Then
               RTOS_Boot False, , , True, True
            Else
               RTOS_Boot False, , , True, False
            End If
    
        End If
    End If

    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_RTOS", "RTOS_Shmoo_Reboot")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function RTOS_RunScenario(Optional TestName As String, Optional Cmd1 As String, Optional Cmd2 As String, Optional Cmd3 As String, Optional Cmd4 As String, _
            Optional Cmd5 As String, Optional Cmd1TimeOut As Double = 0#, Optional Cmd2TimeOut As Double = 0#, Optional Cmd3TimeOut As Double = 0#, _
            Optional Cmd4TimeOut As Double = 0#, Optional Cmd5TimeOut As Double = 0#, Optional SELSRAM_DSSC As String, Optional Interpose_PrePat As String, _
            Optional pmode As String, Optional RampStep As Integer, Optional SetupCMD_Time As Double, Optional Vbump As Boolean = True) As Long
On Error GoTo errHandler
            
    
    ReDim GlobalMergeAry(TheExec.sites.Existing.Count - 1) 'for txt data collection

    Dim CmdList As Variant 'String
    Dim CmdListStatus As New SiteLong
    Dim TResult_coreup As New SiteLong
    Dim TResult_selsrm As New SiteLong
    Dim CZSetupName As String
    Dim powerPin As String
    Dim SupplyVoltage As Long
    Dim LogTimes As Boolean
    Shmoo_Pattern = TestName
    Dim instancename As String: instancename = TheExec.DataManager.instancename
    Dim DevChar_Setup As String
    Dim CMDTotalTT As Double
    
    Dim MasterSite As New SiteBoolean
    Dim SlaveSite As New SiteBoolean
    Dim RestoreSite As New SiteBoolean
    Dim MasterCMDOnly As Boolean
    Dim site As Variant 'Carter, 20240304
    MasterCMDOnly = g_MasterCMDOnly
    RestoreSite = TheExec.sites.Active
    For Each site In TheExec.sites
        If site < g_HardwareSiteCount Then
            MasterSite(site) = RestoreSite(site)
        Else
            SlaveSite(site) = RestoreSite(site)
        End If
    Next site
    
    g_Vbump_function = True 'Using for SELSRAM
    TestName = instancename
    For Each site In TheExec.sites
        TheExec.sites(site).FlagState("F_Rtos_func_coreup") = logicFalse
        TheExec.sites(site).FlagState("F_Rtos_func_SELSRM") = logicFalse
        TheExec.sites(site).FlagState("F_Rtos_func_SC06") = logicFalse
    Next site

    If Cmd1 <> "" Then CmdList = Cmd1
    If Cmd2 <> "" Then CmdList = CmdList + Cmd2
    If Cmd3 <> "" Then CmdList = CmdList + Cmd3
    If Cmd4 <> "" Then CmdList = CmdList + Cmd4
    If Cmd5 <> "" Then CmdList = CmdList + Cmd5
    
    CMDTotalTT = Cmd1TimeOut + Cmd2TimeOut + Cmd3TimeOut + Cmd4TimeOut + Cmd5TimeOut

    CMDTotalTT = CMDTotalTT + SetupCMD_Time
    TheExec.Datalog.DatalogSuspended = False
    TheExec.enableWord("RTOSRamp") = True
    
'    If TheExec.Flow.IsCharacterizing = True Then
    If TheExec.DevChar.Setups.IsRunning = False Then
       g_RTOS_FirstSetp = True
    Else
       DevChar_Setup = TheExec.DevChar.Setups.ActiveSetupName
       If TheExec.DevChar.results(DevChar_Setup).StartTime Like "1/1/0001*" Or TheExec.DevChar.results(DevChar_Setup).StartTime Like "0001/1/1*" Or g_LastRTOSPoint = True Then g_Vbump_function = False 'HS mod for ids issue
       If TheExec.DevChar.results(DevChar_Setup).StartTime Like "1/1/0001*" Or TheExec.DevChar.results(DevChar_Setup).StartTime Like "0001/1/1*" Or g_LastRTOSPoint = True Then Exit Function
       g_LastRTOSPoint = False
    End If
'    End If
    
    If g_RTOS_FirstSetp = True Then
       g_RTOSRampStep = 9
        If RampStep <> 0 Then
            If RampStep Mod 2 = 0 Then
                g_RTOSRampStep = RampStep + 1
            Else
                g_RTOSRampStep = RampStep
            End If
        End If
        
        'Add rampup case

        TheHdw.PinLevels.ApplyPower
        Shmoo_Save_core_power_per_site_for_Vbump ' store voltage into global variable
        TheHdw.DCVS.pins("All_Power").Voltage.Output = tlDCVSVoltageMain '' use safe voltage to Selsram
       
         '====================Update for Pmode + Interpose case =====================
        Dim Pmode_Voltage As String::   Pmode_Voltage = vbNullString
        Dim Forcepin_ary() As String, Forcepin_cnt As Long, k As Long ' Update for Pmode + Interpose case
       
        If pmode <> "" Then
            g_CharInputString_Voltage_Dict.RemoveAll
            Decide_Pmode_ForceVoltage pmode, "CorePower", Pmode_Voltage
            Call SetForceCondition(Pmode_Voltage & ";STOREPREPAT")
        End If
       
        If Interpose_PrePat <> "" Then
            Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
        End If
        
        If Pmode_Voltage <> "" And Interpose_PrePat = "" Then
            Interpose_PrePat = Pmode_Voltage
            Getforcecondition_VDD g_ForceCond_VDD, Interpose_PrePat
        ElseIf Pmode_Voltage <> "" And Interpose_PrePat <> "" Then
            Interpose_PrePat = Interpose_PrePat & ";" & Pmode_Voltage
            Getforcecondition_VDD g_ForceCond_VDD, Interpose_PrePat
            TheExec.DataManager.DecomposePinList g_ForceCond_VDD, Forcepin_ary, Forcepin_cnt
            Interpose_PrePat = vbNullString
            For k = 0 To Forcepin_cnt - 1
                Interpose_PrePat = Interpose_PrePat & ";" & UCase(Forcepin_ary(k)) & ":V:" & CDbl(g_CharInputString_Voltage_Dict(UCase(Forcepin_ary(k))))
            Next k
            Interpose_PrePat = mid(Interpose_PrePat, 2, Len(Interpose_PrePat))
        Else
            Getforcecondition_VDD g_ForceCond_VDD, Interpose_PrePat
        End If
         '====================Update for Pmode + Interpose case =====================
        g_dyanmicDSSCbits = vbNullString
        If SELSRAM_DSSC <> "" Then
        
            If UCase(SELSRAM_DSSC) Like "SELSRM*" Or UCase(SELSRAM_DSSC) Like "SELSRAM*" Then
                SELSRAM_DSSC = Replace(UCase(SELSRAM_DSSC), "SELSRAM", vbNullString)
                SELSRAM_DSSC = Replace(UCase(SELSRAM_DSSC), "SELSRM", vbNullString)
            ElseIf UCase(SELSRAM_DSSC) Like "DSELSRM*" Or UCase(SELSRAM_DSSC) Like "DSELSRAM*" Then
                SELSRAM_DSSC = Replace(UCase(SELSRAM_DSSC), "DSELSRAM", vbNullString)
                SELSRAM_DSSC = Replace(UCase(SELSRAM_DSSC), "DSELSRM", vbNullString)
                Call InverStr(SELSRAM_DSSC)
            End If
            
            Dim SrcBitAry() As Variant
            ReDim SrcBitAry(Len(SELSRAM_DSSC) - 1) As Variant
            Shmoo_Save_core_power_per_site_for_Vbump
        Else
            SELSRAM_DSSC = "SSSSS"
        End If
    End If
    
    Dim Shmoo_Apply_Pin As String, pin_count As Long
    Get_Shmoo_Set_Pin Shmoo_Apply_Pin, g_ForceCond_VDD, pin_count

    'Select Sram Start
    Dim uniquesBit As Boolean ', site As Variant
    Dim RTOS_SelAry() As Variant 'String
    ReDim RTOS_SelAry(TheExec.sites.Existing.Count - 1)
    
    For Each site In TheExec.sites.Selected
        RTOS_SelAry(site) = Decide_Switching_Bit_RTOS(SELSRAM_DSSC, g_ApplyLevelTimingValt, "RTOS", Shmoo_Apply_Pin, g_Globalpointval, g_ForceCond_VDD, g_CharInputString_Voltage_Dict)
    Next site
     
    If g_RTOSNwireChar = True Then
       If g_RTOS2DFirstPoint = True Then
          RTOS_Boot False, , , True, False
       Else
          RTOS_Boot False, , , True, True
          g_RTOS_FirstSetp = True
       End If
    End If

    'Selsram Command
    If MasterCMDOnly = True Then
        TheExec.sites.Selected = MasterSite
        TResult_selsrm = SendCmd(RTOS_SelAry, 0.1)
        TheExec.sites.Selected = RestoreSite
    Else
        TResult_selsrm = SendCmd(RTOS_SelAry, 0.1)
    End If
    
    If MasterCMDOnly = True Then
        For Each site In TheExec.sites
            If site < g_HardwareSiteCount Then
                TResult_selsrm(site + g_HardwareSiteCount) = TResult_selsrm(site)
            End If
        Next site
    End If
    TheExec.flow.TestLimit TResult_selsrm, 1, 1, , , , , , "RTOS_SELSRAM_COMMAND "  ', , , , , , TestName
    For Each site In TheExec.sites
        If TResult_selsrm <> "1" Then
            TheExec.sites(site).FlagState("F_Rtos_func_SELSRM") = logicTrue
        End If
    Next site
    
    If g_RTOS_FirstSetp = True Then
       If TheExec.enableWord("RTOSRamp") = True Then
          RTOS_Voltage_Rampdown
       Else
          Shmoo_Restore_Power_per_site_Vbump Shmoo_Apply_Pin
       End If
    Else
       g_VDDForce = vbNullString
       Shmoo_Restore_Power_per_site_Vbump Shmoo_Apply_Pin
    End If
    g_RTOS_FirstSetp = False
    
    TheHdw.Wait 0.005

    CmdListStatus = 0
    
    'Coreup Command
    If MasterCMDOnly = True Then
        TheExec.sites.Selected = MasterSite
        TResult_coreup = SendCmd("core up acc;fs mount SPI1;", 0.2)
        TheExec.sites.Selected = RestoreSite
    Else
        TResult_coreup = SendCmd("core up acc;fs mount SPI1;", 0.2)
    End If
    
    If MasterCMDOnly = True Then
        For Each site In TheExec.sites
            If site < g_HardwareSiteCount Then
                TResult_coreup(site + g_HardwareSiteCount) = TResult_coreup(site)
            End If
        Next site
    End If
    TheExec.flow.TestLimit TResult_coreup, 1, 1, , , , , , "RTOS_COREUP_COMMAND"
    For Each site In TheExec.sites
        If TResult_coreup <> "1" Then
            TheExec.sites(site).FlagState("F_Rtos_func_coreup") = logicTrue
        End If
    Next site
    
    'Scenario Command
    If MasterCMDOnly = True Then
        TheExec.sites.Selected = MasterSite
        If CmdList <> "" Then Set CmdListStatus = SendCmd(CmdList, CMDTotalTT, False)
        TheExec.sites.Selected = RestoreSite
    Else
        If CmdList <> "" Then Set CmdListStatus = SendCmd(CmdList, CMDTotalTT, False)
    End If

    If MasterCMDOnly = True Then
        For Each site In TheExec.sites
            If site < g_HardwareSiteCount Then
                CmdListStatus(site + g_HardwareSiteCount) = CmdListStatus(site)
            End If
        Next site
    End If
    ''''''''''''''''''''''''''''
'    TheExec.Flow.TestLimit CmdListStatus, 1, 1       ', , , , , , TestName
    
    Dim TnameCombShmooInfo As String
    TnameCombShmooInfo = vbNullString
    Dim TestNumberRTOS As Long
    
    If TheExec.DevChar.Setups.IsRunning = True Then
        Call TPmode_Char_on
        For Each site In TheExec.sites.Active
            TestNumberRTOS = TheExec.sites.item(site).TestNumber
        Next site
        For Each site In TheExec.sites
            Call PrintEachPoint_TestName(TnameCombShmooInfo)
        Next site
        TnameCombShmooInfo = TheExec.DataManager.instancename & TnameCombShmooInfo
        TheExec.flow.TestLimit CmdListStatus, 1, 1, , , , , , TnameCombShmooInfo, , , , , , , , TestNumberRTOS ', , , , , , TestName
        For Each site In TheExec.sites.Active
            TheExec.sites.item(site).TestNumber = TestNumberRTOS + 1
        Next site
        '''''==================================HS=====shmoo format by Motti request===============
    Else '''Modified the branch for non-bincut instances. Validated on JC-Chop, 20200717.
        TheExec.flow.TestLimit CmdListStatus, 1, 1        ', , , , , , TestName
    End If
      
    RTOS_UART_Print instancename, CmdListStatus
    g_RTOS2DFirstPoint = False
    g_RTOSNwireChar = False
    
    'Print shmoo voltage information''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If TheExec.DevChar.Setups.IsRunning = True Then
        Dim i As Integer
        Dim active_setup As String, curr_axis As Variant, curr_track As Variant, apply_Pin As String, apply_Pin_arry() As String, pin_count1 As Long
        active_setup = TheExec.DevChar.Setups.ActiveSetupName
        '''''''''''''TheExec.Datalog.WriteComment "Before ALT, VDD_PCPU0:" & TheHdw.DCVS.Pins("VDD_PCPU0").Voltage.Value & "V"
        '===========================================================
        For Each curr_axis In TheExec.DevChar.Setups(active_setup).Shmoo.axes.list
            ''exit for if any axis is not power pin -by SY
            If TheExec.DevChar.Setups(active_setup).Shmoo.axes(curr_axis).ApplyTo.pins <> "" Then
                apply_Pin = TheExec.DevChar.Setups(active_setup).Shmoo.axes(curr_axis).ApplyTo.pins
        '        Add for store shmoo global spec to avoid direct to apply Vmain used for Vbump function
                Call TheExec.DataManager.DecomposePinList(apply_Pin, apply_Pin_arry, pin_count1)
                For i = 0 To pin_count1 - 1
                   TheExec.Datalog.WriteComment "Scenario voltage, " & apply_Pin_arry(i) & "                       " & TheHdw.DCVS.pins(apply_Pin_arry(i)).Voltage.value & "V"
                Next i
                For Each curr_track In TheExec.DevChar.Setups(active_setup).Shmoo.axes(curr_axis).TrackingParameters.list
                    apply_Pin = TheExec.DevChar.Setups(active_setup).Shmoo.axes(curr_axis).TrackingParameters.item(curr_track).ApplyTo.pins
                    Call TheExec.DataManager.DecomposePinList(apply_Pin, apply_Pin_arry, pin_count1)
                    For i = 0 To pin_count1 - 1
                        TheExec.Datalog.WriteComment "Scenario voltage, " & apply_Pin_arry(i) & "                       " & TheHdw.DCVS.pins(apply_Pin_arry(i)).Voltage.value & "V"
                    Next i
                Next curr_track
            End If
        Next curr_axis
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    g_Vbump_function = False
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_RTOS", "RTOS_RunScenario")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function RTOS_UART_Print(instancename As String, Optional TResult As SiteLong)
On Error GoTo errHandler


    Dim site As Variant
    Dim PowerVolt As Double
    Dim powerPin As String
    Dim FName As String
    Dim OutputFilePath As String
    Dim day_code As String
    Dim SResult As String
    Dim CZSetupName As String
            
    Dim ByteCount As Long
    Dim asciiChar() As String
    Dim numericalVal() As Long
    
    Dim i As Long
    Dim j As Long
    
    Dim iPos As Long
    Dim FF_Count As Long
    Dim CR_Count As Long
            
    If TheExec.enableWord("UARTOutPut") = True Then '#16_ELSE_CASE_CHK
        day_code = CStr(Year(Now)) & right("0" & CStr(Month(Now)), 2) & right("0" & CStr(Day(Now)), 2)
        day_code = day_code & right("0" & CStr(Hour(Now)), 2) & right("0" & CStr(Minute(Now)), 2)
        
        For Each site In TheExec.sites
            If TResult(site) = 1 Then
                SResult = "Pass"
            Else
                SResult = "Fail"
            End If
            If TheExec.DevChar.Setups.IsRunning = True Then
                CZSetupName = TheExec.DevChar.Setups.ActiveSetupName
                powerPin = TheExec.DevChar.Setups(CZSetupName).Shmoo.axes.item(tlDevCharShmooAxis_X).ApplyTo.pins
                PowerVolt = Format(TheHdw.DCVS.pins(powerPin).Voltage.Alt.value * 1000, "0")
                OutputFilePath = ".\UART_Output\" & "Shmooing_Site" & site & "_" & "X_" & XCoord(site) & "_" & "Y_" & YCoord(site) & _
                                 "_" & instancename & "_UARToutput_" & day_code & "_" & powerPin & "_" & CStr(PowerVolt) & "mV" & "_" & SResult & ".txt"
            Else
                OutputFilePath = ".\UART_Output\" & "Site" & site & "_" & "X_" & XCoord(site) & "_" & "Y_" & _
                                 YCoord(site) & "_" & instancename & "_UARToutput_" & day_code & "_" & SResult & ".txt"
            End If
            
            Dim strLen As Long
            strLen = Len(GlobalMergeAry(site))
            
            FName = OutputFilePath
            Open FName For Append As #4
                Print #4, instancename
                For i = 1 To strLen
                Dim tempStr As String
                tempStr = mid(GlobalMergeAry(site), i, 1)
'                    Print #4, Mid(GlobalMergeAry(site), i, 1);
                    If Not (Asc(tempStr) = 255) Then
                        If Asc(tempStr) = 10 Or Asc(tempStr) = 13 Then  ''
                            Print #4, vbCrLf
                        Else
                            If Asc(tempStr) = 62 Then
                                Print #4, tempStr & vbCrLf;
    '                            Print #4, vbCrLf
                            Else
                                Print #4, tempStr;
                            End If
                        End If
                    End If
                Next i
            Close #4
            
        Next site
    Else
    End If
    ReDim GlobalMergeAry(TheExec.sites.Existing.Count - 1)
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_RTOS", "RTOS_UART_Print")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function SendCmd(CmdStr As Variant, Optional ExtendedWait As Double = 0#, Optional CheckPassFail As Boolean = False) As SiteLong
On Error GoTo errHandler
''ZHHUANGF

    Dim CharArray() As String
    Dim dataArray() As Long
    
    Dim PerSiteDataArray() As New SiteLong
    
    Dim StringLength As Long
    Dim StringLengthPlus1 As Long
    Dim i As Long
    Dim dspData As New PinListData
    
    Dim DUTResponse As New DSPWave
    Dim TResult As New SiteLong
    
    Dim LongestCmd As Long
    Dim UniqueCmdPerSite As Boolean
    Dim site As Variant

    UniqueCmdPerSite = IsArray(CmdStr)
    LongestCmd = 0
    
    If UniqueCmdPerSite Then
        For Each site In TheExec.sites.Selected
            If LongestCmd < Len(CmdStr(site)) Then LongestCmd = Len(CmdStr(site))
        Next site
        StringLength = LongestCmd
    Else
        StringLength = Len(CmdStr)
    End If
    
    StringLengthPlus1 = StringLength + 1
    
    ReDim CharArray(StringLength - 1)
    ReDim dataArray(StringLength)
    ReDim PerSiteDataArray(StringLength)
    
    TheHdw.Protocol.ports("UART_RX").Enabled = True
    TheHdw.Protocol.ports("UART_TX").Enabled = True
    
    
    If UniqueCmdPerSite Then
        For Each site In TheExec.sites.Selected
            For i = 1 To StringLength
                If i <= Len(CmdStr(site)) Then
                    CharArray(i - 1) = mid(CmdStr(site), i, 1)
                    PerSiteDataArray(i - 1) = Asc(CharArray(i - 1))
                Else
                    PerSiteDataArray(i - 1) = Asc(" ")
                End If
            Next i
            PerSiteDataArray(StringLength) = 13 ' carriage return
        Next site
    Else
        For i = 1 To StringLength
            CharArray(i - 1) = mid(CmdStr, i, 1)
            dataArray(i - 1) = Asc(CharArray(i - 1))
        Next i
        dataArray(StringLength) = 13 ' carriage return
        ' dataArray(StringLengthPlus1) = 10 ' carriage return
    End If
    
    TheHdw.Protocol.ports("UART_TX").NWire.MaxHoldUntilTimeout.value = 0.11
    
    TheHdw.Protocol.ports("UART_TX").NWire.CMEM.MoveMode = tlNWireCMEMMoveMode_Databus
    
    TheHdw.Protocol.ports("UART_TX").Modules("UART_read_response_extended").start '***
    
    If UniqueCmdPerSite Then
        For i = 0 To StringLength
            With TheHdw.Protocol.ports("UART_RX").NWire.Frames("UART_Snd")
                .Fields("Data_in").value = PerSiteDataArray(i)
                .Execute
            End With
            TheHdw.Protocol.ports("UART_RX").IdleWait
        Next i
    Else
        For i = 0 To StringLength
            With TheHdw.Protocol.ports("UART_RX").NWire.Frames("UART_Snd")
                .Fields("Data_in").value = dataArray(i)
                .Execute
            End With
            TheHdw.Protocol.ports("UART_RX").IdleWait
        Next i
    End If
    
    Set dspData = TheHdw.Protocol.ports("UART_TX").NWire.CMEM.DSPWave '***

    If ExtendedWait > 0.001 Then
        TheHdw.Wait ExtendedWait
    Else
        TheHdw.Wait 0.02
    End If
    TheHdw.Protocol.ports("UART_TX").Enabled = False
    TheHdw.Protocol.ports("UART_RX").Enabled = False
    

        TheHdw.Wait 0.001

'----------------------------------------add to avoid 2 interations-----------------------
    Dim Prompt_Idx As New SiteLong
    Dim UseCmdLength As Boolean

             If TheExec.DataManager.instancename Like "*S094*" Or TheExec.DataManager.instancename Like "*S095*" Or TheExec.DataManager.instancename Like "*S096*" Or TheExec.DataManager.instancename Like "*IDS*" Then
                If UniqueCmdPerSite = True Then
                    For Each site In TheExec.sites                      ''''''''''''''''''''''''''''''''''''''''HSLIU  S094
                        If CStr(CmdStr(site)) Like "pmgr sram -s*" Or CStr(CmdStr(site)) Like "core up acc;" Then
                            Prompt_Idx(site) = 1
                        Else
                            Prompt_Idx(site) = 0
                        End If
                    Next site
                Else
                    For Each site In TheExec.sites                      ''''''''''''''''''''''''''''''''''''''''HSLIU  S094
                            If CStr(CmdStr) Like "pmgr sram -s*" Or CStr(CmdStr) Like "core up acc;" Then
                                Prompt_Idx = 1
                            Else
                                Prompt_Idx = 0
                            End If
                    Next site
                End If
            ElseIf TheExec.DataManager.instancename Like "*UID_FUSE*" Then
                For Each site In TheExec.sites
                    Prompt_Idx(site) = 2
                Next site
             Else
                Prompt_Idx = 1
             End If

    'Add Offline Simulated Data
    If TheExec.TesterMode = testModeOffline Then
        Dim PassArrayOffline(8) As Long
        Dim FailArrayOffline(8) As Long
        Dim dspdata_offline As New DSPWave
        Dim RandomVal As Double
        
        For Each site In TheExec.sites.Selected
            RandomVal = Rnd      '0<=RandomVal<1
            If RandomVal < 0.8 Or TheExec.enableWord("Golden_Default") = True Then
                PassArrayOffline(0) = 80 'P
                PassArrayOffline(1) = 65 'A
                PassArrayOffline(2) = 83 'S
                PassArrayOffline(3) = 83 'S
                PassArrayOffline(4) = CLng(Asc(" ")) 'space
                PassArrayOffline(5) = 65 'A
                PassArrayOffline(6) = 84 'T
                PassArrayOffline(7) = 69 'E
                PassArrayOffline(8) = 62 '>
                
                dspdata_offline.data = PassArrayOffline
            Else
                FailArrayOffline(0) = 70 'F
                FailArrayOffline(1) = 65 'A
                FailArrayOffline(2) = 73 'I
                FailArrayOffline(3) = 76 'L
                FailArrayOffline(4) = CLng(Asc(" ")) 'space
                FailArrayOffline(5) = 65 'A
                FailArrayOffline(6) = 84 'T
                FailArrayOffline(7) = 69 'E
                FailArrayOffline(8) = 62 '>
                
                dspdata_offline.data = FailArrayOffline
            End If
        Next site
        rundsp.ProcessDUTResponse dspdata_offline, TResult, Prompt_Idx, StringLength, UseCmdLength
        For Each site In TheExec.sites.Selected
                DUTResponse(site) = dspdata_offline(site).COPY
        Next site
    Else
        rundsp.ProcessDUTResponse dspData, TResult, Prompt_Idx, StringLength, UseCmdLength
        
        DUTResponse = dspData.COPY
    End If
            
    Call LogDUTResponse(DUTResponse, TResult)

    Set SendCmd = TResult

    If CheckPassFail Then   '#16_ELSE_CASE_CHK
        TheExec.flow.TestLimit TResult, 1, 1, , , , , , CmdStr, , , , , , , tlForceNone
    Else
    End If
       
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_RTOS", "SendCmd")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function LogDUTResponse(DUTResponse As DSPWave, Optional TResult As SiteLong, Optional OutputToDebuglog As Boolean, Optional OutputToDatalog As Boolean)
On Error GoTo errHandler

    Dim site As Variant
    Dim ByteCount As Long

    Dim asciiChar() As String
    Dim numericalVal() As Long
    
    Dim i As Long
    Dim j As Long
    
    Dim iPos As Long
    Dim FF_Count As Long
    Dim CR_Count As Long
    Dim print_file_cmdStr As String


    OutputToDatalog = True: OutputToDebuglog = False

    If TheExec.flow.enableWord("UARTToDatalog") = True Then '#16_ELSE_CASE_CHK
        OutputToDatalog = True
    Else
    End If
    If TheExec.flow.enableWord("UARTToDebuglog") = True Then    '#16_ELSE_CASE_CHK
        OutputToDebuglog = True
    Else
    End If
    
    ReDim Preserve GlobalMergeAry(TheExec.sites.Existing.Count - 1)
    
    For Each site In TheExec.sites.Selected
        j = 0
        ByteCount = DUTResponse.SampleSize
        
        ReDim asciiChar(ByteCount - 1)
        ReDim numericalVal(ByteCount - 1)
        
        For i = 0 To ByteCount - 1
            asciiChar(i) = vbNullString
            numericalVal(i) = DUTResponse.Element(i)    '#16_ELSE_CASE_CHK
            If (numericalVal(i) <> 255) And (numericalVal(i) <> 10) Then
                asciiChar(j) = Chr(numericalVal(i))
                j = j + 1
            Else
            End If
        Next i
        
        GlobalMergeAry(site) = GlobalMergeAry(site) & Join(asciiChar(), vbNullString)
        
        If OutputToDatalog Then '#16_ELSE_CASE_CHK
            TheExec.Datalog.WriteComment "****************************************"
            TheExec.Datalog.WriteComment "Site : " + CStr(site)
            TheExec.Datalog.WriteComment "****************************************"
            
            WriteToDatalog asciiChar, j
    
        ElseIf OutputToDebuglog Then
            TheExec.AddOutput "****************************************"
            TheExec.AddOutput "Site : " + CStr(site)
            TheExec.AddOutput "****************************************"
            WriteToOutputWindow asciiChar, j
        Else
        End If
        
    Next site

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_RTOS", "LogDUTResponse")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub SendCmdOnly(CmdStr As String)
On Error GoTo errHandler

    Dim CharArray() As String
    Dim StringLength As Long
    Dim i As Long
    Dim dataArray() As Long
    
    StringLength = Len(CmdStr)
    ReDim CharArray(StringLength - 1)
    ReDim dataArray(StringLength)
    
    For i = 1 To StringLength
        CharArray(i - 1) = mid(CmdStr, i, 1)
        dataArray(i - 1) = Asc(CharArray(i - 1))
    Next i
    dataArray(StringLength) = 13 ' carriage return

    For i = 0 To StringLength    ' leave newline char to read module
        With TheHdw.Protocol.ports("UART_RX").NWire.Frames("UART_Snd")
            .Fields("Data_in").value = dataArray(i)
            .Execute
        End With
'        thehdw.Protocol.ports("UART_PA").IdleWait
    Next i
    
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_RTOS", "SendCmdOnly")
    If AbortTest Then Exit Sub Else Resume Next
End Sub


Public Sub WriteToOutputWindow(dataArray() As String, CharCount As Long)
On Error GoTo errHandler

    Dim OutLine As String
    Dim i As Long


    OutLine = vbNullString
    For i = 0 To CharCount - 1
        If (Asc(dataArray(i)) = 10) Or (Asc(dataArray(i)) = 13) Then
            TheExec.AddOutput OutLine
            OutLine = vbNullString
        Else
            OutLine = OutLine + dataArray(i)
'            If (Asc(dataArray(i)) = 62) Then
'                theexec.AddOutput OutLine
'                OutLine = ""
'                i = CharCount   'UBound(DataArray)
'            End If
        End If
    Next i
    
    If Len(OutLine) > 0 Then    '#16_ELSE_CASE_CHK
        TheExec.AddOutput OutLine
    Else
    End If
        
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_RTOS", "WriteToOutputWindow")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub WriteToDatalog(dataArray() As String, CharCount As Long)
On Error GoTo errHandler

Dim OutLine As String
Dim i As Long


    OutLine = vbNullString
    For i = 0 To CharCount - 1
        If (Asc(dataArray(i)) = 10) Or (Asc(dataArray(i)) = 13) Then
            TheExec.Datalog.WriteComment OutLine
            OutLine = vbNullString
        Else
            OutLine = OutLine + dataArray(i)
'            If (Asc(dataArray(i)) = 62) Then
'                TheExec.Datalog.WriteComment OutLine
'                OutLine = ""
'                i = CharCount   'UBound(DataArray)
'            End If
        End If
    Next i
    
    If Len(OutLine) > 0 Then    '#16_ELSE_CASE_CHK
        TheExec.Datalog.WriteComment OutLine
    Else
    End If
        
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_RTOS", "WriteToDatalog")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function ReloadUARTModules() As Long
On Error GoTo errHandler


    TheHdw.Protocol.ports("UART_TX").ModuleFiles.UnloadAll
    TheHdw.Protocol.ports("UART_RX").ModuleFiles.UnloadAll

    TheHdw.Protocol.ports("UART_TX").Enabled = True
    TheHdw.Protocol.ports("UART_RX").Enabled = True
    
    TheHdw.Protocol.ports("UART_TX").NWire.CMEM.MoveMode = tlNWireCMEMMoveMode_Databus
   
    With TheHdw.Protocol.ports("UART_TX")
        If (Not .ModuleFiles.Contains("VBT_UART_TX_module")) Then   '#16_ELSE_CASE_CHK
            Call .ModuleFiles.Load("VBT_UART_TX_module", False, False, True)
        Else
        End If
    End With
        
   
    With TheHdw.Protocol.ports("UART_RX")
        If (Not .ModuleFiles.Contains("VBT_UART_RX_module")) Then   '#16_ELSE_CASE_CHK
            Call .ModuleFiles.Load("VBT_UART_RX_module", False, False, True)
        Else
        End If
    End With
    
    TheHdw.Protocol.ports("UART_TX").Enabled = False
    TheHdw.Protocol.ports("UART_RX").Enabled = False
    TheExec.flow.enableWord("Enable_Reload_UART") = False
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_RTOS", "ReloadUARTModules")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function RTOS_Voltage_Rampdown()
On Error GoTo errHandler
    Dim p_ary() As String, p_cnt As Long, i As Long, InstName As String
    Dim step As Integer
    Dim StepNum As Integer
    Dim site As Variant 'Carter, 20240304
    StepNum = g_RTOSRampStep
    
    TheExec.DataManager.DecomposePinList "CorePower", p_ary, p_cnt

    For step = 1 To StepNum
        For i = 0 To p_cnt - 1
            If TheExec.DataManager.ChannelType(p_ary(i)) <> "N/C" Then  '#16_ELSE_CASE_CHK
               If step Mod 2 = 1 Then
                  For Each site In TheExec.sites
                      TheHdw.DCVS.pins(p_ary(i)).Voltage.Alt.value = g_ApplyLevelTimingVmain.pins(p_ary(i)).value - ((g_ApplyLevelTimingVmain.pins(p_ary(i)).value - g_RTOS_SceVoltage.pins(p_ary(i)).value) / StepNum) * step
                  Next site
               ElseIf step Mod 2 = 0 Then
                  For Each site In TheExec.sites
                      TheHdw.DCVS.pins(p_ary(i)).Voltage.Main.value = g_ApplyLevelTimingVmain.pins(p_ary(i)).value - ((g_ApplyLevelTimingVmain.pins(p_ary(i)).value - g_RTOS_SceVoltage.pins(p_ary(i)).value) / StepNum) * step
                  Next site
               End If
            Else
            End If
        Next i
        If step Mod 2 = 1 Then  '#16_ELSE_CASE_CHK
           TheHdw.DCVS.pins("All_Power").Voltage.Output = tlDCVSVoltageAlt
           TheHdw.Wait 20 * 0.000001
        ElseIf step Mod 2 = 0 Then
           TheHdw.DCVS.pins("All_Power").Voltage.Output = tlDCVSVoltageMain
           TheHdw.Wait 20 * 0.000001
        Else
        End If
    Next step
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_RTOS", "RTOS_Voltage_Rampdown")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function RTOS_Voltage_RampUp()
On Error GoTo errHandler
    Dim p_ary() As String, p_cnt As Long, i As Long, InstName As String
    Dim step As Integer
    Dim StepNum As Integer
    Dim site As Variant 'Carter, 20240304
    
    StepNum = g_RTOSRampStep
        
    TheExec.DataManager.DecomposePinList "CorePower", p_ary, p_cnt
    
    For step = 1 To StepNum
        For i = 0 To p_cnt - 1
            If TheExec.DataManager.ChannelType(p_ary(i)) <> "N/C" Then  '#16_ELSE_CASE_CHK
               If step Mod 2 = 1 Then
                  For Each site In TheExec.sites
                      TheHdw.DCVS.pins(p_ary(i)).Voltage.Main.value = g_RTOS_SceVoltage.pins(p_ary(i)).value + ((g_ApplyLevelTimingVmain.pins(p_ary(i)).value - g_RTOS_SceVoltage.pins(p_ary(i)).value) / StepNum) * step
                  Next site
               ElseIf step Mod 2 = 0 Then
                  For Each site In TheExec.sites
                      TheHdw.DCVS.pins(p_ary(i)).Voltage.Alt.value = g_RTOS_SceVoltage.pins(p_ary(i)).value + ((g_ApplyLevelTimingVmain.pins(p_ary(i)).value - g_RTOS_SceVoltage.pins(p_ary(i)).value) / StepNum) * step
                  Next site
               End If
            Else
            End If
        Next i
        If step Mod 2 = 1 Then  '#16_ELSE_CASE_CHK
           TheHdw.DCVS.pins("All_Power").Voltage.Output = tlDCVSVoltageMain
           TheHdw.Wait 20 * 0.000001
        ElseIf step Mod 2 = 0 Then
           TheHdw.DCVS.pins("All_Power").Voltage.Output = tlDCVSVoltageAlt
           TheHdw.Wait 20 * 0.000001
        Else
        End If
    Next step
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_RTOS", "RTOS_Voltage_RampUp")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Decide_Switching_Bit_RTOS(digSrc_EQ As String, Optional DC_Level As PinListData, Optional BlockType As String, Optional shmoo_pin As String, Optional ShmooPinsVoltage As PinListData, Optional ForcePin As String, Optional SetForceVoltage As Dictionary) As String
On Error GoTo errHandler

    Dim site As Variant
    Dim logicPin As String
    Dim SramPin As String
    Dim DSSC_Switching_Voltage As New PinListData
    Dim Sdomain() As Long
    Dim SramValue As Double
    Dim DSSCSelSrmOpposite As Long
    Dim BlockTypeNum As Long
    Dim i As Integer, j As Integer
    Dim ReturnString() As String
    Dim MasterSlaveString() As String
    Dim SRAM_temp As String
  
    On Error GoTo errHandler
    BlockTypeNum = -1
    
    ReDim ReturnString(Len(digSrc_EQ) - 1)
    ReDim MasterSlaveString(2 * Len(digSrc_EQ) - 1)
    Decide_DSSC_Switching_Voltage DSSC_Switching_Voltage, DC_Level, shmoo_pin, ShmooPinsVoltage, ForcePin, SetForceVoltage
    '///find blocktype
  
    Dim l_Selsram_index As Long
    Dim DigSrc_wav As New DSPWave 'dummy dspwave
    Dim sl_ReturnString() As New SiteLong
    Dim s_Statement As String
    
    l_Selsram_index = SelSRAM_Index_Select(SelsramMapping, "RTOS", "")
    Set DigSrc_wav = Nothing
    DigSrc_wav.CreateConstant 0, Len(digSrc_EQ)
    ReDim sl_ReturnString(Len(digSrc_EQ) - 1)
   
    If l_Selsram_index <> -1 Then
        Call SelSRAM_DigSrc_Bit(l_Selsram_index, digSrc_EQ, DSSC_Switching_Voltage, DigSrc_wav, ReturnString, sl_ReturnString)
    For i = 0 To UBound(ReturnString)
        If g_MasterCMDOnly = True Then
            SRAM_temp = Replace(Replace(ReturnString(i), "1", "ffff"), "0", "0000")
            MasterSlaveString(i) = SelsramMapping(BlockTypeNum).COMMENT(i) & SRAM_temp
            MasterSlaveString(i + Len(digSrc_EQ)) = Replace(SelsramMapping(BlockTypeNum).COMMENT(i), "d0", "d1") & SRAM_temp
            SRAM_temp = vbNullString
        Else
                ReturnString(i) = SelsramMapping(l_Selsram_index).COMMENT(i) & Replace(Replace(ReturnString(i), "1", "ffff"), "0", "0000")
        End If
    Next i
      
    If g_MasterCMDOnly = True Then
        Decide_Switching_Bit_RTOS = Join(MasterSlaveString, ";")
        Decide_Switching_Bit_RTOS = Decide_Switching_Bit_RTOS & ";"
    Else
    Decide_Switching_Bit_RTOS = Join(ReturnString, ";")
        Decide_Switching_Bit_RTOS = Decide_Switching_Bit_RTOS & ";"
    End If
    
    g_RTOS_SceVoltage = DSSC_Switching_Voltage.COPY
  Else
        s_Statement = "We could not find from the SelSRAM Mapping Table properly."
        ''Standardize error message
  End If
  
  Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_RTOS", "Decide_Switching_Bit_RTOS")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function RTOS_Freq_Measurement(frequency_meas_pins As String, starttimer As Double, MoniotrTime As Double)
On Error GoTo errHandler

    Dim tmpstoptimer As Double
    Dim CounterValue As New PinListData
    Dim FctrtimerStart As Double
    Dim FctrtimerStop As Double
    Dim CounterValAry() As New SiteLong
    Dim i As Double
    Dim site As Variant 'Carter, 20240304
    
    tmpstoptimer = TheExec.Timer(starttimer)
    TheExec.Datalog.WriteComment "Seuptime " & tmpstoptimer & " Sec!!!"
    
TheHdw.Digital.pins(frequency_meas_pins).Levels.value(chVoh) = 0.3
TheHdw.Digital.pins(frequency_meas_pins).Levels.value(chVol) = 0.3
    FctrtimerStart = TheExec.Timer(0)
    FctrtimerStop = TheExec.Timer(FctrtimerStart)
    i = 0
    
    While FctrtimerStop < MoniotrTime
    
        With TheHdw.Digital.pins(frequency_meas_pins).FreqCtr
            .EventSource = VOH
            .EventSlope = Positive
            .Interval = 0.01
            .Enable = IntervalEnable
            .Clear
            .start
            CounterValue = .Read()
        End With
        ReDim Preserve CounterValAry(i)
        For Each site In TheExec.sites
            If CounterValue.pins(frequency_meas_pins).value(site) <> 0 Then
                TheExec.Datalog.WriteComment "Site: " & site & ", CounterNumber: " & CounterValue.pins(frequency_meas_pins).value(site) & ", Frequency: " & Format((CounterValue.pins(frequency_meas_pins).value(site) / 0.01) / 1000000, "000.00") & " MHz , pins = " & frequency_meas_pins
            ElseIf FctrtimerStop > MoniotrTime - 0.05 Then
                TheExec.Datalog.WriteComment "Site: " & site & ", CounterNumber: " & CounterValue.pins(frequency_meas_pins).value(site) & ", Frequency: " & Format((CounterValue.pins(frequency_meas_pins).value(site) / 0.01) / 1000000, "000.00") & " MHz , pins = " & frequency_meas_pins
            End If
        Next site
    
        CounterValAry(i) = CounterValue.pins(frequency_meas_pins)
        FctrtimerStop = TheExec.Timer(FctrtimerStart)
        i = i + 1
    Wend
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_RTOS", "RTOS_Freq_Measurement")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function RTOS_DigCap_Setting(patt As String, DigCapPin As String, DigCapSignal As String, SampleSize As Long)
On Error GoTo errHandler
      
    With TheHdw.DSSC.pins(DigCapPin).Pattern(patt).Capture
        .Signals.Add DigCapSignal
        If LCase(glb_TesterType) = "jaguar" Then    '#16_ELSE_CASE_CHK
            .Signals(DigCapSignal).offset = 0
        Else
        End If
        .Signals.item(DigCapSignal).SampleSize = SampleSize
        .Signals(DigCapSignal).LoadSettings
    End With
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_RTOS", "RTOS_DigCap_Setting")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function RTOS_RunScenario_T(StartOfBodyF As InterposeName, PrePatF As InterposeName, PreTestF As InterposeName, _
                        PostTestF As InterposeName, PostPatF As InterposeName, EndOfBodyF As InterposeName, _
                        StartOfBodyFArgs As String, PrePatFArgs As String, PreTestFArgs As String, _
                        PostTestFArgs As String, PostPatFArgs As String, EndOfBodyFArgs As String, _
                        Util1Pins As PinList, Util0Pins As PinList, DriveLoPins As PinList, DriveHiPins As PinList, _
                        DriveZPins As PinList, FloatPins As PinList, DisablePins As PinList, Step_ As SubType, _
                        Optional RepeatableSequence As Boolean, Optional Boot_Init As Boolean = False, _
                        Optional TestName As String, Optional PreSetup As String, Optional Cmd1 As String, Optional Cmd2 As String, Optional Cmd3 As String, Optional Cmd4 As String, _
                        Optional Cmd5 As String, Optional Cmd1TimeOut As Double = 0#, Optional Cmd2TimeOut As Double = 0#, Optional Cmd3TimeOut As Double = 0#, _
                        Optional Cmd4TimeOut As Double = 0#, Optional Cmd5TimeOut As Double = 0#, Optional SELSRAM_DSSC As String, Optional Interpose_PrePat As String, _
                        Optional pmode As String, Optional ForceCMD As String, Optional RampStep As Integer, Optional SetupCMD_Time As Double, Optional Vbump As Boolean = True, _
                        Optional CSW_MeasPin As String) As Long
On Error GoTo errHandler
    
    RTOS_RunScenario_T = TL_SUCCESS    ' be optimistic
    If Not TheExec.flow.IsRunning Then Exit Function
    
    Dim CmdList As Variant 'String
    Dim CmdListStatus As New SiteLong
    Dim TResult_coreup As New SiteLong
    Dim TResult_selsrm As New SiteLong
    Dim Reboot_Flag As Boolean
    Dim CZSetupName As String
    Dim powerPin As String
    Dim SupplyVoltage As Long
    Dim LogTimes As Boolean
    Shmoo_Pattern = TestName
    Dim instancename As String: instancename = TheExec.DataManager.instancename
    Dim DevChar_Setup As String
    Dim CMDTotalTT As Double
    Dim uniquesBit As Boolean 'site As Variant
    Dim RTOS_SelAry() As Variant 'String
    
    Dim MasterSite As New SiteBoolean
    Dim SlaveSite As New SiteBoolean
    Dim RestoreSite As New SiteBoolean
    Dim MasterCMDOnly As Boolean
    Dim site As Variant 'Carter, 20240304
    MasterCMDOnly = g_MasterCMDOnly
    RestoreSite = TheExec.sites.Active
    For Each site In TheExec.sites
        If site < g_HardwareSiteCount Then
            MasterSite(site) = RestoreSite(site)
        Else
            SlaveSite(site) = RestoreSite(site)
        End If
    Next site
    
    ReDim RTOS_SelAry(TheExec.sites.Existing.Count - 1)
    g_Vbump_function = True 'Using for SELSRAM
    TestName = instancename
    For Each site In TheExec.sites
        TheExec.sites(site).FlagState("F_Rtos_func_coreup") = logicFalse
        TheExec.sites(site).FlagState("F_Rtos_func_SELSRM") = logicFalse
        TheExec.sites(site).FlagState("F_Rtos_func_SC06") = logicFalse
    Next site

    If Cmd1 <> "" Then  '#16_ELSE_CASE_CHK
        CmdList = Cmd1
    Else
    End If
    If Cmd2 <> "" Then  '#16_ELSE_CASE_CHK
        CmdList = CmdList + Cmd2
    Else
    End If
    If Cmd3 <> "" Then  '#16_ELSE_CASE_CHK
        CmdList = CmdList + Cmd3
    Else
    End If
    If Cmd4 <> "" Then  '#16_ELSE_CASE_CHK
        CmdList = CmdList + Cmd4
    Else
    End If
    If Cmd5 <> "" Then  '#16_ELSE_CASE_CHK
        CmdList = CmdList + Cmd5
    Else
    End If
    
    CMDTotalTT = Cmd1TimeOut + Cmd2TimeOut + Cmd3TimeOut + Cmd4TimeOut + Cmd5TimeOut
    CMDTotalTT = CMDTotalTT + SetupCMD_Time
    TheExec.Datalog.DatalogSuspended = False
    TheExec.enableWord("RTOSRamp") = False
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    If Step_ = subAllBody Or Step_ = subPrebody Or _
       m_InterposeFunctionsSet = False Then

        ' Register interpose function names with the flow controller, which may need to invoke them
        Call tl_SetInterpose(TL_C_PREPATF, PrePatF.value, PrePatFArgs, _
                             TL_C_POSTPATF, PostPatF.value, PostPatFArgs, _
                             TL_C_PRETESTF, PreTestF.value, PreTestFArgs, _
                             TL_C_POSTTESTF, PostTestF.value, PostTestFArgs)

        m_InterposeFunctionsSet = True
    End If

    ' PreBody
    If Step_ = subAllBody Or Step_ = subPrebody Then
        ' Get timing and levels info
        FetchContext
        
        RTOSTest_Inst = TheExec.DataManager.instancename
        
        ' Set up the test
        Call PreBody(DriveHiPins, DriveLoPins, DriveZPins, DisablePins, Util0Pins, Util1Pins, RepeatableSequence)
        g_RTOS_FirstSetp = True
                
                If BootUpInstanceName = "" Then BootUpInstanceName = "RTOS_Boot"
        If Boot_Init = True Or TheExec.flow.IsCharacterizing = True Then        'CHAR: Need to default enable   #16_ELSE_CASE_CHK
            TheExec.flow.instance(BootUpInstanceName).Execute
'            RTOS_Boot False, , , True
'            g_Reboot_Flag = False
        Else
        End If
        
        TheExec.Datalog.Setup.Shared.ascii.Columns.EnableCustomWidths = True
        TheExec.Datalog.Setup.Shared.ascii.Columns.Parametric.TestName.Width = 150
        TheExec.Datalog.ApplySetup  'must need to apply after datalog setup
    End If ' PreBody
    
    ' Body
    If Step_ = subAllBody Or Step_ = subBody Then
        ' Perform the test
        Call Interpose(StartOfBodyF, StartOfBodyFArgs)  ' Run StartOfBody interpose function, if specified
        TheHdw.DCVS.pins(Core_Power_DCVS_pins).Voltage.Output = tlDCVSVoltageMain '' use safe voltage to Selsram
        If g_Reboot_Flag = True Then    '#16_ELSE_CASE_CHK
            TheExec.flow.instance(BootUpInstanceName).Execute
'            g_Reboot_Flag = False
        Else
        End If
    If g_RTOS_FirstSetp = True Then '#16_ELSE_CASE_CHK
       g_RTOSRampStep = 9
        If RampStep <> 0 Then   '#16_ELSE_CASE_CHK
            If RampStep Mod 2 = 0 Then
                g_RTOSRampStep = RampStep + 1
            Else
                g_RTOSRampStep = RampStep
            End If
        Else
        End If
        
        'Add rampup case
        If TheExec.enableWord("RTOSRamp") = True Then   '#16_ELSE_CASE_CHK
            RTOS_Voltage_RampUp
        Else
        End If

        Shmoo_Save_core_power_per_site_for_Vbump ' store voltage into global variable
       
         '====================Update for Pmode + Interpose case =====================
        Dim Pmode_Voltage As String::   Pmode_Voltage = vbNullString
        Dim Forcepin_ary() As String, Forcepin_cnt As Long, k As Long ' Update for Pmode + Interpose case
       
        If pmode <> "" Then '#16_ELSE_CASE_CHK
            g_CharInputString_Voltage_Dict.RemoveAll
            Decide_Pmode_ForceVoltage pmode, "All_Power", Pmode_Voltage
            Call SetForceCondition(Pmode_Voltage & ";STOREPREPAT")
        Else
        End If
       
        If Interpose_PrePat <> "" Then  '#16_ELSE_CASE_CHK
            Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
        Else
        End If
        
        If Pmode_Voltage <> "" And Interpose_PrePat = "" Then
            Interpose_PrePat = Pmode_Voltage
            Getforcecondition_VDD g_ForceCond_VDD, Interpose_PrePat
        ElseIf Pmode_Voltage <> "" And Interpose_PrePat <> "" Then
            Interpose_PrePat = Interpose_PrePat & ";" & Pmode_Voltage
            Getforcecondition_VDD g_ForceCond_VDD, Interpose_PrePat
            TheExec.DataManager.DecomposePinList g_ForceCond_VDD, Forcepin_ary, Forcepin_cnt
            Interpose_PrePat = vbNullString
            For k = 0 To Forcepin_cnt - 1
                Interpose_PrePat = Interpose_PrePat & ";" & UCase(Forcepin_ary(k)) & ":V:" & CDbl(g_CharInputString_Voltage_Dict(UCase(Forcepin_ary(k))))
            Next k
            Interpose_PrePat = mid(Interpose_PrePat, 2, Len(Interpose_PrePat))
        Else
            Getforcecondition_VDD g_ForceCond_VDD, Interpose_PrePat
        End If
         '====================Update for Pmode + Interpose case =====================
        Call Body(FloatPins)
        
        g_dyanmicDSSCbits = vbNullString
        If SELSRAM_DSSC <> "" Then
        
            If UCase(SELSRAM_DSSC) Like "SELSRM*" Or UCase(SELSRAM_DSSC) Like "SELSRAM*" Then   '#16_ELSE_CASE_CHK
                SELSRAM_DSSC = Replace(UCase(SELSRAM_DSSC), "SELSRAM", vbNullString)
                SELSRAM_DSSC = Replace(UCase(SELSRAM_DSSC), "SELSRM", vbNullString)
            ElseIf UCase(SELSRAM_DSSC) Like "DSELSRM*" Or UCase(SELSRAM_DSSC) Like "DSELSRAM*" Then
                SELSRAM_DSSC = Replace(UCase(SELSRAM_DSSC), "DSELSRAM", vbNullString)
                SELSRAM_DSSC = Replace(UCase(SELSRAM_DSSC), "DSELSRM", vbNullString)
                Call InverStr(SELSRAM_DSSC)
            Else
            End If
            
            Dim SrcBitAry() As Variant
            ReDim SrcBitAry(Len(SELSRAM_DSSC) - 1) As Variant
            Shmoo_Save_core_power_per_site_for_Vbump
        Else
            SELSRAM_DSSC = "SSSSSSSSSSS"
        End If
    Else
    End If
    
    If PreSetup <> "" Then
        SendCmd PreSetup, 0.1, False
    End If
    
    Dim Shmoo_Apply_Pin As String, pin_count As Long
    Get_Shmoo_Set_Pin Shmoo_Apply_Pin, g_ForceCond_VDD, pin_count

    For Each site In TheExec.sites.Selected
        RTOS_SelAry(site) = Decide_Switching_Bit_RTOS(SELSRAM_DSSC, g_ApplyLevelTimingValt, "RTOS", Shmoo_Apply_Pin, g_Globalpointval, g_ForceCond_VDD, g_CharInputString_Voltage_Dict)
            TheExec.Datalog.WriteComment "=================================================================================================="
            TheExec.Datalog.WriteComment "Selsrm CMD of site" & site & ": " & RTOS_SelAry(site)
            TheExec.Datalog.WriteComment "=================================================================================================="
    Next site

    TResult_selsrm = SendCmd(RTOS_SelAry, 0.2)
    
    If MasterCMDOnly = True Then    '#16_ELSE_CASE_CHK
        For Each site In TheExec.sites
            If site < g_HardwareSiteCount Then  '#16_ELSE_CASE_CHK
                TResult_selsrm(site + g_HardwareSiteCount) = TResult_selsrm(site)
            Else
            End If
        Next site
    Else
    End If
    
    TheExec.flow.TestLimit TResult_selsrm, 1, 1, , , , , , "RTOS_SELSRAM_COMMAND "  ', , , , , , TestName
    For Each site In TheExec.sites
        If TResult_selsrm <> "1" Then   '#16_ELSE_CASE_CHK
            TheExec.sites(site).FlagState("F_Rtos_func_SELSRAM") = logicTrue
        Else
        End If
    Next site
     
    
    '//Change to Valt mode
    If g_RTOS_FirstSetp = True Then
       If TheExec.enableWord("RTOSRamp") = True Then
          RTOS_Voltage_Rampdown
       Else
          Shmoo_Restore_Power_per_site_Vbump Shmoo_Apply_Pin
       End If
    Else
       g_VDDForce = vbNullString
       Shmoo_Restore_Power_per_site_Vbump Shmoo_Apply_Pin
    End If
    g_RTOS_FirstSetp = False
    
    TheHdw.Wait 0.005

    CmdListStatus = 0
    
       TResult_coreup = SendCmd("core up acc;fs mount SPI1;", 0.5)
    
    If MasterCMDOnly = True Then    '#16_ELSE_CASE_CHK
        For Each site In TheExec.sites
            If site < g_HardwareSiteCount Then  '#16_ELSE_CASE_CHK
                TResult_coreup(site + g_HardwareSiteCount) = TResult_coreup(site)
            Else
            End If
        Next site
    Else
    End If
    
    TheExec.flow.TestLimit TResult_coreup, 1, 1, , , , , , "RTOS_COREUP_COMMAND"
    For Each site In TheExec.sites
        If TResult_coreup <> "1" Then   '#16_ELSE_CASE_CHK
            TheExec.sites(site).FlagState("F_Rtos_func_coreup") = logicTrue
        Else
        End If
    Next site
    
    If CmdList <> "" Then   '#16_ELSE_CASE_CHK
        Set CmdListStatus = SendCmd(CmdList, CMDTotalTT, False)
    Else
    End If

    'For multidevice result'''''
    If g_MasterCMDOnly = True Then  '#16_ELSE_CASE_CHK
        For Each site In TheExec.sites
            If site < g_HardwareSiteCount Then  '#16_ELSE_CASE_CHK
                CmdListStatus(site + g_HardwareSiteCount) = CmdListStatus(site)
            Else
            End If
        Next site
    Else
    End If
    
    Dim TnameCombShmooInfo As String
    TnameCombShmooInfo = vbNullString
    Dim TestNumberRTOS As Long
    
    If TheExec.DevChar.Setups.IsRunning = True Then
        Call TPmode_Char_on
        For Each site In TheExec.sites.Active
            TestNumberRTOS = TheExec.sites.item(site).TestNumber
        Next site
        For Each site In TheExec.sites
            Call PrintEachPoint_TestName(TnameCombShmooInfo)
        Next site
        TnameCombShmooInfo = RTOSTest_Inst & TnameCombShmooInfo
        TheExec.flow.TestLimit CmdListStatus, 1, 1, , , , , , TnameCombShmooInfo, , , , , , , , TestNumberRTOS ', , , , , , TestName
        For Each site In TheExec.sites.Active
            TheExec.sites.item(site).TestNumber = TestNumberRTOS + 1
        Next site
        '''''==================================HS=====shmoo format by Motti request===============
    Else '''Modified the branch for non-bincut instances. Validated on JC-Chop, 20200717.
        TheExec.flow.TestLimit CmdListStatus, 1, 1        ', , , , , , TestName
        For Each site In TheExec.sites
            If CmdListStatus <> "1" Then
                TheExec.sites(site).FlagState("F_Rtos_func_SC06") = logicTrue
            End If
        Next site
    End If
        
    If TheExec.DevChar.Setups.IsRunning = True Or TheExec.flow.enableWord("RTOS_REBOOT") = True Then
        For Each site In TheExec.sites
            If g_Reboot_Flag = True Then
                Exit For
            ElseIf g_Reboot_Flag = False Then
                If TResult_selsrm(site) * TResult_coreup(site) * CmdListStatus(site) <> 1 Then
                    g_Reboot_Flag = True
                                        Exit For
                End If
            Else
            End If
    Next site
    End If
        
    RTOS_UART_Print instancename, CmdListStatus
    Dim i As Integer
    Dim Loop_Cnt As Long
    Loop_Cnt = 1    'Loop count control by customer
    'Add for CSW measurement
    If CSW_MeasPin <> "" Then
        Dim CSW_PinAry() As String
        Dim CSW_PinCnt As Long
        Dim RestoreAry() As Double
        Dim CSW_Data As New PinListData
        Dim CSW_SampleSize As Long
        Dim InstName As String
        Dim tmp_pinIdx As Long
        CSW_SampleSize = 64
        TheExec.DataManager.DecomposePinList CSW_MeasPin, CSW_PinAry, CSW_PinCnt
        ReDim RestoreAry(CSW_PinCnt - 1) As Double
        
        For tmp_pinIdx = 0 To CSW_PinCnt - 1
            InstName = GetInstrument(CSW_PinAry(tmp_pinIdx), 0)
            Select Case InstName
            Case "DC-07"
                TheExec.Datalog.WriteComment "The Instrument " & InstName & " is nto support CSW!!!"
            Case "VHDVS", "HexVS", "VSM", "VS-5A", "VS-800mA"
                If InstName = "VHDVS" Then CSW_SampleSize = 1
                RestoreAry(tmp_pinIdx) = TheHdw.DCVS.pins(CSW_PinAry(tmp_pinIdx)).CurrentRange
                DCVS_MeasureCurrent_AutoRange CSW_PinAry(tmp_pinIdx), True
                For i = 1 To Loop_Cnt
                    CSW_Data = TheHdw.DCVS.pins(CSW_MeasPin).Meter.Read(tlStrobe, CSW_SampleSize)
                    If TheExec.DevChar.Setups.IsRunning = True Then
                        TheExec.flow.TestLimit CSW_Data, , , , , , unitAmp, , TnameCombShmooInfo & "_CSW"
    '                    TheExec.Flow.TestLimit CSW_Data, , , , , , unitAmp, , TnameCombShmooInfo & "_CSW", , , Format(TheHdw.DCVS.Pins("vdd_cpu").Voltage.value, "0.000"), unitVolt
                    Else
                        TheExec.flow.TestLimit CSW_Data, , , , , , unitAmp, , RTOSTest_Inst & "_CSW"
    '                    TheExec.Flow.TestLimit CSW_Data, , , , , , unitAmp, , RTOSTest_Inst & "_CSW", , , Format(TheHdw.DCVS.pins("vdd_cpu").Voltage.value, "0.000"), unitVolt
                    End If
                    TheHdw.Wait 0.01    ' Need to adjust
                Next i
                TheHdw.DCVS.pins(CSW_PinAry(i)).SetCurrentRanges RestoreAry(i), RestoreAry(i)
            Case "HSD-U"
                TheExec.Datalog.WriteComment "The Instrument " & InstName & " is nto support CSW!!!"
            Case Else
            End Select
        Next tmp_pinIdx
    Else
    End If
    
    'Print shmoo voltage information''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If TheExec.DevChar.Setups.IsRunning = True Then '#16_ELSE_CASE_CHK
        
        Dim active_setup As String, curr_axis As Variant, curr_track As Variant, apply_Pin As String, apply_Pin_arry() As String, pin_count1 As Long
        active_setup = TheExec.DevChar.Setups.ActiveSetupName
        '===========================================================
        For Each curr_axis In TheExec.DevChar.Setups(active_setup).Shmoo.axes.list
            ''exit for if any axis is not power pin -by SY
            If TheExec.DevChar.Setups(active_setup).Shmoo.axes(curr_axis).ApplyTo.pins <> "" Then   '#16_ELSE_CASE_CHK
                apply_Pin = TheExec.DevChar.Setups(active_setup).Shmoo.axes(curr_axis).ApplyTo.pins
        '        Add for store shmoo global spec to avoid direct to apply Vmain used for Vbump function
                Call TheExec.DataManager.DecomposePinList(apply_Pin, apply_Pin_arry, pin_count1)
                For i = 0 To pin_count1 - 1
                                        If gl_GetInstrument_Dic.Exists(LCase(apply_Pin_arry(i))) Then _
                    TheExec.Datalog.WriteComment "Scenario voltage, " & apply_Pin_arry(i) & "                       " & TheHdw.DCVS.pins(apply_Pin_arry(i)).Voltage.value & "V"
                Next i
                For Each curr_track In TheExec.DevChar.Setups(active_setup).Shmoo.axes(curr_axis).TrackingParameters.list
                    apply_Pin = TheExec.DevChar.Setups(active_setup).Shmoo.axes(curr_axis).TrackingParameters.item(curr_track).ApplyTo.pins
                    Call TheExec.DataManager.DecomposePinList(apply_Pin, apply_Pin_arry, pin_count1)
                    For i = 0 To pin_count1 - 1
                                                If gl_GetInstrument_Dic.Exists(LCase(apply_Pin_arry(i))) Then _
                        TheExec.Datalog.WriteComment "Scenario voltage, " & apply_Pin_arry(i) & "                       " & TheHdw.DCVS.pins(apply_Pin_arry(i)).Voltage.value & "V"
                    Next i
                Next curr_track
            Else
            End If
        Next curr_axis
    Else
    End If
        
        Call Interpose(EndOfBodyF, EndOfBodyFArgs)      ' Run EndOfBody interpose function, if specified
    
    End If ' Body
    
    ' PostBody
    If Step_ = subAllBody Or Step_ = subPostbody Then
        Call PostBody(DriveHiPins, DriveLoPins, DriveZPins, FloatPins)
        g_Vbump_function = False
        'g_Reboot_Flag = False
    
        TheExec.Datalog.Setup.Shared.ascii.Columns.EnableCustomWidths = True
        TheExec.Datalog.Setup.Shared.ascii.Columns.Parametric.TestName.Width = 70
        TheExec.Datalog.ApplySetup
        
    End If ' PostBody
            
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_RTOS", "RTOS_RunScenario_T")
    If AbortTest Then Exit Function Else Resume Next
End Function


' ===============
' Private Helpers
' ===============

' This test needs to know timing and levels sheet names.
' Fetch them from the Context Manager
Private Sub FetchContext()
On Error GoTo errHandler
    Dim a(0 To 4) As String

    ' For compatibility with 7.01.01 and earlier:
    ' In earlier versions, a contextmgr bug made using a MemberIndex > 0 act like the CurrentlyAppliedContext parameter was False.
    ' This caused "" to be returned for the output parameters...so that ApplyLevelsTiming was NOT called for 2nd & later members of a test group
    
    Dim MemberIndex As Long
    MemberIndex = TheExec.DataManager.MemberIndex
    
    Dim UseCurrentContext As Boolean
    UseCurrentContext = (MemberIndex = 0)
    
    Call m_STDSvcClient.Dmgr.ContextMgr.GetInstanceContextInformation(TheExec.DataManager.instancename, MemberIndex, _
                a(0), a(1), m_TimeSetSheet, a(2), a(3), a(4), m_LevelsSheet, True, UseCurrentContext)

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_RTOS", "FetchContext")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

' =====================
' Private work routines
' =====================

' Set up the test by applying digital timing and voltage levels and
' establishing initial driver states in preparation for executing a pattern.
Private Sub PreBody(DriveHiPins As PinList, DriveLoPins As PinList, DriveZPins As PinList, _
                    DisablePins As PinList, Util0Pins As PinList, Util1Pins As PinList, _
                    RepeatableSequence As Boolean)
On Error GoTo errHandler

    Dim ConnectAllPins As Boolean, LoadLevels As Boolean, LoadTiming As Boolean
    Dim RelayMode As tlRelayMode
    ' Close Pin-Electronics, High-Voltage, & Power Supply Relays,
    '   of pins noted on the active Pin Levels sheet
    'Call TheHdw.PinLevels.ConnectAllPins
    
    ' Set drive state on specified utility bits.
    If NonBlank(Util0Pins) Then Call tl_SetUtilState(Util0Pins, 0)
    If NonBlank(Util1Pins) Then Call tl_SetUtilState(Util1Pins, 1)
    
    ' Apply voltage levels defined on the Pin Levels sheet
    If NonBlank(m_LevelsSheet) Then LoadLevels = True

    ' Load digital timing values into the hardware
    If NonBlank(m_TimeSetSheet) Then LoadTiming = True

    ' Close Pin-Electronics, High-Voltage, & Power Supply Relays,
    '   of pins noted on the active levels sheet, if needed
    ConnectAllPins = True


      ' ApplyLevelTiming will
    '   Not power down instruments and power supplies
    '   Close Pin-Electronics, High-Voltage, & Power Supply Relays,
    '       of pins noted on the active levels sheet
    '   Optionally load Timing and Levels information
    '   Set init-state driver conditions on specified pins
    '       Setting init state causes the pin to drive the specified value.  Init
    '       state is set once, during the prebody, before the first pattern burst.
    '       Default is to leave the pin driving whatever value it last drove during
    '       the previous pattern burst.
    If (RepeatableSequence) Then
        Call TheHdw.PinLevels.ConnectAllPins
        
        TheHdw.SettleWait 1#
        
        If LoadLevels Or LoadTiming Then
            Call TheHdw.LevelsAndTiming.ApplyRepeatableSequence
        End If
    Else
        Call TheHdw.Digital.ApplyLevelsTiming(ConnectAllPins, LoadLevels, LoadTiming, tlPowered)
    End If
    
    If NonBlank(DriveLoPins) Then Call tl_SetInitState(DriveLoPins, chInitLo)
    If NonBlank(DriveHiPins) Then Call tl_SetInitState(DriveHiPins, chInitHi)
    If NonBlank(DriveZPins) Then Call tl_SetInitState(DriveZPins, chInitoff)

    ' Set start-state driver conditions on specified pins.
    ' Start state determines the driver value the pin is set to as each pattern burst starts.
    ' Default is to have start state automatically selected appropriately
    '   depending on the Format of the first vector of each pattern burst.
    If NonBlank(DriveLoPins) Then Call tl_SetStartState(DriveLoPins, chStartLo)
    If NonBlank(DriveHiPins) Then Call tl_SetStartState(DriveHiPins, chStartHi)
    If NonBlank(DriveZPins) Then Call tl_SetStartState(DriveZPins, chStartOff)

    If NonBlank(DisablePins) Then Call tl_SetDisableState(DisablePins)

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_RTOS", "PreBody")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

' Perform the test
Private Sub Body(FloatPins As PinList)
On Error GoTo errHandler
    ' Disconnect specified DUT pins from tester pin-electronics and other resources
    If NonBlank(FloatPins) Then Call tl_SetFloatState(FloatPins)

    ' Test logic goes here....
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_RTOS", "Body")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

' Clean up
Private Sub PostBody(DriveHiPins As PinList, DriveLoPins As PinList, DriveZPins As PinList, FloatPins As PinList)
On Error GoTo errHandler
    Dim DriverNonePins As String
    
    ' Clear previously registered interpose function names
    Call tl_ClearInterpose(TL_C_PREPATF, TL_C_POSTPATF, TL_C_PRETESTF, TL_C_POSTTESTF)
    m_InterposeFunctionsSet = False

    ' Return channels to the default start-state condition, as needed
    DriverNonePins = tl_tm_CombineCslStrings(DriveHiPins, DriveLoPins)
    DriverNonePins = tl_tm_CombineCslStrings(DriveZPins, DriverNonePins)
    If NonBlank(DriverNonePins) Then Call tl_SetStartState(DriverNonePins, chstartNone)

    ' Return specified DUT pins, if any, to connection with tester pin-electronics & power
    If NonBlank(FloatPins) Then Call tl_ConnectTester(FloatPins)

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_RTOS", "PostBody")
    If AbortTest Then Exit Sub Else Resume Next
End Sub


Public Function RTOS_Command_UID_EFUSE(Optional Cmd1 As String, Optional Cmd1TimeOut As Double = 0#, Optional PwrPin As String, Optional vpwr As Double)
 
On Error GoTo errHandler
    ReDim GlobalMergeAry(TheExec.sites.Existing.Count - 1) 'for txt data collection
    Dim CmdList As Variant 'String
    Dim CmdListStatus As New SiteLong
    Dim powerPin As String
    Dim instancename As String: instancename = TheExec.DataManager.instancename
    Dim CMDTotalTT As Double
 
        '' ====== clchend 2022 0316
    Dim m_timer1 As Double
    m_timer1 = TheExec.Timer(0)
    '' ============================
    If Cmd1 <> "" Then CmdList = Cmd1
    CMDTotalTT = Cmd1TimeOut
 
    'Scenario Run Conditions
    CmdListStatus = 0
    TheExec.Datalog.DatalogSuspended = False
    If (PwrPin <> "") Then Call TurnOnEfusePwrPins(PwrPin, vpwr)     '' efuse power on
    If CmdList <> "" Then Set CmdListStatus = SendCmd(CmdList, CMDTotalTT, False)
    If (PwrPin <> "") Then Call TurnOffEfusePwrPins(PwrPin, vpwr)    '' efuse power off
 
    TheExec.flow.TestLimit CmdListStatus, 1, 1       ', , , , , , TestName
    RTOS_UART_Print instancename, CmdListStatus
    TheExec.Datalog.WriteComment "@@@@@ RTOS_Command_UID_EFUSE   Finish  Time: " & TheExec.Timer(m_timer1)
    Exit Function

'''    By Alan 2022/5/18 email
'''    General Notes:
'''1. Current RTOS tests are boot in FUNC mode.
'''2. UID Fusing requires the DUT to boot in UID mode.
'''3. All existing test scenarios expected work in UID mode as well.  Intent is to run all RTOS production scenarios in UID mode.
'''4. UID mode only works when chip is in CFG A00.  Therefore, CFG fusing must happen after UID Fusing.
'''5. Once UID is fused, it would not be possible to boot into UID mode again.  For T0TX or RMA, we will need to boot in FUNC mode.
'''
'''Setup:
'''1. Board_ID<x> pins need to set to "0".  Current RTOS test configure them to "1"
'''2. Pin CLKREQ = "0" -> UID mode.  CLKREQ = "1" -> FUNC mode
'''
'''Instructions:
'''1. Apply bootstrap pin settings as listed in Setup section.
'''2. Run RTOS boot using the attached binary.  If boot successful, you can confirm UID mode by the following:
'''Board ID: 0x00 BootConfig: 0x0 BootMode: UID_MODE
'''3. Run command: "prov uid init"
'''4. Upon completion of step 3, set VDD12_EFUSE2 = 1.8V (prepare for UID Fusing)
'''5. Run command: "prov uid fuse"
'''6. Upon completion of step 5, set VDD12_EFUSE2 = 0.0V
'''7. Proceed with other scenario tests.
''''by Mark 2022/11/18 email
''''From: Mark Jackson <mjackson@apple.com>
''''Sent: Friday, November 18, 2022 7:37 PM
''''Please remember, we must use "prov uid fuse sec" command for lead lot and production screens at FT.
''''The "sec" parameter is now required for production material.
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function
