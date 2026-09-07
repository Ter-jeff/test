Attribute VB_Name = "VBT_LIB_Digital_Mbist"
Option Explicit
'Revision History:
'V0.0 initial bring up
Public MemArray_SocMbist_srvA(10000) As Long ''for Soc mbist RSCR
Public MemStrArray_SocMbist_srvA(10000) As String ''for Soc mbist RSCR
Public MemArray_SocMbist_srvB(10000) As Long ''for Soc mbist RSCR
Public MemStrArray_SocMbist_srvB(10000) As String ''for Soc mbist RSCR

Public MemArray_CpuMbist(10000) As Long ''for Cpu mbist RSCR
Public MemStrArray_CpuMbist(10000) As String ''for Cpu mbist RSCR

Public MemArray_GpuMbist(10000) As Long ''for Gpu mbist RSCR
Public MemStrArray_GpuMbist(10000) As String ''for Gpu mbist RSCR

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

Public BlockSelectLoopCount As Long
Public Function Init_RSCR()
    'soc rscr table read in
    Dim k As Integer
    
    If Flag_RSCR_INIT = False Then
        For k = 0 To 10000
        'CHWu modify 10/14 to remove useless RSCR file
            'MemArray_SocMbist_srvA(k) = Worksheets("SocRSCR_srvA").Cells(k + 1, 4).Value
            'MemStrArray_SocMbist_srvA(k) = Worksheets("SocRSCR_srvA").Cells(k + 1, 8).Value
            'MemArray_SocMbist_srvB(k) = Worksheets("SocRSCR_srvB").Cells(k + 1, 4).Value
            'MemStrArray_SocMbist_srvB(k) = Worksheets("SocRSCR_srvB").Cells(k + 1, 8).Value
            MemArray_CpuMbist(k) = Worksheets("CpuRSCR").Cells(k + 1, 4).Value
            MemStrArray_CpuMbist(k) = Worksheets("CpuRSCR").Cells(k + 1, 8).Value
            'MemArray_GpuMbist(k) = Worksheets("GpuRSCR").Cells(k + 1, 4).Value
            'MemStrArray_GpuMbist(k) = Worksheets("GpuRSCR").Cells(k + 1, 8).Value
        Next k
        'TheExec.Datalog.WriteComment RepeatChr("*", 120)
        TheExec.Datalog.WriteComment "print: RSCR table initialized complete"
    End If
    Flag_RSCR_INIT = True
End Function

Public Function Mbist_RSCR(Shift_Pat As Pattern)

    Dim SampleNum As Integer
    Dim CNumber_plus As Integer
    Dim testS As String
    Dim testS1 As String
    Dim full_str As String
    Dim BISTData(119) As Double
    Dim Mbist_repair_cycle As Long
    Dim capt As CaptType
    Dim numcap As New SiteLong
    Dim pre_trig As Long
    Dim PatData As New PinListData
    Dim PinData As New PinListData
    Dim PinPF As New PinListData
    Dim Failed_Pins() As String
    Dim maxDepth As Integer
    Dim HRAM_PFVar As Variant
    Dim HRAM_EXPECTVar As Variant
    Dim HRAM_DUTVar As Variant
    Dim RVal As New SiteDouble
    Dim file_name As String
    pre_trig = 0
    Dim k As Long
    Dim TestPatName As String, rtnPatternNames() As String, rtnPatternCount As Long
    Dim patt As Variant
    Dim kk  As Long
    Dim sne_str As String
    Dim patGup As String
    Dim mem_location As String, i As Long
    Dim AllSitePass As Boolean
    Dim BurstResult As New SiteLong
    Dim Site As Variant
    Dim InstanceName As String
    
    SampleNum = 120
    CNumber_plus = 0 ' pattern has dummy cycle
    ''''''''''''''''''
    InstanceName = LCase(TheExec.DataManager.InstanceName)
    
    TheHdw.Digital.Patgen.HaltMode = tlHaltOnHRAMFull
    maxDepth = TheHdw.Digital.HRAM.maxDepth
    TheHdw.Digital.HRAM.Size = maxDepth
    TheHdw.Digital.HRAM.CaptureType = captFail
    TheHdw.Digital.HRAM.SetTrigger trigFail, False, 0

    
    TheExec.Datalog.WriteComment "Mbist repair information shift start"
    
    GetPatListFromPatternSet Shift_Pat.Value, rtnPatternNames, rtnPatternCount

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered 'SEC DRAM

    'RSCR site loop
    For Each Site In TheExec.Sites
    
       For Each patt In rtnPatternNames
            TheHdw.Patterns(Shift_Pat).Load
            TheHdw.Patterns(Shift_Pat).start ""
            TheHdw.Digital.Patgen.HaltWait
            'TheExec.Datalog.WriteComment patt
       Next patt
         numcap(Site) = TheHdw.Digital.HRAM.CapturedCycles
    
        If numcap(Site) = 0 Then
            TheExec.Datalog.WriteComment ("Site " & Site & "," & "Fail cycle at: NA" & ", Prime")
        Else
        
            RVal(Site) = TheHdw.Digital.HRAM.PatGenInfo(numcap(Site) - 1, pgCycle)
        '    For Each Site In TheExec.Sites
            For i = 0 To numcap(Site) - 1
        '        PatData = TheHdw.Digital.pins(FailPin_str).HRAM.PatData(0, 1, numcap)
                PinData = TheHdw.Digital.Pins("JTAG_TDO").HRAM.PinData(0, 1, numcap(Site))
        '        PinPF = TheHdw.Digital.pins(FailPin_str).HRAM.PinPF(0, 1, numcap)
        
        
                '//MEMORY_CL52  cycle 2421 to Cycle 3140
                '//MEMORY_CL51  cycle cycle 3141 to Cycle 7620
                '//MEMORY_CL27 cycle 8901 to cycle 9299
                '//MEMORY_CL26 cycle 10001 to cycle 10399
                '//MEMORY_CL17 cycle 11741 to cycle 12139
                '//MEMORY_CL16 cycle 12841 to cycle 13239
                'Mbist_repair_cycle
                Mbist_repair_cycle = TheHdw.Digital.HRAM.PatGenInfo(i, pgCycle)
                'Mbist_repair_cycle = Mbist_repair_cycle + 1 'no shift
                
                mem_location = "Not Match"
                
                'Array selection
                If InstanceName Like "socmbist*srva*" Then
                    For k = 0 To 10000
                        If Mbist_repair_cycle = MemArray_SocMbist_srvA(k) Then
                            mem_location = MemStrArray_SocMbist_srvA(k)
                        End If
                    Next k
                ElseIf InstanceName Like "socmbist*srvb*" Then
                    For k = 0 To 10000
                        If Mbist_repair_cycle = MemArray_SocMbist_srvB(k) Then
                            mem_location = MemStrArray_SocMbist_srvB(k)
                        End If
                    Next k
                ElseIf InstanceName Like "cpubist*" Then
                    For k = 0 To 10000
                        If Mbist_repair_cycle = MemArray_CpuMbist(k) Then
                            mem_location = MemStrArray_CpuMbist(k)
                        End If
                    Next k
                ElseIf InstanceName Like "gfxmbist*" Or InstanceName Like "gpumbist*" Then
                    Mbist_repair_cycle = Mbist_repair_cycle '+ 1 'GPU shift 1 cycle
                    For k = 0 To 10000
                        If Mbist_repair_cycle = MemArray_GpuMbist(k) Then
                            mem_location = MemStrArray_GpuMbist(k)
                        End If
                    Next k
                End If
           
             TheExec.Datalog.WriteComment ("Site " & Site & "," & "Fail cycle at: " & Mbist_repair_cycle & ",Mem Location: " & mem_location) ' first fail cycle
            'theexec.Datalog.WriteComment ("Site " & Site & "," & "Fail cycle at: " & TheHdw.Digital.HRAM.PatGenInfo(i, pgCycle) + 1 & ",Mem:" & mem_location) ' first fail cycle
            Next i
        End If
        

    Next Site
    
    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode 'recover halt on opcode

    TheExec.Datalog.WriteComment "Mbist repair information shift end"

    DebugPrintFunc Shift_Pat.Value

End Function



Public Function PrintToDatlog(PrintInfo As String)
    Dim Site As Variant

For Each Site In TheExec.Sites
    TheExec.Datalog.WriteComment "****************** Site(" & Site & "), " & PrintInfo & "****************"
Next Site

End Function


Public Function TurnOnEfusePwrPins_MBIST(FusePower As String)

    'Escalate VDD18_EFUSE0 and VDD18_EFUSE1 according to Fiji
    'test plan (slower than 1.8v/30us)
    
    DCVS_PowerOn_I_Meter FusePower, 1.8, 0.2, 0.001, 0.002, 10, 0.018   'use 18 ms to power up

End Function


Public Function TurnOffEfusePwrPins_MBIST(FusePower As String)

    'Decline VDD18_EFUSE0 and VDD18_EFUSE1 according to Fiji
    'test plan (slower than 1.8v/30us)
    
    Dim CurrentVoltage As Double
    
    CurrentVoltage = TheHdw.DCVS.Pins(FusePower).Voltage.Main.Value
    DCVS_PowerOff_I_Meter FusePower, CurrentVoltage, 0.2, 0.001, 0.002, 10, 0.018   'use 18 ms to power down


End Function


Public Function BlockSelect()

'For Rhea, CpuMbist has 12 block, use loop and enable word to production
    Dim i As Long
    Dim EnableStr As String
    Dim EnableExStr As String
    EnableStr = ""
    EnableExStr = ""
    
    For i = 0 To 11 'enable word initial

        If i < 10 Then
            EnableStr = "Enable_B0" & i
        Else
            EnableStr = "Enable_B" & i
        End If

        TheExec.Flow.EnableWord(EnableStr) = False
        TheExec.Datalog.WriteComment "print: Turn off EnableStr= " & EnableStr
    Next i
    
    'If TheExec.Flow.EnableWord("Enable_All_BXX") = True Then BlockSelectLoopCount = 0    'initial
        
    For i = 0 To 11
        If BlockSelectLoopCount = i Then
            If i < 10 Then
                EnableStr = "Enable_B0" & i
''                If i = 0 Then
''                    EnableExStr = ""
''                Else
''                    EnableExStr = "Enable_B0" & i - 1
''                End If
            Else
                EnableStr = "Enable_B" & i
''                If i < 11 Then
''                    EnableExStr = "Enable_B0" & i - 1
''                Else
''''                    If i = 0 Then
''''                        EnableExStr = ""
''''                    Else
''                    EnableExStr = "Enable_B" & i - 1
''''                    End If
''                End If
                    
            End If
        End If
    Next i

    'TheExec.Flow.EnableWord("Enable_All_BXX") = False
    TheExec.Flow.EnableWord(EnableStr) = True
    'TheExec.Flow.EnableWord(EnableExStr) = False
    
    TheExec.Datalog.WriteComment "print: Turn on EnableStr= " & EnableStr
    'TheExec.Datalog.WriteComment "print: Turn on EnableStr= " & EnableStr &     ' " , Turn off EnableExStr=" & EnableExStr
    BlockSelectLoopCount = BlockSelectLoopCount + 1

End Function
Public Function BlockSelectLoopCount_Initial()
'first run before loop

    Dim i As Long
    Dim EnableStr As String
    'Dim EnableExStr As String
    EnableStr = ""
    'EnableExStr = ""
    
    BlockSelectLoopCount = 0    'initial
        
''    For i = 0 To 11
''
''        If i < 10 Then
''            EnableStr = "Enable_B0" & i
''        Else
''            EnableStr = "Enable_B" & i
''        End If
''
''        TheExec.Flow.EnableWord(EnableStr) = False
''        TheExec.Datalog.WriteComment "print: Turn off EnableStr= " & EnableStr
''    Next i


End Function

Public Function MbistRetentionLevelWait(mS_Time As Double)
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered 'SEC DRAM
    TheHdw.Wait mS_Time * ms
    'TheExec.Datalog.WriteComment "*************************************************"
    'TheExec.Datalog.WriteComment "*print: MbistRetention wait " & mS_Time & " ms*"
    Call Print_Retention_Power
    'TheExec.Datalog.WriteComment "*************************************************"
    TheExec.Flow.TestLimit mS_Time, 0, 9999, tlSignGreaterEqual, tlSignLessEqual, ScaleType:=scaleNone, unit:=unitCustom, Tname:="MbistRet", customUnit:="ms"   'BurstResult=1:Pass
    
    
End Function
'
'        TheExec.Datalog.WriteComment outstr
'
'End Function
'
Public Function Print_Retention_Power()
        Dim outstr As String
        Dim vdd_cpu_val  As Double
        Dim vdd_gpu_val  As Double
        Dim vdd_cpu_sram_val  As Double
        Dim vdd_gpu_sram_val  As Double
        Dim vdd_sram_val As Double
        Dim vdd_soc_val  As Double
        Dim vdd_fixed_val  As Double
        Dim vdd_low_val As Double
        Dim VDD_Bin_name As String
        Dim siteloop_bypass As Boolean
        Dim Site As Variant
       ' VDD_Bin_name = VddBinName(IndexBinningFlow)
        
            outstr = ""
            vdd_cpu_val = TheHdw.DCVS.Pins("VDD_CPU").Voltage.Main
            vdd_gpu_val = TheHdw.DCVS.Pins("VDD_GPU").Voltage.Main
            vdd_soc_val = TheHdw.DCVS.Pins("VDD_SOC").Voltage.Main
            vdd_cpu_sram_val = TheHdw.DCVS.Pins("VDD_CPU_SRAM").Voltage.Main
            vdd_gpu_sram_val = TheHdw.DCVS.Pins("VDD_GPU_SRAM").Voltage.Main
            vdd_fixed_val = TheHdw.DCVS.Pins("VDD_FIXED").Voltage.Main
            vdd_low_val = TheHdw.DCVS.Pins("VDD_LOW").Voltage.Main
''            If CurrentJobName Like "*cp*" Then
''                vdd_cpu_sram_val = TheHdw.DCVS.Pins("VDD_CPU_SRAM").Voltage.Main
''                vdd_gpu_sram_val = TheHdw.DCVS.Pins("VDD_GPU_SRAM").Voltage.Main
''                vdd_fixed_val = TheHdw.DCVS.Pins("VDD_FIXED").Voltage.Main
''                vdd_low_val = TheHdw.DCVS.Pins("VDD_LOW").Voltage.Main
''            Else
''                vdd_cpu_sram_val = TheHdw.DCVS.Pins("VDD_SRAM").Voltage.Main
''                vdd_gpu_sram_val = TheHdw.DCVS.Pins("VDD_SRAM").Voltage.Main
''                vdd_sram_val = TheHdw.DCVS.Pins("VDD_SRAM").Voltage.Main
''            End If
            
        For Each Site In TheExec.Sites
''            If CurrentJobName Like "*cp*" Then
                outstr = "Site" & ":" & Site & ",VDD_CPU=" & Format(vdd_cpu_val, "0.000") & "," & _
                "VDD_GPU=" & Format(vdd_gpu_val, "0.000") & "," & _
                "VDD_SOC=" & Format(vdd_soc_val, "0.000") & "," & _
                "VDD_CPU_SRAM=" & Format(vdd_cpu_sram_val, "0.000") & "," & _
                "VDD_GPU_SRAM=" & Format(vdd_gpu_sram_val, "0.000") & "," & _
                "VDD_FIXED=" & Format(vdd_fixed_val, "0.000") & "," & _
                "VDD_Low=" & Format(vdd_low_val, "0.000")
''            Else
''                outstr = VDD_Bin_name & "," & Site & ",VDD_CPU=" & Format(vdd_cpu_val, "0.000") & "," & _
''                "VDD_GPU=" & Format(vdd_gpu_val, "0.000") & "," & _
''                "VDD_SOC=" & Format(vdd_soc_val, "0.000") & "," & _
''                "VDD_SRAM=" & Format(vdd_sram_val, "0.000")
''            End If
            TheExec.Datalog.WriteComment outstr
       Next Site
''      DCVS_Mbist_PowerUp "All_HEXVS"
''      DCVS_Mbist_PowerUp "All_IDS_UVS_HI"

End Function

Public Function DCVS_Mbist_PowerUp(PowerPinList As String, Optional WaitConnectTime As Double = 0.001, Optional DebugFlag As Boolean = True)
'power up sequence at flow start
    Dim CurrentChans As String
    Dim Site As Variant
    Dim Pins() As String, PinCnt As Long
    Dim PowerPin As Variant
    Dim PowerName As String
    Dim TempString As String
    Dim Vmain As Double
    Dim Irange As Double
    Dim Step As Integer
    Dim RiseTime As Double
    Dim PowerSequence As Double
    Dim i As Long
    Dim CurrentVolt As Double
    
    On Error GoTo errHandler
    CurrentVolt = 0
    'CurrentChans = TheExec.CurrentChanMap 'obtain FT or CP channel map information

    TheExec.Datalog.WriteComment vbCrLf & "print: Power up start, Power pins: " & PowerPinList
    
    'TheHdw.DCVS.Pins(PowerPinList).Voltage.Main = 0  'reset to 0V
        
    'TheHdw.Digital.Pins(DisconnectPinList).Disconnect
    'TheExec.Datalog.WriteComment "print: Power up digital disconnect, Digital pins: " & DisconnectPinList
    
    TheExec.DataManager.DecomposePinList PowerPinList, Pins(), PinCnt
        
    For i = 0 To PinCnt - 1
        For Each PowerPin In Pins
            TempString = ""
            PowerName = CStr(PowerPin)
            
            'get power sequence global spec
            TempString = PowerName & "_PowerSequence_GLB"
            PowerSequence = TheExec.specs.Globals(TempString).ContextValue
            CurrentVolt = Format(TheHdw.DCVS.Pins(PowerPin).Voltage.Main, "0.00")
            If PowerSequence = i Then
                If TheExec.DataManager.ChannelType(PowerPin) <> "N/C" Then 'check CP for FT form NC pins
                    'get nomial voltage spec value
                    TempString = PowerName & "_GLB"
                    Vmain = TheExec.specs.Globals(TempString).ContextValue
                    
                    'get Ifold limit spec value
                    TempString = PowerName & "_Ifold_GLB"
                    Irange = TheExec.specs.Globals(TempString).ContextValue
                    
                    'auto calculate steps
                    Step = (Vmain - CurrentVolt) / 0.1 '0.1v per step
                    If Step = 0 Then Step = 10  'default value
                    RiseTime = Step * ms '1ms per step

                    DCVS_PowerOn_I_Meter PowerName, Vmain, Irange, WaitConnectTime, WaitConnectTime, Step, RiseTime
                    
                    If DebugFlag = True Then    'debugprint
                        TheExec.Datalog.WriteComment "print: Pin " & PowerName & ", Vmain " & Format(Vmain, "0.00") & ", Irange " & Irange & ", Step " & Step & ", RiseTime " & RiseTime * 1000 & " ms" & ", PowerSequence " & PowerSequence
                    End If
                Else
                    If DebugFlag = True Then    'debugprint
                        TheExec.Datalog.WriteComment "print: Pin " & PowerName & " not turn on by 'NC pin', PowerSequence " & PowerSequence & " ,Warning!!!"
                    End If
                End If
            'sequence 99, means disconnect pins
            ElseIf i = (PinCnt - 1) And PowerSequence = 99 Then
                TheHdw.DCVS.Pins(PowerPin).Disconnect
                If DebugFlag = True Then    'debugprint
                    TheExec.Datalog.WriteComment "print: Pin " & PowerName & " is 'NA' pin, disconnect, PowerSequence " & PowerSequence
                End If
            End If
        Next PowerPin
    Next i
    
    TheExec.Datalog.WriteComment "print: Power up finished" & vbCrLf
    
''    TheExec.Datalog.WriteComment "print: disconnect power monitor digital pins 'All_Power_Monitor_Digital'"
''    TheHdw.Digital.Pins("All_Power_Monitor_Digital").Disconnect
    Exit Function
    
errHandler:
        ErrorDescription ("DCVS_PowerUp")
        If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function auto_Mbist_Initialize()
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_Mbist_Initialize"

    ''''Initialize, 20151111 New
    gB_BIRA_MC000_NRS_flag = False
    gB_BIRA_MC010_NRS_flag = False
    gB_BIRA_MC051_NRS_flag = False
    gB_BIRA_MC051_flag = False
    gB_BIRA_MC000_flag = False
    
    ''Dim sheetName As String
    If (gL_1st_MbistSheetRead = 0) Then
        Call auto_parse_MBIST_ChkList("CPU")
        TheExec.Datalog.WriteComment funcName + " :: CpuMbist :: " + FormatNumeric(gS_CpuMbist_sheetName, -30)
        
        ''Call auto_parse_MBIST_ChkList("GPU")
        ''TheExec.Datalog.WriteComment funcName + " :: GpuMbist :: " + FormatNumeric(gS_GpuMbist_sheetName, -30)
        
        ''Call auto_parse_MBIST_ChkList("SOC")
        ''TheExec.Datalog.WriteComment funcName + " :: SocMbist :: " + FormatNumeric(gS_SocMbist_sheetName, -30)
        ''''debug purpose if True
        
        '''----- CHWUD 110715 Show all Group information -----
        '''Call auto_MbistReadCategory
        
        gL_1st_MbistSheetRead = 1
    Else
        ''''debug purpose if True
        If (False) Then
            Call auto_MbistReadCategory
        End If
    End If
    
    Call UpdateDLogColumns(30)
    TheExec.Flow.TestLimit 1, 1, 1, , , , , , Tname:="Mbist_Initialize", PinName:="ChkList_Sheet"
    Call UpdateDLogColumns__False
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function
''''This Function is used to do the initial check roughly, not every parameters inside the Category.

''''===============================================================
''''CpuMbist Data Structure
''''===============================================================
''''CpuMbist_PMode.Category(i).Name
''''                          .PattName
''''
''''CpuMbist_Block.Category(i).Name
''''                          .PattName
''''---------------------------------------------------------------
''''---------------------------------------------------------------
''''CpuMbist.Category(i).Name  (GroupName)
''''                    ---------------------
''''                    .ChangeFlag
''''                    ---------------------
''''                    .PMode_RAW
''''                    .PMode(j).Name
''''                    .PMode(j).PattName
''''                    ---------------------
''''                    .Block_RAW
''''                    .Block(k).Name
''''                    .Block(k).PattName
''''                    ---------------------
''''                    .DebugMode(n).Name
''''                    .DebugMode(n).PMode
''''                    .DebugMode(n).PMode_PattName
''''                    .DebugMode(n).Block
''''                    .DebugMode(n).Block_PattName
''''---------------------------------------------------------------
''''CpuMbist_VoltSet.Category(i).Name
''''                            .DCCate
''''                            .DCSele
''''                            .ACCate
''''                            .ACSele
''''===============================================================
''''
Public Function auto_MbistReadCategory()

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_MbistReadCategory"
    
    Dim Site As Variant
    Dim m_idx As Integer

    Dim i As Long
    Dim j As Long

    Dim m_tmpStr As String
    Dim m_dlgStr As String
    Dim m_grpname As String
    Dim m_ChangeFlag As String
    Dim m_PMname As String
    Dim m_PM_pattname As String
    Dim m_PM_pattRawname As String
    Dim m_BMname As String
    Dim m_BM_pattname As String
    Dim m_BM_pattRawname As String
    
    Dim m_DMname As String
    
    Dim m_VMname As String
    Dim m_DCCate As String
    Dim m_DCSele As String
    Dim m_ACCate As String
    Dim m_ACSele As String

    If (gB_findCpuMbist_flag = True) Then
        PrintDataLog ""
        PrintDataLog "------------------------------------------------------------------------------------------"
        ''''Voltage Setting
        m_dlgStr = FormatNumeric("CpuMbist_VoltMode", -20) + ", " + _
                   FormatNumeric("DCCategory", -20) + ", " + FormatNumeric("DCSelector", -12) + ", " + _
                   FormatNumeric("ACCategory", -20) + ", " + FormatNumeric("ACSelector", -12)
        PrintDataLog m_dlgStr
        For i = 0 To UBound(CpuMbist_VoltSet.Category)
            m_VMname = CpuMbist_VoltSet.Category(i).Name
            m_DCCate = CpuMbist_VoltSet.Category(i).DCCate
            m_DCSele = CpuMbist_VoltSet.Category(i).DCSele
            m_ACCate = CpuMbist_VoltSet.Category(i).ACCate
            m_ACSele = CpuMbist_VoltSet.Category(i).ACSele
            m_dlgStr = FormatNumeric("   " + m_VMname, -20) + ", " + _
                       FormatNumeric(m_DCCate, -20) + ", " + FormatNumeric("  " + m_DCSele, -12) + ", " + _
                       FormatNumeric(m_ACCate, -20) + ", " + FormatNumeric("  " + m_ACSele, -12)
            PrintDataLog m_dlgStr
        Next i
        PrintDataLog "------------------------------------------------------------------------------------------"
        PrintDataLog ""
    
        m_dlgStr = FormatNumeric("CpuMbist_PMode", -15) + ", " + FormatNumeric("PMode_PattName", -30)
        PrintDataLog m_dlgStr
        For i = 0 To UBound(CpuMbist_PMode.Category)
            m_PMname = CpuMbist_PMode.Category(i).Name
            m_PM_pattname = CpuMbist_PMode.Category(i).PattName
            m_PM_pattRawname = CpuMbist_PMode.Category(i).PattRawName
            m_dlgStr = FormatNumeric("   " + m_PMname, -15) + ", " + FormatNumeric(m_PM_pattname, -30) + ", " + FormatNumeric(m_PM_pattRawname, -30)
            PrintDataLog m_dlgStr
        Next i
        PrintDataLog "------------------------------------------------------------------------------------------"
        PrintDataLog ""

        m_dlgStr = FormatNumeric("CpuMbist_Block", -15) + ", " + FormatNumeric("Block_PattName", -30)
        PrintDataLog m_dlgStr
        For i = 0 To UBound(CpuMbist_Block.Category)
            m_BMname = CpuMbist_Block.Category(i).Name
            m_BM_pattname = CpuMbist_Block.Category(i).PattName
            m_BM_pattRawname = CpuMbist_Block.Category(i).PattRawName
            m_dlgStr = FormatNumeric("   " + m_BMname, -15) + ", " + FormatNumeric(m_BM_pattname, -30) + ", " + FormatNumeric(m_BM_pattRawname, -30)
            PrintDataLog m_dlgStr
        Next i
        
        PrintDataLog "------------------------------------------------------------------------------------------"
        PrintDataLog ""
        
        For i = 0 To UBound(CpuMbist.Category)
            m_grpname = CpuMbist.Category(i).Name
            m_dlgStr = "CpuMbist GroupName(" + CStr(i) + ") :: "
            PrintDataLog m_dlgStr + m_grpname
            PrintDataLog Space(Len(m_dlgStr)) + "------------------------------------------------------------------------------------------"
            
            m_ChangeFlag = CpuMbist.Category(i).ChangeFlag
            PrintDataLog Space(Len(m_dlgStr)) + FormatNumeric("ChangeFlag", -15) + " = " + m_ChangeFlag
            PrintDataLog Space(Len(m_dlgStr)) + "------------------------------------------------------------------------------------------"
            
            m_tmpStr = CpuMbist.Category(i).PMode_RAW
            PrintDataLog Space(Len(m_dlgStr)) + FormatNumeric("PMode_RAW", -15) + " = " + m_tmpStr
            For j = 0 To UBound(CpuMbist.Category(i).PMode)
                m_PMname = CpuMbist.Category(i).PMode(j).Name
                m_PM_pattname = CpuMbist.Category(i).PMode(j).PattName
                m_PM_pattRawname = CpuMbist.Category(i).PMode(j).PattRawName
                m_tmpStr = ""
                m_tmpStr = Format(j, "00")
                m_tmpStr = FormatNumeric("PMode(" + m_tmpStr + ")", -15) + " = "
                m_tmpStr = m_tmpStr + FormatNumeric(m_PMname, -5) + ", " + FormatNumeric(m_PM_pattname, -30) + ", " + FormatNumeric(m_PM_pattRawname, -30)
                PrintDataLog Space(Len(m_dlgStr)) + m_tmpStr
            Next j
            PrintDataLog Space(Len(m_dlgStr)) + "------------------------------------------------------------------------------------------"

            m_tmpStr = CpuMbist.Category(i).Block_RAW
            PrintDataLog Space(Len(m_dlgStr)) + FormatNumeric("Block_RAW", -15) + " = " + m_tmpStr
            For j = 0 To UBound(CpuMbist.Category(i).Block)
                m_BMname = CpuMbist.Category(i).Block(j).Name
                m_BM_pattname = CpuMbist.Category(i).Block(j).PattName
                m_BM_pattRawname = CpuMbist.Category(i).Block(j).PattRawName
                m_tmpStr = ""
                m_tmpStr = Format(j, "00")
                m_tmpStr = FormatNumeric("Block(" + m_tmpStr + ")", -15) + " = "
                m_tmpStr = m_tmpStr + FormatNumeric(m_BMname, -5) + ", " + FormatNumeric(m_BM_pattname, -30) + ", " + FormatNumeric(m_BM_pattRawname, -30)
                PrintDataLog Space(Len(m_dlgStr)) + m_tmpStr
            Next j
            PrintDataLog Space(Len(m_dlgStr)) + "------------------------------------------------------------------------------------------"
            

            For j = 0 To UBound(CpuMbist.Category(i).DebugMode)
                m_DMname = CpuMbist.Category(i).DebugMode(j).Name
                ''PrintDataLog Space(Len(m_dlgStr)) + FormatNumeric("DebugMode", -15) + " = " + m_DMname
                
                m_PMname = CpuMbist.Category(i).DebugMode(j).PMode
                m_PM_pattname = CpuMbist.Category(i).DebugMode(j).PMode_PattName
                m_PM_pattRawname = CpuMbist.Category(i).DebugMode(j).PMode_PattRawName
                
                m_BMname = CpuMbist.Category(i).DebugMode(j).Block
                m_BM_pattname = CpuMbist.Category(i).DebugMode(j).Block_PattName
                m_BM_pattRawname = CpuMbist.Category(i).DebugMode(j).Block_PattRawName

                m_tmpStr = ""
                m_tmpStr = Format(j, "00")
                m_tmpStr = FormatNumeric("DebugMode(" + m_tmpStr + ")", -15) + " = " + m_DMname
                PrintDataLog Space(Len(m_dlgStr)) + m_tmpStr

                m_tmpStr = ""
                m_tmpStr = Format(j, "00")
                m_tmpStr = FormatNumeric("Debug_PMode(" + m_tmpStr + ")", -15) + " = "
                m_tmpStr = m_tmpStr + FormatNumeric(m_PMname, -5) + ", " + FormatNumeric(m_PM_pattname, -30) + ", " + FormatNumeric(m_PM_pattRawname, -30)
                PrintDataLog Space(Len(m_dlgStr)) + m_tmpStr

                m_tmpStr = ""
                m_tmpStr = Format(j, "00")
                m_tmpStr = FormatNumeric("Debug_Block(" + m_tmpStr + ")", -15) + " = "
                m_tmpStr = m_tmpStr + FormatNumeric(m_BMname, -5) + ", " + FormatNumeric(m_BM_pattname, -30) + ", " + FormatNumeric(m_BM_pattRawname, -30)
                PrintDataLog Space(Len(m_dlgStr)) + m_tmpStr
                
                If (j <> UBound(CpuMbist.Category(i).DebugMode)) Then PrintDataLog ""
            Next j
            PrintDataLog Space(Len(m_dlgStr)) + "------------------------------------------------------------------------------------------"
        Next i
        PrintDataLog ""

        If (gB_findPwrPin_flag) Then
            Dim m_pwrpin As String
            Dim m_vstart As Double
            Dim m_vstop As Double
            Dim m_vstep As Double
            Dim m_enable As String
            
            For i = 0 To UBound(CpuMbist_Power.Category)
                m_pwrpin = CpuMbist_Power.Category(i).PwrPin
                m_vstart = CpuMbist_Power.Category(i).V_Start
                m_vstop = CpuMbist_Power.Category(i).V_Stop
                m_vstep = CpuMbist_Power.Category(i).V_Step
                m_enable = CpuMbist_Power.Category(i).Enable
                m_tmpStr = "Power Pin " + FormatNumeric(m_pwrpin, 8) + ", Search From " + FormatNumeric(m_vstart, -6) + _
                                                        " To " + FormatNumeric(m_vstop, -6) + _
                                                        ", STEP " + FormatNumeric(m_vstep, -6) + _
                                                        ", Enable = " + FormatNumeric(m_enable, -5)
                PrintDataLog m_tmpStr
            Next i
            PrintDataLog "------------------------------------------------------------------------------------------"
            PrintDataLog ""
        End If



    End If
    

    If (gB_findGpuMbist_flag = True) Then
    End If
    
    If (gB_findSocMbist_flag = True) Then
    End If
    
    Call UpdateDLogColumns(30)
    TheExec.Flow.TestLimit 1, 1, 1, Tname:="MbistReadCategory"
    Call UpdateDLogColumns__False
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''PM: Performance Mode
''''BM: Block Mode
''''DM: Debug Mode
Public Function auto_Mbist_SetLoopCNT_DM_PM_BM(bistType As String) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_Mbist_SetLoopCNT_DM_PM_BM"
    
    Dim Site As Variant
    Dim m_grpidx As Long
    Dim m_tmpUB As Long
    Dim m_loopcnt_DM As Long ''''Debug       Mode LoopCount
    Dim m_loopcnt_PM As Long ''''Performance Mode LoopCount
    Dim m_loopcnt_BM As Long ''''Block       Mode LoopCount
    
    Dim m_LP_DM_siteVar As String
    Dim m_LP_PM_siteVar As String
    Dim m_LP_BM_siteVar As String

    m_LP_DM_siteVar = "LP_DM" ''''Debug       Mode LoopIndex
    m_LP_PM_siteVar = "LP_PM" ''''Performance Mode LoopIndex
    m_LP_BM_siteVar = "LP_BM" ''''Block       Mode LoopIndex

    Dim m_grpname As String
    Dim m_DMname As String
    Dim m_PMname As String
    Dim m_BMname As String
    
    Dim m_dlgStr As String
    Dim m_grpName_siteVar As String
    Dim m_DM_siteVar As String
    Dim m_PM_siteVar As String
    Dim m_BM_siteVar As String
    
    Dim m_debug_enableFlag As Boolean
    Dim m_pinName As String
    
    Dim m_lolmt As Variant
    Dim m_hilmtG As Variant
    Dim m_hilmtD As Variant
    Dim m_hilmtP As Variant
    Dim m_hilmtB As Variant
    
    ''''<MUST> 20151106, Initialize
    G_TestName = ""
    gB_enable_NewMbist_flag = True
    
    ''''Initialize, 20151111 New
    gB_BIRA_MC000_NRS_flag = False
    gB_BIRA_MC010_NRS_flag = False
    gB_BIRA_MC051_NRS_flag = False
    gB_BIRA_MC051_flag = False
    gB_BIRA_MC000_flag = False

    bistType = UCase(bistType)
    
    m_DM_siteVar = "LCount_DM"
    m_PM_siteVar = "LCount_PM"
    m_BM_siteVar = "LCount_BM"

    m_grpName_siteVar = "GroupName"
    
    ''''<MUST> Initialize after the GroupName is assigned
    For Each Site In TheExec.Sites.Existing
        TheExec.Sites(Site).SiteVariableValue(m_LP_DM_siteVar) = 0
        TheExec.Sites(Site).SiteVariableValue(m_LP_PM_siteVar) = 0
        TheExec.Sites(Site).SiteVariableValue(m_LP_BM_siteVar) = 0
    Next Site

    If (bistType = "CPU") Then
        m_pinName = "CpuMbist"

        For Each Site In TheExec.Sites
            m_grpname = TheExec.Sites(Site).SiteVariableValue(m_grpName_siteVar)
            m_grpidx = CpuMbist_Index(m_grpname)
            Exit For ''''because all sites have the same group name
        Next Site
        
        ''''Performance Mode LoopCount
        m_tmpUB = UBound(CpuMbist.Category(m_grpidx).PMode)
        m_PMname = UCase(CpuMbist.Category(m_grpidx).PMode(m_tmpUB).Name)
        If (m_tmpUB = 0 And m_PMname = "NA") Then
            m_loopcnt_PM = 0
        Else
            m_loopcnt_PM = 1 + m_tmpUB
        End If
        
        ''''Block Mode LoopCount
        m_tmpUB = UBound(CpuMbist.Category(m_grpidx).Block)
        m_BMname = UCase(CpuMbist.Category(m_grpidx).Block(m_tmpUB).Name)
        If (m_tmpUB = 0 And m_BMname = "NA") Then
            m_loopcnt_BM = 0
        Else
            m_loopcnt_BM = 1 + m_tmpUB
        End If
        
        ''''Debug Mode LoopCount
        m_tmpUB = UBound(CpuMbist.Category(m_grpidx).DebugMode)
        m_DMname = UCase(CpuMbist.Category(m_grpidx).DebugMode(m_tmpUB).Name)
        If (m_tmpUB = 0 And m_DMname = "NA") Then
            m_loopcnt_DM = 0
            m_debug_enableFlag = False
        Else
            m_debug_enableFlag = True
            ''''The purpose is due to the Flow For-Loop Setting algorithm
            ''''----------------------------------------
            '''' For LP_DM=0; LP_DM<=LCount_DM; LP_DM++
            '''' Next
            ''''----------------------------------------
            If ((m_tmpUB + 1) = 0) Then m_loopcnt_DM = 0
            If ((m_tmpUB + 1) = 1) Then m_loopcnt_DM = 0
            If ((m_tmpUB + 1) > 1) Then m_loopcnt_DM = (m_tmpUB + 1) - 1
            
            ''''<NOTICE> When there is the DebugMode, the loopcount(PM,BM) MUST be updated to '1' only
            m_loopcnt_PM = 1
            m_loopcnt_BM = 1
        End If

        m_hilmtG = 1 + CVar(UBound(CpuMbist.Category))
        m_hilmtD = CVar(UBound(CpuMbist.Category(m_grpidx).DebugMode))
        m_hilmtP = 1 + CVar(UBound(CpuMbist.Category(m_grpidx).PMode))
        m_hilmtB = 1 + CVar(UBound(CpuMbist.Category(m_grpidx).Block))
        
    ElseIf (bistType = "GPU") Then
        m_pinName = "GpuMbist"
    ElseIf (bistType = "SOC") Then
        m_pinName = "SocMbist"
    End If

    ''''<MUST>
    gS_Mbist_GroupName = m_grpname
    gL_Mbist_GroupName_Index = m_grpidx
    
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment m_pinName + " " + gS_Mbist_GroupName + " SetLoopCount::"
    ''''Setup SiteVariable for all existing Sites
    For Each Site In TheExec.Sites.Existing
        m_dlgStr = ""
        TheExec.Sites(Site).SiteVariableValue(m_DM_siteVar) = m_loopcnt_DM
        TheExec.Sites(Site).SiteVariableValue(m_PM_siteVar) = m_loopcnt_PM
        TheExec.Sites(Site).SiteVariableValue(m_BM_siteVar) = m_loopcnt_BM
        
        m_dlgStr = FormatNumeric("Site(" + CStr(Site) + ") " + gS_Mbist_GroupName, -10 - (Len(gS_Mbist_GroupName)))
        If (m_debug_enableFlag) Then
            ''''Case:: Enable Debug
            m_dlgStr = m_dlgStr + ", " + FormatNumeric(m_DM_siteVar, 12) + " = " + FormatNumeric(m_loopcnt_DM, -3) + " (RealCnt=" + CStr(m_loopcnt_DM + 1) + ") "
        Else
            m_dlgStr = m_dlgStr + ", " + FormatNumeric(m_DM_siteVar, 12) + " = " + FormatNumeric(m_loopcnt_DM, -3)
        End If
        m_dlgStr = m_dlgStr + ", " + FormatNumeric(m_PM_siteVar, 12) + " = " + FormatNumeric(m_loopcnt_PM, -3)
        m_dlgStr = m_dlgStr + ", " + FormatNumeric(m_BM_siteVar, 12) + " = " + FormatNumeric(m_loopcnt_BM, -3)
        
        TheExec.Datalog.WriteComment m_dlgStr
    Next Site
    TheExec.Datalog.WriteComment ""
    
If (False) Then
    ''''-------------------------------------------------------------------------------------------------
    '''' Setup Test Limit For Datalog
    ''''-------------------------------------------------------------------------------------------------
    Call UpdateDLogColumns(30)
    m_lolmt = 0
    TheExec.Flow.TestLimit gL_Mbist_GroupName_Index, m_lolmt, m_hilmtG, , , , , , Tname:=m_grpName_siteVar, PinName:=gS_Mbist_GroupName
    
    If (m_debug_enableFlag) Then
        ''''Case:: Enable Debug
        TheExec.Flow.TestLimit m_loopcnt_DM, m_lolmt, m_hilmtD, , , , , , Tname:=m_DM_siteVar, PinName:=m_pinName
        TheExec.Flow.TestLimit m_loopcnt_PM, m_lolmt, m_hilmtD, , , , , , Tname:=m_PM_siteVar, PinName:=m_pinName
        TheExec.Flow.TestLimit m_loopcnt_BM, m_lolmt, m_hilmtD, , , , , , Tname:=m_BM_siteVar, PinName:=m_pinName
    Else
        ''''Case:: Disable Debug
        TheExec.Flow.TestLimit m_loopcnt_DM, m_lolmt, m_hilmtD, , , , , , Tname:=m_DM_siteVar, PinName:=m_pinName
        TheExec.Flow.TestLimit m_loopcnt_PM, m_lolmt, m_hilmtP, , , , , , Tname:=m_PM_siteVar, PinName:=m_pinName
        TheExec.Flow.TestLimit m_loopcnt_BM, m_lolmt, m_hilmtB, , , , , , Tname:=m_BM_siteVar, PinName:=m_pinName
    End If
    Call UpdateDLogColumns__False
    ''''-------------------------------------------------------------------------------------------------
    TheExec.Datalog.WriteComment ""
End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_SetLoopCNT_CpuMbist()

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_SetLoopCNT_CpuMbist"
    
    Call auto_Mbist_SetLoopCNT_DM_PM_BM("CPU")
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_SetLoopCNT_GpuMbist()

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_SetLoopCNT_GpuMbist"
    
    Call auto_Mbist_SetLoopCNT_DM_PM_BM("GPU")
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_SetLoopCNT_SocMbist()

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_SetLoopCNT_SocMbist"
    
    Call auto_Mbist_SetLoopCNT_DM_PM_BM("SOC")
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''PM: Performance Mode
''''BM: Block Mode
''''DM: Debug Mode
Public Function auto_Mbist_ReSet_LoopCNT(bistType As String) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_Mbist_ReSet_LoopCNT"
    
    Dim Site As Variant
    Dim m_loopcnt_DM As Long ''''Debug       Mode LoopCount
    Dim m_loopcnt_PM As Long ''''Performance Mode LoopCount
    Dim m_loopcnt_BM As Long ''''Block       Mode LoopCount
    
    Dim m_LP_DM_siteVar As String
    Dim m_LP_PM_siteVar As String
    Dim m_LP_BM_siteVar As String

    Dim m_grpname As String
    Dim m_DMname As String
    Dim m_PMname As String
    Dim m_BMname As String
    
    Dim m_dlgStr As String
    Dim m_DM_siteVar As String
    Dim m_PM_siteVar As String
    Dim m_BM_siteVar As String
    Dim m_pinName As String

    ''''<MUST> 20151107, Initialize
    G_TestName = ""
    gB_enable_NewMbist_flag = False
    
    ''''Initialize, 20151111 New
    gB_BIRA_MC000_NRS_flag = False
    gB_BIRA_MC010_NRS_flag = False
    gB_BIRA_MC051_NRS_flag = False
    gB_BIRA_MC051_flag = False
    gB_BIRA_MC000_flag = False

    bistType = UCase(bistType)

    m_LP_DM_siteVar = "LP_DM" ''''Debug       Mode LoopIndex
    m_LP_PM_siteVar = "LP_PM" ''''Performance Mode LoopIndex
    m_LP_BM_siteVar = "LP_BM" ''''Block       Mode LoopIndex
    
    m_DM_siteVar = "LCount_DM"
    m_PM_siteVar = "LCount_PM"
    m_BM_siteVar = "LCount_BM"
    
    If (bistType = "CPU") Then
        m_pinName = "CpuMbist"
    ElseIf (bistType = "GPU") Then
        m_pinName = "GpuMbist"
    ElseIf (bistType = "SOC") Then
        m_pinName = "SocMbist"
    End If
    
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment m_pinName + " " + gS_Mbist_GroupName + " ReSet LoopCount::"
    ''''<MUST> Reset after the all For-Loop in the Flow
    For Each Site In TheExec.Sites.Existing
        TheExec.Sites(Site).SiteVariableValue(m_LP_DM_siteVar) = 0
        TheExec.Sites(Site).SiteVariableValue(m_LP_PM_siteVar) = 0
        TheExec.Sites(Site).SiteVariableValue(m_LP_BM_siteVar) = 0
        
        TheExec.Sites(Site).SiteVariableValue(m_DM_siteVar) = 0
        TheExec.Sites(Site).SiteVariableValue(m_PM_siteVar) = 0
        TheExec.Sites(Site).SiteVariableValue(m_BM_siteVar) = 0
        
        m_dlgStr = ""
        m_dlgStr = FormatNumeric("Site(" + CStr(Site) + ") " + gS_Mbist_GroupName, -10 - (Len(gS_Mbist_GroupName)))
        m_dlgStr = m_dlgStr + ", " + FormatNumeric(m_DM_siteVar, 10) + " = 0"
        m_dlgStr = m_dlgStr + ", " + FormatNumeric(m_PM_siteVar, 10) + " = 0"
        m_dlgStr = m_dlgStr + ", " + FormatNumeric(m_BM_siteVar, 10) + " = 0"
        m_dlgStr = m_dlgStr + ", Reset String G_TestName, gB_enable_NewMbist_flag = False"  ''''20151110 update message to print out
        TheExec.Datalog.WriteComment m_dlgStr
    Next Site
    TheExec.Datalog.WriteComment ""
    
If (False) Then
    ''''-------------------------------------------------------------------------------------------------
    m_pinName = m_pinName + "_" + gS_Mbist_GroupName + "_Reset_Var"
    Call UpdateDLogColumns(30)
    TheExec.Flow.TestLimit 0, 0, 0, , , , , , Tname:=m_pinName
    Call UpdateDLogColumns__False
    ''''-------------------------------------------------------------------------------------------------
    TheExec.Datalog.WriteComment ""
End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_ReSet_LoopCNT_CpuMbist()

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_ReSet_LoopCNT_CpuMbist"
    
    Call auto_Mbist_ReSet_LoopCNT("CPU")
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_ReSet_LoopCNT_GpuMbist()

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_ReSet_LoopCNT_GpuMbist"
    
    Call auto_Mbist_ReSet_LoopCNT("GPU")
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_ReSet_LoopCNT_SocMbist()

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_ReSet_LoopCNT_SocMbist"
    
    Call auto_Mbist_ReSet_LoopCNT("SOC")
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function Mbist_Initialize()
    On Error GoTo errHandler
        Dim FileExists As Boolean
        
        '---------------------------------------------------- for mbist loop module
        If mbist_sheet_init <> True Then
            If TheExec.EnableWord("Mbist_FingerPrint") = True Then
                Call Init_MBISTFailBlock
            End If
            init_MBIST_ChkList_block_loop ("SOC")
            init_MBIST_ChkList_block_loop ("CPU")
            mbist_sheet_init = True
        End If
        
        create_flag_sheet = False  '//print flag list
        index_flag_y = 0
        index_flag_x = 0
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''delet file
        File_path = ".\Mbist_Block_loop_flag_list.csv"
        FileExists = (Dir(File_path) <> "")
        If FileExists Then
               SetAttr File_path, vbNormal
               Kill File_path
        End If
        '----------------------------------------------------
    
    
    Exit Function
errHandler:
   If AbortTest Then Exit Function Else Resume Next
End Function

''''=========================================================
Public Function init_MBIST_ChkList_block_loop(bistType As String)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_parse_MBIST_ChkList_modify"

    Dim SheetName As String
    Dim mysheet As Worksheet
    Dim myCell As Object
    Dim offCell As Object
    Dim myCell_Header As Object

    Dim myCellA1 As Object
    Dim m_A1_rowCnt As Long

    Dim i As Long
    Dim j As Long
    Dim M As Long
    Dim n As Long
    Dim p As Long
    Dim g As Long
    Dim pre_store_nu As Long
    Dim store_nu As Long
    Dim store_string As String
    Dim findout_inst As Boolean
    
    Dim m_cellCnt As Long
    Dim m_cellStr As String
    Dim m_offcolStr As String
    Dim m_lastrow As Long
    Dim m_lastNCnt As Long

    Dim find_1stHeader As Boolean
    Dim find_AllHeader As Boolean

    Dim idx_END As Long
    Dim idx_Instance_BM_Decision_row As Long
    Dim idx_Instance_BM_Decision_column As Long
    Dim idx_BM_Pattern_column As Long
    Dim idx_Block_Name_column As Long
    Dim idx_Pattern_Block_column As Long

    Dim idx_BinFlag_Decision_column As Long
    Dim idx_Name_for_BinFlag_column As Long
    Dim idx_BinFlag_Mid_Name_column As Long
    Dim idx_BinFlag_with_PM_BM_column As Long

    Dim m_cnt As Long
    Dim m_idx As Long
    Dim m_pattname As String
    Dim m_pattRawname As String
    Dim m_debug_PMode As String
    Dim m_debug_Block As String

    Dim m_InsCnt As Long
    Dim m_BM_pat_Cnt As Long

    Dim m_patArr() As String
    Dim m_patcount As Long
    Dim m_Block_Cnt As Long
    Dim m_Block_cata_Cnt As Long
    Dim m_Pat_Cnt As Long

    Dim number As Long
    Dim Character As String
    Dim pre_Character As String
    Dim counter As Long
    Dim counter01 As Long
    Dim sheet_type As Long
    Dim confirm_type As Long
    
    Dim block_type As String
    Dim pre_block_type As String
    
    Dim Loc_dash As Integer
    ''''''''''''''''''''''''''''''''''''
    Dim instance_name() As String
    ReDim instance_name(100)
    Dim ins_true_name() As String
    ReDim ins_true_name(100)

    Dim instance_flag() As Boolean
    ReDim instance_flag(100)
    Dim pattern_name() As String
    ReDim pattern_name(100)
    Dim block_name() As String
    ReDim block_name(100)
    Dim block_name_count() As Long
    ReDim block_name_count(100)

    Dim block_count_name() As String
    ReDim block_count_name(100, 100)

'''    Dim block_type_pat() As String
'''    ReDim block_type_pat(50, 50, 50)
    
    Dim inst_group_name() As String
    ReDim inst_group_name(100)
    Dim flag_pm_bm() As Boolean
    ReDim flag_pm_bm(100)

    Dim flag_match_name() As String
    ReDim flag_match_name(100)
    Dim flag_mid_name() As String
    ReDim flag_mid_name(100)

    Dim block_nu As Double
    Dim block_type_nu As Double
    Dim pat_number As Double

    Dim m_flag_range As Long
    m_flag_range = 0
    ''''-------------------------------
    DebugPrtImm = False
    DebugPrtDlog = False
    ''''-------------------------------
    counter = 0
    bistType = UCase(bistType)
    gB_findPwrPin_flag = False ''''Initial
    '===================================================
    If (bistType = "CPU") Then
        gB_findCpuMbist_flag = False
        SheetName = gS_CpuMbist_sheetName
        sheet_type = CPU_sheet
    ElseIf (bistType = "GPU") Then

    ElseIf (bistType = "SOC") Then
        gB_findSocMbist_flag = False
        SheetName = gS_SocMbist_sheetName
        sheet_type = SOC_sheet
    End If
    '===================================================
    Set mysheet = Sheets(SheetName)
    Set myCellA1 = mysheet.Range("A1")
    m_A1_rowCnt = 0
    Set myCell = mysheet.Range("A1")
    m_cellStr = UCase(Trim(myCell.Value))

    find_1stHeader = False
    find_AllHeader = False
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Do
        '======================================1......To find the 1st Word "Test*Instance*Decision"
        'Here Search Cell from Left to Right
        If (find_1stHeader = False) Then
            m_cellCnt = 0
            Do While (m_cellCnt < 5)
                If (m_cellStr Like UCase("Test*Instance*Decision")) Then
                    idx_Instance_BM_Decision_row = myCell.row          '//y
                    idx_Instance_BM_Decision_column = myCell.Column    '//x
                    find_1stHeader = True
                    Exit Do
                End If
                Set myCell = myCell.Offset(rowOffset:=0, columnOffset:=1)
                m_cellStr = UCase(Trim(myCell.Value))
                m_cellCnt = m_cellCnt + 1
            Loop
        End If

        '======================================2......To find the following Header Words
        'By each Header, get the related parameters.
        If (find_1stHeader) Then
            If (bistType = "CPU") Then
                Set myCell = myCellA1.Offset(rowOffset:=idx_Instance_BM_Decision_row - 1, columnOffset:=0)
            ElseIf (bistType = "GPU") Then
            ElseIf (bistType = "SOC") Then
                Set myCell = myCellA1.Offset(rowOffset:=idx_Instance_BM_Decision_row - 1, columnOffset:=0)
            End If
            m_cellStr = UCase(Trim(myCell.Value))  '//remove sapce(front and end)
            M = 0
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            Do While (UCase(m_offcolStr) <> "END")
                Set offCell = myCell.Offset(rowOffset:=0, columnOffset:=M) ''''search cell from Left to Right
                m_offcolStr = UCase(Trim(offCell.Value))
                M = M + 1
                ''''-------------------------
                ''''Column Sequence
                ''''-------------------------
                If (m_offcolStr Like UCase("Test*Instance*Decision")) Then
                    idx_Instance_BM_Decision_column = M
                ElseIf (m_offcolStr Like UCase("*BM*Pattern*")) Then
                    idx_BM_Pattern_column = M
                ElseIf (m_offcolStr Like UCase("Block*Name*")) Then
                    idx_Block_Name_column = M
                ElseIf (m_offcolStr Like UCase("*Pattern_Block*")) Then
                    idx_Pattern_Block_column = M
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ElseIf (m_offcolStr Like UCase("*BinFlag*Decision*Type*")) Then
                    idx_BinFlag_Decision_column = M
                ElseIf (m_offcolStr Like UCase("*Match*Name*for*BinFlag*")) Then
                    idx_Name_for_BinFlag_column = M
                ElseIf (m_offcolStr Like UCase("*BinFlag*Mid*Name*")) Then
                    idx_BinFlag_Mid_Name_column = M
                ElseIf (m_offcolStr Like UCase("*BinFlag*with*PM/BM*")) Then
                    idx_BinFlag_with_PM_BM_column = M
                End If
            Loop ''''end of Do While (m_offcolStr <> "END")
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            If (m_offcolStr = ("END")) Then
                idx_END = M
                find_AllHeader = True
            End If
        End If

        m_A1_rowCnt = m_A1_rowCnt + 1
        Set myCell = myCellA1.Offset(rowOffset:=m_A1_rowCnt, columnOffset:=0)
        m_cellStr = UCase(Trim(myCell.Value))
    Loop While (find_AllHeader = False)

    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    If (find_AllHeader) Then
        If (bistType = "CPU") Then
            Set myCell_Header = myCellA1.Offset(rowOffset:=idx_Instance_BM_Decision_column - 1, columnOffset:=0)
        ElseIf (bistType = "GPU") Then
        ElseIf (bistType = "SOC") Then
            Set myCell_Header = myCellA1.Offset(rowOffset:=idx_Instance_BM_Decision_column - 1, columnOffset:=0)
        End If
        m_cellStr = UCase(Trim(myCell_Header.Value))
        'DebugPrintLog "4...(find_AllHeader=True) Row=" & myCell_Header.Row & ", Column=" & myCell_Header.Column & ", Cell=" & myCell_Header.Value & " (m_cellStr=" + m_cellStr + ")"

        ''''initialize -----------------------------------------
        If (bistType = "CPU") Then
             ReDim Mbist(sheet_type).Block(100)
        ElseIf (bistType = "GPU") Then
        ElseIf (bistType = "SOC") Then
             ReDim Mbist(sheet_type).Block(100)
        End If
        ''''----------------------------------------------------
        ''''Then get the following parameter per Header
        M = 0
        Do While (M <= idx_END)
            M = M + 1   ''''Column direction
            n = 0       ''''index and row direction

            Set myCell = myCell_Header.Offset(rowOffset:=0, columnOffset:=(M - 1)) ''''rowOffset MUST be always '0'
            m_cellStr = UCase(Trim(myCell.Value))
            m_lastrow = myCell.End(xlDown).row    '//y range
            If (bistType = "CPU") Then
                m_lastNCnt = m_lastrow - idx_Instance_BM_Decision_row
            ElseIf (bistType = "GPU") Then
            ElseIf (bistType = "SOC") Then
                m_lastNCnt = m_lastrow - idx_Instance_BM_Decision_row
            End If

            Select Case (M)
            Case idx_END
                Exit Do     ''''end
            '=================================================================================Test Instance Name for PM/BM Decision
            Case idx_Instance_BM_Decision_column
                Do While (n < m_lastNCnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.Value))
                    If (bistType = "CPU") Then
                        instance_name(n) = m_cellStr
                    ElseIf (bistType = "GPU") Then
                    ElseIf (bistType = "SOC") Then
                        instance_name(n) = m_cellStr
                    End If
                    n = n + 1
                Loop

                m_InsCnt = n
                ReDim Preserve instance_name(n - 1)    'redim and hold orignal data
                n = 0:: counter = 0:: i = 0:: p = 0
            '=================================================================================Choose PM/BM Pattern
             Case idx_BM_Pattern_column
                Do While (n < m_lastNCnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.Value))
                    If (m_cellStr = "BM") Then
                       instance_flag(n) = True
                       ins_true_name(counter) = instance_name(n)
                       counter = counter + 1
                    Else
                       instance_flag(n) = False
                    End If
                    n = n + 1
                Loop

                ReDim Preserve instance_flag(n - 1)
                ReDim Preserve ins_true_name(counter - 1)
                m_BM_pat_Cnt = n
                n = 0:: counter = 0:: i = 0:: p = 0
            '=================================================================================Block Name
             Case idx_Block_Name_column
                Do While (n < m_lastNCnt)
                    number = 0: Character = "": block_type = ""
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.Value))
                    block_type = m_cellStr
                    Call Separate_nu_char(m_cellStr, number, Character)
                    ReDim Preserve Mbist(sheet_type).Block(i).block_type_pat(number)
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    n = n + 1
                    If (pre_Character <> Character) Then
                        block_name(counter01) = Character
                        block_count_name(counter01, block_name_count(counter01)) = m_cellStr
                        block_name_count(counter01) = block_name_count(counter01) + 1
                        If (pre_Character <> "") Then
                            counter = counter + 1
                        End If
                        counter01 = counter01 + 1
                        pre_Character = Character
                    Else
                        block_count_name(counter01 - 1, block_name_count(counter01 - 1)) = m_cellStr
                        block_name_count(counter01 - 1) = block_name_count(counter01 - 1) + 1
                    End If
                    
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    If (pre_block_type = m_cellStr Or pre_block_type = "") Then
                        ReDim Preserve Mbist(sheet_type).Block(counter).block_type_pat(number).Pat(p)
                        Mbist(sheet_type).Block(counter).block_type_pat(number).Pat(p) = myCell.Offset(rowOffset:=0, columnOffset:=1)
                        pre_block_type = m_cellStr
                        p = p + 1
                    Else
                        i = i + 1
                        p = 0
                        ReDim Preserve Mbist(sheet_type).Block(counter).block_type_pat(number)
                        ReDim Preserve Mbist(sheet_type).Block(counter).block_type_pat(number).Pat(p)
                        Mbist(sheet_type).Block(counter).block_type_pat(number).Pat(p) = myCell.Offset(rowOffset:=0, columnOffset:=1)
                        pre_block_type = m_cellStr
                        p = p + 1
                    End If
                    'p = p + 1
                Loop
                   p = p + 1
'''                ReDim Preserve block_name(counter - 1)
'''                ReDim Preserve block_name_count(counter - 1)
                ReDim Preserve block_name(counter01 - 1)
                ReDim Preserve block_name_count(counter01 - 1)
                m_Block_Cnt = n     '//Block amount
                m_Block_cata_Cnt = counter01
                n = 0:: counter = 0:: i = 0:: p = 0
            '=================================================================================Pattern_Block
             Case idx_Pattern_Block_column
                pre_Character = "":: Character = ""
                Do While (n < m_lastNCnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.Value))
                    pattern_name(n) = m_cellStr
                    n = n + 1
                Loop
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ReDim Preserve pattern_name(n - 1)
                m_Pat_Cnt = n
                n = 0:: counter = 0:: i = 0:: p = 0
            '=================================================================================BinFlag Decision Type(TestInstanceName, GroupName)
             Case idx_BinFlag_Decision_column
                Do While (n < m_lastNCnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.Value))
                    inst_group_name(n) = m_cellStr
                    n = n + 1
                Loop
                ReDim Preserve inst_group_name(n - 1)
                m_flag_range = n
                n = 0:: counter = 0:: i = 0:: p = 0
            '=================================================================================Match Name for BinFlag
             Case idx_Name_for_BinFlag_column
                Do While (n < m_lastNCnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.Value))
                    flag_match_name(n) = m_cellStr
                    n = n + 1
                Loop
                ReDim Preserve flag_match_name(n - 1)
                n = 0:: counter = 0:: i = 0:: p = 0
            '=================================================================================BinFlag Mid Name
             Case idx_BinFlag_Mid_Name_column
                Do While (n < m_lastNCnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.Value))
                    flag_mid_name(n) = m_cellStr
                    n = n + 1
                Loop
                ReDim Preserve flag_mid_name(n - 1)
                n = 0:: counter = 0:: i = 0:: p = 0
            '=================================================================================BinFlag with PM/BM (Yes/No)
             Case idx_BinFlag_with_PM_BM_column
                Do While (n < m_lastNCnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.Value))
                    If (UCase(m_cellStr) Like UCase("Yes")) Then
                        flag_pm_bm(n) = True
                    Else
                        flag_pm_bm(n) = False
                    End If

                    n = n + 1
                Loop
                ReDim Preserve flag_pm_bm(n - 1)
                n = 0:: counter = 0:: i = 0:: p = 0
            '=================================================================================
             Case Else
                'DebugPrintLog "6...Empty Column(" & M & ") !!!"
             End Select
            '=================================================================================

        Loop ''''end of Do While (m <= idx_END)
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++for BM & instance & pattern
         i = 0:: n = 0:: p = 0:: g = 0
         ReDim Preserve Mbist(sheet_type).Block(m_Block_cata_Cnt - 1)
         For i = 0 To m_Block_cata_Cnt - 1
            Mbist(sheet_type).Block(i).block_name = block_name(i)
            Mbist(sheet_type).Block(i).block_count = block_name_count(i)
            '--------------------------------------------------------------
            ReDim Preserve Mbist(sheet_type).Block(i).block_count_name(block_name_count(i) - 1)
            For n = 0 To block_name_count(i) - 1
                Mbist(sheet_type).Block(i).block_count_name(n) = block_count_name(i, n)
            Next n
            '--------------------------------------------------------------
            'Mbist(sheet_type).Block(i).ins_name = ins_true_name(i)
            '--------------------------------------------------------------
            p = 0
            For n = 0 To m_Pat_Cnt - 1
                For counter = 0 To block_name_count(i) - 1
                    If (UCase(pattern_name(n)) Like UCase("*" + Mbist(sheet_type).Block(i).block_count_name(counter)) + "*") Then
                        ReDim Preserve Mbist(sheet_type).Block(i).pat_name(p)
                        ReDim Preserve Mbist(sheet_type).Block(i).pat_tested(p)
                        Mbist(sheet_type).Block(i).pat_name(p) = pattern_name(n)
                        Mbist(sheet_type).Block(i).pat_tested(p) = False
                        p = p + 1
                        Mbist(sheet_type).Block(i).pat_count = p
                        Exit For
                    End If
                Next counter
            Next n
            '--------------------------------------------------------------
            For p = 0 To UBound(Mbist(sheet_type).Block(i).block_type_pat)
                For g = 0 To UBound(Mbist(sheet_type).Block(i).block_type_pat(p).Pat)
                    pre_store_nu = 0
                    ReDim Preserve Mbist(sheet_type).Block(i).block_type_pat(p).instance(g)
                    For n = 0 To UBound(ins_true_name)
                        ins_true_name(n) = Trim(ins_true_name(n))
                        Mbist(sheet_type).Block(i).block_type_pat(p).Pat(g) = Trim(Mbist(sheet_type).Block(i).block_type_pat(p).Pat(g))
                        Loc_dash = InStr(1, ins_true_name(n), Mbist(sheet_type).Block(i).block_type_pat(p).Pat(g))
                        If (Loc_dash > 0) Then
                            If (pre_store_nu < Len(Mbist(sheet_type).Block(i).block_type_pat(p).Pat(g))) Then
                                store_string = ins_true_name(n)
                                Mbist(sheet_type).Block(i).block_type_pat(p).instance(g) = ins_true_name(n)
                            End If
                            pre_store_nu = Len(Mbist(sheet_type).Block(i).block_type_pat(p).Pat(g))
                        End If
                    Next n
                Next g
            Next p
            
            For p = 0 To UBound(Mbist(sheet_type).Block(i).block_type_pat)
              For g = 0 To UBound(Mbist(sheet_type).Block(i).block_type_pat(p).Pat)
                 If (p > 0) Then
                    ReDim Preserve Mbist(sheet_type).Block(i).block_type_pat(p).instance(g)
                    Mbist(sheet_type).Block(i).block_type_pat(p).instance(g) = Mbist(sheet_type).Block(i).block_type_pat(0).instance(g)
                 End If
              Next g
            Next p
            '--------------------------------------------------------------
         Next i

         ReDim Preserve Mbist(sheet_type).Block(i - 1)
         i = 0:: n = 0:: p = 0
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++for flag
         For i = 0 To m_flag_range - 1
            If (inst_group_name(i) = UCase("TestInstanceName") And flag_pm_bm(i) = True) Then
                ReDim Preserve mbist_match(sheet_type).inst_nu(p)
                mbist_match(sheet_type).inst_nu(p).binflag_match_name = flag_match_name(i)
                mbist_match(sheet_type).inst_nu(p).binflag_mid_name = flag_mid_name(i)
                p = p + 1
            End If
         Next i
         mbist_match(sheet_type).inst_count = p
         i = 0:: n = 0:: p = 0
         '//flag=front_pp+performance+NV+Blockname
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

    End If ''''end of If (find_AllHeader) Then
    
'    If Mbist(sheet_type).Block(0).block_type_pat(1).instance(0) Is Nothing Then
'       i = 0:: n = 0:: p = 0
'    End If
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

