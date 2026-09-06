Attribute VB_Name = "VBT_LIB_Digital_UART"
Option Explicit

Public Function UARTTest_CMEM_T(patt As Pattern) As Long
On Error GoTo errHandler
Dim dspData As New PinListData
Dim RegCapWav As New DSPWave
Dim TempStr(2000) As String
Dim tempLong(2000) As Long
Dim OutputFile As String
Dim plotWaves As Boolean
Dim Fres As Double
Dim StartTime As Double
Dim endTime As Double
Dim LoopCnt As Long
Dim i As Integer
Dim site As Variant
        
    plotWaves = False
    
    LoopCnt = 1800
    
    'Call TheHdw.Patterns(".\Pattern\FIJI_index54_M1_0814_modify_GLC40_mod_scenario54_gpio26_AI6_Loop.pat").Load
    'Call TheHdw.Patterns(".\Pattern\FIJI_index54_M1_0814_modify_GLC40_mod_scenario54_gpio26_AI6_Loop.pat").Start
    
    OutputFile = ".\UART_Output\UARToutput.txt"
      
    Set dspData = TheHdw.Protocol.ports("UART_PA").NWire.CMEM.DSPWave
    
    
    TheHdw.Protocol.ports("UART_PA").Modules("ReadUART").Select
    TheHdw.Protocol.ports("UART_PA").Modules.StartSelected
    
    For i = 0 To 15
        Call TheHdw.Patterns(patt).start
        TheHdw.Digital.Patgen.HaltWait
    Next i
    
    RegCapWav = dspData.Copy
    
    If TheExec.TesterMode = testModeOffline Then
        Call RegCapWav.FileImport(".\Working_UART_Bytes_8bit_Capture.wav", File_Image_wav)
    End If
    
    If plotWaves = True Then
        For Each site In TheExec.sites.Active
            RegCapWav(site).Plot "UART Capture " & site
        Next site
    End If
        
    For i = 0 To LoopCnt
    
        tempLong(i) = RegCapWav.Element(i)
        
        If tempLong(i) <> 255 Then
            TempStr(i) = Chr(tempLong(i))
        End If
        
    Next i
        
    Open OutputFile For Output As #4
    
    For i = 0 To LoopCnt
        Print #4, TempStr(i);
    Next i
    
    Close #4
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_UART", "UARTTest_CMEM_T")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function UART_LoopBackTest() As Long
On Error GoTo errHandler
Dim dspData As New PinListData
Dim RegCapWav As New DSPWave
Dim TempStr(2000) As String
Dim tempLong(2000) As Long
Dim OutputFilePath As String
Dim plotWaves As Boolean
Dim Fres As Double
Dim StartTime As Double
Dim endTime As Double
Dim LoopCnt As Long
Dim i As Integer
   
    plotWaves = False
    
    LoopCnt = 20
    
    'Call TheHdw.Patterns(".\Pattern\FIJI_index54_M1_0814_modify_GLC40_mod_scenario54_gpio26_AI6_Loop.pat").Load
    'Call TheHdw.Patterns(".\Pattern\FIJI_index54_M1_0814_modify_GLC40_mod_scenario54_gpio26_AI6_Loop.pat").Start

    
    Call TheHdw.Digital.ApplyLevelsTiming(True, True, True)
    
    TheHdw.Wait (20 * ms)
    
''    Thehdw.Protocol.Ports("UART_RX_Port").Enabled = False
''    Thehdw.Protocol.Ports("UART_TX_Port").Enabled = False
''    Thehdw.Protocol.Ports("UART_RX_Port").Enabled = True
''    Thehdw.Protocol.Ports("UART_TX_Port").Enabled = True

''    Set dspData = TheHdw.Protocol.Ports("UART_RX_Port").NWire.CMEM.DSPWave
    
    TheHdw.Protocol.ports("UART_RX_Port").Modules("READ_UART_Module").Select
    TheHdw.Protocol.ports("UART_TX_Port").Modules("WRITE_UART_Module").Select
    TheHdw.Protocol.ports("UART_RX_Port,UART_TX_Port").Modules.StartSelected
    
''    Set dspData = TheHdw.Protocol.Ports("UART_RX_Port").NWire.CMEM.DSPWave
    
''    Thehdw.Protocol.Ports("UART_RX_Port").IdleWait
''    Thehdw.Protocol.Ports("UART_TX_Port").IdleWait
''    Thehdw.Protocol.Ports("UART_RX_Port").Enabled = False
''    Thehdw.Protocol.Ports("UART_TX_Port").Enabled = False
    
    Set dspData = TheHdw.Protocol.ports("UART_RX_Port").NWire.CMEM.DSPWave
    
    
''    For i = 0 To 15
''        Call TheHdw.Patterns(patt).Start
''        TheHdw.digital.Patgen.HaltWait
''    Next i
    
    RegCapWav = dspData.Copy
    
    If TheExec.TesterMode = testModeOffline Then
        Call RegCapWav.FileImport(".\Working_UART_Bytes_8bit_Capture.wav", File_Image_wav)
    End If
    
    Dim site As Variant
    If plotWaves = True Then
        For Each site In TheExec.sites.Active
            RegCapWav(site).Plot "UART Capture " & site
        Next site
    End If
        
    Dim DateCodePath As String

    For Each site In TheExec.sites.Active
        
        DateCodePath = Year(Now) & Month(Now) & Day(Now) & Hour(Now) & Minute(Now) & Second(Now)
        
        For i = 0 To LoopCnt
            tempLong(i) = RegCapWav(site).Element(i)
    
            If tempLong(i) <> 255 Then
                TempStr(i) = Chr(tempLong(i))
            End If
        Next i
        
        OutputFilePath = ".\UART_Output\" & "Site" & site & "_UARToutput_" & DateCodePath & ".txt"
         
        Open OutputFilePath For Output As #4
    
        For i = 0 To LoopCnt
            Print #4, TempStr(i);
        Next i
    
        Close #4
    Next site
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_UART", "UART_LoopBackTest")
    If AbortTest Then Exit Function Else Resume Next
End Function




Public Function ReadCMEM() As Long
On Error GoTo errHandler
    
    TheHdw.Protocol.ports("UART_PA").Enabled = True
    
    'Call UART_read_n_byte_DSP(2000)
    Call UART_read_n_byte_DSP(100000)   ' Roger
        
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_UART", "ReadCMEM")
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function UART_read_n_byte_DSP(n As Long) As Long
On Error GoTo errHandler
Dim i As Long

'Initialize arrays
'ReDim readRegData(N)
'tempDSPwave.CreateConstant 0, N + 1, DspLong

    For i = 0 To n
            
        Call UARTReadRegDSP
    '''    Call TheHdw.Protocol.Ports("UART_RX_Port").IdleWait
        
    Next i
  
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_UART", "UART_read_n_byte_DSP")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function UART_write_n_byte_DSP(n As Long, WriteString As String) As Long
On Error GoTo errHandler
Dim i As Long

'Initialize arrays
ReDim writeRegData(n)
'tempDSPwave.CreateConstant 0, N + 1, DspLong

    Dim ASCII_Array() As Long
    Dim StringLength As Long
    Dim SingleChar As String
    StringLength = Len(WriteString)
    
    ReDim ASCII_Array(StringLength - 1) As Long
    
    For i = 0 To StringLength - 1
        SingleChar = mid(WriteString, i + 1, 1)
        ASCII_Array(i) = Asc(SingleChar)
    Next i
    
    
    For i = 0 To n
            
        Call UARTWriteRegDSP(ASCII_Array(i))
    ''    Call TheHdw.Protocol.Ports("UART_TX_Port").IdleWait
        
    Next i
      
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_UART", "UART_write_n_byte_DSP")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function PreLoad_PA_Modules(ReadWriteFlag As String, ByteNum As Long, PortName As String, Optional WriteString As String) As Long
On Error GoTo errHandler

    
    TheHdw.Protocol.ports(PortName).Enabled = True
     
    TheHdw.Protocol.ModuleRecordingEnabled = True
    
    
    Dim PAModuleName As String
    PAModuleName = ReadWriteFlag & "_UART_Module"
    
    With TheHdw.Protocol.ports(PortName)
    
        If .Modules.IsRecorded(PAModuleName, False, False) = False Then
    
'''            Call ReadCMEM
'''            TheHdw.Protocol.Ports("UART_PA").Enabled = True

            If UCase(ReadWriteFlag) = "READ" Then
                Call UART_read_n_byte_DSP(ByteNum)
                
            ElseIf UCase(ReadWriteFlag) = "WRITE" Then
                
                Call UART_write_n_byte_DSP(ByteNum, WriteString)
                
            End If
            
            Call .Modules.StopRecording
    
        End If
        
    End With
    
'    TheHdw.Protocol.Ports("UART_TX_Port").ModuleFiles.UnloadAll
'    TheHdw.Protocol.Ports("UART_RX_Port").ModuleFiles.UnloadAll
        
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_UART", "PreLoad_PA_Modules")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function VaryFreq_PA_UART(ClockPort As String, ClkFreq As Double, ACSpec As String) As Long
On Error GoTo errHandler

Dim site As Variant

    For Each site In TheExec.sites
        TheHdw.Protocol.ports(ClockPort).Halt
        TheHdw.Protocol.ports(ClockPort).Enabled = False
    Next site
    
    Call TheExec.Overlays.ApplyUniformSpecToHW(ACSpec, ClkFreq)
 
    TheHdw.Wait 0.003
    TheHdw.Protocol.ports(ClockPort).Enabled = True
    TheHdw.Protocol.ports(ClockPort).NWire.ResetPLL
    
    TheHdw.Wait 0.001

    Call TheHdw.Protocol.ports(ClockPort).NWire.Frames("RunFreeClock").Execute
    TheHdw.Protocol.ports(ClockPort).IdleWait
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_UART", "VaryFreq_PA_UART")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function UARTReadRegDSP() As Long
On Error GoTo errHandler

    Dim Status As Long
    
    Status = TL_SUCCESS
    
    With TheHdw.Protocol.ports("UART_PA").NWire.Frames("UART_Rcv")
        .Execute tlNWireExecutionType_CaptureInCMEM
    End With

    UARTReadRegDSP = Status

    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_UART", "UARTReadRegDSP")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function UARTWriteRegDSP(ASCII_Num As Long) As Long
On Error GoTo errHandler

    Dim Status As Long
    
    Status = TL_SUCCESS
    
    With TheHdw.Protocol.ports("UART_TX_Port").NWire.Frames("UART_TX")
        .Fields("Data").value = ASCII_Num
        .Execute
    End With

    UARTWriteRegDSP = Status

    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_UART", "UARTWriteRegDSP")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function UARTTest_CMEM_T_Update(patt As Pattern) As Long
On Error GoTo errHandler
    Dim dspData As New PinListData
    Dim RegCapWav As New DSPWave
    'Dim TempStr(2000) As String        ori 20160516
    'Dim tempLong(2000) As Long        ori 20160516
    
    Dim TempStr() As String
    Dim tempLong() As Long
    
    Dim OutputFilePath As String
    Dim plotWaves As Boolean
    Dim Fres As Double
    Dim StartTime As Double
    Dim endTime As Double
    Dim LoopCnt As Long
    Dim i As Long
    
    
    Dim site As Variant


    Set dspData = TheHdw.Protocol.ports("UART_PA").NWire.CMEM.DSPWave
    
''    TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic
    
    TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug
        
    plotWaves = False
    
    'loopCnt = 2000 ' ori 20160516
    LoopCnt = 5000 * 3
    
      
    TheHdw.Wait (20 * ms)
''    Set dspData = TheHdw.Protocol.Ports("UART_PA").NWire.CMEM.DSPWave

'''    TheHdw.Protocol.Ports("UART_PA").NWire.CMEM.Reset
    
    TheHdw.Protocol.ports("UART_PA").Modules("READ_UART_Module").Select
    TheHdw.Protocol.ports("UART_PA").Modules.StartSelected
    
    TheHdw.Wait (20 * ms)

    Call TheHdw.Patterns(patt).Load
    
    For i = 0 To 0
'''''         2014/03/21 added for FS_mount mode, hard reset ===================================
'''''        TheHdw.Patterns("SPI_Pre_PAT_HighZ").Load
'''''        TheHdw.Patterns("SPI_Pre_PAT_HighZ").start
        TheHdw.DCVS.Pins("SPI_PWR").Voltage.Main.value = 0#
        TheHdw.Wait 0.02 ''' maybe 10mS, org 20mS
        TheHdw.DCVS.Pins("SPI_PWR").Voltage.Main.value = 1.8
        TheHdw.Wait 0.02 ''' maybe 10mS, org 20mS
        
        Call TheHdw.Patterns(patt).start
        TheHdw.Digital.Patgen.HaltWait
    Next i
    
''    TheHdw.Protocol.Ports("UART_PA").IdleWait
    
    TheHdw.Wait (500 * ms)
''    TheHdw.Protocol.Ports("UART_PA").Halt
    TheHdw.Protocol.ports("UART_PA").Enabled = False
    
    TheHdw.Wait (20 * ms)
    
'''    Set dspData = TheHdw.Protocol.Ports("UART_PA").NWire.CMEM.DSPWave
    
    RegCapWav = dspData.Copy
    
    TheHdw.Wait (20 * ms)
    
''    If TheExec.TesterMode = testModeOffline Then
''        Call RegCapWav.FileImport(".\Working_UART_Bytes_8bit_Capture.wav", File_Image_wav)
''    End If
''
''    Dim Site As Variant
''    If plotWaves = True Then
''        For Each Site In TheExec.Sites.Active
''            RegCapWav(Site).Plot "UART Capture " & Site
''        Next Site
''    End If
    
    Dim DateCodePath As String
    For Each site In TheExec.sites.Active
    
    ReDim TempStr(5000 * 3)
    ReDim tempLong(5000 * 3)

        DateCodePath = Year(Now) & Month(Now) & Day(Now) & Hour(Now) & Minute(Now) & Second(Now)
        LoopCnt = RegCapWav(site).SampleSize
        For i = 0 To LoopCnt - 1
            tempLong(i) = RegCapWav(site).Element(i)
            
            If tempLong(i) <> 255 Then
                TempStr(i) = Chr(tempLong(i))
            End If
        Next i
        'OutputFilePath = ".\UART_Output\" & "Site" & Site & " & "_"& Cstr(TheExec.DataManager.InstanceName) & "_UARToutput_" & DateCodePath & ".txt"
        OutputFilePath = ".\UART_Output\" & "Site" & site & "_" & TheExec.DataManager.instancename & "_UARToutput_" & DateCodePath & ".txt"
        TheExec.Datalog.WriteComment "Site" & site & "_" & TheExec.DataManager.instancename & "_UARToutput"
        Open OutputFilePath For Output As #4
        For i = 0 To LoopCnt - 1
            Print #4, TempStr(i);
        Next i
        Close #4
        
        Dim MaxLimit As Integer '20160601 display the number of bytes being captured by UART port as customer request
        MaxLimit = 15000#
        If LoopCnt >= MaxLimit Then
            TheExec.Datalog.WriteComment "Site" & site & "_" & TheExec.DataManager.instancename & "_UARToutput==>" & " Warning! Log Data Has Reached the MaxLimit " & MaxLimit & " Bytes"
        Else
            TheExec.Datalog.WriteComment "Site" & site & "_" & TheExec.DataManager.instancename & "_UARToutput==>" & " Log Data is " & LoopCnt & " Bytes"
        End If
        
    Next site
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_UART", "UARTTest_CMEM_T_Update")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub ReStartFRC()
On Error GoTo errHandler

    Dim site As Variant

    For Each site In TheExec.sites
        If TheHdw.Protocol.ports(XI0_Diff_Port).Enabled = True Then
            TheHdw.Protocol.ports(XI0_Diff_Port).Halt
            TheHdw.Protocol.ports(XI0_Diff_Port).Enabled = False
        End If
        If TheHdw.Protocol.ports(RT_CLK32768_Port).Enabled = True Then
            TheHdw.Protocol.ports(RT_CLK32768_Port).Halt
            TheHdw.Protocol.ports(RT_CLK32768_Port).Enabled = False
        End If
    Next site
    TheHdw.Digital.Pins(REFCLK_XI0).initState = chInitoff
    TheHdw.Digital.Pins(REFCLK_RT_CLK32768).initState = chInitoff
    TheHdw.Wait 0.005
    
''    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
''    Call TheExec.Overlays.ApplyUniformSpecToHW(ACSpec, ClkFreq)
 
    TheHdw.Wait 0.005
    TheHdw.Protocol.ports(XI0_Diff_Port).Enabled = True
    TheHdw.Protocol.ports(XI0_Diff_Port).NWire.ResetPLL
    
    TheHdw.Wait 0.005

    Call TheHdw.Protocol.ports(XI0_Diff_Port).NWire.Frames("RunFreeClock").Execute
    TheHdw.Protocol.ports(XI0_Diff_Port).IdleWait
    
    TheHdw.Wait 0.005
    TheHdw.Protocol.ports(RT_CLK32768_Port).Enabled = True
    TheHdw.Protocol.ports(RT_CLK32768_Port).NWire.ResetPLL
    
    TheHdw.Wait 0.005

    Call TheHdw.Protocol.ports(RT_CLK32768_Port).NWire.Frames("RunFreeClock").Execute
    TheHdw.Protocol.ports(RT_CLK32768_Port).IdleWait

    Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_UART", "ReStartFRC")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

