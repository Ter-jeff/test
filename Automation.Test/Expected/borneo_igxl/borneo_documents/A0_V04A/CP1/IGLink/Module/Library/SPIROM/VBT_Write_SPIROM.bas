Attribute VB_Name = "VBT_Write_SPIROM"
#Const isUFP = True
Option Explicit

Global Uniform_256KB_Sector As Boolean
Global Total_Count As Long
Const Mem32MByte As Long = 268435456
Const Mem16MByte As Long = 134217728
Const Mem8MByte As Long = 67108864
'Const Addr4096MByte As LongLong = 4294967296#
Const Addr128MByte As Long = 134217728
Const Addr32MByte As Long = 33554432
Const Addr16MByte As Long = 16777216
Const Addr8MByte As Long = 8388608
Const MbytSize As Long = 1048576
'1 MByte Data Capture Per Pattern Run
Const Glb_CaptureSizePerPat = 1048576 * 4

Public SPI_Binary_Erase_Data As New DSPWave
Public SPI_Binary_Write_Data As New DSPWave
Public SPI_Flag_Ary() As New SiteBoolean
Public SPI_Flag_temp_Ary() As New SiteLong

Public CheckSumResult As Long
Public SPI_Flag_Sum As New SiteLong
Public SpiromCodeFile As String
Public Glb_RomSize As Long
Public Glb_SetEraseAddr As Boolean
Public EraseCMD As Long
Public ResultPass As New SiteBoolean
Public Glb_EraseTime As Long
Public Function SPIROM_Continuity(RelayOnPins As PinList) As Long
On Error GoTo errHandler
    
    Dim PPMUMeasure As New PinListData
        
'''    TheHdw.Utility.pins("K80_SPI0_SCLK, K84_SPI0_MISO").State = tlUtilBitOn
    TheHdw.Utility.Pins(RelayOnPins).State = tlUtilBitOn
    TheHdw.Digital.Pins("SPIROM_PINS").Disconnect
                       
    TheHdw.Wait 0.05
               
    With TheHdw.PPMU.Pins("SPIROM_PINS")
         .Connect
         .ForceI -0.0002, -2 * mA, -0.2, -1
         .Gate = tlOn
    End With
        
    TheHdw.Wait 0.05

    TheHdw.PPMU.Pins("SPIROM_PINS").test
    
    TheHdw.Wait 0.05
    
    With TheHdw.PPMU.Pins("SPIROM_PINS")
        .ForceI 0
        .Gate = tlOff
        .Disconnect
    End With
    
    TheHdw.Digital.Pins("SPIROM_PINS").Connect
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_Write_SPIROM", "SPIROM_Continuity")
    If AbortTest Then Exit Function Else Resume Next
End Function




Public Function Get_RomSize(RomFileName1 As String) As Long
On Error GoTo errHandler
    
    Dim Rom_Data As Byte
    Dim DspArray(258) As Long
    Dim TotalDspArray() As Long
    Dim Count As Long
    Dim SPI_DSSCCap_Signal As New DSPWave
    Dim result As New SiteDouble
    Dim WaveCount As Long
    Dim RomSize As Long
    
    Const SrcDataSize = 259
    Total_Count = 0
    WaveCount = 0
    Close #1

    Open RomFileName1 For Binary As #1
'=============================================================
    While (Not EOF(1)) And (Total_Count < 33554432)     '256MBits
        For Count = 3 To SrcDataSize - 1
            DspArray(Count) = 2 ^ 8 - 1
        Next Count
        Count = 3
        While ((Not EOF(1)) And (Count < SrcDataSize))
            Get #1, Total_Count + 1, Rom_Data
            DspArray(Count) = Rom_Data
            Count = Count + 1
            Total_Count = Total_Count + 1
        Wend
        WaveCount = WaveCount + 1
    Wend
'======================================================== =====
    Close #1
    
    Get_RomSize = Total_Count
    Glb_RomSize = Total_Count

    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_Write_SPIROM", "Get_RomSize")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function PageProgram(PatName As Pattern, RomFileName1 As String, RelayOnPins As PinList, DigSrcPin As PinList) As Long
On Error GoTo errHandler
    
    Dim Rom_Data As Byte
    Dim DspArray(258) As Long
    Dim TotalDspArray() As Long
    Dim Count As Long
    Dim DspSrcWave As New DSPWave
    Dim DspSrcWave1 As New DSPWave
    Dim DspRefWave As New DSPWave
    Dim DspRefWave1 As New DSPWave
    Dim SrcWaveName As String
    Dim SPI_DSSCCap_Signal As New DSPWave
    Dim result As New SiteDouble
    Dim WaveCount As Long
    Dim PatCount As Long
    Dim PattArray() As String
    Dim patt As String
    
    Const SrcDataSize = 259
    Total_Count = 0
    WaveCount = 0
    Close #1
    
''    PatName = Worksheets("PatSets_SPIROM").Cells(11, 5).Value
  
    TheHdw.Utility.Pins(RelayOnPins).State = tlUtilBitOn
   
    ' Level & Timing

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    Call PATT_GetPatListFromPatternSet(PatName.value, PattArray, PatCount)
    patt = CStr(PattArray(0))
    TheHdw.Patterns(patt).Load
    TheHdw.Digital.Patgen.TimeOut = 10
    
    'Address
    Dim Address(31) As String
    Dim Temp_data As String
    Dim DecToBin As String
    Dim i As Integer
    Dim tempA As New SiteLong
    Dim tempB As New SiteLong
    Dim tempC As New SiteLong
    Dim tempD As New SiteLong
    Dim temp0_7 As New SiteLong
    Dim temp8_15 As New SiteLong
    Dim temp16_23 As New SiteLong
    Dim temp24_31 As New SiteLong
    Dim site As Variant
    Dim AddrByte1 As Long
    Dim AddrByte2 As Long
    Dim AddrByte3 As Long
    Dim TempArray(31) As Long
        
    
    Open RomFileName1 For Binary As #1
      
    '=============================================================
        AddrByte1 = 0
        AddrByte2 = 0
        AddrByte3 = 0
    While (Not EOF(1)) And (Total_Count < Addr8MByte)
        For Count = 3 To SrcDataSize - 1
            DspArray(Count) = 2 ^ 8 - 1
        Next Count
        
        Count = 3
        
        While ((Not EOF(1)) And (Count < SrcDataSize))
            Get #1, Total_Count + 1, Rom_Data
            DspArray(Count) = Rom_Data
            Count = Count + 1
            Total_Count = Total_Count + 1
            'Debug.Print Rom_Data
        Wend
        
        Count = 0
        
        DspArray(0) = AddrByte3
        DspArray(1) = AddrByte2
        DspArray(2) = AddrByte1
'        If AddrByte3 = 16 Then Stop
        AddrByte2 = AddrByte2 + 1
        If AddrByte2 > 255 Then '#16_ELSE_CASE_CHK
            AddrByte3 = AddrByte3 + 1
            AddrByte2 = 0
        Else
        End If
        
        'WaveDefinitions
        DspSrcWave.data = DspArray
        'DspSrcWave.Plot
        Call theexec.WaveDefinitions.CreateWaveDefinition("SrcData", DspSrcWave, True)

        'Insert DSSC Loading & Pattern Loading & Pattern Execution.
        With TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source
'        With TheHdw.DSSC.Pins(DigSrcPin).Pattern(PatName).Source
            .Signals.Add "DSSC_Src"
            .Signals.DefaultSignal = "DSSC_Src"
            .Signals("DSSC_Src").WaveDefinitionName = "SrcData"
            .Signals.item("DSSC_Src").sampleSize = 259
            .Signals("DSSC_Src").Amplitude = 1
            If LCase(glb_TesterType) = "jaguar" Then
                .Signals("DSSC_Src").LoadSettings
            End If
        End With

        TheHdw.Patterns(patt).test pfNever, 0
        TheHdw.Digital.Patgen.HaltWait

        WaveCount = WaveCount + 1
    Wend
        
    
    Close #1
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_Write_SPIROM", "PageProgram")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function PageProgram_4Byte(PatName As Pattern, RomFileName1 As String, RelayOnPins As PinList, DigSrcPin As PinList) As Long
On Error GoTo errHandler
    
    Dim Rom_Data As Byte
    Dim DspArray(259) As Long
    Dim TotalDspArray() As Long
    Dim Count As Long
    Dim DspSrcWave As New DSPWave
    Dim DspSrcWave1 As New DSPWave
    Dim DspRefWave As New DSPWave
    Dim DspRefWave1 As New DSPWave
    Dim SrcWaveName As String
    Dim SPI_DSSCCap_Signal As New DSPWave
    Dim result As New SiteDouble
    Dim WaveCount As Long
    Dim PatCount As Long
    Dim PattArray() As String
    Dim patt As String
    
    Const SrcDataSize = 260
    Total_Count = 0
    WaveCount = 0
    Close #1
    
''    PatName = Worksheets("PatSets_SPIROM").Cells(11, 5).Value
  
    TheHdw.Utility.Pins(RelayOnPins).State = tlUtilBitOn
   
    ' Level & Timing

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    Call PATT_GetPatListFromPatternSet(PatName.value, PattArray, PatCount)
    patt = CStr(PattArray(0))
    TheHdw.Patterns(patt).Load
    TheHdw.Digital.Patgen.TimeOut = 10
    
    'Address
    Dim Address(31) As String
    Dim Temp_data As String
    Dim DecToBin As String
    Dim i As Integer
    Dim tempA As New SiteLong
    Dim tempB As New SiteLong
    Dim tempC As New SiteLong
    Dim tempD As New SiteLong
    Dim temp0_7 As New SiteLong
    Dim temp8_15 As New SiteLong
    Dim temp16_23 As New SiteLong
    Dim temp24_31 As New SiteLong
    Dim site As Variant
    Dim AddrByte1 As Long
    Dim AddrByte2 As Long
    Dim AddrByte3 As Long
    Dim AddrByte4 As Long
    Dim TempArray(31) As Long
        
    
    Open RomFileName1 For Binary As #1
      
    '=============================================================
        AddrByte1 = 0
        AddrByte2 = 0
        AddrByte3 = 0
        AddrByte4 = 0
    While (Not EOF(1)) And (Total_Count < Addr8MByte)
        For Count = 4 To SrcDataSize - 1
            DspArray(Count) = 2 ^ 8 - 1
        Next Count
        
        Count = 4
        
        While ((Not EOF(1)) And (Count < SrcDataSize))
            Get #1, Total_Count + 1, Rom_Data
            DspArray(Count) = Rom_Data
            Count = Count + 1
            Total_Count = Total_Count + 1
            'Debug.Print Rom_Data
        Wend
        
        Count = 0
        
        DspArray(0) = AddrByte4
        DspArray(1) = AddrByte3
        DspArray(2) = AddrByte2
        DspArray(3) = AddrByte1
'        If AddrByte3 = 16 Then Stop
        AddrByte2 = AddrByte2 + 1
        If AddrByte2 > 255 Then '#16_ELSE_CASE_CHK
            AddrByte3 = AddrByte3 + 1
            AddrByte2 = 0
            If AddrByte3 > 255 Then
                AddrByte4 = AddrByte4 + 1
                AddrByte3 = 0
            Else
            End If
        Else
        End If
        
        'WaveDefinitions
        DspSrcWave.data = DspArray
        'DspSrcWave.Plot
        Call theexec.WaveDefinitions.CreateWaveDefinition("SrcData", DspSrcWave, True)

        'Insert DSSC Loading & Pattern Loading & Pattern Execution.
        With TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source
'        With TheHdw.DSSC.Pins(DigSrcPin).Pattern(PatName).Source
            .Signals.Add "DSSC_Src"
            .Signals.DefaultSignal = "DSSC_Src"
            .Signals("DSSC_Src").WaveDefinitionName = "SrcData"
            .Signals.item("DSSC_Src").sampleSize = 260
            .Signals("DSSC_Src").Amplitude = 1
            If LCase(glb_TesterType) = "jaguar" Then
                .Signals("DSSC_Src").LoadSettings
            End If
        End With

        TheHdw.Patterns(patt).test pfNever, 0
        TheHdw.Digital.Patgen.HaltWait

        WaveCount = WaveCount + 1
    Wend
        
    
    Close #1
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_Write_SPIROM", "PageProgram_4Byte")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function SPIROM_p2p_short_Power(PowerPins As PinList, ForceV As Double) As Long
On Error GoTo errHandler
'
   ''++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    'Testing Method:  Force 0.1V , measure smaller than 199ma,set clamp to 200ma, if higher than 199 ma then fail
    ''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Dim HexVSMeasure As New PinListData
    Dim hdvsMeasure As New PinListData
    Dim PowerMeasure As New PinListData
'    Dim PowerPins As String
    Dim p As Variant, Pin_Ary() As String, p_cnt As Long
    Dim site As Variant 'Carter, 20240304
'    PowerPins = AllHexVsPins & "," & AllUVSPins
       
'    On Error GoTo errHandler
    
'    TheHdw.digital.ApplyLevelsTiming False, True, False, tlPowered 'SEC DRAM
'    pwr_on_i_meter_DCVS AllHexVsPins.Value, 0#, 0.05, 0.01, 0.002, 10, 0.002  'set Force voltage and Current/Meter Range

'    No not need to add the code to avoid SPIROM power pin alarm 20170209
''    If TheExec.EnableWord("HardIP_Alarm_off") = True Then
''    '' 20160419 - Debug Alarm off
''    TheHdw.DCVS.Pins("spi_1v8").Alarm(tlDCVSAlarmAll) = tlAlarmOff
''    End If
    DCVS_PowerOn_I_Meter PowerPins.value, 0#, 200 * uA, 0.01, 0.002, 10, 0.002
    TheHdw.Wait 0.01
    theexec.DataManager.DecomposePinList PowerPins, Pin_Ary, p_cnt
    
    For Each p In Pin_Ary
        TheHdw.DCVS.Pins(p).Voltage.Main.value = ForceV
        TheHdw.Wait 0.05
        DCVS_MeterRead DCVS_UVS256, CStr(p), 10, PowerMeasure
        
        'offline mode simulation  20160328
        If theexec.TesterMode = testModeOffline Then    '#16_ELSE_CASE_CHK
            For Each site In theexec.sites
                PowerMeasure.Pins(p).value(site) = -0.005 + Rnd() * 0.0001
            Next site
        Else
        End If
        
        theexec.Flow.TestLimit resultVal:=PowerMeasure, ForceVal:=ForceV, ForceUnit:=unitVolt, ForceResults:=tlForceFlow
        
        If theexec.sites.Active.Count = 0 Then  '#16_ELSE_CASE_CHK
            Exit Function 'chihome
        Else
        End If
        TheHdw.DCVS.Pins(p).Voltage.Main.value = 0#
    Next p

    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_Write_SPIROM", "SPIROM_p2p_short_Power")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function SPIROM_Erase_Universal(PatName As Pattern, DigSrcPin As PinList, RelayOnPins As PinList, Optional EraseTime As Long = 200) As Long
On Error GoTo errHandler
    
    Dim site As Variant
    Dim DspRefWave As New DSPWave
    Dim DspSrcWave As New DSPWave
    Dim DspArray() As Long
    Dim CountSecond As Long
    Dim patt As String
    Dim PatCount As Long
    Dim PattArray() As String
    
    If Glb_SetEraseAddr = True Then
        ReDim DspArray(3) As Long
        DspArray(0) = EraseCMD
        DspArray(1) = 0
        DspArray(2) = 0
        DspArray(3) = 0
    Else
        ReDim DspArray(0) As Long
        DspArray(0) = EraseCMD
    End If
    
    TheHdw.Utility.Pins(RelayOnPins).State = tlUtilBitOn
    'Level & Timing
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    For Each site In theexec.sites
        CountSecond = theexec.Specs.AC.item("TCK_Freq_VAR").CurrentValue(site)
        Exit For
    Next site
    
    Call PATT_GetPatListFromPatternSet(PatName.value, PattArray, PatCount)
    patt = CStr(PattArray(0))
    TheHdw.Patterns(patt).Load
    
    If theexec.TesterMode = testModeOffline Then    '#16_ELSE_CASE_CHK
        EraseTime = 1
    Else
    End If
    If EraseTime > Glb_EraseTime Or EraseTime = 0 Then  '#16_ELSE_CASE_CHK
        EraseTime = Glb_EraseTime
    Else
    End If
    TheHdw.Digital.Patgen.TimeOut = EraseTime + 10
    
    'DigSrc Setup
    TheHdw.Digital.Patgen.counter(tlPgCounter2) = 8 * (UBound(DspArray) + 1) - 1
    TheHdw.Digital.Patgen.counter(tlPgCounter4) = CountSecond
    TheHdw.Digital.Patgen.counter(tlPgCounter1) = EraseTime
    DspSrcWave.data = DspArray
    Call theexec.WaveDefinitions.CreateWaveDefinition("SrcData", DspSrcWave, True)
    DspRefWave = DspSrcWave
        'Insert DSSC Loading & Pattern Loading & Pattern Execution.
    With TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source
        .Signals.Add "DSSC_Src"
        .Signals.DefaultSignal = "DSSC_Src"
        .Signals("DSSC_Src").WaveDefinitionName = "SrcData"
        .Signals.item("DSSC_Src").sampleSize = UBound(DspArray) + 1
        .Signals("DSSC_Src").Amplitude = 1
        If LCase(glb_TesterType) = "jaguar" Then    '#16_ELSE_CASE_CHK
            .Signals("DSSC_Src").LoadSettings
        Else
        End If
    End With
    
'    TheHdw.Patterns(Patt).start
    TheHdw.Patterns(patt).test pfAlways, 0
    TheHdw.Digital.Patgen.HaltWait
    
    theexec.Datalog.WriteComment vbNullString
    theexec.Datalog.WriteComment "print: Chip Data Erase Finished!!"

    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_Write_SPIROM", "SPIROM_Erase_Universal")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function SPIROM_SectorErase_Universal(PatName As Pattern, DigSrcPin As PinList, RelayOnPins As PinList) As Long
On Error GoTo errHandler
    
    Dim DspSrcWave As New DSPWave
    Dim Addr_Idx As Long
    Dim Addr_Ary(2) As Long
    Dim AddrByte1 As Long
    Dim AddrByte2 As Long
    Dim AddrByte3 As Long
    Dim patt As String
    Dim PatCount As Long
    Dim PattArray() As String
    Dim EraseCnt As Long
    Dim EraseSetp As Long
    
    AddrByte1 = 0
    AddrByte2 = 0
    AddrByte3 = 0
    
    TheHdw.Utility.Pins(RelayOnPins).State = tlUtilBitOn
    'Level & Timing
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    Call PATT_GetPatListFromPatternSet(PatName.value, PattArray, PatCount)
    patt = CStr(PattArray(0))
    TheHdw.Patterns(patt).Load
    TheHdw.Digital.Patgen.TimeOut = 10
    
    'Erase 16 bit Sector Data per time
    If Glb_RomSize <> 0 Then    '#16_ELSE_CASE_CHK
        If Glb_RomSize Mod 65536 = 0 Then
            EraseSetp = Glb_RomSize \ 65536
        Else
            EraseSetp = (Glb_RomSize \ 65536) + 1
        End If
    Else
    End If
    
    For Addr_Idx = 0 To EraseSetp
    
        If Addr_Idx = 0 Then
            AddrByte3 = 0
        Else
            AddrByte3 = AddrByte3 + 1
        End If
        
        Addr_Ary(0) = AddrByte3
        Addr_Ary(1) = AddrByte2
        Addr_Ary(2) = AddrByte1
        
        'WaveDefinitions
        DspSrcWave.data = Addr_Ary
        'DspSrcWave.Plot
        Call theexec.WaveDefinitions.CreateWaveDefinition("SrcData", DspSrcWave, True)

        'Insert DSSC Loading & Pattern Loading & Pattern Execution.
        With TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source
            .Signals.Add "DSSC_Src_Addr"
            .Signals.DefaultSignal = "DSSC_Src_Addr"
            .Signals("DSSC_SRC_Addr").WaveDefinitionName = "SrcData"
            .Signals.item("DSSC_Src_Addr").sampleSize = 3
            .Signals("DSSC_SRC_Addr").Amplitude = 1
            If LCase(glb_TesterType) = "jaguar" Then
                .Signals("DSSC_SRC_Addr").LoadSettings
            End If
        End With

        TheHdw.Patterns(patt).test pfNever, 0
        TheHdw.Digital.Patgen.HaltWait
        TheHdw.Wait 0.1
    Next Addr_Idx
    theexec.Datalog.WriteComment vbNullString
    theexec.Datalog.WriteComment "print: Sector Data Erase Bit Counts: " & EraseSetp * 65536 & " Finished!!"

    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_Write_SPIROM", "SPIROM_SectorErase_Universal")
    If AbortTest Then Exit Function Else Resume Next
End Function



Public Function SPIROM_InitialRead_Universal(PatName As Pattern, DigSrcPin As PinList, DigCapPin As PinList, RelayOnPins As PinList, RelayOffPins As PinList, Optional CompareByteCount As Long) As Long
On Error GoTo errHandler
Dim x As Long
Dim RomFileName1 As String
Dim Var_1 As Integer
Dim SPI_Flag As New SiteLong
Dim site As Variant

    x = CompareByteCount * MbytSize
    ResultPass = False
    'Offline Simulation
    If theexec.TesterMode = testModeOffline Then    '#16_ELSE_CASE_CHK
        For Each site In theexec.sites
            SPI_Flag = 2
            theexec.Flow.TestLimit SPI_Flag, 1, 1, , , , , , "CheckSum_First"
        Next site
        Exit Function
    Else
    End If

    If theexec.Flow.enableWord("SPIROM_1_Write") = True Then    '#16_ELSE_CASE_CHK
        RomFileName1 = Worksheets("SpiromCodeFile").Cells(1, 2).value
        gS_SPI_Version = RomFileName1
        CheckSum PatName, DigSrcPin, DigCapPin, RomFileName1, x, RelayOnPins
    Else
    End If
    
    If theexec.Flow.enableWord("SPIROM_2_Write") = True Then    '#16_ELSE_CASE_CHK
        RomFileName1 = Worksheets("SpiromCodeFile").Cells(2, 2).value
        gS_SPI_Version = RomFileName1
        CheckSum PatName, DigSrcPin, DigCapPin, RomFileName1, x, RelayOnPins
    Else
    End If
    
    If theexec.Flow.enableWord("SPIROM_3_Write") = True Then    '#16_ELSE_CASE_CHK
        RomFileName1 = Worksheets("SpiromCodeFile").Cells(3, 2).value
        gS_SPI_Version = RomFileName1
        CheckSum PatName, DigSrcPin, DigCapPin, RomFileName1, x, RelayOnPins
    Else
    End If
    
    If CheckSumResult = 1 Then
        Set SPI_Binary_Erase_Data = Nothing
        Set SPI_Binary_Write_Data = Nothing
        For Each site In theexec.sites
            SPI_Flag = 1
            theexec.Flow.TestLimit SPI_Flag, 1, 1, , , , , , "CheckSum_First"
        Next site
        
        For Each site In theexec.sites
            If write_spirom = True Then
                theexec.Datalog.WriteComment ("site(" & site & ") has been writen")
            End If
            write_spirom = False
        Next site
        ''control whether the program execute again or not.
        theexec.Datalog.WriteComment ("print: SPIROM version check pass with " & gS_SPI_Version)
        theexec.Flow.enableWord("Write_SPIROM") = False
        TheHdw.Utility.Pins(RelayOffPins).State = tlUtilBitOff
    Else
        For Each site In theexec.sites
            SPI_Flag = 0
            theexec.Flow.TestLimit SPI_Flag, 1, 1, , , , , , "CheckSum_First"
        Next site
        
        For Each site In theexec.sites
            If ResultPass = False Then
                theexec.Datalog.WriteComment "Site: " & site & " RomData Check Fail !!!"
            End If
        Next site
        theexec.Datalog.WriteComment ("print: SPIROM version check fail with " & gS_SPI_Version)
        'TheExec.Datalog.WriteComment ("The SPIROM function is not successful.")
    End If
  
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_Write_SPIROM", "SPIROM_InitialRead_Universal")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Read_RomCode(RomFileName1 As String, RomSize As Long, CapNumOfTimes As String) As Long
On Error GoTo errHandler

    Dim TotalDspArray() As Long
    Dim DspArray1() As Long
    Dim DspArray2() As Long
    Dim Count As Long
    Dim Rom_Data As Byte
    Dim StartAddr As Long
    Dim Cap_1st As String
    Dim Cap_2nd As String
    Dim Cap_3rd As String
    Dim Cap_4th As String
    Dim Cap_5th As String
    Dim Cap_6th As String
    Dim Cap_7th As String
    Dim Cap_8th As String
    Dim CapSignalAry() As String
    Dim CapIdx As Long
    Dim SPI_Binary_Signal1 As New DSPWave
    Dim SPI_Binary_Signal2 As New DSPWave
    Dim SPI_Binary_Erase1 As New DSPWave
    Dim SPI_Binary_Erase2 As New DSPWave

    ReDim DspArray1(Glb_CaptureSizePerPat - 1) As Long
    CapSignalAry = Split(CapNumOfTimes, "_")
    
    CapIdx = CLng(CapSignalAry(UBound(CapSignalAry)))
    StartAddr = CapIdx * Glb_CaptureSizePerPat
    
    
    Total_Count = StartAddr
    
    For Count = 0 To Glb_CaptureSizePerPat - 1
        DspArray1(Count) = 2 ^ 8 - 1
    Next Count

    SPI_Binary_Erase1.data = DspArray1
    SPI_Binary_Erase_Data = SPI_Binary_Erase1.Copy

    Open RomFileName1 For Binary As #1
                      
        'Create data array for reference
        Count = 0
        While ((Not EOF(1)) And (Count < Glb_CaptureSizePerPat))
            Get #1, Total_Count + 1, Rom_Data
            If EOF(1) = False Then  '#16_ELSE_CASE_CHK
                While ((Not EOF(1)) And (Count < Glb_CaptureSizePerPat))
                    Get #1, Total_Count + 1, Rom_Data
                    DspArray1(Count) = Rom_Data
                    Count = Count + 1
                    Total_Count = Total_Count + 1
                    'Debug.Print Rom_Data
                Wend
            Else
            End If
        Wend
        SPI_Binary_Signal1.data = DspArray1
        SPI_Binary_Write_Data = SPI_Binary_Signal1.Copy
    
    Close #1
    
    Read_RomCode = Total_Count
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_Write_SPIROM", "Read_RomCode")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function InitRead(PatName As Pattern, DigSrcPin As PinList, DigCapPin As PinList, RomSize As Long, RomFileName1 As String, StartingAddr As Integer, RelayOnPins As PinList) As Long
On Error GoTo errHandler
    
    Dim Rom_Data As Byte
    Dim TotalDspArray() As Long
    Dim DspArray1() As Long
    Dim DspArray2() As Long
    Dim Count As Long
    Dim SrcWaveName As String
    Dim StartSTR As String
    Dim result As New SiteDouble
    Dim WaveCount As Long
    Dim SrcDataSize As Long
    Dim DSSCCapSize As Long
    Dim DspSrcWave As New DSPWave
    Dim DspWaveCap As New DSPWave
    Dim DspRefWave As New DSPWave
    Dim DSSCCap_MByte_Signal As New DSPWave
    Dim DSSCCap_All_Signal_SUM As New DSPWave
    Dim DspArray() As Long
    
    Dim PatCount As Long
    Dim PattArray() As String
    Dim patt As String

    TheHdw.Utility.Pins(RelayOnPins).State = tlUtilBitOn
    'Level & Timing
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.Wait 0.5

    Call PATT_GetPatListFromPatternSet(PatName.value, PattArray, PatCount)
    patt = CStr(PattArray(0))
    TheHdw.Patterns(patt).Load
    TheHdw.Digital.Patgen.TimeOut = 10
    
    DSSCCapSize = Glb_CaptureSizePerPat
    DspWaveCap.Clear
                       
    Dim AddrIndx As Integer
    Dim AddrSize As Integer
    Dim ReadCmd As Long
    Dim Addr4 As Long: Addr4 = 0
    Dim Addr3 As Long: Addr3 = 0
    Dim PatternTestCnt As Long
    
    PatternTestCnt = Glb_CaptureSizePerPat \ MbytSize
    If RomSize <= Addr16MByte Then
        TheHdw.Digital.Patgen.counter(tlPgCounter0) = 16 - 1
        ReadCmd = 11 '3Byte Fast Read(0Bh)
        ReDim DspArray(1) As Long
        Addr3 = (0 + StartingAddr * 16) * PatternTestCnt
        DspArray(0) = ReadCmd
        DspArray(1) = Addr3
    ElseIf RomSize > Addr16MByte And RomSize <= Addr128MByte Then
        TheHdw.Digital.Patgen.counter(tlPgCounter0) = 24 - 1
        ReadCmd = 12 '4Byte Fast Read(0Ch)
        ReDim DspArray(2) As Long
        Addr3 = (0 + StartingAddr * 16) * PatternTestCnt
        If Addr3 > 255 Then
            Addr4 = Addr3 \ 256
            Addr3 = Addr3 Mod 256
        End If
            DspArray(0) = ReadCmd
            DspArray(1) = Addr4
            DspArray(2) = Addr3
    Else
        theexec.Datalog.WriteComment "Please Check the ROM File Size!!"
    End If
    
    TheHdw.Digital.Patgen.counter(tlPgCounter1) = (8 * DSSCCapSize) - 1   'Capture Byte Count
    DspSrcWave.data = DspArray
    Call theexec.WaveDefinitions.CreateWaveDefinition("SrcData", DspSrcWave, True)
    DspRefWave = DspSrcWave
        'Insert DSSC Loading & Pattern Loading & Pattern Execution.
    With TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source
        .Signals.Add "DSSC_Src"
        .Signals.DefaultSignal = "DSSC_Src"
        .Signals("DSSC_Src").WaveDefinitionName = "SrcData"
        .Signals.item("DSSC_Src").sampleSize = UBound(DspArray) + 1
        .Signals("DSSC_Src").Amplitude = 1
        If LCase(glb_TesterType) = "jaguar" Then
            .Signals("DSSC_Src").LoadSettings
        End If
    End With
    'Setup DigCap
    DSSCCap_MByte_Signal.Clear
    With TheHdw.DSSC.Pins(DigCapPin).Pattern(patt).Capture
        .Signals.Add "DSSC_Cap_All"
        If LCase(glb_TesterType) = "jaguar" Then
            .Signals("DSSC_Cap_All").offset = 0
        End If
        .Signals.item("DSSC_Cap_All").sampleSize = DSSCCapSize
        .Signals("DSSC_Cap_All").LoadSettings
    End With
    TheHdw.Wait 0.05
    TheHdw.Patterns(patt).start
    'thehdw.Patterns(patt).test pfNever, 0
    TheHdw.Digital.Patgen.HaltWait
             
    DSSCCap_MByte_Signal = TheHdw.DSSC.Pins(DigCapPin).Pattern(patt).Capture.Signals("DSSC_Cap_All").DSPWave
    DspWaveCap = DSSCCap_MByte_Signal
    
    TheHdw.Wait 0.05
    
    Dim Different_Value As New DSPWave
    Dim Abs_Different_Value As New DSPWave
    Dim IndexOfMaximumValue As Long
    Dim IndexOfMinimumValue As Long
    Dim MaximumValue As Double
    Dim MinimumValue As Double
    Dim site As Variant
    Dim Sum As New DSPWave
    'Capture Data Compare ===================================================
    For Each site In theexec.sites
        Sum = DspWaveCap
        Different_Value = SPI_Binary_Write_Data.Subtract(Sum)
        Abs_Different_Value = Different_Value.Abs
        MaximumValue = Abs_Different_Value.CalcMaximumValue(IndexOfMaximumValue)
        ReDim Preserve SPI_Flag_Ary(StartingAddr) As New SiteBoolean
        ReDim Preserve SPI_Flag_temp_Ary(StartingAddr) As New SiteLong
        If MaximumValue = 0 Then
          SPI_Flag_Ary(StartingAddr) = True                 'SPI-ROM PROGRAMMED with current ROM CODE
          SPI_Flag_temp_Ary(StartingAddr) = 1
        Else
          SPI_Flag_Ary(StartingAddr) = False                'SPI-ROM Erase
          SPI_Flag_temp_Ary(StartingAddr) = 2
        End If
    Next site
    '========================================================================
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_Write_SPIROM", "InitRead")
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function SPIROM_CheckSum_Universal(PatName As Pattern, DigSrcPin As PinList, DigCapPin As PinList, RelayOnPins As PinList, RelayOffPins As PinList, Optional CompareByteCount As Long = 8) As Long
On Error GoTo errHandler
    Dim x As Long
    Dim RomFileName1 As String
    Dim SPI_Flag As New SiteLong
    Dim site As Variant

    x = CompareByteCount * MbytSize
    ResultPass = False
    'Offline Simulation
    If theexec.TesterMode = testModeOffline Then    '#16_ELSE_CASE_CHK
        For Each site In theexec.sites
            SPI_Flag = 1
            theexec.Flow.TestLimit SPI_Flag, 1, 1, , , , , , "CheckSum_Second"
        Next site
        Exit Function
    Else
    End If
    
    If theexec.Flow.enableWord("SPIROM_1_Write") = True Then    '#16_ELSE_CASE_CHK
        RomFileName1 = Worksheets("SpiromCodeFile").Cells(1, 2).value
        gS_SPI_Version = RomFileName1
        CheckSum PatName, DigSrcPin, DigCapPin, RomFileName1, x, RelayOnPins
    Else
    End If
    
    If theexec.Flow.enableWord("SPIROM_2_Write") = True Then    '#16_ELSE_CASE_CHK
        RomFileName1 = Worksheets("SpiromCodeFile").Cells(2, 2).value
        gS_SPI_Version = RomFileName1
        CheckSum PatName, DigSrcPin, DigCapPin, RomFileName1, x, RelayOnPins
    Else
    End If
    
    If theexec.Flow.enableWord("SPIROM_3_Write") = True Then    '#16_ELSE_CASE_CHK
        RomFileName1 = Worksheets("SpiromCodeFile").Cells(3, 2).value
        gS_SPI_Version = RomFileName1
        CheckSum PatName, DigSrcPin, DigCapPin, RomFileName1, x, RelayOnPins
    Else
    End If
    
    If CheckSumResult = 1 Then
        Set SPI_Binary_Erase_Data = Nothing
        Set SPI_Binary_Write_Data = Nothing
        For Each site In theexec.sites
            SPI_Flag = 1
            theexec.Flow.TestLimit SPI_Flag, 1, 1, , , , , , "CheckSum_Second"
        Next site
        
        For Each site In theexec.sites
            If write_spirom = True Then
                theexec.Datalog.WriteComment ("site(" & site & ") has been writen")
            End If
            write_spirom = False
        Next site
        ''control whether the program execute again or not.
        theexec.Datalog.WriteComment ("print: Auto-trim finished, SPIROM version check pass " & gS_SPI_Version)
        theexec.Flow.enableWord("Write_SPIROM") = False
        TheHdw.Utility.Pins(RelayOffPins).State = tlUtilBitOff
    Else
        For Each site In theexec.sites
            SPI_Flag = 0
            theexec.Flow.TestLimit SPI_Flag, 1, 1, , , , , , "CheckSum_Second"
        Next site
        
        For Each site In theexec.sites
            If ResultPass = False Then
                theexec.Datalog.WriteComment "Site: " & site & " RomData Check Fail !!!"
            End If
        Next site
        theexec.Datalog.WriteComment ("print: Auto-trim finished, SPIROM version check fail " & gS_SPI_Version)
        'SPI_Version = SPIROM_code_version & " checksum failed"
        theexec.Datalog.WriteComment ("The SPIROM function is not successful.")
    End If
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_Write_SPIROM", "SPIROM_CheckSum_Universal")
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function SPIROM_Program_Universal(PatName As Pattern, DigSrcPin As PinList, RelayOnPins As PinList) As Long
On Error GoTo errHandler

    Dim x As Long
    Dim RomFileName1 As String

    If theexec.Flow.enableWord("SPIROM_1_Write") = True Then    '#16_ELSE_CASE_CHK
      RomFileName1 = Worksheets("SpiromCodeFile").Cells(1, 2).value
        
        theexec.Datalog.WriteComment (vbNullString)
        theexec.Datalog.WriteComment (RomFileName1 + " start to program.")
        Call PageProgram(PatName, RomFileName1, RelayOnPins, DigSrcPin)
        theexec.Datalog.WriteComment (RomFileName1 + " has been programmed.")
    Else
    End If
    
    If theexec.Flow.enableWord("SPIROM_2_Write") = True Then   '#16_ELSE_CASE_CHK
      RomFileName1 = Worksheets("SpiromCodeFile").Cells(2, 2).value
        
        theexec.Datalog.WriteComment (vbNullString)
        theexec.Datalog.WriteComment (RomFileName1 + " start to program.")
        Call PageProgram(PatName, RomFileName1, RelayOnPins, DigSrcPin)
        theexec.Datalog.WriteComment (RomFileName1 + " has been programmed.")
    Else
    End If

    If theexec.Flow.enableWord("SPIROM_3_Write") = True Then    '#16_ELSE_CASE_CHK
      RomFileName1 = Worksheets("SpiromCodeFile").Cells(3, 2).value
        
        theexec.Datalog.WriteComment (vbNullString)
        theexec.Datalog.WriteComment (RomFileName1 + " start to program.")
        Call PageProgram(PatName, RomFileName1, RelayOnPins, DigSrcPin)
        theexec.Datalog.WriteComment (RomFileName1 + " has been programmed.")
    Else
    End If

    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_Write_SPIROM", "SPIROM_Program_Universal")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function SPIROM_Program_4ByteAddr_Universal(PatName As Pattern, DigSrcPin As PinList, RelayOnPins As PinList) As Long
On Error GoTo errHandler

    Dim x As Long
    Dim RomFileName1 As String

    If theexec.Flow.enableWord("SPIROM_1_Write") = True Then    '#16_ELSE_CASE_CHK
      RomFileName1 = Worksheets("SpiromCodeFile").Cells(1, 2).value
        
        theexec.Datalog.WriteComment (vbNullString)
        theexec.Datalog.WriteComment (RomFileName1 + " start to program.")
        Call PageProgram_4Byte(PatName, RomFileName1, RelayOnPins, DigSrcPin)
        theexec.Datalog.WriteComment (RomFileName1 + " has been programmed.")
    Else
    End If
    
    If theexec.Flow.enableWord("SPIROM_2_Write") = True Then    '#16_ELSE_CASE_CHK
      RomFileName1 = Worksheets("SpiromCodeFile").Cells(2, 2).value
        
        theexec.Datalog.WriteComment (vbNullString)
        theexec.Datalog.WriteComment (RomFileName1 + " start to program.")
        Call PageProgram_4Byte(PatName, RomFileName1, RelayOnPins, DigSrcPin)
        theexec.Datalog.WriteComment (RomFileName1 + " has been programmed.")
    Else
    End If

    If theexec.Flow.enableWord("SPIROM_3_Write") = True Then    '#16_ELSE_CASE_CHK
      RomFileName1 = Worksheets("SpiromCodeFile").Cells(3, 2).value
        
        theexec.Datalog.WriteComment (vbNullString)
        theexec.Datalog.WriteComment (RomFileName1 + " start to program.")
        Call PageProgram_4Byte(PatName, RomFileName1, RelayOnPins, DigSrcPin)
        theexec.Datalog.WriteComment (RomFileName1 + " has been programmed.")
    Else
    End If

    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_Write_SPIROM", "SPIROM_Program_4ByteAddr_Universal")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function SPIROM_Device_Check_Universal(PatName As Pattern, DigCapPin As PinList, RelayOnPins As PinList, SPIROM_MSize As Long)
On Error GoTo errHandler

    Dim site As Variant
    Dim RomSize As New SiteLong
    Dim RomFileName1 As String
    Dim DSSCCap_Device_Signal As New DSPWave
    Dim patt As String
    Dim PattArray() As String
    Dim PatCount As Long
    Dim MemoryLimit As Long
    Dim ItemName As String

    If theexec.RunMode = runModeProduction Then '#16_ELSE_CASE_CHK
        theexec.Flow.enableWord("Write_SPIROM_SectorErase") = False
        theexec.Flow.enableWord("Write_SPIROM") = True
    Else
    End If
    'Assign Specify SPIROM Information
    Select Case SPIROM_MSize
    Case 128
        MemoryLimit = Addr128MByte
    Case 32
        MemoryLimit = Addr32MByte
    Case 8
        MemoryLimit = Addr8MByte
    Case Else
    End Select
    ItemName = "SPIROM_Testing"
    ''''''''''''''''''''''''''''''
    Call PATT_GetPatListFromPatternSet(PatName.value, PattArray, PatCount)
    patt = CStr(PattArray(0))
          
    Dim CurrPath As String
    CurrPath = Application.ActiveWorkbook.path
    'CurrPath = TheExec.TestProgram.path
    If theexec.Flow.enableWord("SPIROM_1_Write") = True Then    '#16_ELSE_CASE_CHK
        RomFileName1 = Worksheets("SpiromCodeFile").Cells(1, 2).value
        CurrPath = CurrPath & mid(RomFileName1, 2)
                 
        If RomFileName1 = "" Then   '#16_ELSE_CASE_CHK
            For Each site In theexec.sites
                theexec.Flow.TestLimit 0, 1, 1, , , , , , "RomFileError"
                theexec.Datalog.WriteComment ("Romcode File Did NOT Specify. Plese Check SpiromCodeFile Sheet  ")
            Next site
            Exit Function
        Else
        End If
        'check if the file exists or not
        If Dir(RomFileName1) = "" Then  '#16_ELSE_CASE_CHK
            For Each site In theexec.sites
                theexec.Flow.TestLimit 0, 1, 1, , , , , , "RomFileError"
                theexec.Datalog.WriteComment (" Binary File Does NOT Exist. Please Check Binary Directory  ")
            Next site
            Exit Function
        Else
        End If
        
        For Each site In theexec.sites
            RomSize = Get_RomSize(RomFileName1)
        Next site
    Else
    End If
          
    If theexec.Flow.enableWord("SPIROM_2_Write") = True Then    '#16_ELSE_CASE_CHK
        RomFileName1 = Worksheets("SpiromCodeFile").Cells(2, 2).value
        CurrPath = CurrPath & mid(RomFileName1, 2)
                
        If RomFileName1 = "" Then   '#16_ELSE_CASE_CHK
            For Each site In theexec.sites
                theexec.Flow.TestLimit 0, 1, 1, , , , , , "RomFileError"
                theexec.Datalog.WriteComment ("Romcode File Did NOT Specify. Plese Check SpiromCodeFile Sheet  ")
            Next site
            Exit Function
        Else
        End If
                 
        'check if the file exists or not
        If Dir(RomFileName1) = "" Then  '#16_ELSE_CASE_CHK
            For Each site In theexec.sites
                theexec.Flow.TestLimit 0, 1, 1, , , , , , "RomFileError"
                theexec.Datalog.WriteComment (" Binary File Does NOT Exist. Please Check Binary Directory  ")
            Next site
            Exit Function
        Else
        End If
        For Each site In theexec.sites
            RomSize = Get_RomSize(RomFileName1)
        Next site
    Else
    End If
                   
    If theexec.Flow.enableWord("SPIROM_3_Write") = True Then    '#16_ELSE_CASE_CHK
        RomFileName1 = Worksheets("SpiromCodeFile").Cells(3, 2).value
        CurrPath = CurrPath & mid(RomFileName1, 2)
                
        If RomFileName1 = "" Then   '#16_ELSE_CASE_CHK
            For Each site In theexec.sites
                theexec.Flow.TestLimit 0, 1, 1, , , , , , "RomFileError"
                theexec.Datalog.WriteComment ("Romcode File Did NOT Specify. Plese Check SpiromCodeFile Sheet  ")
            Next site
            Exit Function
        Else
        End If
                 
        'check if the file exists or not
        If Dir(RomFileName1) = "" Then  '#16_ELSE_CASE_CHK
            For Each site In theexec.sites
                theexec.Flow.TestLimit 0, 1, 1, , , , , , "RomFileError"
                theexec.Datalog.WriteComment (" Binary File Does NOT Exist. Please Check Binary Directory  ")
            Next site
            Exit Function
        Else
        End If
        For Each site In theexec.sites
            RomSize = Get_RomSize(RomFileName1)
        Next site
    Else
    End If
                   
    TheHdw.Utility.Pins(RelayOnPins).State = tlUtilBitOn
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlUnpowered
    TheHdw.Patterns(patt).Load

    'Setup DigCap
    With TheHdw.DSSC.Pins(DigCapPin).Pattern(patt).Capture
        .Signals.Add "DSSC_Cap_All"
        If LCase(glb_TesterType) = "jaguar" Then    '#16_ELSE_CASE_CHK
            .Signals("DSSC_Cap_All").offset = 0
        Else
        End If
        .Signals.item("DSSC_Cap_All").sampleSize = 3
        .Signals("DSSC_Cap_All").LoadSettings
    End With
    
    TheHdw.Patterns(PatName).start
    TheHdw.Digital.Patgen.HaltWait
    
    DSSCCap_Device_Signal = TheHdw.DSSC.Pins(DigCapPin).Pattern(patt).Capture.Signals("DSSC_Cap_All").DSPWave
    Dim DeviceInfo() As Long
    For Each site In theexec.sites
        DeviceInfo = DSSCCap_Device_Signal(site).data
        'For Offline Simulation
        If theexec.TesterMode = testModeOffline Then    '#16_ELSE_CASE_CHK
            ReDim DeviceInfo(2) As Long
            Select Case SPIROM_MSize
            Case 128
                DeviceInfo(0) = CLng("&H" & "20")
                DeviceInfo(1) = CLng("&H" & "BB")
                DeviceInfo(2) = CLng("&H" & "21")
            Case 32
                DeviceInfo(0) = CLng("&H" & "20")
                DeviceInfo(1) = CLng("&H" & "BB")
                DeviceInfo(2) = CLng("&H" & "19")
            Case 8
                DeviceInfo(0) = CLng("&H" & "EF")
                DeviceInfo(1) = CLng("&H" & "65")
                DeviceInfo(2) = CLng("&H" & "17")
            Case Else
            End Select
        Else
        End If
            
        Select Case Hex(DeviceInfo(0))
        Case "20" 'Micron Device 32MByte/128MByte   #16_ELSE_CASE_CHK
            If Hex(DeviceInfo(1)) = "BB" And Hex(DeviceInfo(2)) = "19" Then
                Glb_SetEraseAddr = False
                Glb_EraseTime = 200
                EraseCMD = 199
                theexec.Datalog.WriteComment "The Site(" & CStr(site) & ") is 32M SPIROM Device."
                theexec.Flow.TestLimit 32 * MbytSize, MemoryLimit, MemoryLimit, , , , , , ItemName
            ElseIf Hex(DeviceInfo(1)) = "BB" And Hex(DeviceInfo(2)) = "21" Then
                Glb_SetEraseAddr = True
                Glb_EraseTime = 920
                EraseCMD = 196
                theexec.Datalog.WriteComment "The Site(" & CStr(site) & ") is 128M SPIROM Device."
                theexec.Flow.TestLimit 128 * MbytSize, MemoryLimit, MemoryLimit, , , , , , ItemName
            Else
            End If
            
        Case "EF" 'Winbond Device 8MByte    #16_ELSE_CASE_CHK
            If Hex(DeviceInfo(1)) = "65" And Hex(DeviceInfo(2)) = "17" Then
                Glb_SetEraseAddr = False
                Glb_EraseTime = 160
                EraseCMD = 199
                theexec.Datalog.WriteComment "The Site(" & CStr(site) & ") is 8M SPIROM Device."
                theexec.Flow.TestLimit 8 * MbytSize, MemoryLimit, MemoryLimit, , , , , , ItemName
            Else
            End If
            
        Case Else
            theexec.Flow.TestLimit 0, MemoryLimit, MemoryLimit, , , , , , ItemName
            theexec.Datalog.WriteComment "Error Detect The Type of Site(" & CStr(site) & ") is SPIROM Device."
        End Select
        
        theexec.Datalog.WriteComment "Site(" & site & ")" & "Manufacture ID: " & Hex(DeviceInfo(0)) & ", Device ID: " & Hex(DeviceInfo(1)) & " " & Hex(DeviceInfo(2))
    Next site

    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_Write_SPIROM", "SPIROM_Device_Check_Universal")
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function CheckSum(PatName As Pattern, DigSrcPin As PinList, DigCapPin As PinList, RomFileName1 As String, x As Long, RelayOnPins As PinList) As Long
On Error GoTo errHandler
    
    Dim Cap_1st As String
    Dim Cap_2nd As String
    Dim Cap_3rd As String
    Dim Cap_4th As String
    Dim Cap_5th As String
    Dim Cap_6th As String
    Dim Cap_7th As String
    Dim Cap_8th As String
    Dim i As Long
    Dim Cap_Idx As Long
    Dim CapSize As Long
    Dim CapSignal() As Long
    Dim site As Variant 'Carter, 20240304
    Cap_Idx = -1
    CapSize = Glb_CaptureSizePerPat

    If x \ CapSize = 0 Then
        Cap_Idx = 0
    Else
        If x Mod CapSize = 0 Then
            Cap_Idx = x \ CapSize - 1
        Else
            Cap_Idx = x \ CapSize
        End If
    End If
    
    For i = 0 To Cap_Idx
        Read_RomCode RomFileName1, x, "Cap_" & CStr(i)
        Call InitRead(PatName, DigSrcPin, DigCapPin, x, RomFileName1, CInt(i), RelayOnPins)
        theexec.Datalog.WriteComment "The Number of Captcure Signal is " & i + 1
    Next i
    
    For Each site In theexec.sites
        For i = 0 To Cap_Idx
            If i = 0 Then
                SPI_Flag_Sum = SPI_Flag_temp_Ary(i)
            Else
                SPI_Flag_Sum = SPI_Flag_Sum And SPI_Flag_temp_Ary(i)
            End If
            If SPI_Flag_Sum = 1 Then    '#16_ELSE_CASE_CHK
                ResultPass = True
            Else
            End If
        Next
    Next site
    
    Dim SiteCount2Dec As Long
    Dim ExitingCnt As Long
    Dim j As Long
    ExitingCnt = theexec.sites.Existing.Count
    
    For j = ExitingCnt To 1 Step -1
        SiteCount2Dec = SiteCount2Dec + (2 ^ (j - 1))
    Next j
    CheckSumResult = SiteCount2Dec
    For Each site In theexec.sites
        CheckSumResult = SPI_Flag_Sum(site) And SiteCount2Dec And CheckSumResult
    Next site
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_Write_SPIROM", "CheckSum")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function SPIROM_ROMFile_Check() As Boolean
On Error GoTo errHandler

    Dim RomCodeSheetName As String
    Dim StageAry() As String
    Dim RomFileNameAry() As String
    Dim FileIdx As Long
    Dim StageCnt As Long
    Dim i As Long
    
    RomCodeSheetName = "SpiromCodeFile"
    
    FileIdx = 2
    StageCnt = 1
    While Worksheets(RomCodeSheetName).Cells(StageCnt, FileIdx).value <> ""
        ReDim Preserve RomFileNameAry(StageCnt - 1) As String
        RomFileNameAry(StageCnt - 1) = Worksheets(RomCodeSheetName).Cells(StageCnt, FileIdx).value
        StageCnt = StageCnt + 1
    Wend

    For i = 0 To UBound(RomFileNameAry)
        If RomFileNameAry(i) = "" Or Dir(RomFileNameAry(i)) = "" Then
            theexec.Datalog.WriteComment "Please Call TE to Figure Out Whether the ROM File Exist or Not."  '#16_ELSE_CASE_CHK
            GoTo errHandler
        Else
        End If
    Next i
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_Write_SPIROM", "SPIROM_ROMFile_Check")
    If AbortTest Then Exit Function Else Resume Next
End Function
