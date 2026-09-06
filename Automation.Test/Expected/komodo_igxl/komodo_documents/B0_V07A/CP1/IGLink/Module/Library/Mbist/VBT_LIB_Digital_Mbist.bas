Attribute VB_Name = "VBT_LIB_Digital_Mbist"
#Const isUFP = True
Option Explicit
'Revision History:
'V0.0 initial bring up

Public Type MBIST_BLOCK_TYPE
    MemArray() As Long
    MemStrArray() As String
    MaxRow As Long
    strServerName() As String
    strDecsName() As String
End Type

Public MbistBlock() As MBIST_BLOCK_TYPE
Public MbistBlockName

Public tpEvaPattCycleBlockInfor() As EvaPattMbistCycleBlock

Private Type EvaPattMbistCycleBlock
    strBlaclName As String
    lVector As Long
    lCycle As Long
    strCompare As String
    strFlagName As String   '' Add for MbistFP Binning          201606014 Webster
End Type


Private Type MbistCycleBlock
    strPattName As String
    tpMbistCycleBlock() As EvaPattMbistCycleBlock
    strServerName() As String
    strDecsName() As String
    Dic_VectorIdx As New Dictionary
    Dic_CycleIdx As New Dictionary
End Type

Public tpCycleBlockInfor() As MbistCycleBlock

Public gl_MbistFP_Binout As Boolean      ''201606014 webster
Dim gl_strFlagArr() As String
Private Type FlagInfo
    FlagName As String
    CheckInfo As Boolean
End Type
Public tyFlagInfoArr() As FlagInfo

Public MFP_TableIndex As New Dictionary

Public MbistERT_TargetVol As New Dictionary
Public MbistERT_OriginVol As New Dictionary
Public MbistERT_DropVol As New Dictionary
Public MbistERT_DropVol_PerSite As New Dictionary
Public MbistERT_GroupCurrentVol As New Dictionary
Public MbistERT_GroupPreviousVol As New Dictionary
Public MbistERT_GroupPinNmame As New Dictionary
Public lclPinListData_CurrentCorePower As New PinListData
Public lclPinListData_PreviousCorePower As New PinListData
Public Const Group_Name = "Group_"

Enum RET_RampingDir
    RampUp = 0
    RampDown = 1
End Enum

Public Function Init_RSCR()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim k As Integer

    Const start_row = 2
    Const end_col = 4
    If Flag_RSCR_INIT = False Then

        ' ==================== New Mbist RSCR ====================
        Dim ws As Worksheet
        Dim blockName As String
        Dim blockNum As Long

        Application.ScreenUpdating = False
        blockNum = 0
        Set MbistBlockName = CreateObject("Scripting.Dictionary")
        MbistBlockName.compareMode = 1

        Dim sheetnames() As String

        sheetnames = TheExec.Job.GetSheetNamesOfType(DMGR_SHEET_TYPE_USER)

        Dim indx As Integer

        
        For indx = 0 To UBound(sheetnames)
            If sheetnames(indx) Like "RSCR*" Then
                blockName = Right(sheetnames(indx), Len(sheetnames(indx)) - 5)
                If Not MbistBlockName.Exists(blockName) Then
                    MbistBlockName.Add blockName, blockNum
                    ReDim Preserve MbistBlock(blockNum)
                End If

                Dim MaxRow As Long
                MaxRow = Worksheets(sheetnames(indx)).UsedRange.Rows.Count
                
                Dim maxcol As Long
                maxcol = Worksheets(sheetnames(indx)).UsedRange.Columns.Count
                
                MbistBlock(blockNum).MaxRow = MaxRow
                ReDim MbistBlock(blockNum).MemArray(MaxRow - 2)
                ReDim MbistBlock(blockNum).MemStrArray(MaxRow - 2)

                Dim arr1() As Variant
                Worksheets(sheetnames(indx)).Activate
                arr1 = Worksheets(sheetnames(indx)).range(Cells(start_row, 1), Cells(MaxRow, maxcol)).value
                Dim i As Integer

                For i = 1 To MaxRow - 1
                     'MbistBlock(BlockNum).MemArray(i - 1) = Int(arr1(i, 1))
                     'MbistBlock(BlockNum).MemStrArray(i - 1) = arr1(i, 2) & " " & arr1(i, 3) & " " & arr1(i, 4)
                     MbistBlock(blockNum).MemArray(i - 1) = Int(arr1(i, 2))
                     MbistBlock(blockNum).MemStrArray(i - 1) = arr1(i, 3)
                Next i
                
                ReDim MbistBlock(blockNum).strDecsName(maxcol - 5)
                ReDim MbistBlock(blockNum).strServerName(maxcol - 5)
                For i = 5 To maxcol
                    Dim TempAry() As String
                    If i = 5 Then 'default setting
                        TempAry() = Split(",", ",")
                    ElseIf arr1(1, i) = Empty Then
                        TempAry() = Split(",", ",")
                    Else
                        TempAry() = Split(arr1(1, i), ",")
                        If UBound(TempAry()) = 0 Then
                            ReDim Preserve TempAry(1)
                            TempAry(1) = vbNullString
                        End If
                    End If
                    MbistBlock(blockNum).strServerName(i - 5) = TempAry(0)
                    MbistBlock(blockNum).strDecsName(i - 5) = TempAry(1)
                Next i
                
                blockNum = blockNum + 1
            End If
        Next indx

        ' ==================== New Mbist RSCR ====================

        TheExec.Datalog.WriteComment "print: RSCR table initialized complete"

    End If
    Flag_RSCR_INIT = True
    Application.ScreenUpdating = True

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Mbist", "Init_RSCR") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


'[20230710][T-ALL][Tank] Si use .PatternBurstPassedPerSite to TTR, but she modify before run pattern. It's wrong.
Public Function Mbist_RSCR(Shift_Pat As Pattern, MBIST_BLOCK As String, Optional Server As String, Optional Descri As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18


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
    
    Dim M_maxrow As Long
    Dim M_MbistNum As Long
    Dim NewFmtAry() As String
    ReDim NewFmtAry(TheExec.sites.Existing.Count - 1)
    Dim OldFmtAry() As String
    Dim LineOffset As Long: LineOffset = 0
    Dim NumCapSum As Long: NumCapSum = 0
    Dim EmptyAry As Boolean: EmptyAry = True
    Dim RSCRHead As String: RSCRHead = "1"
    Dim patPassed As New SiteBoolean
    Dim TestNumber As Long
    Dim DecomPatName() As String
    
    SampleNum = 120
    CNumber_plus = 0 ' pattern has dummy cycle
    ''''''''''''''''''

    TheHdw.Digital.Patgen.HaltMode = tlHaltOnHRAMFull
    maxDepth = gl_HRAMmaxDepth            'maxDepth = TheHdw.Digital.HRAM.maxDepth    'Flex UP1600 max depth 512    'Plus org max depth 16k, but use 512 to let speed faster
    TheHdw.Digital.HRAM.size = maxDepth
    TheHdw.Digital.HRAM.CaptureType = captFail
    TheHdw.Digital.HRAM.SetTrigger trigFail, False, 0

    rtnPatternNames = TheExec.DataManager.Raw.GetPatternsInSet(Shift_Pat.value, rtnPatternCount)
    DecomPatName = Split(rtnPatternNames(0), ":")

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered 'SEC DRAM
    

    'RSCR site loop
    For Each Site In TheExec.sites
        If MbistBlockName.Exists(MBIST_BLOCK) Then
            Dim blockNum As Long
            Dim StrServer As String
            Dim StrDesc As String
            Dim RSCR_Flow_Idx As Long
            blockNum = MbistBlockName.item(MBIST_BLOCK)
            RSCR_Flow_Idx = TheExec.sites.item(Site).SiteVariableValue("RSCR_Flow_Idx")
            If Descri <> "" Then StrDesc = Descri Else StrDesc = MbistBlock(blockNum).strDecsName(RSCR_Flow_Idx)
            StrServer = MbistBlock(blockNum).strServerName(RSCR_Flow_Idx)
        End If
        
        TheHdw.Patterns(Shift_Pat).Load
        TheHdw.Patterns(Shift_Pat).start vbNullString
        TheHdw.Digital.Patgen.HaltWait
        patPassed = TheHdw.Digital.Patgen.PatternBurstPassed(Site)     'T-Col TTR approve by Si -- 230413

        numcap(Site) = TheHdw.Digital.HRAM.CapturedCycles

        If numcap(site) = 0 Then
            NewFmtAry(site + LineOffset) = "RSCR3," & RSCRHead & "," & site & "," & StrServer & "," & StrDesc & "," & DecomPatName(1) & ",NA,"
            If EmptyAry = True Then
                ReDim OldFmtAry(0)
                EmptyAry = False
                OldFmtAry(NumCapSum) = "Site " & Site & "," & "Fail cycle at: NA" & ", Prime"
                NumCapSum = NumCapSum + 1
            Else
                ReDim Preserve OldFmtAry(UBound(OldFmtAry) + 1)
                OldFmtAry(NumCapSum) = "Site " & Site & "," & "Fail cycle at: NA" & ", Prime"
                NumCapSum = NumCapSum + 1
            End If
        Else
            If EmptyAry = True Then
                ReDim OldFmtAry(numcap(Site))
                EmptyAry = False
            Else
                ReDim Preserve OldFmtAry(UBound(OldFmtAry) + numcap(Site))
            End If
            
            NewFmtAry(site + LineOffset) = "RSCR3," & RSCRHead & "," & site & "," & StrServer & "," & StrDesc & "," & DecomPatName(1) & ","
            
            RVal(Site) = TheHdw.Digital.HRAM.PatGenInfo(numcap(Site) - 1, pgCycle)
            mem_location = "none"
            
            For i = 0 To numcap(Site) - 1
''                PinData = TheHdw.Digital.Pins("JTAG_TDO").HRAM.PinData(0, 1, numcap(Site))

                '//MEMORY_CL52 cycle 2421 to cycle 3140
                '//MEMORY_CL51 cycle 3141 to cycle 7620
                '//MEMORY_CL27 cycle 8901 to cycle 9299
                '//MEMORY_CL26 cycle 10001 to cycle 10399
                '//MEMORY_CL17 cycle 11741 to cycle 12139
                '//MEMORY_CL16 cycle 12841 to cycle 13239
                'Mbist_repair_cycle
                Mbist_repair_cycle = TheHdw.Digital.HRAM.PatGenInfo(i, pgCycle)
                'Mbist_repair_cycle = Mbist_repair_cycle + 1    'start from cycle 1

                mem_location = "None"
                
                If Len(NewFmtAry(Site + LineOffset)) < 250 Then
                    NewFmtAry(Site + LineOffset) = NewFmtAry(Site + LineOffset) & Mbist_repair_cycle & ","
                Else
                    LineOffset = LineOffset + 1
                    'ReDim Preserve NewFmtAry((TheExec.sites.Existing.Value) + LineOffset)
                    ReDim Preserve NewFmtAry((TheExec.sites.Existing.Count) - 1 + LineOffset)
                    NewFmtAry(site + LineOffset) = "RSCR3," & RSCRHead & "," & site & "," & StrServer & "," & StrDesc & "," & DecomPatName(1) & "," & Mbist_repair_cycle & ","
                End If

                If MbistBlockName.Exists(MBIST_BLOCK) Then
                    M_MbistNum = MbistBlockName.item(MBIST_BLOCK)
                    M_maxrow = MbistBlock(M_MbistNum).MaxRow - 2

                    For k = 0 To M_maxrow
                        If Mbist_repair_cycle = MbistBlock(M_MbistNum).MemArray(k) Then mem_location = MbistBlock(M_MbistNum).MemStrArray(k)
                    Next k
                Else
                    mem_location = "Block-non-define"
                End If
                
                OldFmtAry(NumCapSum) = "Site " & Site & "," & "Fail cycle at: " & Mbist_repair_cycle & ",Mem:" & mem_location
                NumCapSum = NumCapSum + 1
            Next i
        End If

        '///print pattern result 170626
        TestNumber = TheExec.sites.item(Site).TestNumber

        If patPassed Then
            Call TheExec.Datalog.WriteFunctionalResult(Site, TestNumber, logTestPass)
            TheExec.sites.item(site).testResult = sitePass                                   'Fix Flag setting issue by Carter 220905
        Else
            Call TheExec.Datalog.WriteFunctionalResult(Site, TestNumber, logTestFail)
            TheExec.sites.item(site).testResult = siteFail                                   'Fix Flag setting issue by Carter 220905
        End If
        '///print pattern result 170626
    Next Site

    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
    
    For i = 0 To UBound(NewFmtAry)
        If NewFmtAry(i) <> "" Then
            NewFmtAry(i) = left(NewFmtAry(i), Len(NewFmtAry(i)) - 1)
            TheExec.Datalog.WriteComment NewFmtAry(i)
        End If
    Next i
    '===================================================================================
    If TheExec.Flow.EnableWord("RSCR_MP") <> True Then

        TheExec.Datalog.WriteComment "Mbist repair information shift start"
        
        For i = 0 To UBound(OldFmtAry)
            TheExec.Datalog.WriteComment OldFmtAry(i)
        Next i
        
        TheExec.Datalog.WriteComment "Mbist repair information shift end"
    End If
    '===================================================================================
    DebugPrintFunc Shift_Pat.value
    Exit Function

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Mbist", "Mbist_RSCR") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function TurnOnEfusePwrPins_Mbist(FusePower As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    'Escalate VDD18_EFUSE0 and VDD18_EFUSE1 according to Fiji
    'test plan (slower than 1.8v/30us)
    
    DCVS_PowerOn_I_Meter FusePower, 1.8, 0.2, 0.001, 0.002, 10, 0.018   'use 18 ms to power up

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Mbist", "TurnOnEfusePwrPins_Mbist") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function TurnOffEfusePwrPins_Mbist(FusePower As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    'Decline VDD18_EFUSE0 and VDD18_EFUSE1 according to Fiji
    'test plan (slower than 1.8v/30us)
    
    Dim CurrentVoltage As Double
    
    CurrentVoltage = TheHdw.DCVS.Pins(FusePower).Voltage.Main.value
    DCVS_PowerOff_I_Meter FusePower, CurrentVoltage, 0.2, 0.001, 0.002, 10, 0.018   'use 18 ms to power down


Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Mbist", "TurnOffEfusePwrPins_Mbist") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
Public Function MbistRetentionLevelWait_and_lowDown_power(mS_Time As Double, Pwr_pins As PinList, Low_Vol As Double)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    'TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered 'SEC DRAM
    Dim original_vol As Double

    original_vol = TheHdw.DCVS.Pins(Pwr_pins).Voltage.Main.value
    TheHdw.DCVS.Pins(Pwr_pins).Voltage.Main.value = Low_Vol

    TheExec.Datalog.WriteComment "*************************************************"
    TheExec.Datalog.WriteComment "*print: Lower pin: " & Pwr_pins & " to " & Low_Vol & " *"
    TheExec.Datalog.WriteComment "*************************************************"

    DebugPrintFunc ""
    TheHdw.Wait mS_Time * 0.001
    
    
    TheHdw.DCVS.Pins(Pwr_pins).Voltage.Main.value = original_vol

    TheExec.Datalog.WriteComment "*************************************************"
    TheExec.Datalog.WriteComment "*print: MbistRetention wait " & mS_Time & " ms*"
    TheExec.Datalog.WriteComment "*************************************************"


Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Mbist", "MbistRetentionLevelWait_and_lowDown_power") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function MbistRetentionLevelWait(mS_Time As Double, RampStep As Double, Optional RampWaitTime As Double = 0, Optional WaitTimeOnly As Boolean = False, _
                                        Optional RET_TTR As Boolean = False)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    
    ''SWLINZA20171103, for ramp up/down retention voltage

    Dim Retention_Pins_Ary() As String
    Dim Retention_Pins_count As Long
    Dim RampDown_Time As Double: RampDown_Time = RampWaitTime 'RampDown_Time = 0
    Dim RampDown_Step As Double
    Dim Voltage_from_HW As String
    Dim i, j As Integer
    
    Dim Retention_Ramp_Seq As String
    Dim Retention_RampUp_Seq As String
    Dim Ret_SRAMVol_Ary() As Double
    Dim Ret_LogicVol_Ary() As Double
    Dim Ret_SRAMPins_Ary() As String
    Dim Ret_LogicPins_Ary() As String
    
    If RET_TTR = True Then
        If WaitTimeOnly = False Then
            TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
        End If
        '----- Retention Wait time 100 ms ------
        Call MbistRetentionWait(mS_Time)
    End If

    If WaitTimeOnly = False Then
        If RampStep = 0 Then
            RampDown_Step = 20 ' default RampDown_Step = 20
        Else
            RampDown_Step = RampStep
        End If
        
        MbistERT_OriginVol.RemoveAll
        MbistERT_TargetVol.RemoveAll
        MbistERT_DropVol.RemoveAll
        MbistERT_DropVol_PerSite.RemoveAll
        MbistERT_GroupCurrentVol.RemoveAll
        MbistERT_GroupPreviousVol.RemoveAll
        MbistERT_GroupPinNmame.RemoveAll
        
        'Add Start, Leon Li, 20210114, default assign value from DC spec
        Dim lclStr_DcCategory As String
        Dim lclStr_DcSelector As String
        Dim lclStr_AcCategory As String
        Dim lclStr_AcSelector As String
        Dim lclStr_TimeSetSheet As String
        Dim lclStr_EdgeSetSheet As String
        Dim lclStr_LevelsSheet As String
        Dim lclStr_Overlay As String
        
        Dim Retention_SRAMVoltage As String:: Retention_SRAMVoltage = vbNullString
        Dim Retention_SRAMPins As New PinList:: Set Retention_SRAMPins = Nothing
        Dim Retention_LogicVoltage As String:: Retention_LogicVoltage = vbNullString
        Dim Retention_LogicPins As New PinList:: Set Retention_LogicPins = Nothing
'        Dim lclPinListData_PreviousCorePower As New PinListData:: Set lclPinListData_PreviousCorePower = Nothing
'        Dim lclPinListData_CurrentCorePower As New PinListData:: Set lclPinListData_CurrentCorePower = Nothing
        Set lclPinListData_PreviousCorePower = Nothing
        Set lclPinListData_CurrentCorePower = Nothing

        Dim lclStrArr_Pins() As String, lclLng_PinCnt As Long
        Dim lclStr_Prefix As String:: lclStr_Prefix = vbNullString:: lclStr_Prefix = TheExec.DataManager.instanceName

        TheExec.DataManager.DecomposePinList "CorePower", lclStrArr_Pins, lclLng_PinCnt
        TheExec.DataManager.GetInstanceContext lclStr_DcCategory, lclStr_DcSelector, lclStr_AcCategory, lclStr_AcSelector, lclStr_TimeSetSheet, lclStr_EdgeSetSheet, lclStr_LevelsSheet, lclStr_Overlay

        For i = 0 To UBound(lclStrArr_Pins)
            If TheExec.DataManager.ChannelType(lclStrArr_Pins(i)) <> "N/C" Then
                lclPinListData_PreviousCorePower.AddPin lclStrArr_Pins(i)
                lclPinListData_PreviousCorePower.Pins(lclStrArr_Pins(i)).value = Format(TheHdw.DCVS.Pins(lclStrArr_Pins(i)).Voltage.value, "0.000")
                lclPinListData_CurrentCorePower.AddPin lclStrArr_Pins(i)
                lclPinListData_CurrentCorePower.Pins(lclStrArr_Pins(i)).value = Format(TheExec.Specs.DC.item(lclStrArr_Pins(i) & "_VAR").Categories(lclStr_DcCategory).Selectors(lclStr_DcSelector).ContextValue, "0.000")
                                                                                                                                                                                                                                                  
                If lclPinListData_CurrentCorePower.Pins(lclStrArr_Pins(i)).value < lclPinListData_PreviousCorePower.Pins(lclStrArr_Pins(i)).value Then
                    If lclStrArr_Pins(i) Like "*SRAM*" Then
                        Retention_SRAMVoltage = IIf(Retention_SRAMVoltage = "", lclPinListData_CurrentCorePower.Pins(lclStrArr_Pins(i)).value, Retention_SRAMVoltage & "," & lclPinListData_CurrentCorePower.Pins(lclStrArr_Pins(i)).value)
                        Retention_SRAMPins = IIf(Retention_SRAMPins = "", lclStrArr_Pins(i), Retention_SRAMPins & "," & lclStrArr_Pins(i))
                    Else
                        Retention_LogicVoltage = IIf(Retention_LogicVoltage = "", lclPinListData_CurrentCorePower.Pins(lclStrArr_Pins(i)).value, Retention_LogicVoltage & "," & lclPinListData_CurrentCorePower.Pins(lclStrArr_Pins(i)).value)
                        Retention_LogicPins = IIf(Retention_LogicPins = "", lclStrArr_Pins(i), Retention_LogicPins & "," & lclStrArr_Pins(i))
                    End If
                End If
            End If
        Next
        
        '----------------------------------------------------------------------------------------'
        'Due to SELSRM, we need to avoid negative current/alarm happens during ramping
        'We have to consider about Different Retention Sequence for RampUp/Down
        '----------------------------------------------------------------------------------------'
        If Retention_LogicPins = "" And Retention_SRAMPins = "" Then
            TheExec.Datalog.WriteComment "ERROR: No Pin need to do retention, please check DC Spec or WaitTimeOnly Argu."
            TheExec.Datalog.WriteComment "ERROR: If the instance is internal retention, the WaitTimeOnly must fill True."
            Call MbistRetentionWait(mS_Time)
            Exit Function
        End If
        
        Retention_Ramp_Seq = Retention_LogicPins & IIf(Retention_SRAMPins = "" Or Retention_LogicPins = "", "", ",") & Retention_SRAMPins

        'Add Start, Leon Li, 20210114
        '----------------------------------------------------------------------------------------'
        '------ To check counts number between pins and vol-settings, Add Vol in Dictionary
        '------ Expand Vol to all pins, if only one voltage set
        '------ ERROR while count is not 1 but also matched
        '----------------------------------------------------------------------------------------'
        If Retention_SRAMVoltage <> "" And Retention_SRAMPins <> "" Then
            'Temp_Ary() = CDbl(Split(Retention_SRAMVoltage, ","))
            Call Sub_CStrToDblAry(Retention_SRAMVoltage, Ret_SRAMVol_Ary, ",")
            Ret_SRAMPins_Ary() = Split(Retention_SRAMPins, ",")
            Call Sub_SetVol_toAllPins(Ret_SRAMPins_Ary(), Ret_SRAMVol_Ary(), RampDown_Step)
        End If
        
        If Retention_LogicVoltage <> "" And Retention_LogicPins <> "" Then
            'Temp_Ary() = CDbl(Split(Retention_LogicVoltage, ","))
            Call Sub_CStrToDblAry(Retention_LogicVoltage, Ret_LogicVol_Ary, ",")
            Ret_LogicPins_Ary() = Split(Retention_LogicPins, ",")
            Call Sub_SetVol_toAllPins(Ret_LogicPins_Ary(), Ret_LogicVol_Ary(), RampDown_Step)
        End If
                
        '------------------------------------------------'
        '--------- Ramp down for retention voltage ------'
        '------------------------------------------------'
        TheExec.DataManager.DecomposePinList Retention_Ramp_Seq, Retention_Pins_Ary(), Retention_Pins_count
        Call Sub_VoltageRamping(RampDown_Step, RampDown_Time, RampDown)

'        TheExec.DataManager.DecomposePinList Retention_RampDown_Seq, Retention_Pins_Ary(), Retention_Pins_count
'        Call Sub_VoltageRamping(Retention_Pins_Ary(), RampDown_Step, RampDown_Time, RampDown)
        
       
        '--------------------------------------------------------------------'
        '----------- Wait time Procedure/ Chcking HW setting ----------------'
        '--------------------------------------------------------------------'
        Voltage_from_HW = vbNullString
        '--------- Read back retention voltage from HW ------'
        For j = 0 To Retention_Pins_count - 1
            Voltage_from_HW = CombineStringList(Voltage_from_HW, CStr(FormatNumber(TheHdw.DCVS.Pins(Retention_Pins_Ary(j)).Voltage.value, 3)) & " V")
        Next j
            
        '----- Retention Wait time 100 ms ------
        TheHdw.Wait mS_Time * 0.001
        TheExec.Flow.TestLimit mS_Time, PinName:="Wait_Time", unit:=unitCustom, customUnit:="mSec"
        TheExec.Datalog.WriteComment "*************************************************"
        TheExec.Datalog.WriteComment "*print: MbistRetention wait " & mS_Time & " ms*"
        TheExec.Datalog.WriteComment "*print: MbistRetention Pins " & Retention_Ramp_Seq
        TheExec.Datalog.WriteComment "*print: MbistRetention Volt " & Voltage_from_HW
        TheExec.Datalog.WriteComment "*************************************************"
        DebugPrintFunc ""
        
        '------------------------------------------------'
        '--------- Ramp Up for retention voltage --------'
        '------------------------------------------------'
        'TheExec.DataManager.DecomposePinList Retention_RampUp_Seq, Retention_Pins_Ary(), Retention_Pins_count
        'Call Sub_VoltageRamping(Retention_Pins_Ary(), RampDown_Step, RampDown_Time, RampUp)
        Call Sub_VoltageRamping(RampDown_Step, RampDown_Time, RampUp)


    Else
        '----- Retention Wait time 100 ms ------
        Call MbistRetentionWait(mS_Time)
    End If
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Mbist", "MbistRetentionLevelWait") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function Init_MBISTFailBlock() 'Multi Sheets
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    'TheHdw.StartStopwatch

    Dim MBISTFAILBLOCKSHEETS(0) As String
    
    MBISTFAILBLOCKSHEETS(0) = "MBISTFailBlock"
    
'    MBISTFAILBLOCKSHEETS(0) = "MBISTFailBlock_CPU"
'    MBISTFAILBLOCKSHEETS(1) = "MBISTFailBlock_GFX"
'    MBISTFAILBLOCKSHEETS(2) = "MBISTFailBlock_SOC"
    
    
    Dim k As Long
    Dim m As Long
    Dim n As Long
    Dim i As Long
        
    m = 0
    n = 0
    Dim sheetIdx As Integer
    sheetIdx = 0
    Const start_row = 2
    Const end_col = 5
    
    If Flag_MBISTFailBlock_INIT = False Then
        ReDim tpCycleBlockInfor(300)
        Application.ScreenUpdating = False
        For sheetIdx = 0 To UBound(MBISTFAILBLOCKSHEETS)
            
            Dim MaxRow As Long
            
            Dim arr1() As Variant
        
            MaxRow = Worksheets(MBISTFAILBLOCKSHEETS(sheetIdx)).UsedRange.Rows.Count
            
            Dim maxcolumn As Long
            maxcolumn = Worksheets(MBISTFAILBLOCKSHEETS(sheetIdx)).UsedRange.Columns.Count
            
            Worksheets(MBISTFAILBLOCKSHEETS(sheetIdx)).Activate
            
            arr1 = Worksheets(MBISTFAILBLOCKSHEETS(sheetIdx)).range(Cells(start_row, 1), Cells(MaxRow, maxcolumn)).value
            

            
            For k = 1 To MaxRow - 1
                
               ' If arr1(2, 1) = "" Then Exit For
                If (k = 1) And (sheetIdx = 0) Then
                    ReDim tpCycleBlockInfor(m).tpMbistCycleBlock(1800)
                    ReDim tpCycleBlockInfor(m).strDecsName(maxcolumn - 6)
                    ReDim tpCycleBlockInfor(m).strServerName(maxcolumn - 6)
                    tpCycleBlockInfor(m).strPattName = arr1(k, 1)
                Else
                    If tpCycleBlockInfor(m).strPattName = arr1(k, 1) Then
                        If n > 1800 Then
                            ReDim Preserve tpCycleBlockInfor(m).tpMbistCycleBlock(n + 200)
                        End If
                        
                    Else
                        ReDim Preserve tpCycleBlockInfor(m).tpMbistCycleBlock(n - 1)
                        
                        n = 0
                        m = m + 1
                        
                        If m > 300 Then
                            ReDim Preserve tpCycleBlockInfor(m + 20)
                        End If
                        
                        ReDim tpCycleBlockInfor(m).tpMbistCycleBlock(1800)
                        ReDim tpCycleBlockInfor(m).strDecsName(maxcolumn - 6)
                        ReDim tpCycleBlockInfor(m).strServerName(maxcolumn - 6)
                        
                        tpCycleBlockInfor(m).strPattName = arr1(k, 1)
                    End If
                End If
                
                tpCycleBlockInfor(m).tpMbistCycleBlock(n).strBlaclName = arr1(k, 2)
                tpCycleBlockInfor(m).tpMbistCycleBlock(n).lVector = Int(arr1(k, 3))
                tpCycleBlockInfor(m).tpMbistCycleBlock(n).lCycle = Int(arr1(k, 4))
                tpCycleBlockInfor(m).tpMbistCycleBlock(n).strCompare = arr1(k, 5)

                For i = 6 To maxcolumn
                    Dim TempAry() As String
                    If i = 6 Then
                        TempAry() = Split(",", ",")
                    ElseIf arr1(k, i) = Empty Then
                        TempAry() = Split(",", ",")
                    Else
                        TempAry() = Split(arr1(k, i), ",")
                        If UBound(TempAry()) = 0 Then
                            ReDim Preserve TempAry(1)
                            TempAry(1) = vbNullString
                        End If
                    End If
                    tpCycleBlockInfor(m).strServerName(i - 6) = TempAry(0)
                    tpCycleBlockInfor(m).strDecsName(i - 6) = TempAry(1)
                Next i
                    
                n = n + 1
            Next k
    
        Next sheetIdx
        
        ReDim Preserve tpCycleBlockInfor(m).tpMbistCycleBlock(n - 1)
        ReDim Preserve tpCycleBlockInfor(m)
        'TheExec.Datalog.WriteComment RepeatChr("*", 120)
        TheExec.Datalog.WriteComment "print: MBISTFailBlock table initialized complete"
    End If
    Flag_MBISTFailBlock_INIT = True
    
    Application.ScreenUpdating = True
    'Debug.Print " Init_MBISTFailBlock new : " & TheHdw.ReadStopwatch
        
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Mbist", "Init_MBISTFailBlock") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


''' add 20160629  webster
Public Function GetFlagInfoArrIndex(FlagName As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim funcName As String: funcName = "MbistFP_FlagPrintTest"
    Dim lIdxTemp As Long
    
    GetFlagInfoArrIndex = -1
    
    For lIdxTemp = 0 To UBound(tyFlagInfoArr)
        If tyFlagInfoArr(lIdxTemp).FlagName = FlagName Then
            GetFlagInfoArrIndex = lIdxTemp
            Exit For
        End If
    Next lIdxTemp
    
    If GetFlagInfoArrIndex = -1 Then
        TheExec.Datalog.WriteComment "<Warnning> the flag(" & FlagName & ") can not be found in MBISTFailBlock excel sheet"
    End If

    Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Mbist", "GetFlagInfoArrIndex") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function MbistRetentionWait(mS_Time As Double)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    'TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered 'SEC DRAM

    TheHdw.Wait mS_Time * 0.001

    TheExec.Flow.TestLimit mS_Time, PinName:="Wait_Time", unit:=unitCustom, customUnit:="mSec"
    TheExec.Datalog.WriteComment "*************************************************"
    TheExec.Datalog.WriteComment "*print: MbistRetention wait " & mS_Time & " ms*"
    TheExec.Datalog.WriteComment "*************************************************"
    DebugPrintFunc ""

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Digital_Mbist", "MbistRetentionWait") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
