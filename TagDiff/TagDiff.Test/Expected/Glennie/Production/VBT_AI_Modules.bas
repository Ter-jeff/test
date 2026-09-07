Attribute VB_Name = "VBT_AI_Modules"
Option Explicit
Public Enum pinTypes
    None
    VolType
    frqType
End Enum

Public Enum volTypes
    None
    mainType
    altType
End Enum

Public Enum loopTypes
    None
    lFixType
    lVolType
End Enum

Public Enum axes
    aXaxis
    aYaxis
End Enum

Public Enum patTypes
    pInit
    pWrite
    pRead
End Enum

Public Enum pfAxes
    pXaxis
    pYaxis
    pILoops
    pOLoops
End Enum

Public Enum Infos
    iStart
    iEnd
    iBlank
    iLoop
    iXstart
    iYstart
    iCalVol
    iPinVol
    iPinFrq
    iSrcStr
End Enum

Public Type SelSrmPins
    logicPin As String
    SramPin As String
End Type

Public Type SelSrmCodes
    SelSrm0 As String
    SelSrm1 As String
End Type

Public Type SelSrmInfos
    SelSrmPin As SelSrmPins
    SelSrmCode As SelSrmCodes
End Type

Public autoShmooUtility As New autoShmooUtility
Public autoShmoo As AutoShmooMain

Public Function ShmooDIfferentPatterns(patterns As Pattern, SweepPins As String, GlobalForWrite As String, GlobalForRead As String, _
                                        VolForInit As String, VolForWrite As String, VolForRead As String, digSrc As String, pmodeVol As String, blockName As String, Validating_ As Boolean)
    
    ''' preload pattern
    If Validating_ Then
        If patterns.value <> "" Then PrLoadPattern patterns.value
        Exit Function
    End If
    
    ''' exit function for last point while shmoo
    If ShmooEndFunction Then Exit Function
    
    Dim VbumpInfo As VbumpInfos
    Dim preType As patTypes
    Dim pat As Variant
    
    ''' initial
    Set VbumpInfo = New VbumpInfos
    VbumpInfo.Initialize patterns, SweepPins, GlobalForWrite, GlobalForRead, VolForInit, VolForWrite, VolForRead, digSrc, pmodeVol, blockName
    If Not VbumpInfo.PreProcess Then Exit Function
    
    ''' applyTimingLevel
    VbumpInfo.ApplyTimingLevel
    ''' prevent issue
    VbumpInfo.switchToVmainVol
    ''' store voltage for init/auto switch
    If Not VbumpInfo.setVrsPinsVoltage Then Exit Function
    ''' autoswitch case
    If VbumpInfo.IsAutoSwitchCase Then
        Dim autoSwitchInfo As AutoSwitchInfos
        Set autoSwitchInfo = New AutoSwitchInfos
        ''' initial auto switch table
        If Not autoSwitchInfo.Initialize Then Set VbumpInfo = Nothing: Set autoSwitchInfo = Nothing: Exit Function
    Else
        ''' setup digital source
        If VbumpInfo.IsExistSourceCode Then If Not VbumpInfo.SetupDigSrc Then Set VbumpInfo = Nothing: Exit Function
    End If
    
    ''' test pattern
    preType = pWrite ''' it is used to judge voltage for autoswitch
    For Each pat In VbumpInfo.patterns
        If VbumpInfo.IsInitPattern(CStr(pat)) Then VbumpInfo.setVoltageToAxis patTypes.pInit, CStr(pat), preType ''' set voltage for init
        If VbumpInfo.IsWritePattern(CStr(pat)) Then VbumpInfo.setVoltageToAxis patTypes.pWrite ''' set voltage for write
        If VbumpInfo.IsReadPattern(CStr(pat)) Then VbumpInfo.setVoltageToAxis patTypes.pRead ''' set voltage for read
        ''' setup select sram code for auto switch
        If VbumpInfo.IsAutoSwitchCase And VbumpInfo.IsInitPattern(CStr(pat)) Then If Not VbumpInfo.SetupDigSrc(autoSwitchInfo) Then Set VbumpInfo = Nothing: Set autoSwitchInfo = Nothing: Exit Function
        If VbumpInfo.IsSourcePattern(CStr(pat)) Then VbumpInfo.PrintDigSrcInfo CStr(pat) ''' print digital source code
        VbumpInfo.testPattern CStr(pat)
    Next pat
End Function

Public Function RepeatPatternToGetFailLog(patterns As Pattern, _
                                            xAxis As String, xBegs As String, xEnds As String, xStep As Double, _
                                            yAxis As String, yBegs As String, yEnds As String, yStep As Double, _
                                            capXfailcount As Long, capYfailcount As Long, LoopTimes As Long, capLoopsCount As Long, _
                                            showFrqPins As PinList, showVolPins As String, disComparePins As String, _
                                            pinType As pinTypes, VolType As volTypes, loopType As loopTypes, pmodeVol As String, digSrc As String, blockName As String, _
                                            Validating_ As Boolean)
    
    
    Dim FaillogInfo As FaillogInfos
    Dim yCount As Integer
    Dim xCount As Integer
    Dim patcnt As Integer
    Dim iLoopCnt As Long
    Dim oLoopCnt As Long
    Dim yValue As Double
    Dim xValue As Double
    Dim pat As Variant
    
    ''' pre-load pattern
    If Validating_ Then
        If patterns.value <> "" Then Call PrLoadPattern(patterns.value)
        Exit Function
    End If
    
    ''' initial
    Set FaillogInfo = New FaillogInfos
    FaillogInfo.initialFaillog patterns, xAxis, xBegs, xEnds, xStep, yAxis, yBegs, yEnds, yStep, capXfailcount, capYfailcount, LoopTimes, capLoopsCount, _
                                        showFrqPins, showVolPins, disComparePins, pinType, VolType, loopType, pmodeVol, digSrc, blockName
    
    ''' apply timing level
    FaillogInfo.ApplyTimingLevel
    ''' prevent issue
    FaillogInfo.switchToVmainVol
    ''' apply pmode voltage to valt if select sram case
    If Not FaillogInfo.setVrsPinsVoltage Then Exit Function
    ''' set testnumber for increment
    FaillogInfo.setTestNumbers
    ''' initial auto switch table
    If FaillogInfo.IsAutoSwitchCase Then
        Dim autoSwitchInfo As AutoSwitchInfos
        Set autoSwitchInfo = New AutoSwitchInfos
        If Not autoSwitchInfo.Initialize Then Set FaillogInfo = Nothing: Set autoSwitchInfo = Nothing: Exit Function
    Else
        ''' setup digital source
        If FaillogInfo.IsExistSourceCode Then If Not FaillogInfo.SetupDigSrc Then Set FaillogInfo = Nothing: Exit Function
    End If
    ''' print start
    FaillogInfo.PrintInfos Infos.iStart
    
    ''' start y axis
    For yCount = 0 To FaillogInfo.yAxisStepCount
        If FaillogInfo.IsExistAxis(axes.aYaxis) Then FaillogInfo.setValueToAxis yCount ''' apply yaxis value to each pin

        ''' outer loop count start
        For oLoopCnt = FaillogInfo.LoopCountNumber To FaillogInfo.LoopTimes ''' vol point looping
            If ((oLoopCnt < iLoopCnt Or (oLoopCnt = 1 And iLoopCnt = 0)) And FaillogInfo.LoopTimes <> 1) Then FaillogInfo.PrintInfos iLoop, , , oLoopCnt ''' print loop count
            
            ''' start x axis
            For xCount = 0 To FaillogInfo.xAxisStepCount
                If FaillogInfo.IsExistAxis(axes.aXaxis) Then FaillogInfo.setValueToAxis xCount ''' apply xaxis value to each pin

                ''' inner loop count start
                For iLoopCnt = FaillogInfo.LoopCountNumber To FaillogInfo.LoopTimes ''' fix point looping
                    If FaillogInfo.LoopCountNumber <> FaillogInfo.LoopTimes Then FaillogInfo.PrintInfos iLoop, , , iLoopCnt ''' print loop count
                    FaillogInfo.switchToVmainVol ''' switch to safe voltage
                    
                    If FaillogInfo.IsAutoSwitchCase Then If Not FaillogInfo.SetupDigSrc(autoSwitchInfo) Then Set autoSwitchInfo = Nothing: Set FaillogInfo = Nothing: Exit Function ''' re-judge for auto switch
                    'If FaillogInfo.IsExistShowFrqPins Then FaillogInfo.ShowFrequencyOnPins ''' show frequence on specific pins
                    If FaillogInfo.IsExistShowVolPins Then FaillogInfo.ShowVoltageOnPins ''' show voltage on sepecific pins
                    
                    If FaillogInfo.IsExistDisComparePins Then FaillogInfo.DisableComparePins True ''' disable compare pins
                    ''' test pattern
                    If TheExec.enableWord("PrintRSVoltage") Then PrintRSVoltage False, FaillogInfo.SweepPinsString
                    For Each pat In FaillogInfo.patterns
                        If FaillogInfo.IsSelectSramCase And Not FaillogInfo.IsInitPattern(CStr(pat)) Then FaillogInfo.switchToValtVol ''' switch to pmode voltage
                        If FaillogInfo.IsSourcePattern(CStr(pat)) Then FaillogInfo.PrintInfos iSrcStr, , , , CStr(pat) ''' print source code info
                        FaillogInfo.testPattern CStr(pat) ''' run pattern
                    Next pat
                    If FaillogInfo.IsExistDisComparePins Then FaillogInfo.DisableComparePins False ''' enable compare pins
                    If FaillogInfo.IsExistShowFrqPins Then FaillogInfo.ShowFrequencyOnPins ''' show frequence on specific pins
                    
                    FaillogInfo.judgePFResult pILoops ''' judge inner loop pass/fail result
                    If Not FaillogInfo.AnyEnableSites(pILoops, iLoopCnt) Then FaillogInfo.resetToXactiveSites: Exit For
                Next iLoopCnt
                FaillogInfo.judgePFResult pXaxis ''' judge x axis pass/fail result
                ''' if there is not site enable then reset site, else if last loop count then reset site
                If Not FaillogInfo.AnyEnableSites(pXaxis) Then FaillogInfo.resetToYactiveSites: Exit For
                If xCount = FaillogInfo.xAxisStepCount Then FaillogInfo.resetToYactiveSites: FaillogInfo.resetXaxisFailCount
            Next xCount
            FaillogInfo.judgePFResult pOLoops ''' judge outer pass/fail result
            If Not FaillogInfo.AnyEnableSites(pOLoops, oLoopCnt) Then FaillogInfo.resetToYactiveSites: Exit For
        Next oLoopCnt
        FaillogInfo.judgePFResult pYaxis ''' judge y axis pass/fail result
        If Not FaillogInfo.AnyEnableSites(pYaxis) Or yCount = FaillogInfo.yAxisStepCount Then FaillogInfo.resetToActiveSites: Exit For
    Next yCount
    
    ''' faillog end
    FaillogInfo.PrintInfos iEnd
    ''' release
    Call VaryFreq("XI0_Diff_Port", 24000000, "XI0_Diff_Freq_VAR")
    Set FaillogInfo = Nothing
End Function

Public Function MeasFrequencyOnPins(MeasPins As PinList)
    Dim MeasPinsData As New PinListData

    TheHdw.Digital.pins(MeasPins).Levels.DriverMode = tlDriverModeVt
    Call Freq_MeasFreqSetup(MeasPins, 0.01, BOTH)      '' 20150621 - default d_FreqMeasInterval = 0.01
    '' 20150623 - Add Customize Wait Time
    Call HardIP_Freq_MeasFreqStart(MeasPins, 0.01, MeasPinsData, "")          '' 20150621 - default d_FreqMeasInterval = 0.01
'    If FreqLowLimit_ <> 0 Or FreqHighLimit_ <> 0 Then
'        TheExec.Flow.TestLimit MeasPinsData, FreqLowLimit_, FreqHighLimit_, , , , unitHz, , ForceResults:=tlForceNone
'    Else
    TheExec.flow.TestLimit MeasPinsData, , , , , , unitHz, , ForceResults:=tlForceNone
    'glb_TestInstance = theexec.DataManager.instancename
    'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
'    End If
    TheHdw.Digital.pins(MeasPins).Levels.DriverMode = tlDriverModeHiZ
End Function

Public Function charSetup()
    TheExec.DevChar.Configuration.Features(tlDevCharFeature_StoreResultsUntilNextRun) = True
End Function

Public Function CovCond(val As String) As String
    Dim out As String
    out = Replace(Replace(Trim(val), "=", ":V:"), ",", ";")
    CovCond = out
    Debug.Print out
End Function

Sub AddTestNumber()
    buildTestNumber "Flow_Vddbinning", 50000000
    Debug.Print "Done"
End Sub

Public Function buildTestNumber(name As String, ByRef num As Long)
    Dim sht As Worksheet
    Dim Row As Integer
    Dim opcd As String
    Dim Para As String
    
    Set sht = Sheets(name)
    sht.Activate
    For Row = 5 To sht.UsedRange.Rows.Count
        opcd = LCase(sht.Cells(Row, 7).value)
        Para = sht.Cells(Row, 8).value
        If opcd = "call" Then
            If LCase(Para) = LCase("Flow_nWire_Default") Then
                sht.Cells(Row, 7).value = "nop"
            ElseIf LCase(Para) Like LCase("Flow_TMPS*") Then
            ElseIf LCase(Para) Like LCase("*decision*block*") Then
            Else
                buildTestNumber Para, num
                sht.Activate
            End If
        ElseIf LCase(Para) = LCase("Relay_ON_Default") Then
            sht.Cells(Row, 7).value = "nop"
        ElseIf opcd = "test" Then
            sht.Cells(Row, 10).value = num
            num = num + 1000
        End If
    Next Row
End Function

Public Function CorrectBincut(Optional HARDBIN As Integer = 15, Optional COPY As Boolean = False)

    Dim sht As Worksheet
    Dim shts() As String
    Dim name As Variant
    Dim val As String
    Dim pivot As Integer
    Dim Base As Integer: Base = 1
    Dim Row As Integer
    Dim Col As Integer
    Dim left As Integer
    Dim right As Integer
    Dim index As Integer
    Dim high As Integer
    Dim tCopy As Boolean
    Const NON_BINNING_RAIL = "NON_BINNING_RAIL*"
    Const VDD_BINNING_DEF = "VDD_BINNING_DEF*"
    Const COL_SOFT_BIN = "COL_SOFT_BIN*"
    Const BINNING_DOMAIN = "BINNING*DOMAIN*"
    Const OUTSIDE = "*OUTSIDE*"
    Const COMMENT = "COMMENT*"
    Const OTHERS = "ALL OTHERS"
    Const HBV = "HBV*"
    Const QA25 = "*QA*25*"
    Const QA85 = "*QA*85*"
    Const T0TX25 = "*T0TX_ROOM*"
    
    shts = TheExec.job.GetSheetNamesOfType(DMGR_SHEET_TYPE_USER)
    
    For Each name In shts
        If UCase(name) Like NON_BINNING_RAIL Then
            Set sht = Sheets(name)
            sht.Activate
            pivot = 1
            index = -1
            tCopy = COPY

            While Not UCase(Trim(sht.Cells(pivot, 1).value)) Like BINNING_DOMAIN
                pivot = sht.Cells(pivot, 1).End(xlDown).Row
                If pivot = sht.Rows.Count Then Stop
            Wend
            
            ''' base is coming from "non_binning_rail"
            While Base = 1 Or Not UCase(Trim(sht.Cells(pivot + 1, Base).value)) Like COMMENT
                Base = sht.Cells(pivot + 1, Base).End(xlToRight).Column
                If Base = sht.Columns.Count Then Stop
            Wend
            
            Do
                Col = Base
                While Col <= sht.UsedRange.Columns.Count + 1 ''' since last row may not contain comment column, so we need to plus 1
'                    If UCase(Trim(sht.Cells(pivot + 1, col).value)) Like COMMENT Then
'                        Debug.Print "Change name to Comment"
'                    Else
'                        Debug.Print "Add Comment into cell"
'                    End If
                    ''' TODO add/update/remove, we should check the "All Others" and "Binning Domain" location
                    ''' 1.there is comment column
                    ''' 2. there is no comment column
                    ''' 3. there is TD/FUNC column
                    left = Col: right = Col
                    While left >= 0 And Not UCase(sht.Cells(pivot + 1, left).value) = OTHERS
                        left = left - 1
                    Wend
                    
                    While right < sht.UsedRange.Columns.Count And Not UCase(sht.Cells(pivot, right).value) Like BINNING_DOMAIN And Not UCase(Trim(sht.Cells(pivot, right).value)) Like HBV
                        right = right + 1
                    Wend
                    
                    ''' try to copy T0TX for post-bincut sheet
                    If tCopy Then
                        If UCase(Trim(sht.Cells(pivot, right + 2).value)) Like QA25 Then
                            index = left + 2
                            high = pivot
                        ElseIf UCase(Trim(sht.Cells(pivot, right + 2).value)) Like T0TX25 Then
                            tCopy = False
                        End If
                    End If
                    
                    If left = -1 Then
                        ''' can not find "All Others" in header
                        Stop
                    ElseIf right >= sht.UsedRange.Columns.Count Then
                        ''' last column, check if there is comment there
                        If Not UCase(sht.Cells(pivot + 1, right).value) Like COMMENT Then
                            sht.Cells(pivot + 1, left + 1).value = "Comment"
                        End If
                    ElseIf right - left = 1 Then
                        ''' missing comment column
                        sht.Cells(1, right).EntireColumn.Insert shift:=xlToRight
                        sht.Cells(pivot + 1, right).value = "Comment"
                    ElseIf right - left = 2 Then
                        ''' comment column is exist
                        sht.Cells(1, left + 1).Resize(2, 1).ClearContents
                        sht.Cells(pivot + 1, left + 1).value = "Comment"
                    ElseIf right - left > 2 And right - left < 6 Then
                        While right - left <> 1
                            sht.Cells(1, right - 1).EntireColumn.delete shift:=xlToLeft
                            right = right - 1
                        Wend
                        sht.Cells(1, right).EntireColumn.Insert shift:=xlToRight
                        sht.Cells(pivot + 1, right - 1).value = "Comment"
                    Else
                        ''' we should only have maximum column with TD/MBIST/FUNC/COMMENT
                        ''' more than this column number is inconsistent
                        Stop
                    End If
                    
'                    sht.Cells(pivot + 1, Col).value = "Comment"
                    Col = Col + Base
                Wend
                
                Do
                    pivot = sht.Cells(pivot, 1).End(xlDown).Row
                    If pivot = sht.Rows.Count Then Stop
                Loop While Not UCase(Trim(sht.Cells(pivot, 1).value)) Like HBV And Not UCase(Trim(sht.Cells(pivot, 1).value)) = "END"
            Loop While Not pivot = sht.Rows.Count And Not UCase(Trim(sht.Cells(pivot, 1).value)) = "END"
            
            If tCopy And index <> -1 And UCase(sht.name) Like OUTSIDE Then
                sht.Select
                sht.range(sht.Cells(1, index), sht.Cells(1, index + Base - 1)).EntireColumn.Select
                Selection.COPY
                sht.Cells(1, index).EntireColumn.Insert shift:=xlToRight
                Application.CutCopyMode = False
                index = index + Base
                While index <= sht.UsedRange.Columns.Count
                    If UCase(Trim(sht.Cells(high, index + 2).value)) Like QA25 Then
                        sht.Cells(high, index + 2).value = "T0TX_ROOM"
                    ElseIf UCase(Trim(sht.Cells(high, index + 2).value)) Like QA85 Then
                        sht.Cells(high, index + 2).value = "T0TX_HOT"
                    End If
                    index = index + Base
                Wend
            End If
            
        ElseIf UCase(name) Like VDD_BINNING_DEF Then
            Set sht = Sheets(name)
            sht.Activate
            pivot = 1
            
            While Not UCase(Trim(sht.Cells(1, pivot).value)) Like COL_SOFT_BIN
                pivot = sht.Cells(1, pivot).End(xlToRight).Column
                If pivot = sht.Columns.Count Then Stop
            Wend
            pivot = Int(Trim(sht.Cells(1, pivot + 1).value))
            
            For Col = pivot To sht.UsedRange.Columns.Count
                val = Trim(sht.Cells(2, Col).value)
                If UCase(val) = "HARDBIN" Then
                    sht.range(sht.Cells(4, Col), sht.Cells(sht.Cells(4, Col).End(xlDown).Row, Col)).value = HARDBIN
                ElseIf UCase(val) = "SOFTBIN" Then
                    For Row = 4 To sht.Cells(1, Col).End(xlDown).Row
                        If sht.Cells(Row, Col).value >= 8000 Then
                            sht.Cells(Row, Col).value = sht.Cells(Row, Col).value - 1000
                        End If
                    Next Row
                End If
            Next Col
        End If
    Next name
    Debug.Print "Complete the modification"
End Function

Public Function CorrectBinout(Optional job As String = "")
    Dim sht As Worksheet
    Dim shts() As String
    Dim name As Variant
    Dim ref As String
    Dim bin As String
    Dim val As Integer
    Dim val1 As Integer
    Dim mode As String
    Dim fail As String
    Dim Block As String
    Dim header As String
    Dim pivot As Integer
    Dim Base As Integer: Base = 1
    Dim Row As Integer
    Dim Col As Integer
    Dim soft As Dictionary
    Dim soft_rev As Dictionary
    Dim checked As Dictionary
    Dim hard As Dictionary

    Const BINTABLE = "BIN*TABLE*BINCUT"
    Const VDD_BINNING_DEF = "VDD_BINNING_DEF*_#"
    Const COL_SOFT_BIN = "COL_SOFT_BIN*"
    Const BINNINGFAIL = "BINNING*FAIL*"
    Const LVCCFAIL = "LVCC*FAIL"
    Const DEBUG_ = False
    
    shts = TheExec.job.GetSheetNamesOfType(DMGR_SHEET_TYPE_USER)
    For Each name In shts
        ref = VDD_BINNING_DEF
        If job <> "" Then ref = VDD_BINNING_DEF + "_" + UCase(job)
        If UCase(name) Like ref Then
            Set sht = Sheets(name)
            Set soft = Nothing: Set soft = New Dictionary: Set hard = Nothing: Set hard = New Dictionary: Set soft_rev = Nothing: Set soft_rev = New Dictionary: Set checked = New Dictionary
            sht.Activate
            pivot = 1
            
            While Not UCase(Trim(sht.Cells(1, pivot).value)) Like COL_SOFT_BIN
                pivot = sht.Cells(1, pivot).End(xlToRight).Column
                If pivot = sht.Columns.Count Then Stop
            Wend
            pivot = Int(Trim(sht.Cells(1, pivot + 1).value))
            
            For Col = pivot To sht.UsedRange.Columns.Count
                Block = Trim(sht.Cells(1, Col).value): Block = Replace(Block, "spi", "rtos", compare:=vbTextCompare)
                header = Trim(sht.Cells(2, Col).value)
                fail = Trim(sht.Cells(3, Col).value)
                For Row = 4 To sht.Cells(1, Col).End(xlDown).Row
                    mode = sht.Cells(Row, 3).value
                    val = sht.Cells(Row, Col).value
                    If UCase(fail) Like BINNINGFAIL Then
                        bin = UCase("Bin_" + mode + "_" + Block + "_" + "IDS")
                    ElseIf UCase(fail) Like LVCCFAIL Then
                        bin = UCase("Bin_" + mode + "_" + Block + "_" + "BV")
                    Else
                        Stop
                    End If
                    
                    If UCase(header) = "HARDBIN" Then
                        If Not hard.Exists(bin) Then hard.Add bin, val
                        If Not hard.item(bin) = sht.Cells(Row, Col).value Then Stop
                        If LCase(Block) = "rtos" Then
                            ref = Replace(bin, "RTOS", "ILB", compare:=vbTextCompare)
                            If Not hard.Exists(ref) Then hard.Add ref, val
                            If Not hard.item(ref) = sht.Cells(Row, Col).value Then Stop
                        End If
                    ElseIf UCase(header) = "SOFTBIN" Then
                        If Not soft.Exists(bin) Then soft.Add bin, val
                        If Not checked.Exists(bin) Then checked.Add bin, False
                        If Not soft.item(bin) = val Then Stop
                        If LCase(Block) = "rtos" Then
                            ref = Replace(bin, "RTOS", "ILB", compare:=vbTextCompare)
                            If Not soft.Exists(ref) Then soft.Add ref, val
                            If Not soft.item(ref) = sht.Cells(Row, Col).value Then Stop
                        End If
                        
                        If Not soft_rev.Exists(val) Then soft_rev.Add val, bin
                        If soft_rev.Exists(val) And soft_rev.item(val) <> bin Then Stop
                    End If
                Next Row
            Next Col
        End If
    Next name
    
    shts = TheExec.job.GetSheetNamesOfType(DMGR_SHEET_TYPE_BINTABLESSHEET)
    For Each name In shts
        If UCase(name) Like BINTABLE Then
            Set sht = Sheets(name)
            sht.Activate
            
            For Row = 4 To 980 'sht.UsedRange.Rows.Count
                bin = UCase(Trim(sht.Cells(Row, 2).value))
                val = sht.Cells(Row, 5).value
                val1 = sht.Cells(Row, 6).value
                
                If soft.Exists(bin) Then
                    checked(bin) = True
                    If soft.item(bin) <> val Then
                        If Not DEBUG_ Then sht.Cells(Row, 5).value = soft.item(bin)
                        Debug.Print "Update soft bin:" + bin + " " + str(val) + " ->" + str(soft.item(bin))
'                        If soft.item(bin) - val <> 10 Then
'                            Debug.Print bin + ":" + str(val) + "->" + str(soft.item(bin))
'                        End If
                    End If
                    If hard.item(bin) <> val1 Then
                        If Not DEBUG_ Then sht.Cells(Row, 6).value = hard.item(bin)
                        Debug.Print "Update hard bin:" + bin + " " + str(val) + "->" + str(hard.item(bin))
'                        Debug.Print bin + ":" + str(val1) + "->" + str(hard.item(bin))
                    End If
                ElseIf soft_rev.Exists(val) Then
                    If soft_rev.item(val) <> bin Then
                        ''' means we have vddbinning soft bin in bintable but with different bin name
                        Debug.Print "Unexpected name:" + bin + " for soft bin:" + CStr(val)
                    End If
                End If
            Next Row
        End If
    Next
    
    For Each name In checked.Keys
        If Not checked.item(name) Then
            Debug.Print "Missing bin name in bintable: " + name
        End If
    Next name
    
    Set soft = Nothing: Set hard = Nothing: Set soft_rev = Nothing: Set checked = Nothing
    Debug.Print "Complete the change"
End Function

Public Function GenChar(Optional pos As Integer = 3, Optional cond As String = "LV", Optional pivot As Integer = 0) As String
    Dim sht As Worksheet
    Dim Row As Integer
    
    'Set sht = Sheets("TestInst_Vddbinning")
    Set sht = ThisWorkbook.ActiveSheet
    Row = sht.UsedRange.Rows.Count
    If pivot <> 0 Then Row = pivot - 2
    sht.Cells(Row + 2, 46).value = "INIT_NV_PL_SWEEP"
    sht.Cells(Row + 2, 49).value = "INIT" + CStr(pos) + ":5"
    sht.Cells(Row + 2, 50).value = "INIT" + CStr(pos) + ":sgmt0=SELSRAM"
    sht.Cells(Row + 2, 51).value = "INIT" + CStr(pos) + ":JTAG_TDI"
    sht.Cells(Row + 2, 52).value = "INIT" + CStr(pos) + ":sgmt0_5"
    sht.Cells(Row + 2, 53).value = False
    sht.Cells(Row + 2, 55).value = "SelSrmSSSSS"
    sht.Cells(Row + 2, 56).value = "Bincut_X_X_X:" + cond
    sht.Cells(Row + 2, 57).value = False
    
    GenChar = "INIT" + CStr(pos) + ":5" & vbTab & "INIT" + CStr(pos) + ":sgmt0=SELSRAM INIT" + CStr(pos) + ":JTAG_TDI  INIT" + CStr(pos) + ":sgmt0_5   FALSE       SelsrmSSSSS Bincut_X_X_X:" + cond + " FALSE"
End Function

Public Function PatRemove(Optional clusters As String, Optional coverages As String = "", Optional patset As String = "PatSets_BinCut")

    Dim sht As Worksheet
    Dim pat As Variant
    Dim MaxRow As Long
    Dim Row As Long
    Dim idx As Integer
    Dim val As String
    Dim name As String
    Dim info() As String
    Dim found As Boolean
    Dim cluster As New Dictionary
    Dim coverage As New Dictionary
    
    If clusters = "" And coverages = "" Then Exit Function
    
    For Each pat In Split(clusters, ",")
        If Not cluster.Exists(UCase(Trim(pat))) Then
            cluster.Add UCase(Trim(pat)), True
        End If
    Next
    
    For Each pat In Split(coverages, ",")
        If Not coverage.Exists(UCase(Trim(pat))) Then
            coverage.Add UCase(Trim(pat)), True
        End If
    Next pat

    Set sht = Sheets(patset)
    sht.Activate
    sht.Rows(4).Insert
    
    Row = 5
    MaxRow = sht.UsedRange.Rows.Count
    Do While Row <= MaxRow
        If Trim(sht.Cells(Row, 2).value) = "" Then Exit Do
        found = False
        val = UCase(Trim(sht.Cells(Row, 5).value))
        info = Split(val, "_")
        
        Select Case (info(2))
            Case "S":
                If info(4) = "SC" And info(3) Like "PL??" And (cluster.Exists(info(5)) Or coverage.Exists(info(3))) Then
                    found = True
                End If
            Case "C":
                If info(4) = "SC" And info(3) Like "PL??" And (cluster.Exists(info(UBound(info))) Or coverage.Exists(info(3))) Then
                    found = True
                End If
            Case "L":
                If info(4) = "SC" And Not info(3) = "INLP" And (cluster.Exists(info(5)) Or coverage.Exists(info(3))) Then
                    found = True
                End If
        End Select
        
        If found Then
            sht.Rows(Row).delete shift:=xlUp
            Do While Row >= 5
                val = UCase(Trim(sht.Cells(Row - 1, 5).value))
                If val Like "*_PL??_*" Then Exit Do
                sht.Rows(Row - 1).delete shift:=xlUp
                Row = Row - 1
            Loop
        Else
            Row = Row + 1
        End If
    Loop
    
    sht.Rows(4).delete shift:=xlUp
    Debug.Print "Done"
End Function

Private Function GetModuleSuffix(moduleType As Integer) As String
    GetModuleSuffix = IIf(moduleType = 1, ".bas", IIf(moduleType = 2, ".cls", ""))
End Function
Private Function GetModuleFromName(name As String) As Object
    On Error Resume Next
    Set GetModuleFromName = ThisWorkbook.VBProject.VBComponents(Trim(name))
End Function
Public Function ExportModule(Optional exportTarget As String = "ALL", Optional FilePath As String = "D:\thera_tp\Module\Library\")
    Dim module As Object
    Dim suffix As String
    Dim name As Variant

    Select Case UCase(exportTarget)
        Case "ALL":
            For Each module In ThisWorkbook.VBProject.VBComponents
                suffix = GetModuleSuffix(module.type)
                If suffix = "" Then GoTo nextModule
                Debug.Print "Export file: " + module.name + suffix + " to: " + FilePath
                module.Export FilePath + module.name + suffix
nextModule:
            Next module
        Case "AUTOSHMOO"
            'filePath = "D:\thera_tp\ENG\"
            For Each module In ThisWorkbook.VBProject.VBComponents
                If LCase(module.name) Like "autoshmoo*" Then
                    suffix = GetModuleSuffix(module.type)
                    Debug.Print "Export file: " + module.name + suffix + " to: " + FilePath
                    module.Export FilePath + module.name + suffix
                End If
            Next module
        Case "VDDBINNING"
            For Each module In ThisWorkbook.VBProject.VBComponents
                If LCase(module.name) Like "*vdd_binning*" Then
                    suffix = GetModuleSuffix(module.type)
                    Debug.Print "Export file: " + module.name + suffix + " to: " + FilePath
                    module.Export FilePath + module.name + suffix
                End If
            Next module
        Case Else:
            For Each name In Split(exportTarget, ",")
                On Error Resume Next
                Set module = GetModuleFromName(CStr(name))
                If module Is Nothing Then GoTo nextName
                suffix = GetModuleSuffix(module.type)
                If suffix = "" Then GoTo nextName
                Debug.Print "Export file: " + module.name + suffix + " to: " + FilePath
                module.Export FilePath + module.name + suffix
nextName:
            Next name
    End Select

End Function

