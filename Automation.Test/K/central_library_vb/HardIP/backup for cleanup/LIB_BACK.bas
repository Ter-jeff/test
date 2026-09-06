Attribute VB_Name = "LIB_BACK"
Public Function BinDSP2Str(InWave As DSPWave, index As DSPWave) As String
Dim t As Variant
Dim i, j, k As Long
'    For Each T In InWave.Data
'       BinDSP2Str = BinDSP2Str & CStr(T)
'    Next T
    For Each i In index.data
        For j = i To 1 Step -1
            BinDSP2Str = BinDSP2Str & CStr(InWave.Element(k))
            k = k + 1
        Next j
        BinDSP2Str = BinDSP2Str & " "
    Next i
End Function

Public Function DACTrim_DigSrcSetting(Trimname As String, TrimDataWidth As Long, TrimEq As String, digsrc_assignment As String, TrimSrcWf As DSPWave)
Dim TempSemi() As String
Dim Eq() As String
Dim Long_Wf As New DSPWave
Dim TempEqual() As String
Dim TempAnd() As String
Dim TempStr As String
Dim index, IndexColon, IndexEqual, IndexSemi, IndexAnd As Variant
Dim AssinWf() As New DSPWave
Dim assignment As New SiteVariant
Dim i, j, k As Integer
Dim WfDic As New Dictionary
Dim site As Variant
Dim SiteSelect As New SiteBoolean
'AA=001&DicA&1110 ; BB=1111 ; CC=DicC
    SiteSelect = TheExec.sites.Selected
    TheExec.sites.Selected = True
    ReDim AssinWf(0)
    digsrc_assignment = UCase(digsrc_assignment)
    Call EqMapping(TrimEq, digsrc_assignment)
    Eq = Split(TrimEq, "+")
    If Trimname <> "VERIFICATION" And Trimname <> "READEFUSE" Then
        For i = 1 To TrimDataWidth: TempStr = "0" & TempStr: Next i
        assignment = SiteExpand(TempStr)
        Call BinStr2DWave(assignment, AssinWf(0), CStr(Trimname))
        Call WfDic.Add(Trimname, AssinWf(UBound(AssinWf)))
    Else
        For Each index In Eq
            If DACTargetStr.Exists(index) And Not DACTargetStr.Exists("DSPWF@" & index) Then
                Long_Wf.CreateConstant 0, 1
                For Each site In TheExec.sites
                    TempStr = fuseCode(CStr(index))
                    TrimDataWidth = Len(TempStr)
                    Long_Wf.Element(0) = BinStr2Dec(TempStr)
                Next site
                ''OscarLi_Compile,20190629
                'Call rundsp.DSPWf_Dec2Binary(Long_Wf, TrimDataWidth, TrimSrcWf)
                Exit For
            End If
        Next index
        Exit Function
    End If
    Set G_TrimWave = Nothing
    Set RPIndex = Nothing
    If digsrc_assignment <> "" Then
        TempSemi = Split(digsrc_assignment, ";")
        For Each IndexSemi In TempSemi
            TempEqual = Split(IndexSemi, "="): i = 0: ReDim AssinWf(0)
            For Each IndexEqual In TempEqual
                If IndexEqual <> TempEqual(1) Then IndexEqual = TempEqual(1) Else Exit For
                If InStr(1, IndexEqual, "&") > 0 Then
                    TempAnd = Split(IndexEqual, "&")
                    For Each IndexAnd In TempAnd
                        If CStr(IndexEqual) <> Trimname Then
                            If BinStr2Dec(CStr(IndexAnd)) = "OverFlow" Then
                                ReDim Preserve AssinWf(i): i = i + 1
                                AssinWf(i - 1) = GetStoredCaptureData(CStr(IndexAnd))
                            Else
                                ReDim Preserve AssinWf(i): i = i + 1
                                assignment = SiteExpand(IndexAnd)
                                Call BinStr2DWave(assignment, AssinWf(i - 1), CStr(IndexAnd))
                            End If
                        Else
                            ReDim Preserve AssinWf(i): i = i + 1
                            AssinWf(i - 1) = WfDic(Trimname)
                        End If
                    Next IndexAnd
                    If i > 1 Then
                        For i = 0 To UBound(AssinWf) - 1
                            Call rundsp.DSPWf_Merge(AssinWf(i), AssinWf(i + 1), AssinWf(i + 1))
                        Next i
                    End If
                    Call WfDic.Add(TempEqual(0), AssinWf(UBound(AssinWf)))
                Else
                    If CStr(IndexEqual) <> Trimname Then
                        If BinStr2Dec(CStr(IndexEqual)) = "OverFlow" Then
                            ReDim Preserve AssinWf(i): i = i + 1
                            AssinWf(i - 1) = GetStoredCaptureData(CStr(IndexEqual))
                        Else
                            ReDim Preserve AssinWf(i): i = i + 1
                            assignment = SiteExpand(IndexEqual)
                            Call BinStr2DWave(assignment, AssinWf(i - 1), CStr(IndexEqual))
                        End If
                        Call WfDic.Add(TempEqual(0), AssinWf(UBound(AssinWf)))
                    End If
                End If
            Next IndexEqual
        Next IndexSemi
    End If
        
    ReDim AssinWf(0): i = 0
    'RPIndex.CreateConstant 0, 1
    For Each index In Eq
        If index = Trimname Then
            ReDim Preserve AssinWf(i): i = i + 1
            AssinWf(i - 1) = WfDic(Trimname)
            Call rundsp.RPIndexSetting(AssinWf(i - 1), RPIndex, 0.1)
        ElseIf WfDic.Exists(index) Then
            ReDim Preserve AssinWf(i): i = i + 1
            AssinWf(i - 1) = WfDic(index)
            Call rundsp.RPIndexSetting(AssinWf(i - 1), RPIndex, 0)
        ElseIf DACTargetStr.Exists("DSPWF@" & index) Then
            ReDim Preserve AssinWf(i): i = i + 1
            AssinWf(i - 1) = DACTargetStr("DSPWF@" & index)
            Call rundsp.RPIndexSetting(AssinWf(i - 1), RPIndex, 0)
        Else
            ReDim Preserve AssinWf(i): i = i + 1
            AssinWf(i - 1) = GetStoredCaptureData(CStr(index))
            Call rundsp.RPIndexSetting(AssinWf(i - 1), RPIndex, 0)
        End If
        
    Next index

    If i > 1 Then
        For i = 0 To UBound(AssinWf) - 1
            Call rundsp.DSPWf_Merge(AssinWf(i), AssinWf(i + 1), AssinWf(i + 1))
        Next i
        G_TrimWave = AssinWf(UBound(AssinWf))
    Else
        G_TrimWave = AssinWf(UBound(AssinWf))
    End If
    
    TheExec.sites.Selected = SiteSelect
    
End Function

'===========================================================================================
' Check if pattern name provided is a pattern set
' NOTE: If pattern set is true and count > 1 then the elements returned may still be
'          nested pattern sets.  The calling function should recursively call this to
'          ensure that the returned Names resolve to individual patterns
'===========================================================================================
Public Function DecomposePatternSet(TestPat As String, _
                              rtnPatNames() As String, _
                              rtnPatCnt As Long) As Boolean

    Dim PatCnt As Long                          '<- Number of patterns in set
    Dim RawNameData() As String                 '<- Raw pattern name data
    Dim rtnPatNames1() As String
    Dim rtnPatNames2() As String
    Dim i As Long, j As Long
    '___ Init _____________________________________________________________________________
    '    On Error GoTo errhandler
    '___ Check the name ___________________________________________________________________
    '    Individual pattern name or non-pattern string returns an error - thus false
    '--------------------------------------------------------------------------------------
    rtnPatNames = TheExec.DataManager.Raw.GetPatternsInSet(TestPat, PatCnt)
    If (UBound(rtnPatNames) > 0) Then
        If LCase(rtnPatNames(0)) Like "*.pat*" Then
            DecomposePatternSet = True
            rtnPatCnt = UBound(rtnPatNames) + 1
        Else
            rtnPatCnt = 0
            For i = 0 To UBound(rtnPatNames)
                rtnPatNames2 = TheExec.DataManager.Raw.GetPatternsInSet(rtnPatNames(i), PatCnt)
                rtnPatCnt = rtnPatCnt + UBound(rtnPatNames2) + 1
            Next i
            rtnPatNames1 = TheExec.DataManager.Raw.GetPatternsInSet(TestPat, PatCnt)
            ReDim rtnPatNames(rtnPatCnt - 1)    ' modify 827 j
            rtnPatCnt = 0
            For i = 0 To UBound(rtnPatNames1)
                rtnPatNames2 = TheExec.DataManager.Raw.GetPatternsInSet(rtnPatNames1(i), PatCnt)
                For j = 0 To UBound(rtnPatNames2)
                    If LCase(rtnPatNames2(j)) Like "*.pat*" Then
                        rtnPatNames(rtnPatCnt) = rtnPatNames2(j)
                    Else
                        TheExec.ErrorLogMessage TestPat & " in more than 2 level of pattern set"
                    End If
                    rtnPatCnt = rtnPatCnt + 1
                Next j
            Next i
            DecomposePatternSet = True
        End If
    Else
        If LCase(rtnPatNames(0)) Like "*.pat*" Then
            DecomposePatternSet = True
            rtnPatCnt = 1
        Else
            rtnPatNames = TheExec.DataManager.Raw.GetPatternsInSet(rtnPatNames(0), PatCnt)
            rtnPatCnt = UBound(rtnPatNames) + 1
            For j = 0 To UBound(rtnPatNames)
                If LCase(rtnPatNames(j)) Like "*.pat*" Then
                Else
                    TheExec.ErrorLogMessage TestPat & " in more than 2 level of pattern set"
                End If
            Next j
        End If
    End If
    
    Exit Function
    
errHandler:
    DecomposePatternSet = False
    rtnPatCnt = -1
    Exit Function

End Function
Public Function DACTrim_WriteFuncResult(Optional SpecialReserve As String = vbNullString, Optional CodeSearchPatternResult As SiteBoolean, Optional m_testName As String = vbNullString) As Long
    Dim site As Variant
    Dim TestNumber As Long
    Dim FailCount As New PinListData
    Dim AllPins As PinList
    Dim pin As Variant
    Dim Pins() As String
    Dim Pin_Cnt As Long
    
    '' 20150604: Need to modify "All_Digital" to the parameter.
    TheExec.DataManager.DecomposePinList "All_Digital", Pins(), Pin_Cnt
    
    If SpecialReserve <> "" Then
        If SpecialReserve = "DSSC_CODESEARCH" Then
            For Each site In TheExec.sites
                TestNumber = TheExec.sites.item(site).TestNumber
                If CodeSearchPatternResult(site) Then
                    
                    If TheExec.DevChar.Setups.IsRunning = True Then TheExec.sites.item(site).testResult = sitePass

                    ''''20151106 update
                    If (m_testName <> "") Then
                        Call TheExec.Datalog.WriteFunctionalResult(site, TestNumber, logTestPass, , m_testName)
                    Else
                        Call TheExec.Datalog.WriteFunctionalResult(site, TestNumber, logTestPass)
                    End If
                Else
                    
''                    TheExec.Sites.Item(Site).TestResult = siteFail

                    ''''20151106 update
                    If (m_testName <> "") Then
                        Call TheExec.Datalog.WriteFunctionalResult(site, TestNumber, logTestFail, , m_testName)
                    Else
                        Call TheExec.Datalog.WriteFunctionalResult(site, TestNumber, logTestFail)
                    End If
                    '' 20160218 - Modify sequence to let TestResult after WriteFunctionalResult to cover test number increment 2 issue if souce sink time out alarm happen.
                    TheExec.sites.item(site).testResult = siteFail
                    
                    For Each pin In Pins
        
                        If TheExec.DataManager.ChannelType(pin) <> "N/C" Then
                            FailCount = TheHdw.Digital.Pins(pin).FailCount
                            If FailCount <> 0 Then
                                TheExec.Datalog.WriteComment "===> Pin " & pin & " Fail count =" & FailCount
                            End If
                        End If
                    Next pin
                End If
                TheExec.sites.item(site).TestNumber = TestNumber + 1
            Next site
        End If
    
    Else
''        '' 20150604: Need to modify "All_Digital" to the parameter.
''        TheExec.DataManager.DecomposePinList "All_Digital", Pins(), pin_cnt
    
        For Each site In TheExec.sites
            TestNumber = TheExec.sites.item(site).TestNumber
            If TheHdw.Digital.Patgen.PatternBurstPassed(site) Then
                
                If TheExec.DevChar.Setups.IsRunning = True Then TheExec.sites.item(site).testResult = sitePass
                
                ''''20151106 update
                If (m_testName <> "") Then
                    Call TheExec.Datalog.WriteFunctionalResult(site, TestNumber, logTestPass, , m_testName)
                Else
                    Call TheExec.Datalog.WriteFunctionalResult(site, TestNumber, logTestPass)
                End If
                
            Else
                
''                TheExec.Sites.Item(Site).TestResult = siteFail
                
                ''''20151106 update
                If (m_testName <> "") Then
                    Call TheExec.Datalog.WriteFunctionalResult(site, TestNumber, logTestFail, , m_testName)
                Else
                    Call TheExec.Datalog.WriteFunctionalResult(site, TestNumber, logTestFail)
                End If
                '' 20160218 - Modify sequence to let TestResult after WriteFunctionalResult to cover test number increment 2 issue if souce sink time out alarm happen.
                TheExec.sites.item(site).testResult = siteFail
                
                For Each pin In Pins
    
                    If TheExec.DataManager.ChannelType(pin) <> "N/C" Then
                        FailCount = TheHdw.Digital.Pins(pin).FailCount
                        If FailCount <> 0 Then
                            TheExec.Datalog.WriteComment "===> Pin " & pin & " Fail count =" & FailCount
                        End If
                    End If
                Next pin
            End If
            If TheExec.DevChar.Setups.IsRunning = False Then TheExec.sites.item(site).TestNumber = TestNumber + 1
        Next site
    End If
End Function
Public Function StartFRC(PortName As String) As Double
Dim site As Variant
Dim i As Long
Dim PLLLockChecked As New SiteLong
'Dim measf As New PinListData
Dim NotLocked As Boolean
'Dim XI0_REFCLK As String
Dim PortMode As String
Dim PLL_Lock As New SiteLong
Dim PortArray() As String
Dim index As Variant

PortArray = Split(PortName, ",")
For Each index In PortArray
    
    For Each site In TheExec.sites
        If TheHdw.Protocol.ports(index).Enabled = True Then
            TheHdw.Protocol.ports(index).Halt
            TheHdw.Protocol.ports(index).Enabled = False
        End If
    Next site
'    thehdw.Digital.Pins("REFCLK1").InitState = chInitoff
'    thehdw.Digital.Pins("REFCLK2").InitState = chInitoff
'    thehdw.Wait 0.005
    
'    thehdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
 
'    thehdw.Wait 0.005
    TheHdw.Protocol.ports(index).Enabled = True
    TheHdw.Protocol.ports(index).NWire.ResetPLL
    
    TheHdw.Wait 0.005
    
    For Each site In TheExec.sites
        If TheHdw.Protocol.ports(index).NWire.IsPLLLocked = False Then
            PLLLockChecked = 1
            TheExec.Datalog.WriteComment "print: site(" & site & "), PortName(" & index & "), IsPLLLocked = False"
            PLL_Lock(site) = 0
        Else
            PLLLockChecked = 0
            'TheExec.Datalog.WriteComment "print: site(" & site & "), PortName(" & PortName & "), IsPLLLocked = True"
            'TheExec.Datalog.WriteComment "print: site(" & site & "), PortName(" & PortName & "), Wake up finished"
            PLL_Lock(site) = 1
        End If
    Next site
    
    Call TheHdw.Protocol.ports(index).NWire.Frames("RunFreeClock").Execute
    TheHdw.Protocol.ports(index).IdleWait
    
    'TheExec.Flow.TestLimit PLL_Lock, 1, 1, tlSignGreaterEqual, tlSignLessEqual, Tname:="nWirePLL_Lock" 'BurstResult=1:Pass

    TheExec.Datalog.WriteComment vbNullString
    TheExec.Datalog.WriteComment "********** Enable freerunning clock *********"
              
Next index

End Function
Public Function StopFRC(PortName As String) As Double
    Dim site As Variant
    Dim index As Variant
    Dim PortArray() As String
    
    PortArray = Split(PortName, ",")
    For Each index In PortArray
        For Each site In TheExec.sites
            If TheHdw.Protocol.ports(index).Enabled = True Then
                TheHdw.Protocol.ports(index).Halt
                TheHdw.Protocol.ports(index).Enabled = False
            End If
        Next site
    Next index
    TheExec.Datalog.WriteComment vbNullString
    TheExec.Datalog.WriteComment "********** Disable freerunning clock *********"
End Function
Public Function Select_Measure_Pin(TestSeqNum As Integer, MeasPinAry() As String, ByRef Measure_Pin As PinList)
    Dim TestSeqNumIdx As Integer
    TestSeqNumIdx = TestSeqNum
    
    If TestSeqNum > 0 Then
        If UBound(MeasPinAry) = 0 Then TestSeqNumIdx = 0
    End If
    If UBound(MeasPinAry) >= 0 Then
        '' 20150605 - Check with CC
''        If InStr(LCase(MeasPinAry(TestSeqNumIdx)), "idx") <> 0 Then TestSeqNumIdx = Int(Mid(MeasPinAry(TestSeqNumIdx), 4, 1))
        Measure_Pin = MeasPinAry(TestSeqNumIdx)
    End If
End Function
Public Function Select_MeasureI_CurrentRange(TestSeqNum As Integer, MeasPinAry_IRange() As String, ByRef MeasureI_CurrentRange As String)
    Dim TestSeqNumIdx As Integer
    TestSeqNumIdx = TestSeqNum
    
    If TestSeqNum > 0 Then
        If UBound(MeasPinAry_IRange) = 0 Then TestSeqNumIdx = 0
    End If
    If UBound(MeasPinAry_IRange) >= 0 Then
        '' 20150605 - Check with CC
''        If InStr(LCase(MeasPinAry_IRange(TestSeqNumIdx)), "idx") <> 0 Then TestSeqNumIdx = Int(Mid(MeasPinAry_IRange(TestSeqNumIdx), 4, 1))
        MeasureI_CurrentRange = MeasPinAry_IRange(TestSeqNumIdx)
    End If
End Function
Public Function DACTrim_FrequencyMeasureSingleEnd(FreqMeasPins As PinList, _
        d_FreqMeasInterval As Double, Optional MeasFreqWaitTime As String = vbNullString, Optional MeasF_EventSource As FreqCtrEventSrcSel, Optional MeasF_ToTestLimit As PinListData)
        
    'Dim p As Long
    'Dim MeasFreq As New PinListData
    Dim CounterValue As New PinListData
    Dim t_interval As New SiteDouble
        'Call Freq_MeasFreqSetup(FreqMeasPins, d_FreqMeasInterval, MeasF_EventSource)   '' 20150621 - default d_FreqMeasInterval = 0.01
        '' 20150623 - Add Customize Wait Time
        'Call HardIP_Freq_MeasFreqStart(FreqMeasPins, d_FreqMeasInterval, MeasFreq, MeasFreqWaitTime)       '' 20150621 - default d_FreqMeasInterval = 0.01
    
    With TheHdw.Digital.Pins(FreqMeasPins).FreqCtr
        .EventSource = MeasF_EventSource '' VOH
        .EventSlope = Positive
        .Interval = d_FreqMeasInterval
        .Enable = IntervalEnable
        .Clear
    End With
    
    TheHdw.Digital.Pins(FreqMeasPins).FreqCtr.Clear
    TheHdw.Digital.Pins(FreqMeasPins).FreqCtr.start
    
    'For Each Site In TheExec.Sites
    CounterValue = TheHdw.Digital.Pins(FreqMeasPins).FreqCtr.Read
    t_interval = TheHdw.Digital.Pins(FreqMeasPins).FreqCtr.Interval
    MeasF_ToTestLimit = CounterValue.Math.divide(t_interval)
    'Next Site
        

'''''    Dim TestNameInput As String
'''''    TestNameInput = "Freq_meas_"
'''''
'''''''    If UCase(G_TestName) Like UCase("*CPU*") Then
'''''''       TestNameInput = G_TestName
'''''''    End If
'''''
'''''        If Flag_SingleLimit = True Then
'''''''        If CurrentJobName_L Like "*char*" Then
'''''''            TheExec.Flow.TestLimit MeasFreq, Freq_LowLimit, Freq_HighLimit, , , , unitHz, , ForceResults:=tlForceFlow
'''''''        Else
'''''''            '' 20151102 - Modify test name as Chihome
'''''''            '' TheExec.Flow.TestLimit MeasFreq, Freq_LowLimit, Freq_HighLimit, , , , unitHz, , Tname:=TestNameInput + CStr(TestSeqNum) + "_" + "@COND:PATTERN=" + PATT_ExculdePath(Pat), ForceResults:=tlForceFlow
'''''''            If UCase(G_TestName) Like UCase("*CPU*") Then
'''''''                Call UpdateDLogColumns(Len(TestNameInput + CStr(TestSeqNum)) + 6)
'''''''                TheExec.Flow.TestLimit MeasFreq, Freq_LowLimit, Freq_HighLimit, , , , unitHz, , Tname:=TestNameInput + "_Feq" + CStr(TestSeqNum), ForceResults:=tlForceFlow
'''''''                Call UpdateDLogColumns__False
'''''''            Else
'''''''                TheExec.Flow.TestLimit MeasFreq, Freq_LowLimit, Freq_HighLimit, , , , unitHz, , Tname:=TestNameInput + CStr(TestSeqNum), ForceResults:=tlForceFlow
'''''''            End If
'''''''
'''''''        End If
'''''''        Else
'''''            If Mid(TestLimitByPin_VFI, 2, 1) = "T" Then
'''''                If IsDifferentialPin = True Then
'''''                    For p = 0 To MeasFreq.Pins.Count - 1 Step 2 ' freq counter result of differential pins is stored in positive pin
'''''                         ''TheExec.Flow.TestLimit resultVal:=MeasFreq.Pins(p + 1), unit:=unitHz, Tname:=TestNameInput + CStr(TestSeqNum) + "_" + "@COND:PATTERN=" + PATT_ExculdePath(Pat), ForceResults:=tlForceFlow
'''''                         TheExec.Flow.TestLimit resultVal:=MeasFreq.Pins(p + 1), unit:=unitHz, Tname:=TestNameInput + CStr(TestSeqNum), ForceResults:=tlForceFlow
'''''                    Next p
'''''                Else
'''''                    For p = 0 To MeasFreq.Pins.Count - 1
'''''                         '' TheExec.Flow.TestLimit resultVal:=MeasFreq.Pins(p), unit:=unitHz, Tname:=TestNameInput + CStr(TestSeqNum) + "_" + "@COND:PATTERN=" + PATT_ExculdePath(Pat), ForceResults:=tlForceFlow
'''''                         TheExec.Flow.TestLimit resultVal:=MeasFreq.Pins(p), unit:=unitHz, Tname:=TestNameInput + CStr(TestSeqNum), ForceResults:=tlForceFlow
'''''                    Next p
'''''                End If
'''''            Else
'''''                '' TheExec.Flow.TestLimit resultVal:=MeasFreq, unit:=unitHz, Tname:=TestNameInput + CStr(TestSeqNum) + "_" + "@COND:PATTERN=" + PATT_ExculdePath(Pat), ForceResults:=tlForceFlow
'''''                TheExec.Flow.TestLimit resultVal:=MeasFreq, unit:=unitHz, Tname:=TestNameInput + CStr(TestSeqNum), ForceResults:=tlForceFlow
'''''            End If
'''''        End If
'''''
'''''        '' 20151224 - Merge print measured frequency during shmoo if need
'''''        G_MeasFreqForCZ = MeasFreq
End Function
''''Public Function DACTrim_MeasureVolt(MeasureV_Pin As PinList, _
''''                       TestSeqNum As Integer, k As Long, _
''''                       Optional InstSpecialSetting As InstrumentSetup = 0, Optional CUS_STR_MainProgram As String = "", Optional MeasV_ToTestLimit As PinListData) As Long
''''
''''    Dim site As Variant
''''    Dim p As Long
''''
'''''    ''========================================================================================
'''''    '' 20150202 - Range Check
'''''    Dim RangeCheck_HighLimitVal() As Double
'''''    Dim RangeCheck_LowLimitVal() As Double
'''''    Dim TempMeasVal_PerPin(100) As New PinListData
'''''    If Range_Check_Enable_Word = True Then
'''''        Call GetFlowSingleUseLimit(RangeCheck_HighLimitVal, RangeCheck_LowLimitVal)
'''''    End If
'''''    ''========================================================================================
''''
''''    '' 20150114
''''    Dim MeasureV_Pin_IO As String
''''    Dim MeasureV_Pin_UVI80 As String
''''    Dim MeasV_INstType_Num As Integer
''''    MeasV_INstType_Num = 0
''''
''''    Call DiscriminateMeasV_PinType(MeasureV_Pin, MeasureV_Pin_IO, MeasureV_Pin_UVI80, MeasV_INstType_Num)
''''    MeasV_INstType_Num = MeasV_INstType_Num - 1
''''
''''    ReDim MeasVoltage(MeasV_INstType_Num) As New PinListData
''''
''''    Dim index As Integer
''''    index = 0
''''    If MeasureV_Pin_UVI80 <> "" Then
''''        Call DACTrim_SetupAndMeasureVolt_UVI80(MeasureV_Pin_UVI80, MeasVoltage(index))
''''        index = index + 1
''''    End If
''''    If MeasureV_Pin_IO <> "" Then
''''        Call DACTrim_SetupAndMeasureVolt_PPMU(MeasureV_Pin_IO, MeasVoltage(index), InstSpecialSetting)
''''        index = index + 1
''''    End If
''''    '' 20150608 - Merge MeasVoltage to the same pin list data if instrument over 1 type.
''''    'Dim MeasV_ToTestLimit As New PinListData
''''    If MeasV_INstType_Num = 0 Then
''''        If MeasureV_Pin_IO <> "" Then
''''            MeasV_ToTestLimit = MeasVoltage(0)
''''        Else
''''            MeasV_ToTestLimit = MeasVoltage(0)
''''        End If
''''    Else
''''        Call DACTrim_MergePinListData(MeasV_INstType_Num, MeasVoltage, MeasV_ToTestLimit)
''''    End If
''''
''''
'''''''    Dim TestNameInput As String
'''''''    TestNameInput = "Volt_meas_"
'''''''    Dim MipiResult As New PinListData
'''''''   'Already do pin merge for different pin type, need to test limit by pin or only one test limit
'''''''    If Flag_SingleLimit = True Then
'''''''        TheExec.Flow.TestLimit MeasV_ToTestLimit, LowLimitVal, HighLimitVal, , , , unitVolt, , Tname:=TestNameInput + CStr(TestSeqNum), ForceResults:=tlForceFlow
'''''''        '' 20150202 - Range Check
'''''''        If Range_Check_Enable_Word = True Then
'''''''            Call CheckRangesAndClamps(MeasV_ToTestLimit, "V", RangeCheck_HighLimitVal(gl_UseLimitCheck_Counter), RangeCheck_LowLimitVal(gl_UseLimitCheck_Counter))
'''''''            gl_UseLimitCheck_Counter = gl_UseLimitCheck_Counter + 1
'''''''        End If
'''''''
'''''''    Else
'''''''        If Mid(TestLimitByPin_VFI, 1, 1) = "T" Then
'''''''            For p = 0 To MeasV_ToTestLimit.Pins.Count - 1
'''''''                If CurrentJobName_L Like "*char*" Then
'''''''                    TheExec.Flow.TestLimit MeasV_ToTestLimit.Pins(p), , , , , , unitVolt, , , ForceResults:=tlForceFlow
'''''''                Else
'''''''                    TheExec.Flow.TestLimit MeasV_ToTestLimit.Pins(p), , , , , , unitVolt, , Tname:=TestNameInput + CStr(TestSeqNum), ForceResults:=tlForceFlow
'''''''                End If
'''''''
'''''''                '' 20150202 - Range Check
'''''''                If Range_Check_Enable_Word = True Then
'''''''                    TempMeasVal_PerPin(p).AddPin (MeasV_ToTestLimit.Pins(p))
'''''''                    TempMeasVal_PerPin(p).Pins(MeasV_ToTestLimit.Pins(p)) = MeasV_ToTestLimit.Pins(p)
'''''''                    Call CheckRangesAndClamps(TempMeasVal_PerPin(p), "V", RangeCheck_HighLimitVal(gl_UseLimitCheck_Counter), RangeCheck_LowLimitVal(gl_UseLimitCheck_Counter))
'''''''                    gl_UseLimitCheck_Counter = gl_UseLimitCheck_Counter + 1
'''''''                End If
'''''''            Next p
'''''''        Else
'''''''            If CurrentJobName_L Like "*char*" Then
'''''''                TheExec.Flow.TestLimit MeasV_ToTestLimit.Pins(p), , , , , , unitVolt, , , ForceResults:=tlForceFlow
'''''''            Else
'''''''                TheExec.Flow.TestLimit MeasV_ToTestLimit, , , , , , unitVolt, , Tname:=TestNameInput + CStr(TestSeqNum), ForceResults:=tlForceFlow
'''''''            End If
'''''''
'''''''            If CUS_STR_MainProgram = "MIPI_DPHY" Then
'''''''                MipiResult = MeasV_ToTestLimit
'''''''                For p = 0 To MeasV_ToTestLimit.Pins.Count - 1
'''''''                    For Each Site In TheExec.Sites.Active
'''''''
'''''''                        If (MeasV_ToTestLimit.Pins(p).Value(Site) > -0.05 And MeasV_ToTestLimit.Pins(p).Value(Site) < 0.05) Or (MeasV_ToTestLimit.Pins(p).Value(Site) > thehdw.DCVS.Pins("VDDIO18_GRP1").Voltage.Main - 0.05 And MeasV_ToTestLimit.Pins(p).Value(Site) < thehdw.DCVS.Pins("VDDIO18_GRP1").Voltage.Main + 0.05) Then
'''''''                            MipiResult.Pins(p).Value(Site) = 1
'''''''                        Else
'''''''                            MipiResult.Pins(p).Value(Site) = 0
'''''''                        End If
'''''''
'''''''                    Next Site
'''''''                Next p
'''''''                TheExec.Flow.TestLimit MipiResult, 1, 1, , , , unitVolt, , Tname:=TestNameInput + CStr(TestSeqNum), ForceResults:=tlForceFlow
'''''''
'''''''            End If
'''''''
'''''''             '' 20150202 - Range Check
'''''''            If Range_Check_Enable_Word = True Then
'''''''                Call CheckRangesAndClamps(MeasV_ToTestLimit, "V", RangeCheck_HighLimitVal(gl_UseLimitCheck_Counter), RangeCheck_LowLimitVal(gl_UseLimitCheck_Counter))
'''''''                gl_UseLimitCheck_Counter = gl_UseLimitCheck_Counter + 1
'''''''            End If
'''''''        End If
'''''''    End If
''''
''''    '' 20150814 - Move to Customize HardIP module
''''    'Call CUS_VFI_MeasureVolt(CUS_STR_MainProgram, MeasVoltage(0), TestSeqNum, Pat)
''''
''''End Function
Public Function DiscriminateMeasV_PinType(MeasureV_pin As PinList, ByRef MeasureV_Pin_IO As String, ByRef MeasureV_Pin_UVI80 As String, Optional ByRef PinGroupNum As Integer) As Long

    Dim Pin_SplitArray() As String
    Dim i As Integer
    Dim Flag_FirstTime_IO As Boolean
    Dim Flag_FirstTime_UVI80 As Boolean
    
    Flag_FirstTime_IO = False
    Flag_FirstTime_UVI80 = False
    PinGroupNum = 0
    
    Dim ThisPinType As String
    Dim Pins() As String
    Dim NumberPins As Long
    
    Pin_SplitArray = Split(MeasureV_pin, ",")
    
    For i = 0 To UBound(Pin_SplitArray)
        If Pin_SplitArray(i) <> "" Then
        
            '' 20150616 - Pin group must use the same instrument, can not cross instruments.
            ''                  Only use the first pin of pin group to detect instrument type.
            Call TheExec.DataManager.DecomposePinList(Pin_SplitArray(i), Pins(), NumberPins)
            ThisPinType = GetInstrumentType(Pins(0), 0)
                       
            If ThisPinType = "DC-07" Then
                If Flag_FirstTime_UVI80 = False Then
                    MeasureV_Pin_UVI80 = Pin_SplitArray(i)
                    Flag_FirstTime_UVI80 = True
                Else
                    MeasureV_Pin_UVI80 = MeasureV_Pin_UVI80 & "," & Pin_SplitArray(i)
                End If
            ElseIf ThisPinType = "HSD-U" Or ThisPinType = "HSD-M" Then
                If Flag_FirstTime_IO = False Then
                    MeasureV_Pin_IO = Pin_SplitArray(i)
                    Flag_FirstTime_IO = True
                Else
                    MeasureV_Pin_IO = MeasureV_Pin_IO & "," & Pin_SplitArray(i)
                End If
            End If
        End If
    Next i
    If MeasureV_Pin_UVI80 <> "" Then
        MeasureV_Pin_UVI80 = TrimPinNameDuplicate(MeasureV_Pin_UVI80)
        PinGroupNum = PinGroupNum + 1
    End If
    If MeasureV_Pin_IO <> "" Then
        MeasureV_Pin_IO = TrimPinNameDuplicate(MeasureV_Pin_IO)
        PinGroupNum = PinGroupNum + 1
    End If
End Function

Public Function DACTrim_SetupAndMeasureVolt_UVI80(MeasureV_Pin_UVI80 As String, ByRef MeasureVolt As PinListData, Optional b_HighImpedenceMode As Boolean = True) As Long
  
    If MeasureV_Pin_UVI80 <> "" Then
        With TheHdw.DCVI.Pins(MeasureV_Pin_UVI80)
            .Gate = False
            If b_HighImpedenceMode Then
                '' 20150612 - High impedence mode
                ' Only required if force was previously connected
                .Disconnect tlDCVIConnectDefault
                ' Program the DCVI mapped to MyPin to high impedance mode
                .mode = tlDCVIModeHighImpedance
                ' Connect only the sense to use with high impedance mode
                .Connect tlDCVIConnectHighSense
                .Meter.mode = tlDCVIMeterVoltage  '''Change by Martin for TTR 20151230
            Else
                .mode = tlDCVIModeCurrent
                .Connect tlDCVIConnectDefault
                .Voltage = 6
                .Meter.mode = tlDCVIMeterVoltage  '''Change by Martin for TTR 20151230
            End If
            .VoltageRange.Autorange = True
            .Current = 0
            .CurrentRange.Autorange = True
            TheHdw.Wait (5 * ms)
            .Gate = True
        End With
    End If
    
''    TheHdw.DCVI.Pins(MeasureV_Pin_UVI80).Meter.mode = tlDCVIMeterVoltage
   TheHdw.Wait (5 * ms)
    MeasureVolt = TheHdw.DCVI.Pins(MeasureV_Pin_UVI80).Meter.Read(tlStrobe, 10)
    
     '' 20150703 If use HiZ mode to measure volt that have to gate off HiZ and change to mode current
    If b_HighImpedenceMode Then
        With TheHdw.DCVI.Pins(MeasureV_Pin_UVI80)
            .Gate(tlDCVIGateHiZ) = False
            .Disconnect
            .mode = tlDCVIModeCurrent
            '.Voltage = 6
        End With
    End If
    
End Function
'''''Public Function DACTrim_SetupAndMeasureVolt_PPMU(MeasureV_Pin_PPMU As String, ByRef MeasureVolt As PinListData, Optional InstSpecialSetting As InstrumentSetup = 0) As Long
'''''
'''''    '' 20150918 - Check whether have duplicated pins, change measure mrthod from parallel to serial if pin name duplicate.
''''''    Dim FlagDuplicatePins As Boolean
''''''    FlagDuplicatePins = CheckDuplicateInputPins(MeasureV_Pin_PPMU)
'''''    If InstSpecialSetting = DigitalConnectPPMU Then
'''''
'''''        ' Allow simultaneous PE/PPMU connect and report errors.
'''''        TheHdw.PPMU.AllowPPMUFuncRelayConnection True, False
'''''
'''''        ' Connect the digital channel to the PE.
'''''        TheHdw.Digital.Pins(MeasureV_Pin_PPMU).Connect
'''''
'''''        ' Set up PPMU to measure voltage and connect to the PPMU.
'''''        With TheHdw.PPMU.Pins(MeasureV_Pin_PPMU)
'''''            .Gate = tlOff
'''''            .ForceI 0 * uA, 20 * uA ' Force no current while in 20uA range.
'''''            TheHdw.Wait (100 * us)
'''''            .Connect
'''''        End With
'''''
'''''    Else
'''''        TheHdw.Digital.Pins(MeasureV_Pin_PPMU).Disconnect
'''''
'''''        With TheHdw.PPMU.Pins(MeasureV_Pin_PPMU)
'''''            .Gate = tlOff
'''''            .ForceI 0 * uA, 20 * uA
'''''            .Connect
'''''            TheHdw.Wait (100 * us)
'''''            .Gate = tlOn
'''''        End With
'''''
'''''    End If
'''''    '' 20150918 - Check whether have duplicated pins, change measure mrthod from parallel to serial if pin name duplicate.
'''''''    Dim i As Long
'''''''    Dim InputPins() As String
'''''''    InputPins = Split(MeasureV_Pin_PPMU, ",")
'''''''    If FlagDuplicatePins = True Then
'''''''        For i = 0 To UBound(InputPins)
'''''''            MeasureVolt.AddPin(InputPins(i)).Value = thehdw.PPMU.Pins(InputPins(i)).Read(tlPPMUReadMeasurements, 10)
'''''''        Next i
'''''''    Else
'''''
'''''        MeasureVolt = TheHdw.PPMU.Pins(MeasureV_Pin_PPMU).Read(tlPPMUReadMeasurements, 10)
'''''
'''''''    End If
'''''
'''''    If InstSpecialSetting = DigitalConnectPPMU Then
'''''
'''''        With TheHdw.PPMU.Pins(MeasureV_Pin_PPMU)
'''''            .Disconnect
'''''        End With
'''''
'''''        TheHdw.PPMU.AllowPPMUFuncRelayConnection False, False
'''''
'''''    Else
'''''        With TheHdw.PPMU.Pins(MeasureV_Pin_PPMU)
'''''''            .ForceI 0 * uA, 20 * uA
'''''            .Disconnect
'''''            .Gate = tlOff
'''''        End With
'''''
'''''        TheHdw.Wait (0.001 * ms)
'''''        TheHdw.Digital.Pins(MeasureV_Pin_PPMU).Connect
'''''
'''''    End If
'''''
'''''End Function
Public Function DACTrim_MergePinListData(Meas_INstType_Num As Integer, Measurement() As PinListData, ByRef MergedData As PinListData) As Long
    '' 20150608 - Merge Measurement to the same pin list data if instrument over 1 type.
    Dim index As Integer
    Dim p As Integer
    
    For index = 0 To Meas_INstType_Num
        For p = 0 To Measurement(index).Pins.Count - 1
            MergedData.AddPin (Measurement(index).Pins(p))
            MergedData.Pins(Measurement(index).Pins(p)) = Measurement(index).Pins(p)
        Next p
    Next index

End Function
Public Function DACTrim_MeasureCurrent(MeasureI_pin As PinList, MeasureI_Pin_CurrentRange As String, TestSeqNum As Integer, Optional MeasCurrWaitTime As String = vbNullString, Optional CUS_Str_MainProgram As String, Optional MeasI_ToTestLimit As PinListData)
''    Dim MeasCurr As New PinListData
    Dim p As Long
    Dim PinVal As Double
    Dim CurrRange As Double
    Dim site As Variant
    Dim SampleSize As Long
''    ''========================================================================================
''    '' 20150202 - Range Check
''    Dim RangeCheck_HighLimitVal() As Double
''    Dim RangeCheck_LowLimitVal() As Double
''    Dim TempMeasVal_PerPin(100) As New PinListData
''    If Range_Check_Enable_Word = True Then
''        Call GetFlowSingleUseLimit(RangeCheck_HighLimitVal, RangeCheck_LowLimitVal)
''    End If
''    ''========================================================================================
    
    '' 20150616 - Display current range for each pin
''    Dim b_DisplayCurrentRangePerPin As Boolean
''    b_DisplayCurrentRangePerPin = True
    
    '' 20150615 - Move to outside
''    If DisableClock = True Then FreeRunClk_Disable (PortName)
  
    '' 20150615 - Measure current Opt use
    Dim MeasureI_Pin_UVI80 As String
    Dim MeasureI_Pin_HexVS As String
    Dim MeasureI_Pin_UVS256 As String
    
    Dim MeasI_INstType_Num As Integer
    MeasI_INstType_Num = 0
    
    '' 20150616 - TestCase for different instrument
    Dim DUT_TestConditions() As MeasIConditions
    Call DACTrim_MeasureI_Combine2TestCondition(MeasureI_pin, MeasureI_Pin_CurrentRange, DUT_TestConditions)
    
    Dim MI_TestCond_UVI80() As MeasIConditions
    Dim MI_TestCond_HexVS() As MeasIConditions
    Dim MI_TestCond_UVS256() As MeasIConditions
    
''    MeasureI_Pin = "VDD_CPU , ANALOGMUX_OUT_UVI80, VDD_FIXED_LPDP, VDD12_PLL_LPDP, PCIE_ATB0_UVI80, VDD12_PLL, PCIE_ATB0_UVI80, PCIE_ATB1_UVI80, VDD_GPU"
''    MeasureI_Pin = "VDD_CPU"
    Call DiscriminateMeasI_TestCondition(MeasureI_pin, DUT_TestConditions, MI_TestCond_UVI80, MI_TestCond_HexVS, MI_TestCond_UVS256, MeasI_INstType_Num)
    MeasI_INstType_Num = MeasI_INstType_Num - 1

    ReDim MeasCurrent(MeasI_INstType_Num) As New PinListData
        
    Dim index As Integer
    index = 0
    If IsEmptyArray(MI_TestCond_UVI80) = False Then
        Call DACTrim_SetupAndMeasureCurrent_UVI80(MI_TestCond_UVI80, 10, MeasCurrent(index), MeasCurrWaitTime)
        index = index + 1
    End If

    If IsEmptyArray(MI_TestCond_HexVS) = False Then
        Call DACTrim_SetupAndMeasureCurrent_HexVS(MI_TestCond_HexVS, 10, MeasCurrent(index), MeasCurrWaitTime)
        index = index + 1
    End If

    If IsEmptyArray(MI_TestCond_UVS256) = False Then
        'Call CUS_CurrentSampleSize_Setting_UVS256(CUS_STR_MainProgram, SampleSize)
        Call DACTrim_SetupAndMeasureCurrent_UVS256(MI_TestCond_UVS256, SampleSize, MeasCurrent(index), MeasCurrWaitTime, CUS_Str_MainProgram)
        index = index + 1
    End If
    
    'Dim MeasI_ToTestLimit As New PinListData
    If MeasI_INstType_Num = 0 Then
        If IsEmptyArray(MI_TestCond_UVI80) = False Then
            MeasI_ToTestLimit = MeasCurrent(0)
        ElseIf IsEmptyArray(MI_TestCond_HexVS) = False Then
            MeasI_ToTestLimit = MeasCurrent(0)
        ElseIf IsEmptyArray(MI_TestCond_UVS256) = False Then
            MeasI_ToTestLimit = MeasCurrent(0)
        Else
            TheExec.Datalog.WriteComment ("Warning Hint: " & TheExec.DataManager.instancename & "no data for current measurement.")
        End If
    Else
        Call DACTrim_MergePinListData(MeasI_INstType_Num, MeasCurrent, MeasI_ToTestLimit)
    End If
    
'''    Dim TestNameInput As String
    
'''    TestNameInput = "Curr_meas_" + CStr(TestSeqNum)
    'Call CUS_PLL_SRAM(CUS_STR_MainProgram, MeasI_ToTestLimit)
    'Call CUS_UVD_PWR(CUS_STR_MainProgram, MeasI_ToTestLimit)
    
'''    If Flag_SingleLimit = True Then
'''        TheExec.Flow.TestLimit MeasI_ToTestLimit, LowLimitVal, HighLimitVal, , , , unitAmp, , Tname:=TestNameInput, ForceResults:=tlForceFlow
'''
'''        '' 20150202 - Range Check
'''        If Range_Check_Enable_Word = True Then
'''            Call CheckRangesAndClamps(MeasI_ToTestLimit, "I", RangeCheck_HighLimitVal(gl_UseLimitCheck_Counter), RangeCheck_LowLimitVal(gl_UseLimitCheck_Counter))
'''            gl_UseLimitCheck_Counter = gl_UseLimitCheck_Counter + 1
'''        End If
'''
'''    Else
'''        If Mid(TestLimitByPin_VFI, 3, 1) = "T" Then
'''            For p = 0 To MeasI_ToTestLimit.Pins.Count - 1
'''                    TheExec.Flow.TestLimit resultVal:=MeasI_ToTestLimit.Pins(p), Tname:=TestNameInput, unit:=unitAmp, ForceResults:=tlForceFlow
'''
'''                    '' 20150202 - Range Check
'''                    If Range_Check_Enable_Word = True Then
'''                        TempMeasVal_PerPin(p).AddPin (MeasI_ToTestLimit.Pins(p))
'''                        TempMeasVal_PerPin(p).Pins(MeasI_ToTestLimit.Pins(p)) = MeasI_ToTestLimit.Pins(p)
'''                        Call CheckRangesAndClamps(TempMeasVal_PerPin(p), "I", RangeCheck_HighLimitVal(gl_UseLimitCheck_Counter), RangeCheck_LowLimitVal(gl_UseLimitCheck_Counter))
'''                        gl_UseLimitCheck_Counter = gl_UseLimitCheck_Counter + 1
'''                    End If
'''            Next p
'''        Else
'''            TheExec.Flow.TestLimit MeasI_ToTestLimit, , , , , , unitAmp, , Tname:=TestNameInput, ForceResults:=tlForceFlow
'''             '' 20150202 - Range Check
'''            If Range_Check_Enable_Word = True Then
'''                Call CheckRangesAndClamps(MeasI_ToTestLimit, "I", RangeCheck_HighLimitVal(gl_UseLimitCheck_Counter), RangeCheck_LowLimitVal(gl_UseLimitCheck_Counter))
'''                gl_UseLimitCheck_Counter = gl_UseLimitCheck_Counter + 1
'''            End If
'''        End If
'''    End If
'''
'''        '' 20150624 - Use Enum to re-write UVD HexVS to calculate diff value for current measurement
'''    If SpecialCalcValSetting = DIFF_1ST Then
'''        G_MeasI_DIFF_1ST(TestSeqNum) = MeasI_ToTestLimit
'''    End If
'''
'''    If SpecialCalcValSetting = DIFF_2ND Then
'''        MeasI_ToTestLimit = G_MeasI_DIFF_1ST(TestSeqNum).Math.Subtract(MeasI_ToTestLimit)
'''        If LCase(TheExec.CurrentJob) Like "*char*" Then
'''            TheExec.Flow.TestLimit MeasI_ToTestLimit, 0.0222, 0.048, , , , unitAmp, , ForceResults:=tlForceFlow
'''        Else
'''            TheExec.Flow.TestLimit MeasI_ToTestLimit, 0.0148, 0.0353, , , , unitAmp, , Tname:="Diff_" + TestNameInput, ForceResults:=tlForceFlow
'''        End If
'''    End If
'''
'''    If SpecialCalcValSetting = DIFF_PT12 And TestSeqNum = 0 Then
'''        G_MeasI_DIFF_1ST(TestSeqNum) = MeasI_ToTestLimit
'''    End If
'''
'''    If SpecialCalcValSetting = DIFF_PT12 And TestSeqNum = 1 Then
'''        MeasI_ToTestLimit = MeasI_ToTestLimit.Math.Subtract(G_MeasI_DIFF_1ST(0))
'''
'''        If CurrentJobName_L Like "*char*" Then
'''            For p = 0 To MeasI_ToTestLimit.Pins.Count - 1
'''                TheExec.Flow.TestLimit MeasI_ToTestLimit.Pins(p), , , , , , unitAmp, , , ForceResults:=tlForceFlow
'''            Next p
'''        Else
'''            For p = 0 To MeasI_ToTestLimit.Pins.Count - 1
'''                TheExec.Flow.TestLimit MeasI_ToTestLimit.Pins(p), , , , , , unitAmp, , Tname:="Diff_" + TestNameInput, ForceResults:=tlForceFlow
'''            Next p
'''        End If
'''    End If
'''
'''    '' 20150616 - Display current range value
'''    If b_DisplayCurrentRangePerPin = True Then
'''        For Each Site In TheExec.Sites.Active
'''            For p = 0 To MeasI_ToTestLimit.Pins.Count - 1
'''                PinVal = MeasI_ToTestLimit.Pins(p).Value(Site)
'''                If Left(TheExec.DataManager.channelType(MeasI_ToTestLimit.Pins(p).Name), 4) = "DCVS" Then
'''                    CurrRange = thehdw.DCVS.Pins(MeasI_ToTestLimit.Pins(p).Name).CurrentRange.Value
'''                ElseIf Left(TheExec.DataManager.channelType(MeasI_ToTestLimit.Pins(p).Name), 4) = "DCVI" Then
'''
'''                End If
'''            Next p
'''        Next Site
'''    End If
  
End Function
Public Function DACTrim_MeasureI_Combine2TestCondition(MeasureI_pin As PinList, ByVal MeasureI_Pin_CurrentRange As String, ByRef DUT_TestConditions() As MeasIConditions)
    '' 20150621 - Mapping current range to the expected pin
    Dim PinNumber As Integer
    Dim MeasureI_Pin_Array() As String
    Dim MeasureI_Range_Array() As String
    Dim DefaultCurrentRange As Double
    DefaultCurrentRange = 0.02
    Dim Diff_Num As Integer
    Dim i As Integer
    Dim Pins_CPFT() As String
    Dim Pins_CPFT_CurrentRange() As String
    MeasureI_Pin_Array = Split(MeasureI_pin, ",")
    MeasureI_Range_Array = Split(MeasureI_Pin_CurrentRange, ",")
    
    If UBound(MeasureI_Pin_Array) > UBound(MeasureI_Range_Array) Then
        TheExec.Datalog.WriteComment ("Warning Hint: Index number doesn't match between pin name and measure current range, mapping rule is one by one")
        Diff_Num = UBound(MeasureI_Pin_Array) - UBound(MeasureI_Range_Array)
        '' 20150629 - Setup current range to default value if input argument is empty.
        For i = 0 To Diff_Num - 1
            If MeasureI_Pin_CurrentRange <> "" Then
                If UBound(MeasureI_Range_Array) = 0 Then
                    If i = 0 Then
                        DefaultCurrentRange = MeasureI_Pin_CurrentRange
                        MeasureI_Pin_CurrentRange = MeasureI_Pin_CurrentRange & "," & DefaultCurrentRange
                    Else
                        MeasureI_Pin_CurrentRange = MeasureI_Pin_CurrentRange & "," & DefaultCurrentRange
                    End If
                Else
                    MeasureI_Pin_CurrentRange = MeasureI_Pin_CurrentRange & "," & DefaultCurrentRange
                End If
            Else
                If i = 0 Then
                    MeasureI_Pin_CurrentRange = DefaultCurrentRange
                Else
                    MeasureI_Pin_CurrentRange = MeasureI_Pin_CurrentRange & "," & DefaultCurrentRange
                End If
            End If
        Next i
        ReDim MeasureI_Range_Array(UBound(MeasureI_Pin_Array)) As String
        MeasureI_Range_Array = Split(MeasureI_Pin_CurrentRange, ",")
    End If
    
    PinNumber = UBound(MeasureI_Pin_Array)
    
    ReDim TempDUT_TestConditions(PinNumber) As MeasIConditions
    
    Dim b_PinExistFlag As Boolean
    Dim index As Long
    index = 0
    For i = 0 To PinNumber
        b_PinExistFlag = True
        '' 20151117 - Check CP/FT pins if FT pogo not ball out >> Pin name
        If InStr(MeasureI_Pin_Array(i), ":") <> 0 Then
            Pins_CPFT = Split(MeasureI_Pin_Array(i), ":")
            
            If UCase(Pins_CPFT(0)) = "CP" And InStr(UCase(TheExec.CurrentChanMap), "FT") <> 0 Or UCase(Pins_CPFT(0)) = "FT" And InStr(UCase(TheExec.CurrentChanMap), "CP") <> 0 Then
                b_PinExistFlag = False
            Else
            
                If UCase(Pins_CPFT(0)) = "CP" And InStr(UCase(TheExec.CurrentChanMap), "CP") <> 0 Then
                    MeasureI_Pin_Array(i) = Pins_CPFT(1)
                ElseIf UCase(Pins_CPFT(0)) = "FT" And InStr(UCase(TheExec.CurrentChanMap), "FT") <> 0 Then
                    MeasureI_Pin_Array(i) = Pins_CPFT(1)
                Else
                    MeasureI_Pin_Array(i) = Pins_CPFT(0)
                End If
                b_PinExistFlag = True
            End If
        End If
        
        '' 20151117 - Check CP/FT pins if FT pogo not ball out >> Current Range
        If InStr(MeasureI_Range_Array(i), ":") <> 0 Then
           
            Pins_CPFT_CurrentRange = Split(MeasureI_Range_Array(i), ":")
           
            If UCase(Pins_CPFT(0)) = "CP" And InStr(UCase(TheExec.CurrentChanMap), "FT") <> 0 Or UCase(Pins_CPFT(0)) = "FT" And InStr(UCase(TheExec.CurrentChanMap), "CP") <> 0 Then
            Else
            
                If UCase(Pins_CPFT_CurrentRange(0)) = "CP" And InStr(UCase(TheExec.CurrentChanMap), "CP") <> 0 Then
                    '' 20160106 - Give default current range value if current range input argument is CP: or FT:
                    If Pins_CPFT_CurrentRange(1) = "" Then
                        Pins_CPFT_CurrentRange(1) = DefaultCurrentRange
                    End If
                    MeasureI_Range_Array(i) = Pins_CPFT_CurrentRange(1)
                ElseIf UCase(Pins_CPFT_CurrentRange(0)) = "FT" And InStr(UCase(TheExec.CurrentChanMap), "FT") <> 0 Then
                    If Pins_CPFT_CurrentRange(1) = "" Then
                        Pins_CPFT_CurrentRange(1) = DefaultCurrentRange
                    End If
                    MeasureI_Range_Array(i) = Pins_CPFT_CurrentRange(1)
                Else
                    MeasureI_Range_Array(i) = Pins_CPFT_CurrentRange(0)
                End If
            End If
        ElseIf MeasureI_Range_Array(i) = "" Then    ' chihome 20160130, use default current range from tester, not 50mA for power pin
            If TheExec.DataManager.PinType(MeasureI_Pin_Array(i)) = "Power" Then
                DefaultCurrentRange = TheHdw.DCVS.Pins(MeasureI_Pin_Array(i)).CurrentRange
            End If
            MeasureI_Range_Array(i) = DefaultCurrentRange
        End If
        If b_PinExistFlag = True Then
            TempDUT_TestConditions(index).PinName = MeasureI_Pin_Array(i)
            TempDUT_TestConditions(index).CurrentRange = MeasureI_Range_Array(i)
            index = index + 1
        End If
    Next i
    
    '' 20151117 - reorder MeasIConditions to match real ball out
    ReDim DUT_TestConditions(index - 1) As MeasIConditions
    For i = 0 To UBound(DUT_TestConditions)
        DUT_TestConditions(i).PinName = TempDUT_TestConditions(i).PinName
        DUT_TestConditions(i).CurrentRange = TempDUT_TestConditions(i).CurrentRange
    Next i
End Function
Public Function DiscriminateMeasI_TestCondition(MeasureI_pin As PinList, DUT_TestConditions() As MeasIConditions, ByRef TestCond_MeasureI_UVI80() As MeasIConditions, ByRef TestCond_MeasureI_HexVS() As MeasIConditions, ByRef TestCond_MeasureI_UVS256() As MeasIConditions, Optional ByRef PinGroupNum As Integer) As Long

    Dim Pin_SplitArray() As String
    Dim i As Integer
    Dim j As Integer
    Dim MeasureI_Pin_UVI80 As String
    Dim MeasureI_Pin_HexVS As String
    Dim MeasureI_Pin_UVS256 As String
    
    Dim Flag_FirstTime_UVI80 As Boolean
    Dim Flag_FirstTime_HexVS As Boolean
    Dim Flag_FirstTime_UVS256 As Boolean
    Flag_FirstTime_UVI80 = False
    Flag_FirstTime_HexVS = False
    Flag_FirstTime_UVS256 = False
    
    PinGroupNum = 0
    Dim ThisPinType As String
    Dim Pins() As String
    Dim NumberPins As Long
    Dim Pins_CPFT() As String
    
    Pin_SplitArray = Split(MeasureI_pin, ",")
    
    Dim b_PinExistFlag As Boolean
    For i = 0 To UBound(Pin_SplitArray)
        If Pin_SplitArray(i) <> "" Then
            b_PinExistFlag = True
            '' 20151117 - Check CP/FT pins if FT pogo not ball out >> Pin name
            If InStr(Pin_SplitArray(i), ":") <> 0 Then
                Pins_CPFT = Split(Pin_SplitArray(i), ":")
                If UCase(Pins_CPFT(0)) = "CP" And InStr(UCase(TheExec.CurrentChanMap), "FT") <> 0 Or UCase(Pins_CPFT(0)) = "FT" And InStr(UCase(TheExec.CurrentChanMap), "CP") <> 0 Then
                    b_PinExistFlag = False
                Else
                    b_PinExistFlag = True
                    If UCase(Pins_CPFT(0)) = "CP" And InStr(UCase(TheExec.CurrentChanMap), "CP") <> 0 Then
                        Pin_SplitArray(i) = Pins_CPFT(1)
                    ElseIf UCase(Pins_CPFT(0)) = "FT" And InStr(UCase(TheExec.CurrentChanMap), "FT") <> 0 Then
                        Pin_SplitArray(i) = Pins_CPFT(1)
                    Else
                        Pin_SplitArray(i) = Pins_CPFT(0)
                    End If
                End If
            End If
            If b_PinExistFlag Then

                Call TheExec.DataManager.DecomposePinList(Pin_SplitArray(i), Pins(), NumberPins)
                Call DACTrim_NC_Pin(Pins, NumberPins)
                ThisPinType = GetInstrumentType(Pins(0), 0)
                
                If ThisPinType = "DC-07" Then
                    If Flag_FirstTime_UVI80 = False Then
                        MeasureI_Pin_UVI80 = Pin_SplitArray(i)
                        Flag_FirstTime_UVI80 = True
                    Else
                        MeasureI_Pin_UVI80 = MeasureI_Pin_UVI80 & "," & Pin_SplitArray(i)
                    End If
                ElseIf ThisPinType = "HexVS" Then
                    If Flag_FirstTime_HexVS = False Then
                        MeasureI_Pin_HexVS = Pin_SplitArray(i)
                        Flag_FirstTime_HexVS = True
                    Else
                        MeasureI_Pin_HexVS = MeasureI_Pin_HexVS & "," & Pin_SplitArray(i)
                    End If
                ElseIf ThisPinType = "VHDVS" Then
                    If Flag_FirstTime_UVS256 = False Then
                        MeasureI_Pin_UVS256 = Pin_SplitArray(i)
                        Flag_FirstTime_UVS256 = True
                    Else
                        MeasureI_Pin_UVS256 = MeasureI_Pin_UVS256 & "," & Pin_SplitArray(i)
                    End If
                End If
            End If
        End If
    Next i
    
    If MeasureI_Pin_UVI80 <> "" Then
        MeasureI_Pin_UVI80 = TrimPinNameDuplicate(MeasureI_Pin_UVI80)
        PinGroupNum = PinGroupNum + 1
    End If
    If MeasureI_Pin_HexVS <> "" Then
        MeasureI_Pin_HexVS = TrimPinNameDuplicate(MeasureI_Pin_HexVS)
        PinGroupNum = PinGroupNum + 1
    End If
    
    If MeasureI_Pin_UVS256 <> "" Then
        MeasureI_Pin_UVS256 = TrimPinNameDuplicate(MeasureI_Pin_UVS256)
        PinGroupNum = PinGroupNum + 1
    End If
    
    Dim MeasureI_PinArr_UVI80() As String
    Dim MeasureI_PinArr_HexVS() As String
    Dim MeasureI_PinArr_UVS256() As String
    
    If MeasureI_Pin_UVI80 <> "" Then
        MeasureI_PinArr_UVI80 = Split(MeasureI_Pin_UVI80, ",")
        ReDim TestCond_MeasureI_UVI80(UBound(MeasureI_PinArr_UVI80)) As MeasIConditions
        
        For i = 0 To UBound(MeasureI_PinArr_UVI80)
            For j = 0 To UBound(DUT_TestConditions)
                If MeasureI_PinArr_UVI80(i) = DUT_TestConditions(j).PinName Then
                    TestCond_MeasureI_UVI80(i).PinName = DUT_TestConditions(j).PinName
                    TestCond_MeasureI_UVI80(i).CurrentRange = DUT_TestConditions(j).CurrentRange
                    Exit For
                End If
            Next j
        Next i
    End If
    
    If MeasureI_Pin_HexVS <> "" Then
        MeasureI_PinArr_HexVS = Split(MeasureI_Pin_HexVS, ",")
        ReDim TestCond_MeasureI_HexVS(UBound(MeasureI_PinArr_HexVS)) As MeasIConditions
        
        For i = 0 To UBound(MeasureI_PinArr_HexVS)
            For j = 0 To UBound(DUT_TestConditions)
                If MeasureI_PinArr_HexVS(i) = DUT_TestConditions(j).PinName Then
                    TestCond_MeasureI_HexVS(i).PinName = DUT_TestConditions(j).PinName
                    TestCond_MeasureI_HexVS(i).CurrentRange = DUT_TestConditions(j).CurrentRange
                    Exit For
                End If
            Next j
        Next i
    End If
    
    If MeasureI_Pin_UVS256 <> "" Then
        MeasureI_PinArr_UVS256 = Split(MeasureI_Pin_UVS256, ",")
        ReDim TestCond_MeasureI_UVS256(UBound(MeasureI_PinArr_UVS256)) As MeasIConditions
        
        For i = 0 To UBound(MeasureI_PinArr_UVS256)
            For j = 0 To UBound(DUT_TestConditions)
                If MeasureI_PinArr_UVS256(i) = DUT_TestConditions(j).PinName Then
                    TestCond_MeasureI_UVS256(i).PinName = DUT_TestConditions(j).PinName
                    TestCond_MeasureI_UVS256(i).CurrentRange = DUT_TestConditions(j).CurrentRange
                    Exit For
                End If
            Next j
        Next i
    End If
    
    '' 20150622 - Check Ubound of MeasIConditions for each instument type
    Dim PinNum_UVI80 As Integer
    Dim PinNum_HexVS As Integer
    Dim PinNum_UVS256 As Integer
    
    If IsEmptyArray(TestCond_MeasureI_UVI80) = False Then
        PinNum_UVI80 = UBound(TestCond_MeasureI_UVI80) + 1
    End If
    
    If IsEmptyArray(TestCond_MeasureI_HexVS) = False Then
        PinNum_HexVS = UBound(TestCond_MeasureI_HexVS) + 1
    End If
    
    If IsEmptyArray(TestCond_MeasureI_UVS256) = False Then
        PinNum_UVS256 = UBound(TestCond_MeasureI_UVS256) + 1
    End If
    
    If UBound(DUT_TestConditions) + 1 <> (PinNum_UVI80 + PinNum_HexVS + PinNum_UVS256) Then

        TheExec.Datalog.WriteComment ("Warning Hint: Measure I Pin duplicated, please check it")
    End If
End Function
Public Function DACTrim_NC_Pin(ByRef original_ary() As String, ByRef original_pin_cnt As Long)
'get active pins array
    Dim i As Long, j As Long
    Dim p As Variant
    Dim TempArray() As String
    Dim TempPinCnt As Long
    Dim NullArray() As String
    Dim TempString As String
    Dim PowerSequence As Double
    
    On Error GoTo errHandler
    
    If original_pin_cnt <> 0 Then
        i = 0   'init
        For Each p In original_ary
            If TheExec.DataManager.ChannelType(p) <> "N/C" Then i = i + 1
        Next p
        
        'redim
        ReDim TempArray(i - 1)
        
        j = 0   'init
        For Each p In original_ary
            If TheExec.DataManager.ChannelType(p) <> "N/C" Then
                TempArray(j) = p
                j = j + 1
            Else
                j = j
            End If
        Next p
        
        'return array and pin count
        original_ary = TempArray
        original_pin_cnt = j
    End If
    
    Exit Function
errHandler:
    ErrorDescription ("DACTrim_NC_Pin")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function IsEmptyArray(InputArr() As MeasIConditions) As Boolean
 'Public Function IsEmptyArray(InputArr() As DACTrimObj) As Boolean
    On Error GoTo ChkFalse
    
    If UBound(InputArr) >= 0 Then
        IsEmptyArray = False
        Exit Function
    End If
    
ChkFalse:
    IsEmptyArray = True
End Function

Public Function TrimPinNameDuplicate(PinName As String) As String
    Dim i As Integer
    Dim j As Integer
    Dim Pin_SplitArrayNum As Integer
    Dim Pin_SplitArray() As String
    Dim PinStringTemp As String
    Dim PinStringAfterCheck As String

    
    Pin_SplitArray = Split(PinName, ",")
    Pin_SplitArrayNum = UBound(Pin_SplitArray)
    
    ReDim b_FlagDuplicate(Pin_SplitArrayNum) As Boolean
    
    Dim b_FlagFirstTime As Boolean
    b_FlagFirstTime = False
    
    For i = 0 To Pin_SplitArrayNum
        PinStringTemp = Pin_SplitArray(i)
        For j = i + 1 To Pin_SplitArrayNum
            If Pin_SplitArray(j) = PinStringTemp Then
                b_FlagDuplicate(j) = True
            End If
        Next j
    Next i
    
    For i = 0 To Pin_SplitArrayNum
        If b_FlagDuplicate(i) = False Then
            If b_FlagFirstTime = False Then
                PinStringAfterCheck = Pin_SplitArray(i)
                b_FlagFirstTime = True
            Else
                PinStringAfterCheck = PinStringAfterCheck & "," & Pin_SplitArray(i)
            End If
            
        End If
    Next i
    TrimPinNameDuplicate = PinStringAfterCheck
End Function
Public Function DACTrim_SetupAndMeasureCurrent_UVI80(MI_TestCond_UVI80() As MeasIConditions, SampleSize As Long, ByRef measureCurrent As PinListData, Optional CustomizeWaitTime As String = vbNullString)
    '' 20150623 - Suggest use for single pin it can mapping expected current range,
    ''                  if use for pin group it will refer the same current range to pin group by your specified.
    Dim i As Integer
    Dim WaitTime As Double
    Dim MaxWaitTime As Double

    WaitTime = 1 * ms
    MaxWaitTime = 0
    Dim Pins() As String
    Dim NumberPins As Long
    Dim NumTypes As Long
    Dim PowerType() As String
    Dim Factor As Long
    Dim Pins_MeasureI_Together As String
    '' 20150616 - Findout the expected range and wait time also think for merge Mode
    For i = 0 To UBound(MI_TestCond_UVI80)
        MaxWaitTime = WaitTime
        
        Call TheExec.DataManager.DecomposePinList(MI_TestCond_UVI80(i).PinName, Pins(), NumberPins)
        Call TheExec.DataManager.GetChannelTypes(Pins(0), NumTypes, PowerType())
        
        Select Case PowerType(0)
            Case "DCVI"
                Factor = 1
                
            Case "DCVIMerged"
                Factor = 2
                
            Case Else
        End Select
        
        If MI_TestCond_UVI80(i).CurrentRange > 2 * Factor Then
            MI_TestCond_UVI80(i).CurrentRange = 2 * Factor
            WaitTime = 1.6 * ms
        ElseIf MI_TestCond_UVI80(i).CurrentRange > 1 * Factor Then
            MI_TestCond_UVI80(i).CurrentRange = 2 * Factor
            WaitTime = 1.6 * ms
        ElseIf MI_TestCond_UVI80(i).CurrentRange > 0.2 * Factor Then
            MI_TestCond_UVI80(i).CurrentRange = 1 * Factor
            WaitTime = 1.6 * ms
        ElseIf MI_TestCond_UVI80(i).CurrentRange > 0.02 * Factor Then
            MI_TestCond_UVI80(i).CurrentRange = 0.2 * Factor
            WaitTime = 260 * us
        ElseIf MI_TestCond_UVI80(i).CurrentRange > 0.002 * Factor Then
            MI_TestCond_UVI80(i).CurrentRange = 0.02 * Factor
            WaitTime = 1.5 * ms
        ElseIf MI_TestCond_UVI80(i).CurrentRange > 0.0002 * Factor Then
            MI_TestCond_UVI80(i).CurrentRange = 0.002 * Factor
            WaitTime = 11 * ms
        ElseIf MI_TestCond_UVI80(i).CurrentRange > 0.00002 * Factor Then
            MI_TestCond_UVI80(i).CurrentRange = 0.0002 * Factor
            WaitTime = 1.4 * ms
        Else
            MI_TestCond_UVI80(i).CurrentRange = 0.00002 * Factor
            WaitTime = 6 * ms
        End If

        If WaitTime > MaxWaitTime Then
            MaxWaitTime = WaitTime
        End If

        With TheHdw.DCVI.Pins(MI_TestCond_UVI80(i).PinName)
            .Gate = False
            .mode = tlDCVIModeVoltage
            .Voltage = 0
            .VoltageRange.value = 7
            .Current = MI_TestCond_UVI80(i).CurrentRange
            .CurrentRange.value = MI_TestCond_UVI80(i).CurrentRange
            .Connect tlDCVIConnectDefault
            .Gate = True
        End With
    
        With TheHdw.DCVI.Pins(MI_TestCond_UVI80(i).PinName)
            .Meter.mode = tlDCVIMeterCurrent
            .Meter.CurrentRange.value = MI_TestCond_UVI80(i).CurrentRange
        End With

        TheExec.Datalog.WriteComment (TheExec.DataManager.instancename & " =====> Curr_meas Meter I range setting, " & MI_TestCond_UVI80(i).PinName & " =" & TheHdw.DCVI.Pins(MI_TestCond_UVI80(i).PinName).Meter.CurrentRange.value)

        If i = 0 Then
            Pins_MeasureI_Together = MI_TestCond_UVI80(i).PinName
        Else
            Pins_MeasureI_Together = Pins_MeasureI_Together & "," & MI_TestCond_UVI80(i).PinName
        End If
    Next i
    '' 20150623 - Convert customize wait time type string to double if MeasCurrWaitTime specified
    If CustomizeWaitTime <> "" Then
        MaxWaitTime = CDbl(CustomizeWaitTime)
    End If
    TheHdw.Wait (MaxWaitTime)
    
    '' 20150615 - Current measurement
    measureCurrent = TheHdw.DCVI.Pins(Pins_MeasureI_Together).Meter.Read(tlStrobe, SampleSize)

End Function
Public Function DACTrim_SetupAndMeasureCurrent_HexVS(MI_TestCond_HexVS() As MeasIConditions, SampleSize As Long, ByRef measureCurrent As PinListData, Optional CustomizeWaitTime As String = vbNullString, Optional CUS_Str_MainProgram As String)

    '' 20150623 - Suggest use for single pin it can mapping expected current range,
    ''                  if use for pin group it will refer the same current range to pin group by your specified.
    Dim i As Integer
    Dim WaitTime As Double
    Dim MaxWaitTime As Double

    WaitTime = 1 * ms
    MaxWaitTime = 0

    '' 20150922 - Create store test condition function for HexVS
    Dim StoreSourceFoldLimit() As Double
    Dim StoreSinkFoldLimit() As Double
    Dim PinsMaxNum As Long
    PinsMaxNum = UBound(MI_TestCond_HexVS)
    ReDim StoreSourceFoldLimit(PinsMaxNum) As Double
    ReDim StoreSinkFoldLimit(PinsMaxNum) As Double
    ReDim StoreFilterSetting(PinsMaxNum) As Double

    Dim Pins_MeasureI_Together As String
    
    '' 20150616 - Findout the expected range and wait time also think for merge Mode
    For i = 0 To UBound(MI_TestCond_HexVS)
        MaxWaitTime = WaitTime
        ''20150922 - Store test condition
         Call DACTrim_HexVS__MI_StoreCondition(MI_TestCond_HexVS(i).PinName, StoreSourceFoldLimit(i), StoreSinkFoldLimit(i))
        
        If CDbl(MI_TestCond_HexVS(i).CurrentRange) > 60 Then
            MI_TestCond_HexVS(i).CurrentRange = 90
            WaitTime = 100 * us
        ElseIf CDbl(MI_TestCond_HexVS(i).CurrentRange) > 30 Then
            MI_TestCond_HexVS(i).CurrentRange = 60
            WaitTime = 100 * us
        ElseIf CDbl(MI_TestCond_HexVS(i).CurrentRange) > 15 Then
            MI_TestCond_HexVS(i).CurrentRange = 30
            WaitTime = 100 * us
        ElseIf CDbl(MI_TestCond_HexVS(i).CurrentRange) > 1 Then
            MI_TestCond_HexVS(i).CurrentRange = 15
            WaitTime = 100 * us
        ElseIf CDbl(MI_TestCond_HexVS(i).CurrentRange) > 0.1 Then
            MI_TestCond_HexVS(i).CurrentRange = 1
            WaitTime = 1 * ms
        ElseIf CDbl(MI_TestCond_HexVS(i).CurrentRange) > 0.01 Then
            MI_TestCond_HexVS(i).CurrentRange = 0.1
            WaitTime = 10 * ms
        Else
            MI_TestCond_HexVS(i).CurrentRange = 0.01
            WaitTime = 100 * ms
        End If
        
        If WaitTime > MaxWaitTime Then
            MaxWaitTime = WaitTime
        End If

        With TheHdw.DCVS.Pins(MI_TestCond_HexVS(i).PinName)
            .Meter.mode = tlDCVSMeterCurrent
            .SetCurrentRanges CDbl(MI_TestCond_HexVS(i).CurrentRange), CDbl(MI_TestCond_HexVS(i).CurrentRange)
            .Gate = True
        End With
        TheExec.Datalog.WriteComment (TheExec.DataManager.instancename & " =====> Curr_meas Meter I range setting, " & MI_TestCond_HexVS(i).PinName & " =" & TheHdw.DCVS.Pins(MI_TestCond_HexVS(i).PinName).Meter.CurrentRange.value)

        If i = 0 Then
            Pins_MeasureI_Together = MI_TestCond_HexVS(i).PinName
        Else
            Pins_MeasureI_Together = Pins_MeasureI_Together & "," & MI_TestCond_HexVS(i).PinName
        End If

    Next i
    
    '' 20150623 - Convert customize wait time type string to double if MeasCurrWaitTime specified
    If CustomizeWaitTime <> "" Then
        MaxWaitTime = CDbl(CustomizeWaitTime)
    End If
    TheHdw.Wait (MaxWaitTime)

    '' 20150615 - Current measurement
    measureCurrent = TheHdw.DCVS.Pins(Pins_MeasureI_Together).Meter.Read(tlStrobe, SampleSize, 10000, tlDCVSMeterReadingFormatAverage)


    '' 20150922 - Restore source/sink fold limit
    For i = 0 To UBound(MI_TestCond_HexVS)
        Call DACTrim_HexVS__MI_RestoreCondition(MI_TestCond_HexVS(i).PinName, StoreSourceFoldLimit(i), StoreSinkFoldLimit(i))
    Next i
End Function
Public Function DACTrim_HexVS__MI_StoreCondition(MI_Pin As String, ByRef SourceFlodLimit As Double, ByRef SinkFoldLimit As Double) As Long
    
''    Dim NumTypes As Long
''    Dim PowerType() As String
    
    SourceFlodLimit = TheHdw.DCVS.Pins(MI_Pin).CurrentLimit.Source.FoldLimit.level.value
    SinkFoldLimit = TheHdw.DCVS.Pins(MI_Pin).CurrentLimit.Sink.FoldLimit.level.value
''    FilterValue = TheHdw.DCVS.Pins(MI_Pin).Meter.Filter.Value
    
''    Dim Pins() As String
''    Dim pin_cnt As Long
''
''    TheExec.DataManager.DecomposePinList MI_Pin, Pins(), pin_cnt
''
''    If pin_cnt > 1 Then
''        TheExec.Datalog.WriteComment ("Error for HardIP UVS256 current measurement: You can not use the pin group for test pin argument to do current measurement, please use single pin for test pin")
''    End If
''
''    Call TheExec.DataManager.GetChannelTypes(Pins(0), NumTypes, PowerType())
''
''    If UCase(PowerType(0)) Like UCase("*Merged*") Then
''    Else
''        TheHdw.DCVS.Pins(MI_Pin).CurrentRange.Value = 0.02
''        TheHdw.DCVS.Pins(MI_Pin).CurrentLimit.Source.FoldLimit.Level.Value = 0.02
''    End If
End Function
Public Function DACTrim_HexVS__MI_RestoreCondition(MI_Pin As String, ByRef SourceFlodLimit As Double, ByRef SinkFoldLimit As Double)

End Function
Public Function DACTrim_SetupAndMeasureCurrent_UVS256(MI_TestCond_UVS256() As MeasIConditions, SampleSize As Long, ByRef measureCurrent As PinListData, Optional CustomizeWaitTime As String = vbNullString, Optional CUS_Str_MainProgram As String)

    '' 20150623 - Suggest use for single pin it can mapping expected current range,
    ''                  if use for pin group it will refer the same current range to pin group by your specified.
    Dim i As Integer
    Dim WaitTime As Double
    Dim MaxWaitTime As Double

    WaitTime = 1 * ms
    MaxWaitTime = 0

    Dim StoreSourceFoldLimit() As Double
    Dim StoreSinkFoldLimit() As Double
    Dim StoreFilterSetting() As Double
    Dim StoreSrcCurrentRange() As Double
    
    Dim PinsMaxNum As Long
    PinsMaxNum = UBound(MI_TestCond_UVS256)
    
    ReDim StoreSourceFoldLimit(PinsMaxNum) As Double
    ReDim StoreSinkFoldLimit(PinsMaxNum) As Double
    ReDim StoreFilterSetting(PinsMaxNum) As Double
    ReDim StoreSrcCurrentRange(PinsMaxNum) As Double
    
    Dim Pins_MeasureI_Together As String
    
'    For i = 0 To UBound(MI_TestCond_UVS256)
'        '' 20150701 - Store source/sink fold limit and filter setting
'        Call HardIP_UVS256__MI_StoreCondition(MI_TestCond_UVS256(i).PinName, StoreSourceFoldLimit(i), StoreSinkFoldLimit(i), StoreFilterSetting(i))
'    Next i
    '' 20150616 - Findout the expected range and wait time also think for merge Mode.
    For i = 0 To UBound(MI_TestCond_UVS256)
        '' 20150701 - Store source/sink fold limit and filter setting
        Call DACTrim_UVS256__MI_StoreCondition(MI_TestCond_UVS256(i).PinName, StoreSourceFoldLimit(i), StoreSinkFoldLimit(i), StoreFilterSetting(i), StoreSrcCurrentRange(i))
        
        MaxWaitTime = WaitTime

        If CDbl(MI_TestCond_UVS256(i).CurrentRange) > 2.8 Then
            MI_TestCond_UVS256(i).CurrentRange = 5.6
            WaitTime = 30 * us
        ElseIf CDbl(MI_TestCond_UVS256(i).CurrentRange) > 1.4 Then
            MI_TestCond_UVS256(i).CurrentRange = 2.8
            WaitTime = 45 * us
        ElseIf CDbl(MI_TestCond_UVS256(i).CurrentRange) > 0.8 Then
            MI_TestCond_UVS256(i).CurrentRange = 1.4
            WaitTime = 50 * us
        ElseIf CDbl(MI_TestCond_UVS256(i).CurrentRange) > 0.7 Then
            MI_TestCond_UVS256(i).CurrentRange = 0.8
            WaitTime = 100 * us
        ElseIf CDbl(MI_TestCond_UVS256(i).CurrentRange) > 0.4 Then
            MI_TestCond_UVS256(i).CurrentRange = 0.7
            WaitTime = 100 * us
        ElseIf CDbl(MI_TestCond_UVS256(i).CurrentRange) > 0.2 Then
            MI_TestCond_UVS256(i).CurrentRange = 0.4
            WaitTime = 90 * us
        ElseIf CDbl(MI_TestCond_UVS256(i).CurrentRange) > 0.04 Then
            MI_TestCond_UVS256(i).CurrentRange = 0.2
            WaitTime = 210 * us
        ElseIf CDbl(MI_TestCond_UVS256(i).CurrentRange) > 0.02 Then
            MI_TestCond_UVS256(i).CurrentRange = 0.04
            WaitTime = 260 * us
        ElseIf CDbl(MI_TestCond_UVS256(i).CurrentRange) > 0.002 Then
            MI_TestCond_UVS256(i).CurrentRange = 0.02
            WaitTime = 540 * us
        ElseIf CDbl(MI_TestCond_UVS256(i).CurrentRange) > 0.0002 Then
            MI_TestCond_UVS256(i).CurrentRange = 0.002
            WaitTime = 3.5 * ms
        ElseIf CDbl(MI_TestCond_UVS256(i).CurrentRange) > 0.00002 Then
            MI_TestCond_UVS256(i).CurrentRange = 0.0002
            WaitTime = 0.21 * ms
        ElseIf CDbl(MI_TestCond_UVS256(i).CurrentRange) > 0.000004 Then
            MI_TestCond_UVS256(i).CurrentRange = 0.00002
            WaitTime = 4 * ms
        Else
            MI_TestCond_UVS256(i).CurrentRange = 0.000004
            WaitTime = 4 * ms
        End If
        If WaitTime > MaxWaitTime Then
            MaxWaitTime = WaitTime
        End If

        With TheHdw.DCVS.Pins(MI_TestCond_UVS256(i).PinName)
            .Meter.mode = tlDCVSMeterCurrent
            .SetCurrentRanges CDbl(MI_TestCond_UVS256(i).CurrentRange), CDbl(MI_TestCond_UVS256(i).CurrentRange)
            'New Update 20170803, don't update Foldlimit to current range
'            .CurrentLimit.Source.FoldLimit.Level.Value = CDbl(MI_TestCond_UVS256(i).CurrentRange)
            .Gate = True
        End With
        TheExec.Datalog.WriteComment (TheExec.DataManager.instancename & " =====> Curr_meas Meter I range setting, " & MI_TestCond_UVS256(i).PinName & " =" & TheHdw.DCVS.Pins(MI_TestCond_UVS256(i).PinName).Meter.CurrentRange.value)
        TheHdw.DCVS.Pins(MI_TestCond_UVS256(i).PinName).Meter.Filter.value = TheHdw.DCVS.Pins(MI_TestCond_UVS256(i).PinName).Meter.Filter.max / SampleSize
        If i = 0 Then
            Pins_MeasureI_Together = MI_TestCond_UVS256(i).PinName
        Else
            Pins_MeasureI_Together = Pins_MeasureI_Together & "," & MI_TestCond_UVS256(i).PinName
        End If
    Next i
    '' 20150623 - Convert customize wait time type string to double if MeasCurrWaitTime specified
    If CustomizeWaitTime <> "" Then
        MaxWaitTime = CDbl(CustomizeWaitTime)
    End If
    'Call CUS_NAND_IDS_1
    TheHdw.Wait (MaxWaitTime)
    
    TheExec.Datalog.WriteComment "Wait time before measure current = " & MaxWaitTime
    
    measureCurrent = TheHdw.DCVS.Pins(Pins_MeasureI_Together).Meter.Read(tlStrobe, SampleSize)  'UVS only allow one sample
    
    '''20151118 auto decrease thecurrent range
    Dim pin As Variant
    Dim CurrRngeNew As Double
    Dim NegRSTFlag As Boolean
    NegRSTFlag = False
    Dim NewRange As Double
    Dim SumTime As Integer
    Dim site As Variant
    If CurrentJobName_L Like "*char*" Then
'================================
    '   If (TheExec.DataManager.InstanceName Like "*AD00_IDS_JTG_IMX_ALLFV_SI_IDDQ*") = True Or (TheExec.DataManager.InstanceName Like "*RX01_MEA_JTG_IMX_ALLFV_SI_PWR_Fx*") = True Then
          For i = 0 To 4
              For Each pin In measureCurrent.Pins
                    For Each site In TheExec.sites
                        If measureCurrent.Pins(pin).value(site) < 0 Then
                             NewRange = TheHdw.DCVS.Pins(pin).Meter.CurrentRange * 0.09
                             TheHdw.DCVS.Pins(pin).SetCurrentRanges NewRange, NewRange
                             'New Update 20170803, don't update Foldlimit to current range
'                             TheHdw.DCVS.Pins(Pin).CurrentLimit.Source.FoldLimit.Level.Value = TheHdw.DCVS.Pins(Pin).CurrentLimit.Source.FoldLimit.Level.Value
                             TheExec.Datalog.WriteComment "Site " & site & ", Run : " & i + 1 & ", Previous Result : " & measureCurrent.Pins(pin).value(site) & "===============================> Set pin " & pin & ", Range = " & TheHdw.DCVS.Pins(pin).Meter.CurrentRange
                             NegRSTFlag = True
                        End If
                     Next site
                  '  MeasureCurrent = TheHdw.DCVS.Pins(pin).Meter.Read(tlStrobe, 1)
              Next pin
                  
              If NegRSTFlag = True Then
                  TheHdw.Wait 0.2
                   measureCurrent = TheHdw.DCVS.Pins(Pins_MeasureI_Together).Meter.Read(tlStrobe, SampleSize)
                   NegRSTFlag = False
              Else
               i = 5
              End If
          Next i
      ' End If
        
    Else
        'Call HardIP_UVS256_AutoRange(MeasureCurrent, Pins_MeasureI_Together, waitTime, CUS_STR_MainProgram)
        For i = 0 To 200
                For Each pin In measureCurrent.Pins
                      For Each site In TheExec.sites
                            ' thehdw.DCVS.Pins(pin).CurrentRange
                          If measureCurrent.Pins(pin).value(site) < (-TheHdw.DCVS.Pins(pin).CurrentRange / 166.67) Then
                               NewRange = TheHdw.DCVS.Pins(pin).Meter.CurrentRange * 0.09
                               If TheHdw.DCVS.Pins(pin).Meter.CurrentRange > TheHdw.DCVS.Pins(pin).CurrentRange.min Then
                                    TheHdw.DCVS.Pins(pin).SetCurrentRanges NewRange, NewRange
                                    'New Update 20170803, don't update Foldlimit to current range
'                                    TheHdw.DCVS.Pins(Pin).CurrentLimit.Source.FoldLimit.Level.Value = TheHdw.DCVS.Pins(Pin).CurrentLimit.Source.FoldLimit.Level.Value
                                    TheExec.Datalog.WriteComment "Site " & site & ", Run : " & i + 1 & ", Previous Result : " & measureCurrent.Pins(pin).value(site) & "===============================> Set pin " & pin & ", Range = " & TheHdw.DCVS.Pins(pin).Meter.CurrentRange
                                    NegRSTFlag = True
                               End If
                          End If
                       Next site
                       'MeasureCurrent = TheHdw.DCVS.Pins(pin).Meter.Read(tlStrobe, 1)
                Next pin

            If NegRSTFlag = True Then
                TheHdw.Wait (MaxWaitTime)
                measureCurrent = TheHdw.DCVS.Pins(Pins_MeasureI_Together).Meter.Read(tlStrobe, SampleSize)
                NegRSTFlag = False
            Else
                SumTime = i
                i = 200
            End If
        Next i
        ' End If
        
    TheExec.Datalog.WriteComment "wait time  =" & SumTime * 0.001
    
    End If
  '  ===============
    
    '' 20150701 - Restore source/sink fold limit, current range and filter setting
    For i = 0 To UBound(MI_TestCond_UVS256)
      Call DACTrim_UVS256__MI_RestoreCondition(MI_TestCond_UVS256(i).PinName, StoreSourceFoldLimit(i), StoreSinkFoldLimit(i), StoreFilterSetting(i), StoreSrcCurrentRange(i))
    Next i
    
End Function
Public Function DACTrim_UVS256__MI_StoreCondition(MI_Pin As String, ByRef SourceFlodLimit As Double, ByRef SinkFoldLimit As Double, ByRef FilterValue As Double, ByRef SrcCurrentRange) As Long
    
    Dim NumTypes As Long
    Dim PowerType() As String
    
    SourceFlodLimit = TheHdw.DCVS.Pins(MI_Pin).CurrentLimit.Source.FoldLimit.level.value
    SinkFoldLimit = TheHdw.DCVS.Pins(MI_Pin).CurrentLimit.Sink.FoldLimit.level.value
    FilterValue = TheHdw.DCVS.Pins(MI_Pin).Meter.Filter.value
    SrcCurrentRange = TheHdw.DCVS.Pins(MI_Pin).CurrentRange.value
''    Dim Pins() As String
''    Dim pin_cnt As Long
''
''    TheExec.DataManager.DecomposePinList MI_Pin, Pins(), pin_cnt
''
''    If pin_cnt > 1 Then
''        TheExec.Datalog.WriteComment ("Error for HardIP UVS256 current measurement: You can not use the pin group for test pin argument to do current measurement, please use single pin for test pin")
''    End If
''
''    Call TheExec.DataManager.GetChannelTypes(Pins(0), NumTypes, PowerType())

''    If UCase(PowerType(0)) Like UCase("*Merged*") Then
''    Else
''        TheHdw.DCVS.Pins(MI_Pin).CurrentRange.Value = 0.02
''        TheHdw.DCVS.Pins(MI_Pin).CurrentLimit.Source.FoldLimit.Level.Value = 0.02
''    End If

End Function
Public Function DACDigSrcDspWave(patt As String, DigSrcPin As PinList, SignalName As String, InWave As DSPWave)

    Dim site As Variant
    Dim WaveDef As String
    
    WaveDef = "WaveDef"
    ''20150708 - Comment program load
    TheHdw.Patterns(patt).Load  ' 20151211: addedback to fix error re: pattern not being loaded
    
    For Each site In TheExec.sites
    
        TheExec.WaveDefinitions.CreateWaveDefinition WaveDef & site, InWave, True
        
        TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals.Add SignalName
        With TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals(SignalName)
            .WaveDefinitionName = WaveDef & site
            .SampleSize = InWave.SampleSize
            .Amplitude = 1
            .LoadSamples
            .LoadSettings
        End With
            
        TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals.DefaultSignal = SignalName
    Next site
    
    ''20150708 - Comment repeat setting for DefaultSignal
''    TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals.DefaultSignal = SignalName
End Function
Public Function DACTrim_UVS256__MI_RestoreCondition(MI_Pin As String, ByRef SourceFlodLimit As Double, ByRef SinkFoldLimit As Double, ByRef FilterValue As Double, ByRef SrcCurrentRange As Double) As Long
    
    Dim NumTypes As Long
    Dim PowerType() As String
    
    Dim Pins() As String
    Dim Pin_Cnt As Long
    
    Dim ChannelName() As String
    Dim ChannelType As String
    Dim SharedChannels() As String
    Dim NumberChannels As Long
    Dim NumberSharedChannels As Long
    Dim NumberSites As Long
    Dim sites() As Long
    Dim Error As String
    
    TheExec.DataManager.DecomposePinList MI_Pin, Pins(), Pin_Cnt
    
    If Pin_Cnt > 1 Then
        'check share pin
        ChannelType = TheExec.DataManager.ChannelType(Pins(0))
        Call TheExec.DataManager.GetSharedChannelListForSelectedSites(MI_Pin, ChannelType, ChannelName(), SharedChannels(), NumberChannels, NumberSharedChannels, NumberSites, sites(), Error)
        If UBound(SharedChannels) = -1 Then
            TheExec.Datalog.WriteComment ("Error for HardIP UVS256 current measurement: You can not use the pin group for test pin argument to do current measurement, please use single pin for test pin")
        End If
    End If

    
''    Call TheExec.DataManager.GetChannelTypes(Pins(0), NumTypes, PowerType())
''
''    If UCase(PowerType(0)) Like UCase("*Merged*") Then
''    Else
''        If SinkFoldLimit > 0.02 Then
''            TheHdw.DCVS.Pins(MI_Pin).CurrentRange.Value = 0.2
''            TheHdw.DCVS.Pins(MI_Pin).CurrentLimit.Source.FoldLimit.Level.Value = 0.2
''            TheHdw.DCVS.Pins(MI_Pin).CurrentLimit.Sink.FoldLimit.Level.Value = 0.075
''        Else
''            TheHdw.DCVS.Pins(MI_Pin).CurrentRange.Value = 0.02
''            TheHdw.DCVS.Pins(MI_Pin).CurrentLimit.Source.FoldLimit.Level.Value = 0.02
''            TheHdw.DCVS.Pins(MI_Pin).CurrentLimit.Sink.FoldLimit.Level.Value = 0.02
''        End If
''    End If
    
    TheHdw.DCVS.Pins(MI_Pin).CurrentRange.value = CDbl(Format(SrcCurrentRange, "0.00"))
    TheHdw.DCVS.Pins(MI_Pin).CurrentLimit.Source.FoldLimit.level.value = CDbl(Format(SourceFlodLimit, "0.000"))
    TheHdw.DCVS.Pins(MI_Pin).CurrentLimit.Sink.FoldLimit.level.value = CDbl(Format(SinkFoldLimit, "0.000"))

    TheHdw.DCVS.Pins(MI_Pin).Meter.Filter.value = FilterValue
    
End Function
Public Function DisConnectMeasureVoltPins(MeasureV_Pin_PPMU As String, MeasureV_Pin_UVI80 As String) As String
  
    If MeasureV_Pin_PPMU <> "" Then
        TheHdw.PPMU.Pins(MeasureV_Pin_PPMU).Disconnect
    End If
    
    If MeasureV_Pin_UVI80 <> "" Then
        TheHdw.DCVI.Pins(MeasureV_Pin_UVI80).Disconnect tlDCVIConnectDefault
    End If
    
End Function
Public Function DACTrim_DigSrc_Data(Optional TrimTarget As Double, Optional TrimStart As Long, Optional TrimStop As Long, _
                                    Optional TrimSearchMethod As SearchMethod, Optional TrimEq As String, Optional Trimname As String, _
                                    Optional TrimDataWidth As Long, Optional TrimAssign As SiteVariant, Optional outwave As DSPWave, _
                                    Optional MeasVal As PinListData, Optional TestSeqNum As Integer, Optional digsrc_assignment As String, _
                                    Optional InitStart As Long, Optional TrimTargetCondition As TargetCondition) As Boolean
Dim TrimMethod As SearchMethod
Dim TrimIdx() As String
Dim TargetIdx() As String
Dim IniedFlag As Boolean
Dim TempStr As New SiteVariant
Dim TrimEquation() As String
Dim site, index As Variant
Dim TrimWaveStr As New SiteVariant
'Dim TrimEqWaveWaveStr As New SiteVariant
Dim TrimIndex As New SiteVariant
Dim IndexStr As String
Dim IndexVal As Long
Dim TrimSrcDec As New DSPWave
Dim TrimSrcWf As New DSPWave
Dim siteKey As String
Dim SiteSelect As New SiteBoolean
Dim SiteTrimStart As New SiteVariant
Dim SiteTrimStop As New SiteVariant
Dim SiteSTSP(1) As New SiteVariant
Dim STMV() As Double
Dim SPMV() As Double
Dim InitFlag As Boolean
'UCase reset
    Trimname = UCase(Trimname)
    TrimEq = UCase(TrimEq)
    TrimIdx = Split(Trimname, "@")
    TrimSrcDec.CreateConstant 0, 1
    IniedFlag = UBound(TrimIdx)
    If Not IniedFlag Then Call DACTrim_DigSrcSetting(TrimIdx(0), TrimDataWidth, TrimEq, digsrc_assignment, TrimSrcWf)
    If TrimIdx(0) <> "VERIFICATION" And TrimIdx(0) <> "READEFUSE" And TrimIdx(0) <> "" Then DACTrim_DigSrc_Data = True Else GoTo L
'Calc start/stop
    If TrimStart = 0 And TrimStop = 0 And TrimDataWidth > 0 Then
        TrimStop = 2 ^ TrimDataWidth - 1
        Trimname = Trimname & "@" & Dec2BinStr(0, TrimDataWidth)
        TrimIdx = Split(Trimname, "@")
        TrimAssign = SiteExpand(Trimname)
        TrimSrcDec = SiteExpandDSPWf(TrimStart)
        StepIdx = 1
        IndexVal = 0
        If InitStart > 0 Then InitFlag = True
    ElseIf Not IniedFlag Then
        If TrimStart > TrimStop Then StepIdx = -1 Else StepIdx = 1
        Trimname = Trimname & "@" & Dec2BinStr(TrimStart, TrimDataWidth): TrimIdx = Split(Trimname, "@")
        TrimAssign = SiteExpand(Trimname)
        TrimSrcDec = SiteExpandDSPWf(TrimStart)
        IndexVal = TrimStart
        If InitStart > 0 Then InitFlag = True
    End If
    
    If IniedFlag Then
        IndexVal = BinStr2Dec(TrimIdx(1)) + StepIdx: TrimIdx(1) = Dec2BinStr(IndexVal, TrimDataWidth)
        If IndexVal = TrimStop Then DACTrim_DigSrc_Data = False
        Trimname = TrimIdx(0) & "@" & TrimIdx(1)
    End If
    
    If Not InitFlag Then
        If TrimSearchMethod = Binary Then TrimSearchMethod = 292
        If TrimSearchMethod = Interpolation Then TrimSearchMethod = 3331
    Else
        InitFlag = False
        If TrimSearchMethod = linear Then TrimSearchMethod = 991
        If TrimSearchMethod = Binary Then TrimSearchMethod = 992
        If TrimSearchMethod = Interpolation Then TrimSearchMethod = 99331
    End If
    If TrimSearchMethod > 3 Then TrimMethod = GetCustomSearch(TrimSearchMethod, Abs(TrimStart - IndexVal)) Else TrimMethod = TrimSearchMethod
    
    Select Case TrimMethod
        Case DoAll
                If IniedFlag Then
                    SiteSTSP(0) = DACTrimValue.item("SiteSTSP")(0): SiteSTSP(1) = DACTrimValue.item("SiteSTSP")(1)
                    SiteSelect = TheExec.sites.Selected
                Else
                    SiteSTSP(0) = SiteExpand(TrimStart): SiteSTSP(1) = SiteExpand(TrimStop)
                    If DACTrimValue.Exists("SiteSTSP") Then DACTrimValue.Remove ("SiteSTSP")
                    Call DACTrimValue.Add("SiteSTSP", SiteSTSP)
                End If
        Case linear
                If IniedFlag Then
                    SiteSTSP(0) = DACTrimValue.item("SiteSTSP")(0): SiteSTSP(1) = DACTrimValue.item("SiteSTSP")(1)
                    SiteSelect = TheExec.sites.Selected
                Else
                    SiteSTSP(0) = SiteExpand(TrimStart): SiteSTSP(1) = SiteExpand(TrimStop)
                    If DACTrimValue.Exists("SiteSTSP") Then DACTrimValue.Remove ("SiteSTSP")
                    Call DACTrimValue.Add("SiteSTSP", SiteSTSP)
                End If
        Case Binary
                If IniedFlag Then
                    SiteSTSP(0) = DACTrimValue.item("SiteSTSP")(0): SiteSTSP(1) = DACTrimValue.item("SiteSTSP")(1)
                    SiteSelect = TheExec.sites.Selected
                Else
                    SiteSTSP(0) = SiteExpand(TrimStart): SiteSTSP(1) = SiteExpand(TrimStop)
                    If DACTrimValue.Exists("SiteSTSP") Then DACTrimValue.Remove ("SiteSTSP")
                    Call DACTrimValue.Add("SiteSTSP", SiteSTSP)
                End If
        Case Interpolation
                If IniedFlag Then
                    SiteSTSP(0) = DACTrimValue.item("SiteSTSP")(0): SiteSTSP(1) = DACTrimValue.item("SiteSTSP")(1)
                    SiteSelect = TheExec.sites.Selected
                Else
                    SiteSTSP(0) = SiteExpand(TrimStart): SiteSTSP(1) = SiteExpand(TrimStop)
                    If DACTrimValue.Exists("SiteSTSP") Then DACTrimValue.Remove ("SiteSTSP")
                    Call DACTrimValue.Add("SiteSTSP", SiteSTSP)
                End If
        Case Else
                If IniedFlag Then
                    SiteSTSP(0) = DACTrimValue.item("SiteSTSP")(0): SiteSTSP(1) = DACTrimValue.item("SiteSTSP")(1)
                    SiteSelect = TheExec.sites.Selected
                Else
                    SiteSTSP(0) = SiteExpand(TrimStart): SiteSTSP(1) = SiteExpand(TrimStop)
                    If DACTrimValue.Exists("SiteSTSP") Then DACTrimValue.Remove ("SiteSTSP")
                    Call DACTrimValue.Add("SiteSTSP", SiteSTSP)
                End If
    End Select
    Select Case TrimMethod
        Case DoAll
                If IniedFlag Then
                    For Each site In TheExec.sites
                        'siteKey = TrimAssign & "@" & TestSeqNum & "@site" & site
                        IndexVal = Do_All(SiteSTSP(0), SiteSTSP(1), siteKey, TestSeqNum, TrimTarget, TrimStart, TrimStop)
                        If IndexVal <> -1 Then TrimSrcDec.Element(0) = IndexVal: TrimAssign = TrimIdx(0) & "@" & Dec2BinStr(IndexVal, TrimDataWidth) Else SiteSelect = False
                    Next site
                    DACTrimValue.Remove ("SiteSTSP"): Call DACTrimValue.Add("SiteSTSP", SiteSTSP)
                    TheExec.sites.Selected = SiteSelect
                    If Not CBool(TheExec.sites.Selected.Count) Then DACTrim_DigSrc_Data = False
                Else
                    TrimAssign = SiteExpand(TrimIdx(0) & "@" & TrimIdx(1))
                    TrimSrcDec = SiteExpandDSPWf(CLng(BinStr2Dec(TrimIdx(1))))
                    SiteSTSP(0) = SiteExpand(IndexVal + StepIdx)
                    DACTrimValue.Remove ("SiteSTSP"): Call DACTrimValue.Add("SiteSTSP", SiteSTSP)
                End If
        Case linear
                If IniedFlag Then
                    For Each site In TheExec.sites
                        siteKey = TrimAssign & "@" & TestSeqNum & "@site" & site
                        IndexVal = LinearSearch(SiteSTSP(0), SiteSTSP(1), siteKey, TestSeqNum, TrimTarget)
                        If IndexVal <> -1 Then TrimSrcDec.Element(0) = IndexVal: TrimAssign = TrimIdx(0) & "@" & Dec2BinStr(IndexVal, TrimDataWidth) Else SiteSelect = False
                    Next site
                    DACTrimValue.Remove ("SiteSTSP"): Call DACTrimValue.Add("SiteSTSP", SiteSTSP)
                    TheExec.sites.Selected = SiteSelect
                    If Not CBool(TheExec.sites.Selected.Count) Then DACTrim_DigSrc_Data = False
                End If
        Case Binary
                If IniedFlag Then
                    For Each site In TheExec.sites
                        siteKey = TrimAssign & "@" & TestSeqNum & "@site" & site
                        IndexVal = BinarySearch(SiteSTSP(0), SiteSTSP(1), siteKey, TestSeqNum, TrimTarget, TrimStart, TrimStop)
                        If IndexVal <> -1 Then TrimSrcDec.Element(0) = IndexVal: TrimAssign = TrimIdx(0) & "@" & Dec2BinStr(IndexVal, TrimDataWidth) Else SiteSelect = False 'DACTrim_DigSrc_Data = False
                    Next site
                    DACTrimValue.Remove ("SiteSTSP"): Call DACTrimValue.Add("SiteSTSP", SiteSTSP)
                    TheExec.sites.Selected = SiteSelect
                    If Not CBool(TheExec.sites.Selected.Count) Then DACTrim_DigSrc_Data = False
                Else
                    TrimAssign = SiteExpand(TrimIdx(0) & "@" & Dec2BinStr(0.5 * (TrimStart + TrimStop), TrimDataWidth))
                    TrimSrcDec = SiteExpandDSPWf(CLng(0.5 * (TrimStart + TrimStop)))
                End If
        Case Interpolation
                If IniedFlag Then
                    If Abs(IndexVal - TrimStart) = 1 Then
                        TrimAssign = SiteExpand(TrimIdx(0) & "@" & Dec2BinStr(TrimStop, TrimDataWidth))
                        TrimSrcDec = SiteExpandDSPWf(CLng(TrimStop))
                    Else
                        For Each site In TheExec.sites
                            siteKey = TrimAssign & "@" & TestSeqNum & "@site" & site
                            IndexVal = Interpolate(SiteSTSP(0), SiteSTSP(1), siteKey, TestSeqNum, TrimTarget, site)
                            If IndexVal <> -1 Then TrimSrcDec.Element(0) = IndexVal: TrimAssign = TrimIdx(0) & "@" & Dec2BinStr(IndexVal, TrimDataWidth) Else SiteSelect = False 'DACTrim_DigSrc_Data = False
                        Next site
                    End If
                    DACTrimValue.Remove ("SiteSTSP"): Call DACTrimValue.Add("SiteSTSP", SiteSTSP)
                    TheExec.sites.Selected = SiteSelect
                    If Not CBool(TheExec.sites.Selected.Count) Then DACTrim_DigSrc_Data = False
                End If
         Case Transitions
                If IniedFlag Then
                    For Each site In TheExec.sites
                        'siteKey = TrimAssign & "@" & TestSeqNum & "@site" & site
                        IndexVal = TransitionSearch(SiteSTSP(0), SiteSTSP(1), siteKey, TestSeqNum, TrimTarget, TrimStart, TrimStop)
                        If IndexVal <> -1 Then TrimSrcDec.Element(0) = IndexVal: TrimAssign = TrimIdx(0) & "@" & Dec2BinStr(IndexVal, TrimDataWidth) Else SiteSelect = False
                    Next site
                    DACTrimValue.Remove ("SiteSTSP"): Call DACTrimValue.Add("SiteSTSP", SiteSTSP)
                    TheExec.sites.Selected = SiteSelect
                    If Not CBool(TheExec.sites.Selected.Count) Then DACTrim_DigSrc_Data = False
                Else
                    TrimAssign = SiteExpand(TrimIdx(0) & "@" & TrimIdx(1))
                    TrimSrcDec = SiteExpandDSPWf(CLng(BinStr2Dec(TrimIdx(1))))
                    SiteSTSP(0) = SiteExpand(IndexVal + StepIdx)
                    DACTrimValue.Remove ("SiteSTSP"): Call DACTrimValue.Add("SiteSTSP", SiteSTSP)
                End If
         Case 9 'Boundary(TrimStart or TrimStop)
                If IniedFlag Then
                    For Each site In TheExec.sites
                        siteKey = TrimAssign & "@" & TestSeqNum & "@site" & site
                        IndexVal = Boundary(SiteSTSP(0), SiteSTSP(1), siteKey, TestSeqNum, TrimTarget, TrimStart, TrimStop)
                        If IndexVal <> -1 Then TrimSrcDec.Element(0) = IndexVal: TrimAssign = TrimIdx(0) & "@" & Dec2BinStr(IndexVal, TrimDataWidth) Else SiteSelect = False 'DACTrim_DigSrc_Data = False
                    Next site
                    DACTrimValue.Remove ("SiteSTSP"): Call DACTrimValue.Add("SiteSTSP", SiteSTSP)
                    TheExec.sites.Selected = SiteSelect
                    If Not CBool(TheExec.sites.Selected.Count) Then DACTrim_DigSrc_Data = False
                Else
                    If InitStart > 0 Then
                        TrimAssign = SiteExpand(TrimIdx(0) & "@" & Dec2BinStr(InitStart, TrimDataWidth))
                        TrimSrcDec = SiteExpandDSPWf(InitStart)
                    Else
                        TrimAssign = SiteExpand(TrimIdx(0) & "@" & Dec2BinStr(0.5 * (TrimStart + TrimStop), TrimDataWidth))
                        TrimSrcDec = SiteExpandDSPWf(CLng(0.5 * (TrimStart + TrimStop)))
                    End If
                End If
            Case Else
    End Select
    ''OscarLi_Compile,20190629
    'Call rundsp.DSPWf_Dec2Binary(TrimSrcDec, TrimDataWidth, TrimSrcWf)
L:  Call rundsp.TrimSrcRP(G_TrimWave, RPIndex, TrimSrcWf, outwave)

    If TrimIdx(0) = "VERIFICATION" Or TrimIdx(0) = "READEFUSE" Then
        TheExec.Datalog.WriteComment ("======== Setup Dig Src Test ========")
        TheExec.Datalog.WriteComment "DataSequence:" & TrimEq
        If LSBFirst Then
            For Each site In TheExec.sites
                TrimWaveStr = BinDSP2Str(outwave, RPIndex)
                TheExec.Datalog.WriteComment "Site" & site & " SrcData:[LSB:MSB]" & TrimWaveStr
            Next site
        Else
            For Each site In TheExec.sites
                TrimWaveStr = BinDSP2Str(outwave, RPIndex)
                TheExec.Datalog.WriteComment "Site" & site & " SrcData:[MSB:LSB]" & TrimWaveStr
            Next site
        End If
    End If


'    TrimEquation = Split(TrimEq, "+")
'    For Each Index In TrimEquation
'          If Index = TrimIdx(0) Then
'              If LSBFirst Then
'                  For Each site In TheExec.sites
'                     TrimWaveStr = TrimWaveStr & StrReverse(GetTrimCode(CStr(TrimAssign)))
'                  Next site
'              Else
'                  For Each site In TheExec.sites
'                     TrimWaveStr = TrimWaveStr & GetTrimCode(CStr(TrimAssign))
'                  Next site
'              End If
'          ElseIf TrimIdx(0) = "READEFUSE" Then
'          Else
'              ' Get TrimEq Wave
'              If DACTargetStr.Exists(Index) Then
'                  TempStr = DACTargetStr.Item(Index)
'              Else
'                  TempStr = SiteExpand(CStr("@" & Index))
'              End If
'              If LSBFirst Then
'                  For Each site In TheExec.sites
'                      TrimWaveStr = TrimWaveStr & StrReverse(GetTrimCode(CStr(TempStr)))
'                  Next site
'              Else
'                   For Each site In TheExec.sites
'                      TrimWaveStr = TrimWaveStr & GetTrimCode(CStr(TempStr))
'                   Next site
'              End If
'          End If
'    Next Index
'    Call BinStr2DWave(TrimWaveStr, OutWave)
'
'    If TrimIdx(0) = "VERIFICATION" Or Not ByPassTestLimit Then
'        TheExec.Datalog.WriteComment ("======== Setup Dig Src Test ========")
'        TheExec.Datalog.WriteComment "DataSequence:" & TrimEq
'        If LSBFirst Then
'            For Each site In TheExec.sites
'                TheExec.Datalog.WriteComment "Site" & site & " SrcData:[LSB:MSB]" & TrimWaveStr
'            Next site
'        Else
'            For Each site In TheExec.sites
'                TheExec.Datalog.WriteComment "Site" & site & " SrcData:[MSB:LSB]" & TrimWaveStr
'            Next site
'        End If
'    End If



End Function

Public Function CheckInputStringAt(ByRef InputString As String) As Long
    If right(InputString, 1) <> "@" And InputString <> "" Then
        InputString = "@" & InputString
    End If
    If InputString <> "" Then
        InputString = mid(InputString, 2, Len(InputString) - 1)
    End If
End Function

Public Function SiteExpand(InVar As Variant) As SiteVariant
    Dim site As Variant
    Dim ss As New SiteVariant
    For Each site In TheExec.sites
        ss = InVar
    Next site
    Set SiteExpand = ss
End Function
Public Function SiteExpandDSPWf(InVar As Variant) As DSPWave
    Dim site As Variant
    Dim ss As New DSPWave
    ss.CreateConstant 0, 1
    For Each site In TheExec.sites
        ss.Element(0) = InVar
    Next site
    Set SiteExpandDSPWf = ss
End Function


Public Function CalcTarget_V2(MeasVal() As PinListData, TestSeqNum As Integer, Trimname As String, TrimAssign As SiteVariant, LastInterval As SiteDouble, TargetVal As Double, TestSequenceArray() As String, TrimCalcName As String, TrimTargetCondition As TargetCondition)
Dim MeasPin As String
Dim siteKey As String
Dim site As Variant
Dim TargetResult As New SiteVariant
Dim Interval As Double
Dim i As Integer
Dim ValIdx() As Double
Dim SVFlag As Boolean
Dim TrimValuePinlist As New PinListData
Dim TrimValueDSPWV As New DSPWave
Dim SiteSelect As New SiteBoolean
Dim TrimVal_DSP2Dec As New DSPWave

    'TrimName = UCase(TrimName)
    If TrimStoreType <> NoTrimCalcName Then
        If TestSeqNum = 0 Then
            If UBound(TestSequenceArray) = 0 And TestSequenceArray(0) <> "N" And TestSequenceArray(0) <> "" Then
                TestSeqNum = TestSeqNum + 1
            End If
        Else
            TestSeqNum = TestSeqNum + 1
        End If
    End If
    ReDim ValIdx(TestSeqNum)
    
    If Trimname Like "VERIFICATION" Or Trimname Like "READEFUSE" Then Exit Function
    
    If TrimStoreType = SPinListData Then
        TrimValuePinlist = GetStoredMeasurement(TrimCalcName)
        MeasPin = TrimValuePinlist.Pins(0).name
    ElseIf TrimStoreType = SDSPWave Then
        TrimValueDSPWV = GetStoredCaptureData(TrimCalcName)
        Call rundsp.Trim_ConvertToLongAndSerialToParrel(TrimValueDSPWV, TrimVal_DSP2Dec)
    Else
        MeasPin = MeasVal(TestSeqNum).Pins(0).name
    End If
    
    'Store ALL /per code/ValIdx/all measVal
    For Each site In TheExec.sites
        siteKey = TrimAssign & "@" & CStr(TestSeqNum) & "@site" & CStr(site) 'TrimName@TrimCode@TrimSeq@SiteNum
        
        If TestSequenceArray(0) <> "" Then 'check VIF sequence
            For i = 0 To TestSeqNum
                If i = TestSeqNum Then
                    If TrimStoreType = SPinListData Then
                        ValIdx(i) = TrimValuePinlist.Pins(0).value
                    ElseIf TrimStoreType = SDSPWave Then
                        ValIdx(i) = TrimVal_DSP2Dec.Element(0)
                    Else
                        ValIdx(i) = MeasVal(i).Pins(MeasPin).value
                    End If
                Else
                    ValIdx(i) = MeasVal(i).Pins(MeasPin).value
                End If
            Next i
        Else
            If TrimStoreType = SPinListData Then
                ValIdx(0) = TrimValuePinlist.Pins(0).value
            ElseIf TrimStoreType = SDSPWave Then
                ValIdx(0) = TrimVal_DSP2Dec.Element(0)
            End If
        End If
        
        If DACTrimValue.Exists(siteKey) Then
            DACTrimValue.Remove (siteKey)
        End If
        Call DACTrimValue.Add(siteKey, ValIdx)
    Next site
    
    If DACTargetStr.Exists(Trimname) Then
        SiteSelect = TheExec.sites.Selected
        TheExec.sites.Selected = True
        TargetResult = DACTargetStr.item(Trimname) 'TargetResult = TrimName@CodeStr@TestSeqNum@SiteNum
        TheExec.sites.Selected = SiteSelect
    Else
        SiteSelect = TheExec.sites.Selected
        For Each site In TheExec.sites
            ValIdx = DACTrimValue.item(TrimAssign & "@" & CStr(TestSeqNum) & "@site" & CStr(site))
            TargetResult = TrimAssign & "@" & CStr(TestSeqNum) & "@site" & CStr(site) & "@" & ValIdx(TestSeqNum) 'MeasVal(TestSeqNum).Pins(MeasPin).Value
            LastInterval = TargetVal - ValIdx(TestSeqNum) 'MeasVal(TestSeqNum).Pins(MeasPin).Value
        Next site
        SVFlag = True
    End If
    
    'Calc Target
    Select Case TrimTargetCondition
        Case Equal
            For Each site In TheExec.sites
                ValIdx = DACTrimValue.item(TrimAssign & "@" & CStr(TestSeqNum) & "@site" & CStr(site))
                Interval = TargetVal - ValIdx(TestSeqNum) 'MeasVal(TestSeqNum).Pins(MeasPin).Value
                If Abs(LastInterval) > Abs(Interval) Then
                    LastInterval = Interval
                    TargetResult = TrimAssign & "@" & CStr(TestSeqNum) & "@site" & CStr(site) & "@" & ValIdx(TestSeqNum) 'MeasVal(TestSeqNum).Pins(MeasPin).Value
                    SVFlag = True
                End If
            Next site
        
        Case Transition
            For Each site In TheExec.sites
                ValIdx = DACTrimValue.item(TrimAssign & "@" & CStr(TestSeqNum) & "@site" & CStr(site))
                Interval = TargetVal - ValIdx(TestSeqNum) 'MeasVal(TestSeqNum).Pins(MeasPin).Value
                If Not SearchDone Then
                    If Abs(LastInterval) > Abs(Interval) Then
                        TargetResult = TrimAssign & "@" & CStr(TestSeqNum) & "@site" & CStr(site) & "@" & ValIdx(TestSeqNum)
                        SVFlag = True
                    End If
                    If ((LastInterval < 0) Xor (Interval < 0)) Or (Interval = 0) Then
                        SearchDone = True: SiteSelect = False
                        TargetResult = TrimAssign & "@" & CStr(TestSeqNum) & "@site" & CStr(site) & "@" & ValIdx(TestSeqNum)
                        SVFlag = True
                    End If
                    If SVFlag Then LastInterval = Interval
                End If
            Next site
            TheExec.sites.Selected = SiteSelect
        Case TrimTargetCondition = GreaterThanEqual
        Case TrimTargetCondition = LessThanEqual
        Case Else
    End Select
    
    If SVFlag Then
        If DACTargetStr.Exists(Trimname) Then
            DACTargetStr.Remove (Trimname)
        End If
        Call DACTargetStr.Add(Trimname, TargetResult)
    End If
End Function

Public Function Do_All(TrimStart As SiteLong, TrimStop As SiteLong, TrimAssign As String, TestSeqNum As Integer, TrimTarget As Double, TrimST As Long, TrimSP As Long) As Long
Dim MeasVal() As Double
    Do_All = -1
'    If DACTrimValue.Exists(TrimAssign) Then MeasVal = DACTrimValue.Item(TrimAssign)
'    Select Case StepIdx
'        Case 1
'            If TrimStart = TrimST Then
'                Do_All = TrimStart
'                TrimStart = TrimStart + StepIdx
'            Else
                Do_All = TrimStart
                TrimStart = TrimStart + StepIdx
'                If TrimStart = TrimSP Then
'                    Do_All = TrimStart
                'If ((StepIdx) * Do_All > (StepIdx) * TrimStop) Xor RV Then
                If ((StepIdx) * Do_All > (StepIdx) * TrimStop) Then
                    Do_All = -1
                End If
                
                If SearchDone = True Then Do_All = -1 ' for transition
'            End If
'        Case Else
'            If TrimStart = TrimST Then
'                Doll = TrimStart
'                TrimStart = TrimStart + StepIdx
'            Else
'                TrimStart = TrimStart + StepIdx
'                If TrimStart = TrimSP Then
'                    DoAll = TrimStart
'                ElseIf TrimStart + StepIdx = TrimStop Then
'                    DoAll = -1
'                End If
'            End If
'    End Select
End Function

Public Function TransitionSearch(TrimStart As SiteLong, TrimStop As SiteLong, TrimAssign As String, TestSeqNum As Integer, TrimTarget As Double, TrimST As Long, TrimSP As Long) As Long
Dim MeasVal() As Double

    TransitionSearch = -1
    TransitionSearch = TrimStart
    TrimStart = TrimStart + StepIdx
    If ((StepIdx) * TransitionSearch > (StepIdx) * TrimStop) Then
        TransitionSearch = -1
    End If
    If SearchDone = True Then TransitionSearch = -1 ' for transition

End Function
Public Function BinarySearch(TrimStart As SiteLong, TrimStop As SiteLong, TrimAssign As String, TestSeqNum As Integer, TrimTarget As Double, TrimST As Long, TrimSP As Long) As Long
Dim MeasVal() As Double
    BinarySearch = -1
    If DACTrimValue.Exists(TrimAssign) Then MeasVal = DACTrimValue.item(TrimAssign)
    Select Case StepIdx
        Case 1
            If (MeasVal(TestSeqNum) < TrimTarget) Xor RV Then
                TrimStart = BinStr2Dec(GetTrimCode(TrimAssign))
                BinarySearch = Abs(Int(-0.5 * (TrimStart + TrimStop)))
                If (TrimStop = BinarySearch) Xor TrimSP - TrimStart = StepIdx Then BinarySearch = -1 ': Stop
                If TrimStart = BinarySearch Then BinarySearch = -1 ': Stop
            ElseIf (MeasVal(TestSeqNum) > TrimTarget) Xor RV Then
                TrimStop = BinStr2Dec(GetTrimCode(TrimAssign))
                BinarySearch = Abs(Int(0.5 * (TrimStart + TrimStop)))
                If TrimStart = BinarySearch Xor TrimStop - TrimST = StepIdx Then BinarySearch = -1
                If TrimStop = BinarySearch Then BinarySearch = -1
            End If
            If TrimStart > TrimStop Then BinarySearch = -1 ': Stop
        Case Else
            If (MeasVal(TestSeqNum) > TrimTarget) Xor RV Then
                TrimStart = BinStr2Dec(GetTrimCode(TrimAssign))
                BinarySearch = Abs(Int(0.5 * (TrimStart + TrimStop)))
                If TrimStop = BinarySearch Xor TrimSP - TrimStart = StepIdx Then BinarySearch = -1 ': Stop
                If TrimStart = BinarySearch Then BinarySearch = -1 ': Stop
            ElseIf (MeasVal(TestSeqNum) < TrimTarget) Xor RV Then
                TrimStop = BinStr2Dec(GetTrimCode(TrimAssign))
                BinarySearch = Abs(Int(-0.5 * (TrimStart + TrimStop)))
                If TrimStart = BinarySearch Xor TrimStop - TrimST = StepIdx Then BinarySearch = -1
                If TrimStop = BinarySearch Then BinarySearch = -1
            End If
            If TrimStart < TrimStop Then BinarySearch = -1 ': Stop
    End Select
    If TrimDebug Then TheExec.Datalog.WriteComment "BinarySearch : " & TrimStart & " : " & TrimStop & " : " & BinarySearch & " TrimAssign : " & TrimAssign
End Function

Public Function Boundary(TrimStart As SiteLong, TrimStop As SiteLong, TrimAssign As String, TestSeqNum As Integer, TrimTarget As Double, TrimST As Long, TrimSP As Long) As Long
Dim MeasVal() As Double
    Boundary = -1
    If DACTrimValue.Exists(TrimAssign) Then MeasVal = DACTrimValue.item(TrimAssign)
    Select Case StepIdx
        Case 1
            If (MeasVal(TestSeqNum) < TrimTarget) Xor RV Then
                TrimStart = BinStr2Dec(GetTrimCode(TrimAssign))
                Boundary = TrimStop
'                BinarySearch = Abs(Int(-0.5 * (TrimStart + TrimStop)))
'                If (TrimStop = BinarySearch) Xor TrimSP - TrimStart = StepIdx Then BinarySearch = -1 ': Stop
'                If TrimStart = BinarySearch Then BinarySearch = -1 ': Stop
            ElseIf (MeasVal(TestSeqNum) > TrimTarget) Xor RV Then
                TrimStop = BinStr2Dec(GetTrimCode(TrimAssign))
                Boundary = TrimStart
'                BinarySearch = Abs(Int(0.5 * (TrimStart + TrimStop)))
'                If TrimStart = BinarySearch Xor TrimStop - TrimST = StepIdx Then BinarySearch = -1
'                If TrimStop = BinarySearch Then BinarySearch = -1
            End If
            If TrimStart > TrimStop Then Boundary = -1 ': Stop
        Case Else
            If (MeasVal(TestSeqNum) > TrimTarget) Xor RV Then
                TrimStart = BinStr2Dec(GetTrimCode(TrimAssign))
                Boundary = TrimStop
'                BinarySearch = Abs(Int(0.5 * (TrimStart + TrimStop)))
'                If TrimStop = BinarySearch Xor TrimSP - TrimStart = StepIdx Then BinarySearch = -1 ': Stop
'                If TrimStart = BinarySearch Then BinarySearch = -1 ': Stop
            ElseIf (MeasVal(TestSeqNum) < TrimTarget) Xor RV Then
                TrimStop = BinStr2Dec(GetTrimCode(TrimAssign))
                Boundary = TrimStart
'                BinarySearch = Abs(Int(-0.5 * (TrimStart + TrimStop)))
'                If TrimStart = BinarySearch Xor TrimStop - TrimST = StepIdx Then BinarySearch = -1
'                If TrimStop = BinarySearch Then BinarySearch = -1
            End If
            If TrimStart < TrimStop Then Boundary = -1 ': Stop
    End Select
    If TrimDebug Then TheExec.Datalog.WriteComment "Boundary : " & TrimStart & " : " & TrimStop & " : " & Boundary & " TrimAssign : " & TrimAssign
End Function

Public Function LinearSearch(TrimStart As SiteLong, TrimStop As SiteLong, TrimAssign As String, TestSeqNum As Integer, TrimTarget As Double) As Long
Dim MeasVal() As Double
    LinearSearch = -1
    If DACTrimValue.Exists(TrimAssign) Then MeasVal = DACTrimValue.item(TrimAssign)
    Select Case StepIdx
        Case 1
            If (MeasVal(TestSeqNum) < TrimTarget) Xor RV Then
                TrimStart = BinStr2Dec(GetTrimCode(TrimAssign))
                LinearSearch = TrimStart + StepIdx
                If LinearSearch > TrimStop Then LinearSearch = -1
            ElseIf (MeasVal(TestSeqNum) > TrimTarget) Xor RV Then
                TrimStop = BinStr2Dec(GetTrimCode(TrimAssign))
                LinearSearch = TrimStop - StepIdx
                If LinearSearch <= TrimStart Then LinearSearch = -1 '20170612 modify from If LinearSearch = TrimStart Then LinearSearch = -1
            End If
        Case Else
            If (MeasVal(TestSeqNum) > TrimTarget) Xor RV Then
                TrimStart = BinStr2Dec(GetTrimCode(TrimAssign))
                LinearSearch = TrimStart + StepIdx
                If LinearSearch < TrimStop Then LinearSearch = -1
            ElseIf (MeasVal(TestSeqNum) < TrimTarget) Xor RV Then
                TrimStop = BinStr2Dec(GetTrimCode(TrimAssign))
                LinearSearch = TrimStop - StepIdx
                If LinearSearch >= TrimStart Then LinearSearch = -1 '20170612 modify from If LinearSearch = TrimStart Then LinearSearch = -1
            End If
    End Select
'If LinearSearch = -1 Then Stop
    If TrimDebug Then TheExec.Datalog.WriteComment "LinearSearch : " & TrimStart & " : " & TrimStop & " : " & LinearSearch & " TrimAssign : " & TrimAssign
End Function

Public Function Interpolate(TrimStart As SiteLong, TrimStop As SiteLong, TrimAssign As String, TestSeqNum As Integer, TrimTarget As Double, site As Variant) As Long
Dim MeasVal() As Double
Dim STMV() As Double
Dim SPMV() As Double
    Interpolate = -1
    If DACTrimValue.Exists(TrimAssign) Then MeasVal = DACTrimValue.item(TrimAssign)
    Select Case StepIdx
        Case 1
            If (MeasVal(TestSeqNum) < TrimTarget) Xor RV Then
                TrimStart = BinStr2Dec(GetTrimCode(TrimAssign)): If TrimStart = TrimStop Then Interpolate = -1: Exit Function
                TrimAssign = GetTrimName(TrimAssign) & "@" & Dec2BinStr(CLng(TrimStop), Len(GetTrimCode(TrimAssign))) & "@" & TestSeqNum & "@site" & site: SPMV = DACTrimValue.item(TrimAssign)
                If SPMV(TestSeqNum) > TrimTarget Xor RV Then
                    Interpolate = Fix((TrimStop - TrimStart) * (TrimTarget - MeasVal(TestSeqNum)) / (SPMV(TestSeqNum) - MeasVal(TestSeqNum)) + TrimStart)
                End If
                If Interpolate = TrimStart Then Interpolate = Interpolate + StepIdx
                If Interpolate = TrimStop Then Interpolate = Interpolate - StepIdx
            ElseIf (MeasVal(TestSeqNum) > TrimTarget) Xor RV Then
                TrimStop = BinStr2Dec(GetTrimCode(TrimAssign)): If TrimStart = TrimStop Then Interpolate = -1: Exit Function
                TrimAssign = GetTrimName(TrimAssign) & "@" & Dec2BinStr(CLng(TrimStart), Len(GetTrimCode(TrimAssign))) & "@" & TestSeqNum & "@site" & site: STMV = DACTrimValue.item(TrimAssign)
                If STMV(TestSeqNum) < TrimTarget Xor RV Then
                    Interpolate = Fix((TrimStop - TrimStart) * (TrimTarget - STMV(TestSeqNum)) / (MeasVal(TestSeqNum) - STMV(TestSeqNum)) + TrimStart) ' + 1
                End If
                If Interpolate = TrimStop Then Interpolate = Interpolate - StepIdx
                If Interpolate = TrimStart Then Interpolate = Interpolate + StepIdx
            End If
            If Interpolate <= TrimStart Or Interpolate >= TrimStop Then Interpolate = -1
        Case Else
            If (MeasVal(TestSeqNum) > TrimTarget) Xor RV Then
                TrimStart = BinStr2Dec(GetTrimCode(TrimAssign)): If TrimStart = TrimStop Then Interpolate = -1: Exit Function
                TrimAssign = GetTrimName(TrimAssign) & "@" & Dec2BinStr(CLng(TrimStop), Len(GetTrimCode(TrimAssign))) & "@" & TestSeqNum & "@site" & site: SPMV = DACTrimValue.item(TrimAssign)
                If SPMV(TestSeqNum) < TrimTarget Xor RV Then
                    Interpolate = Fix((TrimStop - TrimStart) * (TrimTarget - MeasVal(TestSeqNum)) / (SPMV(TestSeqNum) - MeasVal(TestSeqNum)) + TrimStart)
                End If
                If Interpolate = TrimStart Then Interpolate = Interpolate + StepIdx
                If Interpolate = TrimStop Then Interpolate = Interpolate - StepIdx
            ElseIf (MeasVal(TestSeqNum) < TrimTarget) Xor RV Then
                TrimStop = BinStr2Dec(GetTrimCode(TrimAssign)): If TrimStart = TrimStop Then Interpolate = -1: Exit Function
                TrimAssign = GetTrimName(TrimAssign) & "@" & Dec2BinStr(CLng(TrimStart), Len(GetTrimCode(TrimAssign))) & "@" & TestSeqNum & "@site" & site: STMV = DACTrimValue.item(TrimAssign)
                If STMV(TestSeqNum) > TrimTarget Xor RV Then
                    Interpolate = Fix((TrimStop - TrimStart) * (TrimTarget - STMV(TestSeqNum)) / (MeasVal(TestSeqNum) - STMV(TestSeqNum)) + TrimStart) ' + 1
                End If
                If Interpolate = TrimStop Then Interpolate = Interpolate - StepIdx
                If Interpolate = TrimStart Then Interpolate = Interpolate + StepIdx
            End If
            If Interpolate >= TrimStart Or Interpolate <= TrimStop Then Interpolate = -1
    End Select
    If TrimDebug Then TheExec.Datalog.WriteComment "Interpolate : " & TrimStart & " : " & TrimStop & " : " & Interpolate & " TrimAssign : " & TrimAssign
End Function

Public Function GetTrimName(InputStr As String) As String
Dim TrimIdx() As String
    TrimIdx = Split(InputStr, "@")
    GetTrimName = TrimIdx(0)
End Function

Public Function GetTrimCode(InputStr As String) As String
Dim TrimIdx() As String
    TrimIdx = Split(InputStr, "@")
    GetTrimCode = TrimIdx(1)
End Function
Public Function GetTrimVal(InputStr As String) As Variant
Dim TrimIdx() As String
    TrimIdx = Split(InputStr, "@")
    GetTrimVal = CDbl(TrimIdx(4))
End Function
Public Function TrimLimit_V2(ByVal Trimname As String, DataWdth As Long, code As SiteLong, FinVal As SiteDouble, TestSequenceArray() As String, Meas_ToTestlimit() As PinListData, ByVal TestSeqNum As Integer, Optional MeasPin As String, Optional DigSrcTrimPinName As PinList) As Variant
Dim TargetIdx As New SiteVariant
Dim Long_Wf As New DSPWave
Dim LSB_Wf As New DSPWave
Dim site As Variant
Dim DACTrimTestName As String
Dim Ts, i As Integer
Dim DefaultSet As Boolean
    
    'If TrimName <> "VERIFICATION" Then
    'If DACTargetStr.Exists(TrimName) Then TargetIdx = DACTargetStr.Item(TrimName)
    Long_Wf.CreateConstant 0, 1, DspLong
    For Each site In TheExec.sites
        code = BinStr2Dec(fuseCode(Trimname))
        FinVal = FuseVal(Trimname)
        Long_Wf.Element(0) = code
    Next site
    'End If
    'Call StoreDoAll(DefaultSet)  ' Can be removed if use Bintable
    
    'TrimValueLimit
        If ArrayCheck(Meas_ToTestlimit) Then
                If TheExec.TesterMode = testModeOnline Then
                        If TrimStoreType = NoTrimCalcName Then
                                For i = 0 To UBound(Meas_ToTestlimit) - 1
                                        TheExec.Flow.TestLimit resultVal:=Meas_ToTestlimit(i), Tname:="MeasValue" & i, ForceResults:=tlForceFlow
                                Next i
                        ElseIf TrimStoreType = SDSPWave Then
                                For i = 0 To UBound(Meas_ToTestlimit)
                                        TheExec.Flow.TestLimit resultVal:=Meas_ToTestlimit(i), Tname:="MeasValue" & i, ForceResults:=tlForceFlow
                                Next i
                        ElseIf TrimStoreType = SPinListData Then
                                For i = 0 To UBound(Meas_ToTestlimit)
                                        TheExec.Flow.TestLimit resultVal:=Meas_ToTestlimit(i), Tname:="MeasValue" & i, ForceResults:=tlForceFlow
                                Next i
                        End If
                Else
                        If TrimStoreType = NoTrimCalcName Then
                                For i = 0 To UBound(Meas_ToTestlimit) - 1
                                        TheExec.Flow.TestLimit resultVal:=Meas_ToTestlimit(i), Tname:="MeasValue" & i, ForceResults:=tlForceNone 'eng_forceflow_transfer
                                Next i
                        ElseIf TrimStoreType = SDSPWave Then
                                For i = 0 To UBound(Meas_ToTestlimit)
                                        TheExec.Flow.TestLimit resultVal:=Meas_ToTestlimit(i), Tname:="MeasValue" & i, ForceResults:=tlForceNone 'eng_forceflow_transfer
                                Next i
                        ElseIf TrimStoreType = SPinListData Then
                                For i = 0 To UBound(Meas_ToTestlimit)
                                        TheExec.Flow.TestLimit resultVal:=Meas_ToTestlimit(i), Tname:="MeasValue" & i, ForceResults:=tlForceNone 'eng_forceflow_transfer
                                Next i
                        End If
                End If
        End If
    
'    If TheExec.Flow.EnableWord("CZ2_PRINT_EN") = False Then

                   If TPModeAsCharz_GLB = True Then
                        TheExec.Flow.TestLimit resultVal:=FinVal, ForceResults:=tlForceFlow
                   Else
                        TheExec.Flow.TestLimit resultVal:=FinVal, Tname:="TrimValue", ForceResults:=tlForceFlow
                    End If
'    Else
'        For Each Site In TheExec.sites
'            MeasPin = Meas_ToTestlimit(0).Pins(0).Name
'            If MeasPin <> "" Then Exit For
'        Next Site
'        Report_TestLimit_by_CZ_Format FinVal, Tname:=DACTrimTestName, ForceResults:=tlForceFlow, unit:=unitHz, PinName:=MeasPin
'    End If
    
    'TrimCodeLimit
    If TheExec.TesterMode = testModeOnline Then
'        If TheExec.Flow.EnableWord("CZ2_PRINT_EN") = False Then
                   If TPModeAsCharz_GLB = True Then
                        TheExec.Flow.TestLimit resultVal:=code, ForceResults:=tlForceFlow
                   Else
                        TheExec.Flow.TestLimit resultVal:=code, Tname:="DecimalCode", ForceResults:=tlForceNone 'eng_forceflow_transfer
                    End If
'        Else
'            Report_TestLimit_by_CZ_Format resultVal:=code, Tname:="DecimalCode", ForceResults:=tlForceFlow, MeasType:="C", PinName:=DigSrcTrimPinName
'        End If
    Else
''        For Each Site In TheExec.sites
            TheExec.Flow.TestLimit resultVal:=code, Tname:="DecimalCode", ForceResults:=tlForceFlow
''        Next Site
    End If

    'For HDC future request
    For Each site In TheExec.sites
        If TheExec.sites.item(site).result = tlResultFail Then
            'read init data from efuse and passing into Long_Wf
            'Long_Wf.Element(0) = 0
        End If
    Next site
    ''OscarLi_Compile,20190629
    'Call rundsp.DSPWf_Dec2Binary(Long_Wf, DataWdth, LSB_Wf)
    Call AddStoredCaptureData(Trimname, LSB_Wf)  'For HIP stage
    Call AddDSPWf2Dic("DSPWF@" & Trimname, LSB_Wf) 'For DAC trim stage
    'Call RestoreDoAll(DefaultSet) ' Can be removed if use Bintable
End Function

Public Function TrimCodeLimit(ByVal Trimname As String, DataWdth As Long, code As SiteLong, FinVal As SiteDouble) As Long
Dim TargetIdx As New SiteVariant
Dim Long_Wf As New DSPWave
Dim LSB_Wf As New DSPWave
Dim site As Variant
    Long_Wf.CreateConstant 0, 1, DspLong
    If Trimname = "VERIFICATION" Then Exit Function
    'If DACTargetStr.Exists(TrimName) Then TargetIdx = DACTargetStr.Item(TrimName)
    For Each site In TheExec.sites
        code = BinStr2Dec(fuseCode(Trimname))
        FinVal = FuseVal(Trimname)
        Long_Wf.Element(0) = code
    Next site

    TheExec.Flow.TestLimit resultVal:=code, Tname:="DecimalCode", ForceResults:=tlForceFlow
    ''OscarLi_Compile,20190629
    'Call rundsp.DSPWf_Dec2Binary(Long_Wf, DataWdth, LSB_Wf)
    Call AddStoredCaptureData(Trimname, LSB_Wf)  'For HIP stage
    Call AddDSPWf2Dic("DSPWF@" & Trimname, LSB_Wf) 'For DAC trim stage
        
End Function

Public Function TrimVallimit(TestSequenceArray() As String, Meas_ToTestlimit() As PinListData, ByVal TestSeqNum As Integer)
Dim DACTrimTestName As String
Dim Ts, i As Integer
    
    Ts = UBound(TestSequenceArray)
        For i = 0 To TestSeqNum
            If i = TestSeqNum Then DACTrimTestName = "TrimValue" Else DACTrimTestName = "MeasValue" & i
            If i > Ts Then
                TheExec.Flow.TestLimit Meas_ToTestlimit(i), Tname:=DACTrimTestName, ForceResults:=tlForceFlow
            ElseIf TestSequenceArray(i) <> "N" Then
                TheExec.Flow.TestLimit Meas_ToTestlimit(i), Tname:=DACTrimTestName, ForceResults:=tlForceFlow
            End If
        Next i
End Function
Public Function TrimVallimit_V2(TestSequenceArray() As String, Meas_ToTestlimit() As PinListData, ByVal TestSeqNum As Integer, FinTrimVal As SiteDouble)
Dim DACTrimTestName As String
Dim Ts, i As Integer

     TheExec.Flow.TestLimit FinTrimVal, Tname:="TrimValue", ForceResults:=tlForceFlow
    
'    Ts = UBound(TestSequenceArray)
'        For i = 0 To TestSeqNum
'            If i = TestSeqNum Then DACTrimTestName = "TrimValue" Else DACTrimTestName = "MeasValue" & i
'            If i > Ts Then
'                TheExec.Flow.TestLimit Meas_ToTestlimit(i), Tname:=DACTrimTestName, ForceResults:=tlForceFlow
'            ElseIf TestSequenceArray(i) <> "N" Then
'                TheExec.Flow.TestLimit Meas_ToTestlimit(i), Tname:=DACTrimTestName, ForceResults:=tlForceFlow
'            End If
'        Next i
End Function

Public Function fuseCode(FuseName As String) As Variant
Dim TargetIdx As New SiteVariant
    If FuseName = "VERIFICATION" Then Exit Function
    If DACTargetStr.Exists(FuseName) Then TargetIdx = DACTargetStr.item(FuseName): fuseCode = GetTrimCode(CStr(TargetIdx))
End Function
Public Function FuseVal(FuseName As String) As Variant
Dim TargetIdx As New SiteVariant
    If FuseName = "VERIFICATION" Then Exit Function
    If DACTargetStr.Exists(FuseName) Then TargetIdx = DACTargetStr.item(FuseName): FuseVal = GetTrimVal(CStr(TargetIdx))
End Function

Public Function GetCustomSearch(TrimMethod As SearchMethod, idx As Long) As SearchMethod
    'If DACTrimValue.Exists(TrimMethod) Then TrimMethod = DACTrimValue.Item(TrimMethod) Else Call DACTrimValue.Add(TrimMethod, TrimMethod)
    If idx + 1 < Len(CStr(TrimMethod)) Then
        GetCustomSearch = mid(CStr(TrimMethod), idx + 1, 1)
    Else
        GetCustomSearch = mid(CStr(TrimMethod), Len(CStr(TrimMethod)), 1)
    End If
End Function

Public Function InverseBinStr(InputStr As String, Inverse As Boolean) As String
'1100101 -> 0011010
    If Inverse And Len(InputStr) > 0 Then
        InverseBinStr = CStr(Abs(CLng(left(InputStr, 1)) - 1)) + InverseBinStr(right(InputStr, Len(InputStr) - 1), Inverse)
    Else
        InverseBinStr = InputStr
    End If
End Function

Public Function InverseDec(DecVal As Long, Width As Long) As Variant
'(7,3) -> 0 : (8,3) or (-1,3) -> OverFlow
    If DecVal < 0 Then InverseDec = "OverFlow": Exit Function
    InverseDec = (2 ^ Width - 1) - DecVal
    If InverseDec < 0 Then InverseDec = "OverFlow"
End Function

Public Function MapSearchMethod(SHM As String) As SearchMethod
    SHM = LCase(SHM)
    Select Case SHM
        Case "doall"
            MapSearchMethod = DoAll
        Case "linear"
            MapSearchMethod = linear
        Case "binary"
            MapSearchMethod = Binary
        Case "interpolation"
            MapSearchMethod = Interpolation
        Case Else
            MapSearchMethod = DoAll
    End Select
End Function
Public Function DACTrim_ProcessCalcEquation(Calc_Eqn As String, SplitBySemi As String, MeasToTestlimit As PinListData, TestSeqNum As Integer) As Long
    
    Dim SplitByColon() As String, SplitByLeftPara() As String
    Dim i As Long, j As Long
    Dim KeyWord_Calc As String
    Dim ALG_InterPoseFuncName As String
    Dim ALG_InterPoseArgcName As String
    Dim TestName As String
    Dim Operator As String
    Dim SplitByKeyWord() As String
    Dim ReturnDSPWave As New DSPWave

    '' 20160914 - Equation pins for V,F,I
''    Dim EquationPins As String
    Dim CalcEquationPLD As CALC_EQUATION_PLD
    Dim ReturnPLD As New PinListData
    Dim pld2() As New PinListData
    'ReturnDSPWave.CreateConstant 0, 1, DspDouble
    
    'SplitBySemi = Split(Calc_Eqn, ";")
    
    Dim ALG_SplitBycomma() As String
    Dim ALG_InputKey As String
    Dim ALG_ReturnKey As String
    'For i = 0 To UBound(SplitBySemi)
        SplitByColon = Split(SplitBySemi, ":")
        KeyWord_Calc = UCase(SplitByColon(0))
        TestName = SplitByColon(1)
        TestSeqNum = TestSeqNum + 1
        'ReDim Preserve MeasToTestlimit(TestSeqNum)
        'ReDim Preserve PLD2(2)
        Select Case KeyWord_Calc
            Case "V"
                Call ProcessStringForCalType(SplitByColon(2), CalcEquationPLD)
                Call StandardCalcuation(CalcEquationPLD, MeasToTestlimit)
            
            Case "F"
                Call ProcessStringForCalType(SplitByColon(2), CalcEquationPLD)
                Call StandardCalcuation(CalcEquationPLD, MeasToTestlimit)
            
            Case "I"
                Call ProcessStringForCalType(SplitByColon(2), CalcEquationPLD)
                Call StandardCalcuation(CalcEquationPLD, MeasToTestlimit)
            
''            Case "C"
''                Operator = GetSplitKeyWord(SplitByColon(2))
''                SplitByKeyWord = Split(SplitByColon(2), Operator)
                Call ProcessDSPCalculation(SplitByKeyWord, Operator, ReturnDSPWave)

            Case "ALG"
                SplitByLeftPara = Split(SplitByColon(2), "(")

                ALG_InterPoseFuncName = SplitByLeftPara(0)
                ALG_InterPoseArgcName = Replace(SplitByLeftPara(1), ")", vbNullString)
                ALG_SplitBycomma = Split(ALG_InterPoseArgcName, ",")
                ALG_ReturnKey = ALG_SplitBycomma(0)
''                ALG_InputKey = ALG_SplitBycomma(1)
''                Call AddStoredMeasurement(ALG_InputKey, MeasToTestlimit)
''                TheExec.Flow.SetInterpose 1, ALG_InterPoseFuncName, ALG_InterPoseArgcName    ' key, func name, arguments
''                TheExec.Flow.ExecuteInterpose 1

                Call Interpose(ALG_InterPoseFuncName, ALG_InterPoseArgcName)
                MeasToTestlimit = GetStoredMeasurement(ALG_ReturnKey)
            Case Else
            
        End Select
        'Call PLD2PLD(ReturnPLD, MeasToTestlimit)
        'Call PLD2PLD(ReturnPLD, PLD2(0))
        'MeasToTestlimit(TestSeqNum) = ReturnPLD.Copy
    'Next i
End Function

Public Function AddDSPWf2Dic(keyname As String, ByRef obj As DSPWave)
    keyname = UCase(keyname)
    If DACTargetStr.Exists(keyname) Then
        DACTargetStr.Remove (keyname)
    End If
    DACTargetStr.Add keyname, obj
End Function
Public Function Trim_DiscriminateMeasureV_TestCondition(MeasureV_pin As PinList) As Long
Dim PinName As Variant
Dim PinArray() As String
Dim NewMeasureV_Pin As String
Dim JobName, NJobName As String
    JobName = "CP:": NJobName = "FT:"
    If InStr(1, MeasureV_pin, ":") = 0 Then Exit Function
    If TheExec.CurrentJob Like "FT*" Then JobName = "FT:": NJobName = "CP:"
    PinArray = Split(MeasureV_pin, ",")
    For Each PinName In PinArray
        If InStr(1, PinName, JobName) > 0 Then
            PinName = ReplStr(CStr(PinName), CStr(JobName), vbNullString)
        ElseIf InStr(1, PinName, NJobName) > 0 Then
            PinName = vbNullString
        End If
        If NewMeasureV_Pin = "" Then
            NewMeasureV_Pin = PinName
        Else
            NewMeasureV_Pin = NewMeasureV_Pin & "," & PinName
        End If
    Next PinName
    MeasureV_pin = NewMeasureV_Pin
End Function
Public Function ReplStr(InputStr As String, FindStr As String, ReplaceStr As String, Optional StartIndex As Integer = 1, Optional Count As Integer = -1) As String
    Dim TTempstr As String
    TTempstr = mid(InputStr, 1, StartIndex - 1)
    ReplStr = TTempstr & Replace(InputStr, FindStr, ReplaceStr, StartIndex, Count)
End Function
Public Function EqMapping(Eq As String, RegSrc As String)
    Dim EqArr() As String
    Dim RegSrcArr() As String
    Dim NewEq As String
    Dim EqIndex As Variant
    Dim TempArr() As String
    Dim TempStr As String
    EqArr = Split(Eq, "+")
    RegSrcArr = Split(RegSrc, ";")
    For Each EqIndex In EqArr
        If InStr(1, RegSrc, CStr(EqIndex) & "=") > 0 Then
            TempArr = Filter(RegSrcArr, CStr(EqIndex) & "=")
            TempStr = ReplStr(TempArr(0), CStr(EqIndex) & "=", vbNullString)
            If InStr(1, TempStr, "&") = 0 Then
                If BinStr2Dec(TempStr) = "OverFlow" Then
                    EqIndex = TempStr
                End If
            End If
        End If
        If NewEq = "" Then
            NewEq = Trim(EqIndex)
        Else
            NewEq = NewEq & "+" & Trim(EqIndex)
        End If
    Next EqIndex
    Eq = NewEq
    Debug.Print Eq
End Function

Public Function TrimReturn(q As Integer, SeqCount As Integer, TrimMethod As SearchMethod) As Boolean
    TrimReturn = False
    Select Case TrimMethod
        Case Transitions
            If q < SeqCount Then
                If SearchDone.All(True) = True Then
                    q = SeqCount
                Else
                    TrimReturn = True
                End If
            End If
        Case Else
            If q < SeqCount Then
                TheExec.sites.Selected = True
            End If
    End Select
End Function

Public Function ArrayCheck(arr() As PinListData) As Boolean
Dim ArraySize As Long
On Error GoTo ErrorHandler
    ArraySize = UBound(arr)
    If ArraySize = -1 Then ArrayCheck = False: Exit Function
    ArrayCheck = True
    Exit Function
ErrorHandler:
    ArrayCheck = False
End Function
Public Function BinStr2Dec(binstr As String) As Variant
Dim BinWidth As Long
Dim DecVal As Long
Dim i As Integer
'110101 = 53'
'111...111(32bit) = overflow'
'A11,211 = overflow'
    BinWidth = CLng(Len(binstr)) - 1
    If BinWidth >= 31 Then BinStr2Dec = "OverFlow": Exit Function
    If BinWidth < 0 Then Exit Function
    If left(binstr, 1) <> "1" And left(binstr, 1) <> "0" Then BinStr2Dec = "OverFlow": Exit Function
    If BinStr2Dec <> "Overflow" Then BinStr2Dec = BinStr2Dec(right(binstr, BinWidth)) + CLng(left(binstr, 1)) * (2 ^ (BinWidth))
End Function

Public Function IO_Power_Split(TestPinArrayIV() As String, TestSeqNumIdx As Long, TempArr3() As String, TempStr As String) As Double
    If InStr(TestPinArrayIV(TestSeqNumIdx), ":") > 0 Then
        'Dim TempStr As String
        Dim TempArr1() As String
        Dim TempArr2() As String
        Dim TempArr4() As String
        Dim TempStrPin() As String
        Dim PinCount As Long
        Dim index As Variant
        Dim j As Integer
        ReDim TempArr3(0) As String
        TempStr = TestPinArrayIV(TestSeqNumIdx)
        TempArr2 = Split(TestPinArrayIV(TestSeqNumIdx), ",")
        TestPinArrayIV(TestSeqNumIdx) = vbNullString
        For Each index In TempArr2
            TempStrPin = Split(index, ":")
            If TestPinArrayIV(TestSeqNumIdx) <> "" Then
                TestPinArrayIV(TestSeqNumIdx) = TestPinArrayIV(TestSeqNumIdx) + "," + TempStrPin(0)
            Else
                TestPinArrayIV(TestSeqNumIdx) = TempStrPin(0)
            End If
            Call TheExec.DataManager.DecomposePinList(TempStrPin(0), TempArr4, PinCount)
            For j = 0 To UBound(TempArr4)
                TempArr3(UBound(TempArr3)) = TempArr4(j) + ":" + TempStrPin(1)
                ReDim Preserve TempArr3(UBound(TempArr3) + 1)
            Next j
        Next index
    End If
End Function

Public Function PATT_ExculdePath(Pat As Variant) As String
Dim patt_ary_temp() As String
    patt_ary_temp = Split(Pat, "\")
    PATT_ExculdePath = patt_ary_temp(UBound(patt_ary_temp))

End Function

Public Function Decide_Measure_Pin(TestSeqNum As Integer, MeasPinAry() As String, ByRef Measure_Pin As PinList, Optional k As Long)
    Dim TestSeqNumIdx As Integer
    TestSeqNumIdx = TestSeqNum
    
    If TestSeqNum > 0 Then
        If UBound(MeasPinAry) = 0 Then TestSeqNumIdx = 0
    End If
    If UBound(MeasPinAry) >= 0 Then
''        If InStr(LCase(MeasPinAry(TestSeqNumIdx)), "idx") <> 0 Then TestSeqNumIdx = Int(Mid(MeasPinAry(TestSeqNumIdx), 4, 1))
        Measure_Pin = MeasPinAry(TestSeqNumIdx)
    End If
End Function

Public Function Decide_MeasureI_CurrentRange(TestSeqNum As Integer, MeasPinAry_IRange() As String, ByRef MeasureI_CurrentRange As String, Optional k As Long)
    Dim TestSeqNumIdx As Integer
    TestSeqNumIdx = TestSeqNum
    
    If TestSeqNum > 0 Then
        If UBound(MeasPinAry_IRange) = 0 Then TestSeqNumIdx = 0
    End If
    If UBound(MeasPinAry_IRange) >= 0 Then
        '' 20150605 - Check with CC
''        If InStr(LCase(MeasPinAry_IRange(TestSeqNumIdx)), "idx") <> 0 Then TestSeqNumIdx = Int(Mid(MeasPinAry_IRange(TestSeqNumIdx), 4, 1))
        If InStr(MeasPinAry_IRange(TestSeqNumIdx), ":") <> 0 Then
            If (UBound(Split(MeasPinAry_IRange(TestSeqNumIdx), ":")) >= (k - 1)) Then
                MeasureI_CurrentRange = Split(MeasPinAry_IRange(TestSeqNumIdx), ":")(k - 1)
            Else
                MeasureI_CurrentRange = Split(MeasPinAry_IRange(TestSeqNumIdx), ":")(0)
            End If
        Else
        MeasureI_CurrentRange = MeasPinAry_IRange(TestSeqNumIdx)
    End If
        
    End If
End Function

Public Function PinNameDuplicateCheck(PinName As String) As String
    Dim i As Integer
    Dim j As Integer
    Dim Pin_SplitArrayNum As Integer
    Dim Pin_SplitArray() As String
    Dim PinStringTemp As String
    Dim PinStringAfterCheck As String

    
    Pin_SplitArray = Split(PinName, ",")
    Pin_SplitArrayNum = UBound(Pin_SplitArray)
    ReDim b_FlagDuplicate(Pin_SplitArrayNum) As Boolean
    Dim b_FlagFirstTime As Boolean
    b_FlagFirstTime = False
    
    For i = 0 To Pin_SplitArrayNum
        PinStringTemp = Pin_SplitArray(i)
        For j = i + 1 To Pin_SplitArrayNum
            If Pin_SplitArray(j) = PinStringTemp Then
                b_FlagDuplicate(j) = True
            End If
        Next j
    Next i
    
    For i = 0 To Pin_SplitArrayNum
        If b_FlagDuplicate(i) = False Then
            If b_FlagFirstTime = False Then
                PinStringAfterCheck = Pin_SplitArray(i)
                b_FlagFirstTime = True
            Else
                PinStringAfterCheck = PinStringAfterCheck & "," & Pin_SplitArray(i)
            End If
            
        End If
    Next i
    PinNameDuplicateCheck = PinStringAfterCheck
End Function

Public Function IPF_CZ_PrintFreq(argc As Integer, argv() As String) As Long

    '' 20151114 - Print Freq measurement during shmoo
    Dim site As Variant
    Dim i As Long
    Dim X_SetupName As String
    Dim Y_SetupName As String
    Dim Volt_pointval As Double
    Dim FRC_pointval As Double

''    If UCase(argv(2)) = UCase("PrintFreq") Then
        X_SetupName = TheExec.DevChar.Setups.item(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes.item(tlDevCharShmooAxis_X).StepName
        Y_SetupName = TheExec.DevChar.Setups.item(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes.item(tlDevCharShmooAxis_Y).StepName
        For Each site In TheExec.sites
            Volt_pointval = TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_X).value
            FRC_pointval = TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_Y).value
            If gl_Disable_HIP_debug_log = False Then
                For i = 0 To G_MeasFreqForCZ.Pins.Count - 1
                    TheExec.Datalog.WriteComment ("Site = " & site & ",  " & X_SetupName & "=" & Volt_pointval & " V,  " & Y_SetupName & "=" & FRC_pointval & "Hz,  Pin name = " & G_MeasFreqForCZ.Pins(i) & " , Frequency value is " & G_MeasFreqForCZ.Pins(i).value(site))
                Next i
            End If
        Next site
''    End If
End Function
Public Function SrcVoltFromFlowForLoop(FlowForLoopIntegerName As String, powerPin As PinList, StartStopStep As String) As Long
    
    
    Dim StepIndex As Long
    Dim i As Long
    
    StepIndex = TheExec.Flow.var(FlowForLoopIntegerName).value
    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Get index = " & StepIndex)
    
    Dim StartVolt As Double
    Dim StopVolt As Double
    Dim StepVolt As Double
    Dim ForceVolt As Double
    
    Dim DecomposeString() As String
    
    DecomposeString = Split(StartStopStep, ",")
    StartVolt = CDbl(DecomposeString(0))
    StopVolt = CDbl(DecomposeString(1))
    StepVolt = CDbl(DecomposeString(2))
    
    ForceVolt = StartVolt + StepIndex * StepVolt
    
    If ForceVolt > StopVolt Then
        TheExec.Datalog.WriteComment ("Warning !! Force Volt over Stop Volt")
        Exit Function
    End If
    
    Dim InstName As String
    Dim Pins() As String
    Dim NumberPins As Long
    Call TheExec.DataManager.DecomposePinList(powerPin, Pins(), NumberPins)
     
    InstName = GetInstrument(Pins(0), 0)

    Select Case InstName
        Case "DC-07"
            TheHdw.DCVI.Pins(powerPin).Voltage = ForceVolt
        
        Case "VHDVS"
            TheHdw.DCVS.Pins(powerPin).Voltage.value = ForceVolt
            
        Case "HexVS", "VSM"
            TheHdw.DCVS.Pins(powerPin).Voltage.value = ForceVolt
            
        Case Else
        
    End Select
    
    TheHdw.Wait (1 * ms)
    
    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment (powerPin.value & " output voltage = " & ForceVolt)


End Function

Public Function Rev_BinArray(m_binarr() As Long) As String
    Dim i As Long
    Dim SrcBitBinaryString  As String
    '' 20150811 - Reverse order of binary string
    For i = UBound(m_binarr) To 0 Step -1
        If i = UBound(m_binarr) Then
            SrcBitBinaryString = m_binarr(i)
        Else
            SrcBitBinaryString = SrcBitBinaryString & m_binarr(i)
        End If
        Rev_BinArray = SrcBitBinaryString
    Next i
End Function

Public Function PLL_calibration_calc(ByRef calib_code() As DSPWave, CUS_Str_MainProgram As String) As Long


Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim temp1_dict As New DSPWave
Dim temp2_dict As New DSPWave
Dim calc_data() As New SiteDouble
Dim temp_testname_bin As Long
Dim temp_testname_dec As Long
Dim testname_str() As String
Dim delta_value() As New SiteDouble
Dim target_var As New SiteDouble
Dim temp_delta_value As Integer
Dim temp_cal_code As New DSPWave
Dim RefPLL_calibration_code As New DSPWave
Dim TXPLL_calibration_code As New DSPWave
Dim DPTXPLL_calibration_code As New DSPWave
Dim calibration_target() As String
Dim calibration_target_value As Long

Dim OutputTname_format() As String
Dim TestNameInput As String
Dim site As Variant
calibration_target = Split(CUS_Str_MainProgram, "_")
calibration_target_value = CLng(calibration_target(4))


ReDim testname_str(5)
ReDim calc_data(31)
ReDim delta_value(31)

        ''''calc and print in datalog
        temp1_dict.CreateConstant 0, 1, DspLong
        temp2_dict.CreateConstant 0, 1, DspLong
        RefPLL_calibration_code.CreateConstant 0, 5, DspLong
        temp_cal_code.CreateConstant 0, 1, DspLong
        
            For i = 0 To (UBound(calib_code()) + 1) / 2 - 1
                    For Each site In TheExec.sites
                        temp1_dict.Element(0) = calib_code(i).Element(0)
                        temp2_dict.Element(0) = calib_code(i + 32).Element(0)
                        calc_data(i) = (temp2_dict.Element(0) + temp1_dict.Element(0)) / 2
                    Next site
    
                '''''dec to bin testname
                    temp_testname_dec = i
                        For j = 0 To 4
                          temp_testname_bin = temp_testname_dec Mod 2
                          temp_testname_dec = Fix(temp_testname_dec / 2)
                          testname_str(j) = CStr(temp_testname_bin)
                        Next j
                        testname_str(5) = testname_str(4) & testname_str(3) & testname_str(2) & testname_str(1) & testname_str(0)
                    
                        TestNameInput = Report_TName_From_Instance("C", vbNullString, "F_" & testname_str(5), i, 0)
                
                If gl_Disable_HIP_debug_log = False Then
                    TheExec.Flow.TestLimit resultVal:=calc_data(i), Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling, formatStr:="%.1f" 'eng_forceflow_transfer
                End If
            Next i

            '''' compare the target
            For Each site In TheExec.sites
             temp_delta_value = 5000
                 For k = 0 To (UBound(calib_code()) + 1) / 2 - 1
                    If UCase(glb_TestInstance) Like "*PCIEREFPLL*" Then
                    delta_value(k) = Abs(calibration_target_value - calc_data(k))
                    ElseIf UCase(glb_TestInstance) Like "*PCIETXPLL*" Then
                    delta_value(k) = Abs(calibration_target_value - calc_data(k))
                    ElseIf UCase(glb_TestInstance) Like "*DPTXPLL*" Then
                    delta_value(k) = Abs(calibration_target_value - calc_data(k))
                    End If
                'search min delta
                    If delta_value(k) < temp_delta_value Then
                        temp_delta_value = delta_value(k)
                        target_var = k
                    End If
                 Next k
             Next site
             
             
            If UCase(glb_TestInstance) Like "*PCIEREFPLL*" Then
                
                TestNameInput = Report_TName_From_Instance("C", vbNullString, "PCIE_REFPLL_Fcal_code", 0, 0)
                TheExec.Flow.TestLimit resultVal:=target_var, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
                
                 'store to dictionary
                    For Each site In TheExec.sites
                    temp_cal_code.Element(0) = target_var
                    Next site
                 Call HardIP_Dec2Bin(RefPLL_calibration_code, temp_cal_code, 5)
                 Call AddStoredCaptureData("PCIE_REFPLL_FCAL_BYPASS", RefPLL_calibration_code)
                 
            ElseIf UCase(glb_TestInstance) Like "*PCIETXPLL*" Then
            
                TestNameInput = Report_TName_From_Instance("C", vbNullString, "PCIE_TXPLL_Fcal_code", 0, 0)
                
                TheExec.Flow.TestLimit resultVal:=target_var, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
                
                 'store to dictionary
                    For Each site In TheExec.sites
                    temp_cal_code.Element(0) = target_var
                    Next site
                 Call HardIP_Dec2Bin(TXPLL_calibration_code, temp_cal_code, 5)
                 Call AddStoredCaptureData("PCIE_TXPLL_FCAL_BYPASS", TXPLL_calibration_code)
                 
            ElseIf UCase(glb_TestInstance) Like "*DPTXPLL*" Then
            
                TestNameInput = Report_TName_From_Instance("C", vbNullString, "DP_TXPLL_Fcal_code", 0, 0)
                TheExec.Flow.TestLimit resultVal:=target_var, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
                
                 'store to dictionary
                    For Each site In TheExec.sites
                    temp_cal_code.Element(0) = target_var
                    Next site
                 Call HardIP_Dec2Bin(DPTXPLL_calibration_code, temp_cal_code, 5)
                 Call AddStoredCaptureData("DPTX_PCIEPLL_FCAL_BYPASS", DPTXPLL_calibration_code)
            End If
   
End Function
    
Public Function VFI_ProcessWaitTimeString(MeasI_WaitTime As String, MeasF_WaitTime As String, UVI80_MeasV_WaitTime As String, _
                                        ByRef MeasI_WaitTime_Ary() As String, ByRef MeasF_WaitTime_Ary() As String, ByRef UVI80_MeasV_WaitTime_Ary() As String, TestSequenceArray() As String)
                                        

Dim i As Long
Dim Diff As Long
'If InStr(MeasI_WaitTime, "+") Then
    MeasI_WaitTime_Ary = Split(MeasI_WaitTime, "+")
    If UBound(TestSequenceArray) - UBound(MeasI_WaitTime_Ary) > 0 Then
        For i = 0 To UBound(TestSequenceArray) - UBound(MeasI_WaitTime_Ary) - 1
            If UBound(MeasI_WaitTime_Ary) = -1 Then
                MeasI_WaitTime = MeasI_WaitTime & "+" & vbNullString
            Else
                MeasI_WaitTime = MeasI_WaitTime & "+" & MeasI_WaitTime_Ary(0)
            End If
        Next i
        MeasI_WaitTime_Ary = Split(MeasI_WaitTime, "+")
    End If
'End If

'If InStr(MeasF_WaitTime, "+") Then
    MeasF_WaitTime_Ary = Split(MeasF_WaitTime, "+")
    If UBound(TestSequenceArray) - UBound(MeasF_WaitTime_Ary) > 0 Then
        For i = 0 To UBound(TestSequenceArray) - UBound(MeasF_WaitTime_Ary) - 1
            If UBound(MeasF_WaitTime_Ary) = -1 Then
                MeasF_WaitTime = MeasF_WaitTime & "+" & vbNullString
            Else
                MeasF_WaitTime = MeasF_WaitTime & "+" & MeasF_WaitTime_Ary(0)
            End If
        Next i
        MeasF_WaitTime_Ary = Split(MeasF_WaitTime, "+")
    End If
'End If

'If InStr(UVI80_MeasV_WaitTime, "+") Then
    UVI80_MeasV_WaitTime_Ary = Split(UVI80_MeasV_WaitTime, "+")
    If UBound(TestSequenceArray) - UBound(UVI80_MeasV_WaitTime_Ary) > 0 Then
        For i = 0 To UBound(TestSequenceArray) - UBound(UVI80_MeasV_WaitTime_Ary) - 1
            If UBound(UVI80_MeasV_WaitTime_Ary) = -1 Then
                UVI80_MeasV_WaitTime = UVI80_MeasV_WaitTime & "+" & vbNullString
            Else
                UVI80_MeasV_WaitTime = UVI80_MeasV_WaitTime & "+" & UVI80_MeasV_WaitTime_Ary(0)
            End If
        Next i
        UVI80_MeasV_WaitTime_Ary = Split(UVI80_MeasV_WaitTime, "+")
    End If
'End If








End Function



Public Function VFI_AnalyzedInputStringByAt(ByRef MeasV_PinS As String, ByRef MeasF_PinS_SingleEnd As String, ByRef MeasI_pinS As String, ByRef MeasI_Range As String, ByRef MeasF_PinS_Differential As String, _
    ByRef ForceV_Val As String, ByRef ForceI_Val As String) As Long
    '' 20160201 - Check input argumenets whether have "@" in the first character. Add it If no "@" in the beginning. Then remove it to process fomat.
    '' The purpose is to cover import issue. Ex:++
    
    'Call CheckInputStringByAt(MeasV_PinS)
    'Call CheckInputStringByAt(MeasF_PinS_SingleEnd)
    'Call CheckInputStringByAt(MeasI_pinS)
    'Call CheckInputStringByAt(MeasI_Range)
    'Call CheckInputStringByAt(MeasF_PinS_Differential)
    
    'Call CheckInputStringByAt(ForceV_Val)
    'Call CheckInputStringByAt(ForceI_Val)
    
    
    MeasV_PinS = Replace(MeasV_PinS, "@", vbNullString)
    MeasF_PinS_SingleEnd = Replace(MeasF_PinS_SingleEnd, "@", vbNullString)
    MeasI_pinS = Replace(MeasI_pinS, "@", vbNullString)
    MeasI_Range = Replace(MeasI_Range, "@", vbNullString)
    MeasF_PinS_Differential = Replace(MeasF_PinS_Differential, "@", vbNullString)
    
    ForceV_Val = Replace(ForceV_Val, "@", vbNullString)
    ForceI_Val = Replace(ForceI_Val, "@", vbNullString)
    
End Function

Public Function HardIP_Alarm_off()
    If TheExec.Flow.enableWord("HardIPAlarm") = True Then
        Dim i As Long, j As Long, p As Long
    
         Dim PinAry() As String, pinCnt As Long
    
 '   If TheExec.Flow.EnableWord("HardIPAlarm") = True Then
    
         TheExec.DataManager.DecomposePinList "All_UVS256,VDD_Warm", PinAry(), pinCnt
            For i = 0 To pinCnt - 1
            
         For Each site In TheExec.sites
            TheHdw.DCVS.Pins(PinAry(i)).Alarm(tlDCVSAlarmOpenKelvinDUT) = tlAlarmOff
            TheHdw.DCVS.Pins(PinAry(i)).Alarm(tlDCVSAlarmFoldCurrentLimitTimeout) = tlAlarmOff
            TheHdw.DCVS.Pins(PinAry(i)).Alarm(tlDCVSAlarmSourceFoldCurrentLimitTimeout) = tlAlarmOff
       Next site
       Next i
       
              TheExec.DataManager.DecomposePinList "VDDIO18_AOP,ANALOGMUX_OUT", PinAry(), pinCnt
            For i = 0 To pinCnt - 1
            
        ' For Each Site In TheExec.sites
            TheHdw.DCVI.Pins(PinAry(i)).Alarm(tlDCVSAlarmOpenKelvinDUT) = tlAlarmOff
            TheHdw.DCVI.Pins(PinAry(i)).Alarm(tlDCVIAlarmOpenKelvin) = tlAlarmOff
            TheHdw.DCVI.Pins(PinAry(i)).Alarm(tlDCVIAlarmDGS) = tlAlarmOff
           ' TheHdw.DCVI.Pins(PinAry(i)).Alarm(tlDCVIAlarmGuard) = tlAlarmOff
            TheHdw.DCVI.Pins(PinAry(i)).Alarm(tlDCVIAlarmMode) = tlAlarmOff
      ' Next Site
       Next i
    TheHdw.DIB.LeavePowerOn = False
        If glb_TesterType = "Jaguar" Then
            TheHdw.DCVI.Pins("All_DCVI").Alarm(tlDCVIAlarmDGS) = tlAlarmOff
        End If
    End If
End Function



Public Function Freq_WalkingStrobe_Meas_VOD_Diff(MeasureF_Pin_Differential As PinList, Optional MeasF_WalkingStrobe_StartV As Double, Optional MeasF_WalkingStrobe_EndV As Double, _
    Optional MeasF_WalkingStrobe_StepVoltage As Double, Optional MeasF_WalkingStrobe_BothVohVolDiffV As Double, _
    Optional MeasF_WalkingStrobe_interval As Double, Optional MeasF_WalkingStrobe_miniFreq As Double)
    
    Dim site As Variant
    Dim MeasF_WalkingStrobe_Step As Long
    MeasF_WalkingStrobe_Step = (MeasF_WalkingStrobe_EndV - MeasF_WalkingStrobe_StartV) / MeasF_WalkingStrobe_StepVoltage + 1
    
    Dim MeasFreq_WKStrobe() As New PinListData
    ReDim MeasFreq_WKStrobe(MeasF_WalkingStrobe_Step) As New PinListData
    Dim WalkStrobe_i As Long
    Dim WalkStrobe_j As Long
    ''setup and measure Freq base on VOL and VOH setting.
    Dim WalkingStrobe_stepV As Double
    WalkingStrobe_stepV = (MeasF_WalkingStrobe_EndV - MeasF_WalkingStrobe_StartV) / MeasF_WalkingStrobe_Step
    
    Dim DiffPinGroup() As String
    Dim Pin_Ary() As String
    Dim Pin_Cnt As Long
    Dim i As Long, j As Long, k As Long
    DiffPinGroup = Split(MeasureF_Pin_Differential, ",")
    
    Dim DiffPinGroupPinList As New PinList
    
    Dim MeasurePin As String
    Dim MeasurePin_Opposite As String
    Dim FreqAccessPin As String
    Dim Record_Final_Mid_VOD() As New SiteDouble
    Dim b_UpdateVOD_Flag() As New SiteBoolean
    Dim Val_UpdateToVt As Double
    Dim Default_VOD As Double
    
    
    For i = 0 To UBound(DiffPinGroup)
        TheExec.DataManager.DecomposePinList DiffPinGroup(i), Pin_Ary, Pin_Cnt
        DiffPinGroupPinList.value = DiffPinGroup(i)
        Default_VOD = TheHdw.Digital.Pins(DiffPinGroupPinList).DifferentialLevels.value(chVod)
        ReDim Record_Final_Mid_VOD(Pin_Cnt - 1) As New SiteDouble
        ReDim b_UpdateVOD_Flag(Pin_Cnt - 1) As New SiteBoolean

        For j = 0 To Pin_Cnt - 1
            
            MeasurePin = Pin_Ary(j)
            If InStr(UCase(MeasurePin), "_P") <> 0 Then
                MeasurePin_Opposite = Replace(UCase(MeasurePin), "_P", "_N")
            ElseIf InStr(UCase(MeasurePin), "_N") <> 0 Then
                MeasurePin_Opposite = Replace(UCase(MeasurePin), "_N", "_P")
            End If
            
            If InStr(UCase(MeasurePin), "_N") <> 0 Then
                FreqAccessPin = Replace(UCase(MeasurePin), "_N", "_P")
            Else
                FreqAccessPin = MeasurePin
            End If
            
            TheHdw.Digital.Pins(MeasurePin_Opposite).Disconnect
            TheHdw.Wait 10 * us
            With TheHdw.PPMU.Pins(MeasurePin_Opposite)
                .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_Max_InitialValue_FI_Range
                .Connect
                .Gate = tlOn
                .ForceV 0, 0
            End With
            
            For WalkStrobe_i = 0 To MeasF_WalkingStrobe_Step
                TheHdw.Digital.Pins(DiffPinGroupPinList).DifferentialLevels.value(chVod) = MeasF_WalkingStrobe_StartV + WalkStrobe_i * WalkingStrobe_stepV
                
                Call Freq_MeasFreqSetup(DiffPinGroupPinList, MeasF_WalkingStrobe_interval, VOH)
                Call HardIP_Freq_MeasFreqStart(DiffPinGroupPinList, MeasF_WalkingStrobe_interval, MeasFreq_WKStrobe(WalkStrobe_i), 0)
            Next WalkStrobe_i
                
            ''analyze measurement data to decide which VOH/VOL level shiuld be used for measurement.
            Dim Record_Temp_VOD As Double
            Dim Record_Min_VOD As Double
            Dim Record_Max_VOD As Double
            Dim Record_Mid_VOD As Double
            
            For Each site In TheExec.sites
            
                Record_Min_VOD = 9999
                Record_Max_VOD = -9999
                For WalkStrobe_i = 0 To MeasF_WalkingStrobe_Step
                    If MeasFreq_WKStrobe(WalkStrobe_i).Pins(FreqAccessPin).value(site) > MeasF_WalkingStrobe_miniFreq Then
                        Record_Temp_VOD = MeasF_WalkingStrobe_StartV + WalkStrobe_i * WalkingStrobe_stepV
                        If Record_Temp_VOD > Record_Max_VOD Then Record_Max_VOD = Record_Temp_VOD
                        If Record_Temp_VOD < Record_Min_VOD Then Record_Min_VOD = Record_Temp_VOD
                    End If
                Next WalkStrobe_i
                
                If TheExec.TesterMode = testModeOffline Then
                    Record_Max_VOD = 0.5 + i * 0.1 + j * 0.1 + site * 0.01
                    Record_Min_VOD = 0.5 - i * 0.1 - j * 0.1 - site * 0.02
                End If
                
                If Record_Min_VOD <> 9999 Then
                    Record_Mid_VOD = (Record_Max_VOD + Record_Min_VOD) / 2
                    Record_Final_Mid_VOD(j).value(site) = Record_Mid_VOD
                    b_UpdateVOD_Flag(j)(site) = True
                    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site= " & site & " , " & " Pin= " & MeasurePin & " , " & " Record VOD = " & Record_Mid_VOD & " V"
                Else
                    b_UpdateVOD_Flag(j)(site) = False
                    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site= " & site & " , " & " Pin= " & MeasurePin & " , " & " Record VOD = Default, search fail"
                End If
            Next site
            
            TheHdw.PPMU.Pins(MeasurePin_Opposite).Disconnect
            TheHdw.Digital.Pins(MeasurePin_Opposite).Connect
        Next j
        
        For Each site In TheExec.sites
            If b_UpdateVOD_Flag(0)(site) = True And b_UpdateVOD_Flag(1)(site) = True Then
                Val_UpdateToVt = (Record_Final_Mid_VOD(0).value(site) + Record_Final_Mid_VOD(1).value(site)) / 2
                TheHdw.Digital.Pins(DiffPinGroupPinList).DifferentialLevels.value(chDiff_Vt) = Val_UpdateToVt
                TheHdw.Digital.Pins(DiffPinGroupPinList).DifferentialLevels.value(chVod) = Default_VOD
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site= " & site & " , " & " Pin= " & DiffPinGroupPinList & " , " & " Update Differential Vt = " & Val_UpdateToVt & " V"
            ElseIf gl_Disable_HIP_debug_log = False Then
                TheExec.Datalog.WriteComment "Site= " & site & " , " & " Pin= " & DiffPinGroupPinList & " , " & " Vt = Default, search fail"
            End If
        
        Next site
        
    Next i
End Function


Public Function UP1600_PPMU_Measure_R_SE(MeasurePin As String, ForceVoltStr As String, MeasureCurrRange As Double, Optional RAK_Flag As Enum_RAK = 0, _
                                                                                 Optional ByRef RTN_Imped_Val As PinListData, Optional b_PD_Mode As Boolean = True) As Long

    Dim MeasureValue As New PinListData
    Dim Imped As New PinListData
    Dim pin  As Variant
    Dim site As Variant

    
    Dim ForceVoltVal As Double
    ForceVoltVal = CDbl(ForceVoltStr)
    
    TheHdw.Digital.Pins(MeasurePin).Disconnect
    TheHdw.Wait 10 * us
    
    '' Initial force I to 0 and force V by your specified
    With TheHdw.PPMU.Pins(MeasurePin)
        .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_Max_InitialValue_FI_Range
        .Connect
        .Gate = tlOn
        .ForceV ForceVoltVal, MeasureCurrRange
    End With
    
    TheHdw.Wait 1 * ms
    
    DebugPrintFunc_PPMU CStr(MeasurePin)
    
    MeasureValue = TheHdw.PPMU.Pins(MeasurePin).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
    
    '' Avoid divide 0
    If TheExec.TesterMode = testModeOffline Then
        For Each site In TheExec.sites
            For Each pin In MeasureValue.Pins
                If MeasureValue.Pins(pin).value(site) = 0 Then
                    MeasureValue.Pins(pin).value(site) = 1
                End If
            Next pin
        Next site
    End If
    
    For Each pin In MeasureValue.Pins
        For Each site In TheExec.sites
            If MeasureValue.Pins(pin).value(site) = 0 Then
                MeasureValue.Pins(pin).value(site) = 0.000000000001
            End If
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & ", Pin : " & pin & ", Measure Current = " & MeasureValue.Pins(pin).value(site))
        Next site
    Next pin
 
    '' Print force condition
    Call Print_Force_Condition("r", MeasureValue)
    
    Dim PowerVal As Double
    '' Impedence measurement
    If b_PD_Mode Then
        Imped = MeasureValue.Math.Invert.Multiply(ForceVoltVal).Abs
    Else
        PowerVal = TheHdw.DCVS.Pins("VDDQL_DDR0").Voltage.value
        Imped = MeasureValue.Math.Invert.Multiply(PowerVal - ForceVoltVal).Abs
    End If
    
    If TheExec.TesterMode = testModeOffline Then
        Call SimulateOutputImped(MeasurePin, Imped)
    Else
        If RAK_Flag = R_PathWithContact Then
            '' Compensate resistance after Kelvin for path resistance considerations
            For Each pin In Imped.Pins
                For Each site In TheExec.sites
                    Imped.Pins.item(pin).value(site) = Imped.Pins.item(pin).value(site) - R_Path_PLD.Pins.item(pin).value(site)
                Next site
            Next pin
        End If
    End If
    
    TheHdw.PPMU.Pins(MeasurePin).Disconnect
    TheHdw.Digital.Pins(MeasurePin).Connect
    
    RTN_Imped_Val = Imped
    
End Function

Public Function SubMeasR(CPUA_Flag_In_Pat As Boolean, pin As String, ForceVolt As String, ByRef RTN_Imped_Val As PinListData, Optional b_IsDifferential As Boolean, _
                                              Optional b_PD_Mode As Boolean = True)

    If (CPUA_Flag_In_Pat) Then
        Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
    Else
        Call TheHdw.Digital.Patgen.HaltWait
    End If
    
    Dim Diff_P_Pin As String, Diff_N_Pin As String
    Dim P_Pin_ForceV As Double, N_Pin_ForceV As Double
    Dim SplitForceVolt() As String
    
    If b_IsDifferential = False Then
        Call UP1600_PPMU_Measure_R_SE(pin, ForceVolt, 50 * mA, R_PathWithContact, RTN_Imped_Val, b_PD_Mode)
    Else
        Diff_P_Pin = UCase(pin)
        Diff_N_Pin = Replace(UCase(Diff_P_Pin), "_P", "_N")
        SplitForceVolt = Split(ForceVolt, ",")
        P_Pin_ForceV = CDbl(SplitForceVolt(0))
        N_Pin_ForceV = CDbl(SplitForceVolt(1))
        
        Call UP1600_PPMU_Measure_R_DI(Diff_P_Pin, Diff_N_Pin, P_Pin_ForceV, N_Pin_ForceV, 50 * mA, R_PathWithContact, RTN_Imped_Val)
    End If
    
    If (CPUA_Flag_In_Pat) Then
        Call TheHdw.Digital.Patgen.Continue(0, cpuA)
    Else
        TheHdw.Digital.Patgen.HaltWait
    End If
    
    TheHdw.Digital.Patgen.HaltWait

End Function

Public Function SimulateOutputImped(MeasureR_Pin As String, ByRef MeasureImped As PinListData) As Long
    Dim site As Variant
    For Each site In TheExec.sites.Active
        If site = 0 Then
            MeasureImped.Pins(MeasureR_Pin).value(site) = 46
        ElseIf site = 1 Then
            MeasureImped.Pins(MeasureR_Pin).value(site) = 53
        ElseIf site = 2 Then
''            MeasureImped.Pins(MeasureR_Pin).Value(Site) = 49
        ElseIf site = 3 Then
''            MeasureImped.Pins(MeasureR_Pin).Value(Site) = 52
        End If
    Next site
End Function

Public Function UP1600_PPMU_Measure_R_DI(Measure_P_Pin As String, Measure_N_Pin As String, P_ForceVolt As Double, N_ForceVolt As Double, _
MeasureCurrRange As Double, Optional RAK_Flag As Enum_RAK = 0, Optional ByRef RTN_Imped_Val As PinListData) As Long

    Dim MeasureValue As New PinListData
    Dim Imped As New PinListData
    Dim pin  As Variant
    Dim site As Variant
    
    TheHdw.Digital.Pins(Measure_P_Pin & "," & Measure_N_Pin).Disconnect
    TheHdw.Wait 10 * us
    
    '' Initial force I to 0 and force V by your specified
    With TheHdw.PPMU.Pins(Measure_P_Pin)
        .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_Max_InitialValue_FI_Range
        .Connect
        .Gate = tlOn
        .ForceV P_ForceVolt, MeasureCurrRange
    End With
    
    With TheHdw.PPMU.Pins(Measure_N_Pin)
        .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_Max_InitialValue_FI_Range
        .Connect
        .Gate = tlOn
        .ForceV N_ForceVolt, MeasureCurrRange
    End With

    TheHdw.Wait 1 * ms
    
    DebugPrintFunc_PPMU CStr(Measure_P_Pin)
    
    MeasureValue = TheHdw.PPMU.Pins(Measure_P_Pin).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
    
    '' Avoid divide 0
    If TheExec.TesterMode = testModeOffline Then
        For Each site In TheExec.sites
            For Each pin In MeasureValue.Pins
                If MeasureValue.Pins(pin).value(site) = 0 Then
                    MeasureValue.Pins(pin).value(site) = 1
                End If
            Next pin
        Next site
    End If
    
    For Each pin In MeasureValue.Pins
        For Each site In TheExec.sites
            If MeasureValue.Pins(pin).value(site) = 0 Then
                MeasureValue.Pins(pin).value(site) = 0.000000000001
            End If
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & ", Pin : " & pin & ", Measure Current = " & MeasureValue.Pins(pin).value(site))
        Next site
    Next pin
 
    '' Print force condition
    Call Print_Force_Condition("r", MeasureValue)
 
    '' Impedence measurement
    Imped = MeasureValue.Math.Invert.Multiply(P_ForceVolt - N_ForceVolt)
    
    Dim RAK_Pin_N As String
    
    If TheExec.TesterMode = testModeOffline Then
        Call SimulateOutputImped(Measure_P_Pin, Imped)
    Else
        If RAK_Flag = R_PathWithContact Then
            '' Compensate resistance after Kelvin for path resistance considerations
            For Each pin In Imped.Pins
                RAK_Pin_N = Replace(UCase(pin), "_P", "_N")
                For Each site In TheExec.sites
                    Imped.Pins.item(pin).value(site) = Imped.Pins.item(pin).value(site) - R_Path_PLD.Pins.item(pin).value(site) - R_Path_PLD.Pins.item(RAK_Pin_N).value(site)
                Next site
            Next pin
        End If
    End If
    
    TheHdw.PPMU.Pins(Measure_P_Pin & "," & Measure_N_Pin).Disconnect
    TheHdw.Digital.Pins(Measure_P_Pin & "," & Measure_N_Pin).Connect
    
    RTN_Imped_Val = Imped
    
End Function


Public Function CreateSimulateMDLL_Data(argc As Integer, argv() As String) As Long
    Dim i As Long, j As Long
    Dim site As Variant
    
    Dim DSPWaveLength As Long
    If TheExec.TesterMode = testModeOffline Then
    
        DSPWaveLength = 16
        Dim SimulateDSPWaveBin() As New DSPWave
        ReDim SimulateDSPWaveBin(argc - 1) As New DSPWave
        
        
        For i = 1 To argc - 1
            SimulateDSPWaveBin(i).CreateConstant 0, DSPWaveLength, DspLong
            For Each site In TheExec.sites
                For j = 0 To DSPWaveLength - 1
                    If site = 0 Then
                        If i < 4 Then
                            SimulateDSPWaveBin(i)(site).Element(j) = 0
                        Else
                            If j < 2 Then
                                SimulateDSPWaveBin(i)(site).Element(j) = 0
                            Else
                                SimulateDSPWaveBin(i)(site).Element(j) = 1
                            End If
                        End If
                    Else
                        If i < 3 Then
                            SimulateDSPWaveBin(i)(site).Element(j) = 1
                        ElseIf i >= 3 And i <= 5 Then
                            If j < 2 Then
                                SimulateDSPWaveBin(i)(site).Element(j) = 0
                            Else
                                SimulateDSPWaveBin(i)(site).Element(j) = 1
                            End If
                        Else
                            SimulateDSPWaveBin(i)(site).Element(j) = 0
                        End If
                        
                    End If
                Next j
                
            Next site
            Call AddStoredCaptureData(argv(i), SimulateDSPWaveBin(i))
        Next i
    End If

    Dim Displaystring As String

     For Each site In TheExec.sites
            For i = 1 To argc - 1
            Displaystring = vbNullString
            For j = 0 To DSPWaveLength - 1
                If j = 0 Then
                    Displaystring = SimulateDSPWaveBin(i)(site).Element(j)
                Else
                    Displaystring = Displaystring & SimulateDSPWaveBin(i)(site).Element(j)
                End If
            Next j
           If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Simulate data Site_" & site & " Dict = " & argv(i) & " Content = " & Displaystring)
        Next i
     Next site

End Function

Public Function IPF_Connect_PPMU_ForceV(argc As Long, argv() As String)

    Dim ForceValStr(0) As String
    ForceValStr(0) = argv(1)
    Call HIP_Evaluate_ForceVal(ForceValStr())
    
    argv(0) = Replace(argv(0), "+", ",")
    TheHdw.Digital.Pins(argv(0)).Disconnect
    With TheHdw.PPMU.Pins(argv(0))
        .ForceV CDbl(ForceValStr(0)), 0.05
        .Connect
        .Gate = tlOn
    End With
                                                                                                                                                                                                                                                               
   If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Pin = " & argv(0) & ", Force V = " & argv(1) & ", Measure Current Range = " & TheHdw.PPMU.Pins(argv(0)).MeasureCurrentRange
End Function

Public Function IPF_Connect_Digital(argc As Long, argv() As String)
    argv(0) = Replace(argv(0), "+", ",")
    
    With TheHdw.PPMU.Pins(argv(0))
        .ForceV 0, 0
        .Disconnect
        .Gate = tlOff
    End With
    TheHdw.Digital.Pins(argv(0)).Connect
End Function

Public Function DictDSPToSiteLong(DictName As String, ByRef RTN_sl_Val As SiteLong, DictTrimFuseName As String) As Long
    Dim DSP_Val_Dec As New DSPWave
    Dim site As Variant
    DSP_Val_Dec = GetStoredCaptureData(DictName)
    
    Call AddStoredCaptureData(DictTrimFuseName, DSP_Val_Dec)
    
    For Each site In TheExec.sites
        RTN_sl_Val(site) = DSP_Val_Dec(site).Element(0)
    Next site
    
End Function

Public Function HardIP_Duty_Frequency(FreqMeasPins As PinList, IsDifferentialPin As Boolean, TestSeqNum As Integer, d_FreqMeasInterval As Double, _
        Optional Rtn_MeasFreq As PinListData, Optional b_TestLimitPerPin As Boolean = False, Optional b_SkipTestLimit As Boolean = True)
    
    Dim site As Variant
    Dim p As Long
    Dim MeasFreq As New PinListData
    
    Call Freq_MeasFreqSetup(FreqMeasPins, d_FreqMeasInterval)  '' 20150621 - default d_FreqMeasInterval = 0.001
    '' 20150623 - Add Customize Wait Time
    Call HardIP_Freq_MeasFreqStart(FreqMeasPins, d_FreqMeasInterval, MeasFreq)       '' 20150621 - default d_FreqMeasInterval = 0.001

    Dim TestNameInput As String
    

    TestNameInput = "Freq_Meas_" & CStr(TestSeqNum)


    If Not b_SkipTestLimit Then
        If IsDifferentialPin = True Then
            For p = 0 To MeasFreq.Pins.Count - 1 Step 2 ' freq counter result of differential pins is stored in positive pin
                TheExec.Flow.TestLimit resultVal:=MeasFreq.Pins(p + 1), unit:=unitHz, Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
            Next p
        Else
            If b_TestLimitPerPin = True Then
                For p = 0 To MeasFreq.Pins.Count - 1
                    TheExec.Flow.TestLimit resultVal:=MeasFreq.Pins(p), unit:=unitHz, Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
                Next p
            Else
                TheExec.Flow.TestLimit resultVal:=MeasFreq, unit:=unitHz, Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
            End If
        End If
    End If

        
    '' 20151224 - Merge print measured frequency during shmoo if need
    G_MeasFreqForCZ = MeasFreq
    
    ''20160906 - Return MeasFreq to main program
    Rtn_MeasFreq = MeasFreq
    
End Function

Public Function TrimUVI80_Meas_VFI(Pat As String, TestSequenceArray() As String, srcPin As PinList, code As SiteLong, _
MeasV_Pin As String, MeasValue As SiteDouble, MeasI_Pin As String, MeasureI_Range As Double, _
MeasF_PinS_SingleEnd As PinList, MeasF_Interval As String, MeasF_EventSourceWithTerminationMode As EventSourceWithTerminationMode, _
DigSrc_DataWidth As Long, DigSrc_Sample_Size As Long, DigSrc_Equation As String, digsrc_assignment As String, _
DigCap_Pin As PinList, DigCap_DataWidth As Long, DigCap_Sample_Size As Long, CUS_Str_DigCapData As String, OutDSP As DSPWave, _
TrimCodeSize As Long, Trimname As String, Meas_StoreName As String, Cal_Eqn As String, TrimCal_Name As String, CPUA_Flag_In_Pat As Boolean, Optional Final_Calc As Boolean, Optional b_Trimfinish As Boolean = False, Optional MSB_First_Flag As Boolean = False)

    On Error GoTo err
    
    Dim sigName As String, srcWave As New DSPWave, site As Variant
    
    Dim DigSrcCodeSize As Long
    Dim i As Long, j As Long
    Dim code_bin() As String
    Dim Ts As Variant
    Dim Str_FinalPatName As String
    Dim temp_assignment As String
    Dim cal As New SiteDouble
    Dim out_str() As String
    Dim MeasStoreName_Ary() As String
    Dim TestSeqNum As Long
    ReDim code_bin(TheExec.sites.Existing.Count)
    ReDim out_str(TheExec.sites.Existing.Count)
    Dim TrimCal_value As New PinListData
    Dim TrimCalCap_value As New DSPWave
    Dim TrimCal_Name_array() As String
    Dim MeasV_Pin_split() As String
    TrimCal_Name_array = Split(TrimCal_Name, ":")
    MeasV_Pin_split = Split(MeasV_Pin, "+")
    ''''''''''''''''''''''''''''''''setup store name'''''''''''''''''''''''''''''.
    MeasStoreName_Ary = Split(Meas_StoreName, "+")
    ReDim Preserve MeasStoreName_Ary(UBound(TestSequenceArray))
    Dim Rtn_Meas As New PinListData
    Dim Store_Rtn_Meas() As New PinListData
    Dim SoreMaxNum As Long
    Dim StoreIndex As Long
    ''20170123-Get how many store name in MeasStoreName_Ary
    If Meas_StoreName <> "" Then
        SoreMaxNum = 0
        For i = 0 To UBound(MeasStoreName_Ary)
            If MeasStoreName_Ary(i) <> "" Then
                SoreMaxNum = SoreMaxNum + 1
            End If
        Next i
         ReDim Store_Rtn_Meas(SoreMaxNum - 1) As New PinListData
         StoreIndex = 0
    End If
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    temp_assignment = digsrc_assignment
    
    sigName = "DSSC_Search_Code"
    TheHdw.DSSC.Pins(srcPin).Pattern(Pat).Source.Signals.Add sigName
    srcWave.CreateConstant 0, DigSrc_Sample_Size, DspLong
    For Each site In TheExec.sites
        

        digsrc_assignment = temp_assignment
        code_bin(site) = vbNullString
        
        For i = 0 To TrimCodeSize - 1
            If i = 0 Then
                code_bin(site) = CStr(code(site) And 1)
            Else
                code_bin(site) = code_bin(site) & CStr((code(site) And (2 ^ i)) \ (2 ^ i))
            End If
        Next i
        
        
        
        digsrc_assignment = Replace(digsrc_assignment, Trimname, code_bin(site))
        
        
        Call Create_DigSrc_Data_Trim(srcPin, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, srcWave, site, MSB_First_Flag:=MSB_First_Flag)
        
        If DigSrc_DataWidth = 0 Then
            DigSrc_DataWidth = 4
        End If
        
        out_str(site) = vbNullString
        
        For i = 0 To DigSrc_Sample_Size - 1
        
            If (i Mod DigSrc_DataWidth) = 0 Then
                out_str(site) = out_str(site) & " "
            End If
                out_str(site) = out_str(site) & srcWave.Element(i)

        Next i
        
        'theexec.Datalog.WriteComment "Site " & Site & ",Code " & code_bin(Site) & ",Src Code = " & out_str(Site)
'        For i = 0 To TrimRepeat - 1
'            For j = 0 To TrimCodeSize - 1
'                srcwave.Element(i * TrimCodeSize + j) = srcwave.Element(j)
'            Next j
'        Next i
        
        TheExec.WaveDefinitions.CreateWaveDefinition "WaveDef" & site, srcWave, True
        With TheHdw.DSSC.Pins(srcPin).Pattern(Pat).Source.Signals(sigName)
            .WaveDefinitionName = "WaveDef" & site
            .SampleSize = DigSrc_Sample_Size
            .Amplitude = 1
            .LoadSamples
            If glb_TesterType = "Jaguar" Then .LoadSettings
        End With

    Next site
    TheHdw.DSSC.Pins(srcPin).Pattern(Pat).Source.Signals.DefaultSignal = sigName
    
    If DigCap_Sample_Size <> 0 Then
    
        Call AnalyzePatName(Pat, Str_FinalPatName)
        
        '' 20150812-Modify program to process multiply dig cap pins
        With TheHdw.DSSC.Pins(DigCap_Pin).Pattern(Pat).Capture.Signals
            .Add (Str_FinalPatName & DigCap_Sample_Size & "_" & DigCap_Pin)
            With .item(Str_FinalPatName & DigCap_Sample_Size & "_" & DigCap_Pin)
                .SampleSize = DigCap_Sample_Size    'CaptureCyc * OneCycle
                .LoadSettings
            End With
        End With
        
        'Create capture waveform
        OutDSP = TheHdw.DSSC.Pins(DigCap_Pin).Pattern(Pat).Capture.Signals(Str_FinalPatName & DigCap_Sample_Size & "_" & DigCap_Pin).DSPWave
        
        '' 20150813 - Assign WaveName to the DSPWave to do recognition of post process.
        For Each site In TheExec.sites
            OutDSP(site).info.WaveName = DigCap_Pin
        Next site
        
        ''TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug ''20180827 -- TYCHENGG -- use defaut as automatic
        TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
    
    End If
    
  
    
        
    TheHdw.Patterns(Pat).start
    
    
    
    'Call DebugPrintFunc_PPMU("")
    Dim MeasValue_Temp As New SiteDouble
    Set MeasValue_Temp = MeasValue_Temp.Add(10000000000000#)
    
    Dim MeasV_Flag As Boolean: MeasV_Flag = False
    
    For Each Ts In TestSequenceArray
        If CPUA_Flag_In_Pat = True Then
        TheHdw.Digital.Patgen.FlagWait cpuA, 0
        'thehdw.Wait 10 * ms
        Else
            TheHdw.Digital.Patgen.HaltWait
        End If
        Select Case UCase(Ts)
            Case "V"
                MeasV_Pin = CheckAndReturnArrayData(MeasV_Pin_split, TestSeqNum)
                Call Trim_SetupandmeasureV_UVI80(MeasV_Pin, MeasValue, code, code_bin, out_str, b_Trimfinish)
                If Meas_StoreName <> "" Then
                    If MeasStoreName_Ary(TestSeqNum) <> "" Then
                        Store_Rtn_Meas(StoreIndex).AddPin (MeasV_Pin)
                        Store_Rtn_Meas(StoreIndex).Pins(MeasV_Pin) = MeasValue
                        Call AddStoredMeasurement(MeasStoreName_Ary(TestSeqNum), Store_Rtn_Meas(StoreIndex))
                        StoreIndex = StoreIndex + 1
                    End If
                End If
                For Each site In TheExec.sites.Active
                    If (MeasValue_Temp > MeasValue) Then
                        MeasValue_Temp = MeasValue
                    End If
                Next site
                MeasV_Flag = True
            Case "I"
            
            
                Call Trim_SetupandmeasureI_UVI80(MeasI_Pin, MeasValue, MeasureI_Range, code, code_bin, out_str, b_Trimfinish)
                If Meas_StoreName <> "" Then
                    If MeasStoreName_Ary(TestSeqNum) <> "" Then
                        Rtn_Meas.AddPin (MeasI_Pin)
                        For Each site In TheExec.sites
                            Rtn_Meas.Pins(MeasI_Pin).value(site) = MeasValue(site)
                        Next site
                        Store_Rtn_Meas(StoreIndex) = Rtn_Meas
                        Call AddStoredMeasurement(MeasStoreName_Ary(TestSeqNum), Store_Rtn_Meas(StoreIndex))
                        StoreIndex = StoreIndex + 1
                    End If
                End If
            
               
                 
            Case "F"
            
                If MeasF_Interval = "" Then
                    MeasF_Interval = 0.001
                End If
            
                Call Trim_SetupandmeasureF_UVI80(MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, MeasValue, code, code_bin, out_str, b_Trimfinish)
                
                If Meas_StoreName <> "" Then
                    If MeasStoreName_Ary(TestSeqNum) <> "" Then
                        Rtn_Meas.AddPin (MeasF_PinS_SingleEnd)
                        For Each site In TheExec.sites
                            Rtn_Meas.Pins(MeasF_PinS_SingleEnd).value(site) = MeasValue(site)
                        Next site
                        Store_Rtn_Meas(StoreIndex) = Rtn_Meas
                        Call AddStoredMeasurement(MeasStoreName_Ary(TestSeqNum), Store_Rtn_Meas(StoreIndex))
                        StoreIndex = StoreIndex + 1
                    End If
                End If
                
                
                

            Case "C"
                Dim OutDSP2 As New DSPWave
                Dim OutDSP_Temp As New DSPWave
            
                If TheExec.TesterMode = testModeOffline Then
                    For Each site In TheExec.sites.Active
                        OutDSP.CreateRandom 0, 1, DigCap_Sample_Size, , DspLong
                    Next site
                End If
                
                
                'Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDSP)
                'Call DigCapDataProcessByDSP(CUS_Str_DigCapData, OutDSP, DigCap_Sample_Size, DigCap_DataWidth)
                Call Addstorecapture_Trim(CUS_Str_DigCapData, OutDSP, DigCap_Sample_Size, DigCap_DataWidth)
                If TrimCal_Name <> "" And InStr(TrimCal_Name, "C:") = 0 Then
                    OutDSP_Temp = GetStoredCaptureData(TrimCal_Name)
                Else
                    OutDSP_Temp = OutDSP
                End If
                
                For Each site In TheExec.sites.Active
                    OutDSP_Temp = OutDSP_Temp.ConvertDataTypeTo(DspLong)
                    OutDSP2 = OutDSP_Temp.ConvertStreamTo(tldspParallel, OutDSP_Temp.SampleSize, 0, Bit0IsMsb) '' Convert BInary to Desimal
                    MeasValue(site) = OutDSP2.Element(0)
                   If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site " & site & ",Code " & code_bin(site) & ", Src_code = " & out_str(site) & ", Code(Decimal) = " & MeasValue(site)
                Next site
            
            Case "N"
            
        End Select
        If CPUA_Flag_In_Pat = True Then
        TheHdw.Digital.Patgen.Continue 0, cpuA
        
        TestSeqNum = TestSeqNum + 1
        End If
    Next Ts
    
    If MeasV_Flag = True Then
        Set MeasValue = MeasValue_Temp
    End If
        
    TheHdw.Digital.Patgen.HaltWait
    
    If Final_Calc <> True Then
    If TrimCal_Name <> "" Then
        If Cal_Eqn <> "" Then
            Call ProcessCalcEquation(Cal_Eqn)
            If TrimCal_Name_array(0) = "C" Then
                TrimCalCap_value = GetStoredCaptureData(TrimCal_Name_array(1))
                For Each site In TheExec.sites.Active
                    MeasValue(site) = TrimCalCap_value.Element(0)
                   If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site " & site & ",Code " & code_bin(site) & ", Src_code = " & out_str(site) & ", " & TrimCal_Name_array(1) & " =" & MeasValue(site)
                Next site
                
            Else
                TrimCal_value = GetStoredMeasurement(TrimCal_Name)
                For Each site In TheExec.sites.Active
                    MeasValue(site) = TrimCal_value.Pins(0).value(site)
                    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site " & site & ",Code " & code_bin(site) & ", Src_code = " & out_str(site) & ", " & TrimCal_Name & " =" & MeasValue(site)
                Next site
            End If
        End If
                
    End If
    End If
    
    Exit Function
err:
    Stop
    Resume Next
End Function

Public Function Trim_CIOTXbyTable() As Long
'Added  20190509

    Dim regStr As String
    Dim PatStr() As String
    Dim TrimValue As String
    Dim i, j, k As Integer
    Dim TempCnt As Integer
    Dim RegWidth As Integer
    Dim RegSweeps() As Integer
    Dim OutWf() As New DSPWave
    ReDim OutWf(0) As New DSPWave
    Dim WorkBookName As Workbook
    Dim WorkSheetName As Worksheet
    
    Dim ciotx_sheetnumber As Integer
    Dim sheet_Cnt As Integer
    Dim SheetExist As Boolean: SheetExist = False
    sheet_Cnt = ActiveWorkbook.Sheets.Count
    For i = 1 To sheet_Cnt
        regStr = LCase(Sheets(i).name)
        If LCase(Sheets(i).name) Like "*ciotx_trimtable*" Then
            SheetExist = True
            ciotx_sheetnumber = i
            Exit For
        End If
    Next i
    
    
    If SheetExist = True Then
        Set WorkBookName = Application.ActiveWorkbook
        'Set WorkSheetName = WorkBookName.Sheets("CIOTX_TrimTable")
         Set WorkSheetName = WorkBookName.Sheets(UCase(Sheets(ciotx_sheetnumber).name))
        For i = 1 To CLng(WorkSheetName.UsedRange.Rows.Count)
            If CStr(WorkSheetName.Cells(i, 1)) <> "" Then
                If CStr(WorkSheetName.Cells(i, 1)) Like "*Pat*" Then                        ' Record each register parameter (size/width) from trim table
                    RegWidth = 0
                    PatStr = Split(CStr(WorkSheetName.Cells(i, 1)), ":")
                    ReDim RegSweeps(WorkSheetName.Cells(i, 1).END(xlToRight).Column - 2)
                    For j = 0 To WorkSheetName.Cells(i, 1).END(xlToRight).Column - 2
                        regStr = mid(CStr(WorkSheetName.Cells(i, j + 2)), InStr(1, CStr(WorkSheetName.Cells(i, j + 2)), "["))
                        regStr = WorksheetFunction.Substitute(WorksheetFunction.Substitute(regStr, "[", vbNullString), "]", vbNullString)
                        RegSweeps(j) = CInt(regStr)                                         ' Each regsiter size
                        RegWidth = RegWidth + CInt(regStr)                                  ' Each sweep registers width
                    Next j
                Else
                    OutWf(UBound(OutWf)).CreateConstant 0, CLng(RegWidth), DspLong
                    For j = 0 To WorkSheetName.Cells(i, 1).END(xlToRight).Column - 2        ' Record each register value from trim table
                        TrimValue = CStr(WorkSheetName.Cells(i, j + 2))
                        If TrimValue Like "*x*" Or TrimValue Like "*X*" Then                ' Avoid format error , 0xA0 --> 0A0 , 0X55 --> 055
                            TrimValue = Replace(TrimValue, "x", vbNullString)
                            TrimValue = Replace(TrimValue, "X", vbNullString)
                        End If
                        For k = 0 To RegSweeps(j) - 1                                       ' Format LSB ----> MSB
                            OutWf(UBound(OutWf)).Element(TempCnt) = CLng((CInt(WorksheetFunction.Hex2Dec(CStr(TrimValue))) And 2 ^ k) / 2 ^ k)
                            TempCnt = TempCnt + 1
                        Next k
                    Next j
                    TempCnt = 0
                    AddStoredCaptureData PatStr(1) & "_DigSrcTable_" & CStr(WorkSheetName.Cells(i, 1)), OutWf(UBound(OutWf))
                    ReDim Preserve OutWf(UBound(OutWf) + 1)
                End If
            End If
        Next i
    End If
    
End Function

Public Function Trim_SetupandmeasureF_UVI80(MeasF_PinS_SingleEnd As PinList, MeasF_Interval As String, MeasF_EventSourceWithTerminationMode As EventSourceWithTerminationMode, MeasValue As SiteDouble, code As SiteLong, code_bin() As String, out_str() As String, b_Trimfinish As Boolean)
    Dim site As Variant
    Dim MeasF_EventSource As FreqCtrEventSrcSel
    Dim MeasF_EnableVtMode As Boolean
    
    Call Freq_ProcessEventSourceTerminationMode(MeasF_EventSourceWithTerminationMode, MeasF_EventSource, MeasF_EnableVtMode)
    
    ''''''''''''''''''''''''''''setup measure F'''''''''''''''''''''''''''''''''
    With TheHdw.Digital.Pins(MeasF_PinS_SingleEnd).FreqCtr
        .EventSource = MeasF_EventSource '' VOH
        .EventSlope = Positive
        .Interval = MeasF_Interval
        .Enable = IntervalEnable
        .Clear
    End With
     
   
    
    Dim CounterValue As New SiteDouble
    
    TheHdw.Digital.Pins(MeasF_PinS_SingleEnd).FreqCtr.Clear
    TheHdw.Digital.Pins(MeasF_PinS_SingleEnd).FreqCtr.start
    
'    If CustomizeWaitTime <> "" Then
'        thehdw.Wait (CDbl(CustomizeWaitTime))
'    End If
        
    
''        freq = CounterValue.Math.Divide(interval)
    
    ''''''''''''''''''''''''''offline''''''''''''''''''''''''''''''''''''''''''''
    If TheExec.TesterMode = testModeOffline Then
        'Dim Pin As Variant
        '900000+code*100000

        MeasValue = code.Multiply(-50000).Add(1100000)     '0.000015 * -1 + code * 0.000001
        
        If b_Trimfinish = False And gl_Disable_HIP_debug_log = False Then
            TheExec.Datalog.WriteComment "Trimming"
        ElseIf gl_Disable_HIP_debug_log = False Then
            TheExec.Datalog.WriteComment "TrimResult"
        End If
        
        If gl_Disable_HIP_debug_log = False Then
            For Each site In TheExec.sites.Active
                TheExec.Datalog.WriteComment "Site " & site & ",Code " & code_bin(site) & ", Src_code = " & out_str(site) & ", Frequency = " & MeasValue(site)
            Next site
        End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Else
    
        CounterValue = TheHdw.Digital.Pins(MeasF_PinS_SingleEnd).FreqCtr.Read
        MeasValue = CounterValue.divide(MeasF_Interval)
        
        If b_Trimfinish = False And gl_Disable_HIP_debug_log = False Then
            TheExec.Datalog.WriteComment "Trimming"
        ElseIf gl_Disable_HIP_debug_log = False Then
            TheExec.Datalog.WriteComment "TrimResult"
        End If
        
        If gl_Disable_HIP_debug_log = False Then
            For Each site In TheExec.sites.Active
                TheExec.Datalog.WriteComment "Site " & site & ",Code " & code_bin(site) & ", Src_code = " & out_str(site) & ", Frequency = " & MeasValue(site)
            Next site
        End If
    End If
    
End Function

Public Function Trim_SetupandmeasureI_UVI80(MeasI_Pin As String, MeasValue As SiteDouble, MeasureI_Range As Double, code As SiteLong, code_bin() As String, out_str() As String, b_Trimfinish As Boolean)
 '''''''''''''''''setup UVI80 for measI''''''''''''''''''''''''''''''''''
   Dim Factor As Long
   Dim WaitTime As Double
   Dim site As Variant
   Factor = 1
   
    If MeasureI_Range > 2 * Factor Then
        MeasureI_Range = 2 * Factor
        WaitTime = 1.6 * ms
    ElseIf MeasureI_Range > 1 * Factor Then
        MeasureI_Range = 2 * Factor
        WaitTime = 1.6 * ms
    ElseIf MeasureI_Range > 0.2 * Factor Then
        MeasureI_Range = 1 * Factor
        WaitTime = 1.6 * ms
    ElseIf MeasureI_Range > 0.02 * Factor Then
        MeasureI_Range = 0.2 * Factor
        WaitTime = 260 * us
    ElseIf MeasureI_Range > 0.002 * Factor Then
        MeasureI_Range = 0.02 * Factor
        WaitTime = 1.5 * ms
    ElseIf MeasureI_Range > 0.0002 * Factor Then
        MeasureI_Range = 0.002 * Factor
        WaitTime = 11 * ms
    ElseIf MeasureI_Range > 0.00002 * Factor Then
        MeasureI_Range = 0.0002 * Factor
        WaitTime = 1.4 * ms
    Else
        MeasureI_Range = 0.00002 * Factor
        WaitTime = 6 * ms
    End If
      
    
    With TheHdw.DCVI.Pins(MeasI_Pin)
        .Gate = False
        .mode = tlDCVIModeVoltage
        .Voltage = 0
        .VoltageRange.value = pc_Def_VFI_UVI80_VoltageRange
        .SetCurrentAndRange MeasureI_Range, MeasureI_Range
        .Connect tlDCVIConnectDefault
        .Gate = True
    End With
    
    With TheHdw.DCVI.Pins(MeasI_Pin)
        .Meter.mode = tlDCVIMeterCurrent
        .Meter.CurrentRange.value = MeasureI_Range
    End With
    
    TheHdw.Wait (WaitTime)
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     
    ''''''''''''''''''offline simulate''''''''''''''''''''''''''''''
    If TheExec.TesterMode = testModeOffline Then
        'Dim Pin As Variant
        
        MeasValue = code.Multiply(0.000001).Add(-0.000015)  '0.000015 * -1 + code * 0.000001
        
        
        If gl_Disable_HIP_debug_log = False Then
            If b_Trimfinish = False Then
                TheExec.Datalog.WriteComment "Trimming"
            Else
                TheExec.Datalog.WriteComment "TrimResult"
            End If
                
            For Each site In TheExec.sites.Active
                TheExec.Datalog.WriteComment "Site " & site & ",Code " & code_bin(site) & ", Src_code = " & out_str(site) & ", Current = " & MeasValue(site)
            Next site
        End If
        
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Else
        
        MeasValue = TheHdw.DCVI.Pins(MeasI_Pin).Meter.Read(tlStrobe, 10)
        
        If gl_Disable_HIP_debug_log = False Then

            If b_Trimfinish = False Then
                TheExec.Datalog.WriteComment "Trimming"
            Else
                TheExec.Datalog.WriteComment "TrimResult"
            End If
            
            For Each site In TheExec.sites.Active
                TheExec.Datalog.WriteComment "Site " & site & ",Code " & code_bin(site) & ", Src_code = " & out_str(site) & ", Current = " & MeasValue(site)
            Next site
        End If
    
    End If





End Function

Public Function Trim_SetupandmeasureV_UVI80(MeasV_Pin As String, MeasValue As SiteDouble, code As SiteLong, code_bin() As String, out_str() As String, b_Trimfinish As Boolean)
        Dim site As Variant
 '''''''''''''''setup UVI80 for meas V''''''''''''''''''
 
    Dim Previous_ByPassTestLimit_Flag As Boolean: Previous_ByPassTestLimit_Flag = ByPassTestLimit
    
    ByPassTestLimit = True
'    With TheHdw.DCVI.Pins(MeasV_Pin)
'        .Gate = False
'        .Disconnect tlDCVIConnectDefault
'        .mode = tlDCVIModeHighImpedance
'        .Connect tlDCVIConnectHighSense
'        .Voltage = 6
'        .current = 0
'         'thehdw.Wait 1 * ms
'        .Gate = True
'    End With
'
'    With TheHdw.DCVI.Pins(MeasV_Pin)
'        .Meter.mode = tlDCVIMeterVoltage
'    End With
'    TheHdw.Wait 1 * ms
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim cal As New SiteDouble
    
    ''''''''''''''''''offline simulate''''''''''''''''''''''''''''''
    If TheExec.TesterMode = testModeOffline Then
        'Dim Pin As Variant
        
        For Each site In TheExec.sites.Active
            MeasValue(site) = code * 0.1
        Next site
'        If InStr(glb_TestInstance, "MTRGR_T4P2") <> 0 Then
'            TheExec.Datalog.WriteComment "trimming"
'            cal = MeasValue.Subtract(0.4).Divide(0.7975).Subtract(1)
'            For Each Site In theexec.sites.Active
'                TheExec.Datalog.WriteComment "Site " & Site & ",Code " & code_bin(Site) & ", Src_code = " & out_str(Site) & ", Gain_error = " & cal(Site) & ", Voltage = " & MeasValue(Site)
'            Next Site
'            MeasValue = cal
'        Else
        If gl_Disable_HIP_debug_log = False Then

            If b_Trimfinish = False Then
                TheExec.Datalog.WriteComment "Trimming"
            Else
                TheExec.Datalog.WriteComment "TrimResult"
            End If
            
            For Each site In TheExec.sites.Active
                TheExec.Datalog.WriteComment "Site " & site & ",Code " & code_bin(site) & ", Src_code = " & out_str(site) & ", Voltage = " & MeasValue(site)
            Next site
        End If
'        End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Else
        MeasValue = HardIP_MeasureVolt
'        MeasValue = TheHdw.DCVI.Pins(MeasV_Pin.Value).Meter.Read(tlStrobe, 10)
'        If InStr(glb_TestInstance, "MTRGR_T4P2") <> 0 Then
'            TheExec.DataLog.WriteComment "trimming"
'            cal = MeasValue.Subtract(0.4).Divide(0.7975).Subtract(1)
'            For Each site In TheExec.sites.Active
'                TheExec.DataLog.WriteComment "Site " & site & ",Code " & code_bin(site) & ", Src_code = " & out_str(site) & ", Gain_error = " & cal(site) & ", Voltage = " & MeasValue(site)
'            Next site
'            MeasValue = cal
'        Else
        If gl_Disable_HIP_debug_log = False Then

            If b_Trimfinish = False Then
                TheExec.Datalog.WriteComment "Trimming"
            Else
                TheExec.Datalog.WriteComment "TrimResult"
            End If

            For Each site In TheExec.sites.Active
                TheExec.Datalog.WriteComment "Site " & site & ",Code " & code_bin(site) & ", Src_code = " & out_str(site) & ", Voltage = " & MeasValue(site)
            Next site
        End If
'        End If
    
    End If

'    With TheHdw.DCVI.Pins(MeasV_Pin)
'        .Gate(tlDCVIGateHiZ) = False
'        .Disconnect
'        .mode = tlDCVIModeCurrent
'    End With
'

    ByPassTestLimit = Previous_ByPassTestLimit_Flag

End Function
Public Function Addstorecapture_Trim(CUS_Str_DigCapData As String, OutDspWave As DSPWave, DigCap_Sample_Size As Long, DigCap_DataWidth As Long, Optional CUS_Str_MainProgram As String, _
                        Optional BypassAllDigCapTestLimit As Boolean = False)
    
    Dim site As Variant
    Dim i As Long, j As Long
    Dim Str_PrintBinary As New SiteVariant
    Dim ConvertedDataWf As New DSPWave
    Dim SourceBitStrmWf As New DSPWave
    Dim NoOfSamples As New SiteLong
    
    Dim FlexibleConvertedDataWf() As New DSPWave

    '' 20160328
    Dim TestLimitWithTestName As New PinListData

    ''20170418 - Move out from If DigCap_DataWidth <> 0 And InStr(UCase(CUS_Str_DigCapData), "DSSC_OUT") = 0 And CUS_Str_MainProgram = "" Then
    Dim p As Long
    Dim PinName As String
    Dim DigCapValue As New PinListData
    Dim b_FirstTimeSwitch As Boolean
    
    '' 20160211 - Process format by DigCap_DataWidth, capture word size is fixed
    If DigCap_DataWidth <> 0 And InStr(UCase(CUS_Str_DigCapData), "DSSC_OUT") = 0 Then
    
            Dim CalcOutputDSPWave As New DSPWave
            Dim CalcEyeWidth As New SiteLong
            Dim FinalEyeOutBitNum As Long
            Dim TestLimitForEyeSweep As New DSPWave
        
        ''20170811 - EyeSweep for LPDPRX
        If UCase(CUS_Str_MainProgram) = UCase("LPDPRX_EyeSweep") Then
            For Each site In TheExec.sites
                For i = 1 To OutDspWave.SampleSize
                    Str_PrintBinary(site) = Str_PrintBinary(site) & OutDspWave(site).Element(i - 1)
                    If i Mod (DigCap_DataWidth) = 0 Then
                        Str_PrintBinary(site) = Str_PrintBinary(site) & ","
                    End If
                Next i
                If gl_Disable_HIP_debug_log = False Then Call TheExec.Datalog.WriteComment("Site(" & site & ") DigCap Bit Size = " & DigCap_Sample_Size & ", Data Width = " & DigCap_DataWidth & ", Binary string = " & Str_PrintBinary(site))
            
            Next site
            SourceBitStrmWf = OutDspWave
        
            rundsp.BitWf2Arry SourceBitStrmWf, DigCap_DataWidth, NoOfSamples, ConvertedDataWf
            
'            Dim CalcOutputDSPWave As New DSPWave
'            Dim CalcEyeWidth As New SiteLong
'            Dim FinalEyeOutBitNum As Long
            FinalEyeOutBitNum = DigCap_Sample_Size / 32
            rundsp.LPDPRX_EyeSweep ConvertedDataWf, FinalEyeOutBitNum, CalcOutputDSPWave, CalcEyeWidth
            
            For Each site In TheExec.sites
                Str_PrintBinary(site) = vbNullString
                For i = 1 To CalcOutputDSPWave.SampleSize
                    Str_PrintBinary(site) = Str_PrintBinary(site) & CalcOutputDSPWave(site).Element(i - 1)
    ''                If i Mod (DigCap_DataWidth) = 0 Then
    ''                    Str_PrintBinary(Site) = Str_PrintBinary(Site) & ","
    ''                End If
                Next i
                If gl_Disable_HIP_debug_log = False Then Call TheExec.Datalog.WriteComment("Site(" & site & ") Output eye bits = " & FinalEyeOutBitNum & ", Binary string = " & Str_PrintBinary(site))
                ''20170510 Store Binary String for Eye Diagram
                Eye_Diagram_Binary(TheExec.Flow.var("SrcCodeIndx").value + 31)(site) = Str_PrintBinary(site)
            Next site
'            Dim TestLimitForEyeSweep As New DSPWave
            rundsp.BitWf2Arry SourceBitStrmWf, DigCap_DataWidth, NoOfSamples, TestLimitForEyeSweep
    
            b_FirstTimeSwitch = True
            For Each site In TheExec.sites
                PinName = "EyeCapWord_"
                Exit For
            Next site
            For Each site In TheExec.sites
                For i = 1 To TestLimitForEyeSweep.SampleSize
                    If b_FirstTimeSwitch Then
                        DigCapValue.AddPin (PinName & CStr(i - 1))
                    End If
                    DigCapValue.Pins(PinName & CStr(i - 1)).value(site) = TestLimitForEyeSweep(site).Element(i - 1)
                Next i
                b_FirstTimeSwitch = False
            Next site
    
            If BypassAllDigCapTestLimit = False Then
                For p = 0 To DigCapValue.Pins.Count - 1
                    'TheExec.Flow.TestLimit DigCapValue.Pins(p), 0, 2 ^ DigCap_DataWidth - 1, Tname:=glb_TestInstance & "_EyeCpatureCode_" & p, PinName:="EyeCpatureCode_" & p, ForceResults:=tlForceFlow, ScaleType:=scaleNoScaling, formatstr:="%.0f"
                Next p
                'TheExec.Flow.TestLimit resultVal:=CalcEyeWidth, Tname:=glb_TestInstance & "_EyeWidth", ForceResults:=tlForceFlow
            End If
        '20170811 PCIE Eye Sweep
        ElseIf UCase(CUS_Str_MainProgram) = UCase("PCIE_EyeSweep") Then
            For Each site In TheExec.sites
                For i = 1 To OutDspWave.SampleSize
                    Str_PrintBinary(site) = Str_PrintBinary(site) & OutDspWave(site).Element(i - 1)
                    If i Mod (DigCap_DataWidth) = 0 Then
                        Str_PrintBinary(site) = Str_PrintBinary(site) & ","
                    End If
                Next i
                If gl_Disable_HIP_debug_log = False Then Call TheExec.Datalog.WriteComment("Site(" & site & ") DigCap Bit Size = " & DigCap_Sample_Size & ", Data Width = " & DigCap_DataWidth & ", Binary string = " & Str_PrintBinary(site))
            
            Next site
            SourceBitStrmWf = OutDspWave
        
            rundsp.BitWf2Arry SourceBitStrmWf, DigCap_DataWidth, NoOfSamples, ConvertedDataWf
            
'            Dim CalcOutputDSPWave As New DSPWave
'            Dim CalcEyeWidth As New SiteLong
'            Dim FinalEyeOutBitNum As Long
            FinalEyeOutBitNum = DigCap_Sample_Size / 20
            rundsp.PCIE_EyeSweep ConvertedDataWf, FinalEyeOutBitNum, CalcOutputDSPWave, CalcEyeWidth
            
            For Each site In TheExec.sites
                Str_PrintBinary(site) = vbNullString
                For i = 1 To CalcOutputDSPWave.SampleSize
                    Str_PrintBinary(site) = Str_PrintBinary(site) & CalcOutputDSPWave(site).Element(i - 1)
    ''                If i Mod (DigCap_DataWidth) = 0 Then
    ''                    Str_PrintBinary(Site) = Str_PrintBinary(Site) & ","
    ''                End If
                Next i
               If gl_Disable_HIP_debug_log = False Then Call TheExec.Datalog.WriteComment("Site(" & site & ") Output eye bits = " & FinalEyeOutBitNum & ", Binary string = " & Str_PrintBinary(site))
                ''20170510 Store Binary String for Eye Diagram
                Eye_Diagram_Binary(TheExec.Flow.var("SrcCodeIndx").value + 31)(site) = Str_PrintBinary(site)
            Next site
'            Dim TestLimitForEyeSweep As New DSPWave
            rundsp.BitWf2Arry SourceBitStrmWf, DigCap_DataWidth, NoOfSamples, TestLimitForEyeSweep
    
            b_FirstTimeSwitch = True
            For Each site In TheExec.sites
                PinName = "EyeCapWord_"
                Exit For
            Next site
            For Each site In TheExec.sites
                For i = 1 To TestLimitForEyeSweep.SampleSize
                    If b_FirstTimeSwitch Then
                        DigCapValue.AddPin (PinName & CStr(i - 1))
                    End If
                    DigCapValue.Pins(PinName & CStr(i - 1)).value(site) = TestLimitForEyeSweep(site).Element(i - 1)
                Next i
                b_FirstTimeSwitch = False
            Next site
    
            If BypassAllDigCapTestLimit = False Then
                For p = 0 To DigCapValue.Pins.Count - 1
                    'TheExec.Flow.TestLimit DigCapValue.Pins(p), 0, 2 ^ DigCap_DataWidth - 1, Tname:=glb_TestInstance & "_EyeCpatureCode_" & p, PinName:="EyeCpatureCode_" & p, ForceResults:=tlForceFlow, ScaleType:=scaleNoScaling, formatstr:="%.0f"
                Next p
                'TheExec.Flow.TestLimit resultVal:=CalcEyeWidth, Tname:=glb_TestInstance & "_EyeWidth", ForceResults:=tlForceFlow
            End If
        Else
    
            For Each site In TheExec.sites
                For i = 1 To OutDspWave.SampleSize
                    Str_PrintBinary(site) = Str_PrintBinary(site) & OutDspWave(site).Element(i - 1)
                    If i Mod (DigCap_DataWidth) = 0 Then
                        Str_PrintBinary(site) = Str_PrintBinary(site) & ","
                    End If
                Next i
                If gl_Disable_HIP_debug_log = False Then Call TheExec.Datalog.WriteComment("Site(" & site & ") DigCap Bit Size = " & DigCap_Sample_Size & ", Data Width = " & DigCap_DataWidth & ", Binary string = " & Str_PrintBinary(site))
            
            Next site
            SourceBitStrmWf = OutDspWave
        
            rundsp.BitWf2Arry SourceBitStrmWf, DigCap_DataWidth, NoOfSamples, ConvertedDataWf
            
            
            '' 20160211 - Get pin name from dsp wave
    ''        Dim p As Long
    ''        Dim PinName As String
    ''        Dim DigCapValue As New PinListData
    ''        Dim b_FirstTimeSwitch As Boolean
            b_FirstTimeSwitch = True
            For Each site In TheExec.sites
                PinName = OutDspWave(site).info.WaveName & "_DigCapWord_"
                Exit For
            Next site
            For Each site In TheExec.sites
                For i = 1 To ConvertedDataWf.SampleSize
                    If b_FirstTimeSwitch Then
                        DigCapValue.AddPin (PinName & CStr(i - 1))
                    End If
                    DigCapValue.Pins(PinName & CStr(i - 1)).value(site) = ConvertedDataWf(site).Element(i - 1)
                Next i
                b_FirstTimeSwitch = False
            Next site
            
            If BypassAllDigCapTestLimit = False Then
                For p = 0 To DigCapValue.Pins.Count - 1
                    'TheExec.Flow.TestLimit DigCapValue.Pins(p), 0, 2 ^ DigCap_DataWidth - 1, Tname:=glb_TestInstance & "_CpatureCode_" & p, PinName:="CpatureCode_" & p, ForceResults:=tlForceFlow, ScaleType:=scaleNoScaling, formatstr:="%.0f"
                Next p
            End If
    ''        Call CUS_VFI_MainProgram_ECID(CUS_Str_MainProgram, DigCapValue)

        End If


        
    '' 20160212 - Process format by DSSC_OUT, capture word size is flexible, also parse with/without test name.
    ElseIf InStr(UCase(CUS_Str_DigCapData), "DSSC_OUT") <> 0 Then

        Dim Split_Num() As String
        Dim StartNum As Long
        '' 20151231 - Add rule to check new format that include test name and parse bits
        Dim ParseStringByBits As String
        Dim ParseStringForTestName As String
        Dim DSSC_Out_DecompseByComma() As String
        Dim DSSC_Out_DecompseByColon() As String
        Dim b_DSSC_Out_InvolveTestName As Boolean
        ParseStringByBits = vbNullString
        ParseStringForTestName = vbNullString
        b_DSSC_Out_InvolveTestName = False
        Dim DecomposeTestName() As String
        Dim DecomposeParseDigCapBit() As String
        
        ''20160807 - Add directionary to store DigCap DSPwave
        Dim ParseStringForDirectionary As String
        Dim DecomposeDirectionary() As String
        Dim b_ParseForDirectionary_Switch As Boolean
        b_ParseForDirectionary_Switch = False
        
        Dim b_ParseForGrayCode_Switch As Boolean
        b_ParseForGrayCode_Switch = False
        Dim ParseStringForGrayCode As String
        Dim DecomposeGrayCode() As String
        
        If InStr(UCase(CUS_Str_DigCapData), ":") <> 0 Then
            b_DSSC_Out_InvolveTestName = True
            DSSC_Out_DecompseByComma = Split(CUS_Str_DigCapData, ",")
            For i = 0 To UBound(DSSC_Out_DecompseByComma)
                DSSC_Out_DecompseByColon = Split(DSSC_Out_DecompseByComma(i), ":")
                If UBound(DSSC_Out_DecompseByColon) > 0 Then
                    If ParseStringByBits = "" And ParseStringForTestName = "" Then
                        ParseStringByBits = DSSC_Out_DecompseByColon(0)
                        ParseStringForTestName = DSSC_Out_DecompseByColon(1)
                        If UBound(DSSC_Out_DecompseByColon) = 2 Then    '' Dictionary
                            ParseStringForDirectionary = DSSC_Out_DecompseByColon(2) & ","
                            ParseStringForGrayCode = ","
                        ElseIf UBound(DSSC_Out_DecompseByColon) = 3 Then    '' Dictionary and GrayCode
                            ParseStringForDirectionary = DSSC_Out_DecompseByColon(2) & ","
                            ParseStringForGrayCode = DSSC_Out_DecompseByColon(3) & ","
                        Else
                            ParseStringForDirectionary = ","
                            ParseStringForGrayCode = ","
                        End If
                    Else
                        ParseStringByBits = ParseStringByBits & "," & DSSC_Out_DecompseByColon(0)
                        ParseStringForTestName = ParseStringForTestName & "," & DSSC_Out_DecompseByColon(1)
                        
                        If b_ParseForDirectionary_Switch = False Then
                            b_ParseForDirectionary_Switch = True
                        Else
                            ParseStringForDirectionary = ParseStringForDirectionary & ","
                        End If
                        
                        If b_ParseForGrayCode_Switch = False Then
                            b_ParseForGrayCode_Switch = True
                        Else
                            ParseStringForGrayCode = ParseStringForGrayCode & ","
                        End If
                        
                        If UBound(DSSC_Out_DecompseByColon) = 2 Then    '' Dictionary
                            ParseStringForDirectionary = ParseStringForDirectionary & DSSC_Out_DecompseByColon(2)
                        End If
                        
                        If UBound(DSSC_Out_DecompseByColon) = 3 Then    '' Dictionary
                            ParseStringForDirectionary = ParseStringForDirectionary & DSSC_Out_DecompseByColon(2)
                            ParseStringForGrayCode = ParseStringForGrayCode & DSSC_Out_DecompseByColon(3)
                        End If
                    
                    End If
                End If
            Next i
            
            ''20161220-Remove comma in the last of string
            If right(ParseStringForTestName, 1) = "," Then
                ParseStringForTestName = left(ParseStringForTestName, (Len(ParseStringForTestName) - 1))
            End If
            If right(ParseStringForDirectionary, 1) = "," Then
                ParseStringForDirectionary = left(ParseStringForDirectionary, (Len(ParseStringForDirectionary) - 1))
            End If
            
            ParseStringByBits = "DSSC_OUT," & ParseStringByBits
            DecomposeTestName = Split(ParseStringForTestName, ",")
            DecomposeDirectionary = Split(ParseStringForDirectionary, ",")
            DecomposeGrayCode = Split(ParseStringForGrayCode, ",")
        Else
            ParseStringByBits = CUS_Str_DigCapData
        End If
        
        If right(ParseStringByBits, 1) = "," Then
            ParseStringByBits = left(ParseStringByBits, (Len(ParseStringByBits) - 1))
        End If
        DecomposeParseDigCapBit = Split(ParseStringByBits, ",")
        Dim StrParseDigCapBit As String
        
        For i = 1 To UBound(DecomposeParseDigCapBit)
            If i = 1 Then
                StrParseDigCapBit = DecomposeParseDigCapBit(i)
            Else
                StrParseDigCapBit = StrParseDigCapBit & "," & DecomposeParseDigCapBit(i)
            End If
        Next i
        DecomposeParseDigCapBit = Split(StrParseDigCapBit, ",")
        
        ReDim FlexibleConvertedDataWf(UBound(DecomposeParseDigCapBit)) As New DSPWave
        ''20160823-Store binary dsp wave after processed by DSSC_OUT
        Dim DSPWave_Binary() As New DSPWave
        ReDim DSPWave_Binary(UBound(DecomposeParseDigCapBit)) As New DSPWave
        
        Dim DSPWave_GrayCode() As New DSPWave
        ReDim DSPWave_GrayCode(UBound(DecomposeParseDigCapBit)) As New DSPWave
        
        Dim DSPWave_GrayCodeDec() As New DSPWave
        ReDim DSPWave_GrayCodeDec(UBound(DecomposeParseDigCapBit)) As New DSPWave
        
        Dim StartIndex As Long
        StartIndex = 0
        
        ''20161230-Add copy and site loop to pass data
        For Each site In TheExec.sites
            SourceBitStrmWf = OutDspWave.Copy
        Next site
        
        
        Dim width_Wf  As New DSPWave, OutWf As New DSPWave ', OutBinWf() As New DSPWave
       ' ReDim OutBinWf(UBound(DecomposeParseDigCapBit))
        width_Wf.CreateConstant 0, UBound(DecomposeParseDigCapBit) + 1   'Create space for DSP
        'OutWf.CreateConstant 0, UBound(DecomposeParseDigCapBit) + 1
        
        For Each site In TheExec.sites
            For i = 0 To UBound(DecomposeParseDigCapBit)
                width_Wf.ElementLite(i) = CLng(DecomposeParseDigCapBit(i))  'deliver data to dsp array
            Next i
        Next site
        
        rundsp.Split_Dspwave SourceBitStrmWf, width_Wf, OutWf                   ', OutBinWf
        
        For Each site In TheExec.sites
            For i = 0 To UBound(DecomposeParseDigCapBit)
                FlexibleConvertedDataWf(i).CreateConstant 0, 1
                FlexibleConvertedDataWf(i).Element(0) = OutWf.ElementLite(i)
            Next i
        Next site
       
        ''20160823-Modify dsp function to add one input argument to process DSPwave with binary format and use Directionary to store it.
        For i = 0 To UBound(DecomposeParseDigCapBit)
'            rundsp.FlexibleBitWf2Arry SourceBitStrmWf, StartIndex, CLng(DecomposeParseDigCapBit(i)), FlexibleConvertedDataWf(i), DSPWave_Binary(i)
            
            ''20160823-Store binary DSP wave by using Directionary
''            If DecomposeDirectionary(i) <> "" Then
''                Call AddStoredCaptureData(DecomposeDirectionary(i), DSPWave_Binary(i))
''            End If
            
            If UCase(DecomposeGrayCode(i)) = "GRAYCODE" Then
''                DSPWave_GrayCode(i).CreateConstant 0, DecomposeParseDigCapBit(i), DspLong
''                DSPWave_GrayCodeDec(i).CreateConstant 0, 1, DspLong
                
                Call rundsp.Transfer2GrayCode(DSPWave_Binary(i), DSPWave_GrayCode(i), DSPWave_GrayCodeDec(i))
                
            End If
            
            StartIndex = StartIndex + DecomposeParseDigCapBit(i)
        Next i
        
        ''20161215-Check the dictionary name, re-combine them to one dsp wave and store to dictionary if there have the same dictionary name cross multi-segment (over 24 bit).
        '' Separate dsp wave to different segment if over 24bits, this is for cover STDF display truncation issue.
        Dim CombineDSPBit2Dict As New Dictionary
        Dim keyname As String
        If ParseStringForDirectionary <> "" Then
            CombineDSPBit2Dict.RemoveAll
            For i = 0 To UBound(DecomposeDirectionary)
                If DecomposeDirectionary(i) = "" Then
                    keyname = "EMPTYSPACE_DICT_" & i
                Else
                    keyname = LCase(DecomposeDirectionary(i))
                End If
                If i = 0 Then
                    CombineDSPBit2Dict.Add keyname, CLng(DecomposeParseDigCapBit(i))
    
                Else
                    If CombineDSPBit2Dict.Exists(keyname) Then
                        CombineDSPBit2Dict.item(keyname) = CombineDSPBit2Dict.item(keyname) + CLng(DecomposeParseDigCapBit(i))
                    Else
                        CombineDSPBit2Dict.Add keyname, CLng(DecomposeParseDigCapBit(i))
                    End If
                End If
            Next i
            
            Dim CombineKeys() As Variant
            CombineKeys() = CombineDSPBit2Dict.Keys()
            
            StartIndex = 0
            ReDim AddToDict_DSP_Dec(CombineDSPBit2Dict.Count - 1) As New DSPWave
            ReDim AddToDict_DSP_Bin(CombineDSPBit2Dict.Count - 1) As New DSPWave
            Dim FinalLength As Long
            
            For i = 0 To CombineDSPBit2Dict.Count - 1
                FinalLength = CombineDSPBit2Dict.item(CombineKeys(i))
''                rundsp.FlexibleBitWf2Arry SourceBitStrmWf, StartIndex, FinalLength, AddToDict_DSP_Dec(i), AddToDict_DSP_Bin(i)
                If InStr(CombineKeys(i), "EMPTYSPACE_DICT_") <> 0 Then
                Else
                    For Each site In TheExec.sites
                        AddToDict_DSP_Bin(i) = SourceBitStrmWf.Select(StartIndex, , FinalLength).Copy '.ConvertStreamTo(tldspSerial, FinalLength, 0, Bit0IsMsb)
                    Next site
                    Call AddStoredCaptureData(CStr(CombineKeys(i)), AddToDict_DSP_Bin(i))
                End If
                StartIndex = StartIndex + FinalLength
            Next i
        End If
        
        
        '' Debug use
        Dim BinaryCodeString As String
        Dim GrayCodeString As String
''        Dim j As Long
        For Each site In TheExec.sites
            For i = 0 To UBound(DecomposeParseDigCapBit)
                If UCase(DecomposeGrayCode(i)) = "GRAYCODE" Then
                    BinaryCodeString = vbNullString
                    GrayCodeString = vbNullString
                    For j = 0 To DSPWave_Binary(i).SampleSize - 1
                        BinaryCodeString = BinaryCodeString & DSPWave_Binary(i)(site).Element(j)
                        GrayCodeString = GrayCodeString & DSPWave_GrayCode(i)(site).Element(j)
                    Next j
                    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " DSSC_OUT part " & i & " binary code = " & BinaryCodeString)
                    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " DSSC_OUT part " & i & "   gray code = " & GrayCodeString)
                End If
            Next i
        Next site
                
        '' 20160317 - Test limit for DSSC_OUT
'        If b_DSSC_Out_InvolveTestName = True Then '' Test limit with test name
'            For i = 0 To UBound(DecomposeTestName)
'                If LCase(DecomposeTestName(i)) = "skip" Then
'                Else
'                    TestLimitWithTestName.AddPin (DecomposeTestName(i) & "_" & i)
'                    If UCase(DecomposeGrayCode(i)) = "GRAYCODE" Then
'                        TestLimitWithTestName.Pins(DecomposeTestName(i) & "_" & i).Value = DSPWave_GrayCodeDec(i).Element(0)
'                    Else
'                        TestLimitWithTestName.Pins(DecomposeTestName(i) & "_" & i).Value = FlexibleConvertedDataWf(i).Element(0)
'                    End If
'                    If BypassAllDigCapTestLimit = False Then
'                        If CUS_Str_MainProgram <> "" And InStr(UCase(CUS_Str_MainProgram), UCase("Loop_DigSrc")) <> 0 Then
'                            TheExec.Flow.TestLimit TestLimitWithTestName.Pins(DecomposeTestName(i) & "_" & i), 0, 2 ^ DecomposeParseDigCapBit(i) - 1, Tname:=glb_TestInstance & "_" & DecomposeTestName(i) & "_" & i, ForceResults:=tlForceFlow, ScaleType:=scaleNoScaling, formatstr:="%.0f"
'
'                        ElseIf MTR_CusDigCap <> "" And UCase(MTR_CusDigCap) = "CUS_DIGCAP_VIN" Then
'                            TheExec.Flow.TestLimit TestLimitWithTestName.Pins(DecomposeTestName(i) & "_" & i), 0, 2 ^ DecomposeParseDigCapBit(i) - 1, Tname:=glb_TestInstance & "_" & DecomposeTestName(i) & "_" & MTR_VIN & "_" & i, ForceResults:=tlForceFlow, ScaleType:=scaleNoScaling, formatstr:="%.0f"
'                            MTR_CusDigCap = ""
'
'                        ElseIf TPModeAsCharz_GLB = True Then  ''CZ TP name force flow
'                            TheExec.Flow.TestLimit TestLimitWithTestName.Pins(DecomposeTestName(i) & "_" & i), 0, 2 ^ DecomposeParseDigCapBit(i) - 1, ForceResults:=tlForceFlow, ScaleType:=scaleNoScaling, formatstr:="%.0f"
'
'                        ElseIf CUS_Str_MainProgram = "TMPS_BV" Then  ''TMPS_BV
'                            TheExec.Flow.TestLimit TestLimitWithTestName.Pins(DecomposeTestName(i) & "_" & i), 0, 2 ^ DecomposeParseDigCapBit(i) - 1, Tname:=glb_TestInstance & "_" & DecomposeTestName(i) & "_" & i, ForceResults:=tlForceFlow, ScaleType:=scaleNoScaling, formatstr:="%.0f"
'                        Else
'                            TheExec.Flow.TestLimit TestLimitWithTestName.Pins(DecomposeTestName(i) & "_" & i), 0, 2 ^ DecomposeParseDigCapBit(i) - 1, Tname:=glb_TestInstance & "_" & DecomposeTestName(i) & "_" & i, ForceResults:=tlForceFlow, ScaleType:=scaleNoScaling, formatstr:="%.0f"
'                        End If
'                    End If
'                End If
'            Next i
'
'        Else
'            If BypassAllDigCapTestLimit = False Then
'                For i = 0 To UBound(FlexibleConvertedDataWf)
'                    TheExec.Flow.TestLimit FlexibleConvertedDataWf(i).Element(0), 0, 2 ^ DecomposeParseDigCapBit(i) - 1, PinName:="DSSC_OUT_Code_" & i, Tname:=glb_TestInstance & "_DSSC_OUT_" & CStr(i - 1), ForceResults:=tlForceFlow, ScaleType:=scaleNoScaling, formatstr:="%.0f"
'                Next
'            End If
'        End If
    End If
End Function


Public Function DSSC_Search_par_run_LDO(Pat As String, srcPin As PinList, code As SiteLong, MeasPin As PinList, Res As SiteDouble, TrimCodeSize As Long, NumberOfMeasV As Integer, ByRef MeasV_Name_Array() As String, ByRef MeasValue_Array() As SiteDouble, ByRef TrimPoint() As Long, DigSrc_Sample_Size As Long, DigSrc_Equation As String, digsrc_assignment As String, TrimStoreName As String, MeasV_WaitTime As String)
    Dim sigName As String, srcWave As New DSPWave, site As Variant ': Site = theexec.sites.SiteNumber
    Dim InDSPWave As New DSPWave
'    Dim MeasV_Name_Array() As String: MeasV_Name_Array = Split(MeasV_Name, "+")
    Dim MeasValue As New SiteDouble
    Dim i As Long, j As Long
    Dim Rtn_MeasVolt As New PinListData
'    ByPassTestLimit = True
    Dim FlowTestNme() As String
    Dim HighLimitVal() As Double, LowLimitVal() As Double
    Call GetFlowSingleUseLimit(HighLimitVal, LowLimitVal)
    Dim srcwave_array() As Long: ReDim srcwave_array(TrimCodeSize - 1)
    srcWave.CreateConstant 0, TrimCodeSize, DspLong
    Dim Previous_ByPassTestLimit_Flag As Boolean: Previous_ByPassTestLimit_Flag = ByPassTestLimit
    Dim Previous_Disable_CurrRangeSetting_Print_Flag As Boolean: Previous_Disable_CurrRangeSetting_Print_Flag = glb_Disable_CurrRangeSetting_Print
    

    ByPassTestLimit = True
    glb_Disable_CurrRangeSetting_Print = True
    For Each site In TheExec.sites
        For i = 0 To TrimCodeSize - 1
            If i = 0 Then
                srcwave_array(i) = code And 1
            Else
                srcwave_array(i) = (code And (2 ^ i)) \ (2 ^ i)
            End If
        Next i
    srcWave.data = srcwave_array
    Next site
    
    Call AddStoredCaptureData(TrimStoreName, srcWave)
    Call GeneralDigSrcSetting_LDO(Pat, srcPin, DigSrc_Sample_Size, DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, vbNullString, vbNullString, InDSPWave, vbNullString)
    Dim MinTrimIndex As Long
    For i = 0 To UBound(TrimPoint)
        If TrimPoint(i) = 1 Then
            MinTrimIndex = i
            Exit For
        End If
    Next i
    TheHdw.Patterns(Pat).start
    
    For i = 0 To NumberOfMeasV - 1
        TheHdw.Digital.Patgen.FlagWait cpuA, 0

'        Call HardIP_MeasureVolt(MeasPin, "FFF", NumberOfMeasV, 1, Pat, False, HighLimitVal(0), LowLimitVal(0), FlowTestNme, , "LDO_Trim", , Rtn_MeasVolt, , , MeasV_WaitTime)
        Rtn_MeasVolt = HardIP_MeasureVolt
        Call DebugPrintFunc_PPMU(vbNullString)

        For Each site In TheExec.sites
            MeasValue = Rtn_MeasVolt.Pins(MeasPin.value).value
            MeasValue_Array(i) = Rtn_MeasVolt.Pins(MeasPin.value).value

'            If i = 0 And TrimPoint(i) Then
            If i = MinTrimIndex Then
                Res = Rtn_MeasVolt.Pins(MeasPin.value).value
            ElseIf MeasValue.compare(LessThan, Res) And TrimPoint(i) Then
                Res = Rtn_MeasVolt.Pins(MeasPin.value).value
            End If
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site " & site & ",Code " & code(site) & ", Meas Name : " & MeasV_Name_Array(i) & ", Voltage = " & Rtn_MeasVolt.Pins(MeasPin.value).value
        Next site

        TheHdw.Digital.Patgen.Continue 0, cpuA
    Next i
    TheHdw.Digital.Patgen.HaltWait
    
    ByPassTestLimit = Previous_ByPassTestLimit_Flag
    glb_Disable_CurrRangeSetting_Print = Previous_Disable_CurrRangeSetting_Print_Flag
    
    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment ("ERROR in DSSC_Search_par_run_LDO: " & err.Description)
    DSSC_Search_par_run_LDO = TL_ERROR
End Function
Public Function Rtn_Dic_count(in_dic As Scripting.Dictionary) As Long
    On Error GoTo err

    Rtn_Dic_count = in_dic.Count
    Exit Function
err:
    Rtn_Dic_count = 0
End Function

Public Function Split_GrayDSP_to_Dec(CUS_Str_MainProgram As String, DecomposeParseDigCapBit() As String, DecomposeTestName() As String, SourceBitStrmWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave) As Long
            Dim i As Long
            Dim Index_SignedGray As Long
            Dim Index_UnSignedGray As Long
            Dim Index_2sComplement As Long
            Dim DSSC_SplitBySemiColon() As String: DSSC_SplitBySemiColon = Split(CUS_Str_MainProgram, ";")
            Dim DSSC_SignedGray() As String
            Dim DSSC_UnSignedGray() As String
            Dim DSSC_2sComplement() As String
            Dim DSPSignedGray_StartBit As New DSPWave
            Dim DSPUnSignedGray_StartBit As New DSPWave
            Dim DSP2sComplement_StartBit As New DSPWave
            Dim DSPSignedGray_StartBit_Array() As Long
            Dim DSPUnSignedGray_StartBit_Array() As Long
            Dim DSP2sComplement_StartBit_Array() As Long
            Dim AccumulateParseDigCapBit() As Long: ReDim AccumulateParseDigCapBit(UBound(DecomposeParseDigCapBit)) As Long
            
            For i = 0 To UBound(DSSC_SplitBySemiColon)
                If UCase(DSSC_SplitBySemiColon(i)) Like "SIGNEDGRAY*" Then
                    DSSC_SignedGray = Split(Split(DSSC_SplitBySemiColon(i), ":")(1), ",")
                ElseIf UCase(DSSC_SplitBySemiColon(i)) Like "UNSIGNEDGRAY*" Then
                    DSSC_UnSignedGray = Split(Split(DSSC_SplitBySemiColon(i), ":")(1), ",")
                ElseIf UCase(DSSC_SplitBySemiColon(i)) Like "2SCOMPLEMENT*" Then
                    'DSSC_2sComplement = Split(Split(DSSC_SplitBySemiColon(i), ":")(1), ",")
                End If
            Next i
            
            ReDim DSPSignedGray_StartBit_Array(UBound(DSSC_SignedGray)) As Long
            ReDim DSPUnSignedGray_StartBit_Array(UBound(DSSC_UnSignedGray)) As Long
            'ReDim DSP2sComplement_StartBit_Array(UBound(DSSC_2sComplement)) As Long
            
            For i = 0 To UBound(DecomposeTestName)
                If i = 0 Then
                    AccumulateParseDigCapBit(i) = DecomposeParseDigCapBit(i)
                Else
                    AccumulateParseDigCapBit(i) = AccumulateParseDigCapBit(i - 1) + DecomposeParseDigCapBit(i)
                End If
                
                If DecomposeTestName(i) = DSSC_SignedGray(Index_SignedGray) Then
                    DSPSignedGray_StartBit_Array(Index_SignedGray) = AccumulateParseDigCapBit(i) - DecomposeParseDigCapBit(i)
                    If Index_SignedGray <> UBound(DSSC_SignedGray) Then: Index_SignedGray = Index_SignedGray + 1
                ElseIf DecomposeTestName(i) = DSSC_UnSignedGray(Index_UnSignedGray) Then
                    DSPUnSignedGray_StartBit_Array(Index_UnSignedGray) = AccumulateParseDigCapBit(i) - DecomposeParseDigCapBit(i)
                    If Index_UnSignedGray <> UBound(DSSC_UnSignedGray) Then: Index_UnSignedGray = Index_UnSignedGray + 1
'                ElseIf DecomposeTestName(i) = DSSC_2sComplement(Index_2sComplement) Then
'                    'DSP2sComplement_StartBit_Array(Index_2sComplement) = AccumulateParseDigCapBit(i) - DecomposeParseDigCapBit(i)
'                    'If Index_2sComplement <> UBound(DSSC_2sComplement) Then: Index_2sComplement = Index_2sComplement + 1
                End If
            Next i
            If UBound(DSSC_SignedGray) = 0 And LCase(DSSC_SignedGray(0)) = "nouse" Then DSPSignedGray_StartBit_Array(0) = Instance_Data.DigCap_Sample_Size + 1
            If UBound(DSSC_UnSignedGray) = 0 And LCase(DSSC_UnSignedGray(0)) = "nouse" Then DSPUnSignedGray_StartBit_Array(0) = Instance_Data.DigCap_Sample_Size + 1
            'If UBound(DSSC_2sComplement) = 0 And LCase(DSSC_2sComplement(0)) = "nouse" Then DSP2sComplement_StartBit_Array(0) = Instance_Data.DigCap_Sample_Size + 1
            
            DSPSignedGray_StartBit.data = DSPSignedGray_StartBit_Array
            DSPUnSignedGray_StartBit.data = DSPUnSignedGray_StartBit_Array
            'DSP2sComplement_StartBit.Data = DSP2sComplement_StartBit_Array
            rundsp.Split_Gray_to_Dec DSPSignedGray_StartBit, DSPUnSignedGray_StartBit, SourceBitStrmWf, width_Wf, OutWf
            'rundsp.Split_Gray_2sComplementDSPWave_to_Dec DSPSignedGray_StartBit, DSPUnSignedGray_StartBit, DSP2sComplement_StartBit, SourceBitStrmWf, width_Wf, OutWf
End Function

Public Function Calc_delay(argc As Integer, argv() As String) As Long

    Dim i As Long
    Dim DSPWave_Dict As New DSPWave
    Dim DSPWave_GrayCode As New DSPWave
    Dim DSPWave_GrayCodeDec As New DSPWave
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    
    
    Dim meas_name As String
    Dim site As Variant
    Dim result As New SiteDouble
    Dim meas_val As New SiteDouble

    
    For i = 0 To argc - 1
        meas_name = argv(i)
        meas_val = GetStoredMeasurement(meas_name)
    
            If TheExec.TesterMode = testModeOffline Then
                meas_val = Rnd() * 1000000000000#
            End If
            
            If meas_val_delay_instance <> TheExec.DataManager.instancename Then
                meas_val_first(i) = meas_val
            Else
                For Each site In TheExec.sites
                    If meas_val = 0 Then meas_val = 0.0000000001
                    If meas_val_first(i) = 0 Then meas_val_first(i) = 0.0000000001
                    
                    result = Format(meas_val.Invert.Subtract(meas_val_first(i).Invert).Multiply(0.5), "0.00000000000000000000000000")
                Next
                meas_val_first(i) = meas_val
                TestNameInput = "Time delay F" + CStr(i + 1)
                
                TheExec.Flow.TestLimit resultVal:=result, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scalePico
                
            End If
        
    Next i
    meas_val_delay_instance = TheExec.DataManager.instancename
    
End Function
Public Function Calc_SetFlag(argc As Integer, argv() As String) As Long

    Dim i As Long
    Dim DSPWave_Dict As New DSPWave
    Dim DSPWave_GrayCode As New DSPWave
    Dim DSPWave_GrayCodeDec As New DSPWave
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    
    
    Dim meas_name As String
    Dim site As Variant
    Dim meas_val As New SiteDouble


    For i = 0 To argc - 1
        meas_name = argv(i)
        meas_val = GetStoredMeasurement(meas_name)
        For Each site In TheExec.sites
            If meas_val(site) = 0 Then TheExec.sites(site).FlagState("F_" + meas_name) = logicTrue
        Next
    Next i
    
End Function
Public Function Calc_GrayCode(argc As Integer, argv() As String) As Long

    Dim i As Long
    Dim DSPWave_Dict As New DSPWave
    Dim DSPWave_GrayCode As New DSPWave
    Dim DSPWave_GrayCodeDec As New DSPWave
    Dim TestNameInput As String
    Dim OutputTname_format() As String



    For i = 0 To argc - 1
        DSPWave_Dict = GetStoredCaptureData(argv(i))
        TestNameInput = TestNameInput & argv(i)
        Call rundsp.Transfer2GrayCode(DSPWave_Dict, DSPWave_GrayCode, DSPWave_GrayCodeDec)

        TestNameInput = Report_TName_From_Instance(CalcC, "X", "GrayCode", CInt(i))
        TheExec.Flow.TestLimit resultVal:=DSPWave_GrayCodeDec.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i
    

End Function

Public Function CMRR(argc As Integer, argv() As String) As Long

    Dim i As Long
    Dim DSPWave_Dict As New DSPWave
    Dim DSPWave_GrayCode As New DSPWave
    Dim DSPWave_GrayCodeDec As New DSPWave
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    Dim site As Variant
    Dim CMRR_Value As New SiteDouble
    Dim Voltage_Value As Double
    
    'Voltage_Value = theexec.Specs.DC.item(argv(1)).CurrentValue
    TestNameInput = Report_TName_From_Instance(Calc, "X", "CMRR")
    
    For Each site In TheExec.sites
        'CMRR_Value = GetStoredCaptureData(argv(0))
        Voltage_Value = TheExec.Specs.DC.item(argv(1)).CurrentValue(site)
        CMRR_Value = GetStoredData(argv(0) + "_para")
        
        CMRR_Value = CMRR_Value * 1.25 / (2 ^ 17)
        
        OutputTname_format = Split(TestNameInput, "_")
        OutputTname_format(6) = "CMRR"
        OutputTname_format(7) = CStr(GetStoredData(argv(0) + "_para"))
        OutputTname_format(8) = Replace(CStr(TheExec.Specs.DC.item(argv(1)).CurrentValue(site)), ".", "p")
        TestNameInput = Merge_TName(OutputTname_format)
        CMRR_Value = CMRR_Value / Voltage_Value
        
    Next
    
    TheExec.Flow.TestLimit resultVal:=CMRR_Value, Tname:=TestNameInput, ForceResults:=tlForceFlow

''    For i = 0 To argc - 1
''        For Each Site In TheExec.sites
''            DSPWave_Dict = GetStoredCaptureData(argv(i))
''            TheExec.Flow.TestLimit resultVal:=DSPWave_Dict.ConvertStreamTo(tldspParallel, 21, 0, Bit0IsMsb).Multiply(50000).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
''        Next
''    Next i
    

End Function

Public Function PSRR(argc As Integer, argv() As String) As Long

    Dim i As Long
    Dim DSPWave_Dict As New DSPWave
    Dim DSPWave_GrayCode As New DSPWave
    Dim DSPWave_GrayCodeDec As New DSPWave
    Dim TestNameInput As String
    Dim TestNameInput1 As String
    Dim OutputTname_format() As String
    Dim site As Variant
    Dim PSRR_Value As New SiteDouble
    Dim Voltage_Value As Double
    
    'Voltage_Value = theexec.Specs.DC.item(argv(1)).CurrentValue
    TestNameInput = Report_TName_From_Instance(Calc, "X", "PSRR")
    
    For Each site In TheExec.sites
        'CMRR_Value = GetStoredCaptureData(argv(0))
        Voltage_Value = TheExec.Specs.DC.item(argv(1)).CurrentValue(site)
        PSRR_Value = GetStoredData(argv(0) + "_para")
        
        PSRR_Value = PSRR_Value * 1.25 / (2 ^ 17)
        
        OutputTname_format = Split(TestNameInput, "_")
        OutputTname_format(6) = "PSRR"
        OutputTname_format(7) = CStr(GetStoredData(argv(0) + "_para"))
        OutputTname_format(8) = Replace(CStr(TheExec.Specs.DC.item(argv(1)).CurrentValue(site)), ".", "p")
        TestNameInput = Merge_TName(OutputTname_format)
        OutputTname_format(6) = "VDDIO12_MTR_GR"
        TestNameInput1 = Merge_TName(OutputTname_format)
        'PSRR_Value = PSRR_Value / Voltage_Value
        PSRR_Value = PSRR_Value.Power(-1).Multiply(0.2).Log10.Multiply(20)
    Next
    
    TheExec.Flow.TestLimit resultVal:=Voltage_Value, Tname:=TestNameInput1, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit resultVal:=PSRR_Value, Tname:=TestNameInput, ForceResults:=tlForceFlow

''    For i = 0 To argc - 1
''        For Each Site In TheExec.sites
''            DSPWave_Dict = GetStoredCaptureData(argv(i))
''            TheExec.Flow.TestLimit resultVal:=DSPWave_Dict.ConvertStreamTo(tldspParallel, 21, 0, Bit0IsMsb).Multiply(50000).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
''        Next
''    Next i
    

End Function

Public Function Calc_RXEye(argc As Integer, argv() As String) As Long
    '--- arg list ---
    ' 0:    StepSize
    ' 1:    mdll0_lsw,
    ' 2:    mdll0_msw,
    ' 3:    ddr0_dqs0_sw0,
    ' 4:    ddr0_dqs0_sw1,
    ' 5:    mdll1_lsw,
    ' 6:    mdll1_msw,
    ' 7:    ddr0_dqs1_sw0,
    ' 8:    ddr0_dqs1_sw1


    
    Dim InputKey As String
    Dim Step_Size As Integer
    
    Dim site As Variant
    Dim i As Integer
    
    Dim LSW_dspwave As New DSPWave
    Dim MSW_dspwave As New DSPWave
    Dim Combined_dspwave As New DSPWave
    Dim DecValueDspwave As New DSPWave
    
    Dim mdll_8x8 As New DSPWave
    
    
    Dim LSW_SampleSize As Integer
    Dim MSW_SampleSize As Integer
    Dim SampleSize As Integer
    
    
    '/* ------------------------------ */
    Dim mdll0 As New SiteDouble
    Dim mdll1 As New SiteDouble
    
    Dim dqs0rx_sweep As New SiteLong
    Dim dqs1rx_sweep As New SiteLong
    
    Dim ReportVal As New SiteDouble
    Dim LoVal As Double
    Dim TestNameInput As String
    Dim MaxContinuousOne As New SiteLong
    '/* ------------------------------ */
    
    
    Step_Size = val(argv(0))
    
    
    DecValueDspwave.CreateConstant 0, 1, DspDouble
    

    '/*** --------------------------------------------- ***/
    '/*** ------------------- MDLL0 ------------------- ***/
    '/*** --------------------------------------------- ***/
    
    InputKey = LCase(argv(1))
    LSW_dspwave = GetStoredCaptureData(InputKey)
    InputKey = LCase(argv(2))
    MSW_dspwave = GetStoredCaptureData(InputKey)
    
    For Each site In TheExec.sites
        LSW_SampleSize = LSW_dspwave.SampleSize
        MSW_SampleSize = MSW_dspwave.SampleSize
        SampleSize = LSW_SampleSize + MSW_SampleSize
        Exit For
    Next site
    
'    Call rundsp.CombineDSPWave(LSW_dspwave, MSW_dspwave, LSW_SampleSize, MSW_SampleSize, Combined_dspwave)
'
'    '/* ------------------ update on 2017/09/20 ------------------ */
'
'    '/* --- separate 64 bits data to 8 x 8 bits --- */
'    Call rundsp.ConvertToLongAndSerialToParrel(Combined_dspwave, 8, mdll_8x8)
'
'
    '/* ----- update on 2018//04/17 make one rundsp of " CombineDSPWave and ConvertToLongAndSerialToParrel "--------*/
    Call rundsp.CombineDSPWave_and_ConvertToLongAndSerialToParrel(LSW_dspwave, MSW_dspwave, LSW_SampleSize, MSW_SampleSize, Combined_dspwave, 8, mdll_8x8)
    
    '/* --- Calculate average of  8 x 8 bits --- */
    For Each site In TheExec.sites
        mdll0 = mdll_8x8(site).CalcMean
    Next site
    
    '/* ------------------ update on 2017/09/20 ------------------ */
    
    
    
    InputKey = LCase(argv(3))
    LSW_dspwave = GetStoredCaptureData(InputKey)
    InputKey = LCase(argv(4))
    MSW_dspwave = GetStoredCaptureData(InputKey)
    ''SampleSize = LSW_SampleSize + MSW_SampleSize
    
    Call rundsp.CombineDSPWave(LSW_dspwave, MSW_dspwave, LSW_SampleSize, MSW_SampleSize, Combined_dspwave)
    
    '/*** --------------------------------------------- ***/
    dqs0rx_sweep = 0
    MaxContinuousOne = 0
    For Each site In TheExec.sites
        For i = 0 To SampleSize - 1
            If Combined_dspwave(site).Element(i) = 1 Then
                dqs0rx_sweep = dqs0rx_sweep + 1
            Else
                '/*** Count the number of the first continuous '1' ***/
                'If dqs0rx_sweep > 0 Then
                '    Exit For
                'End If
                
                '/*** Count the number of the Max continuous '1' ***/
                If dqs0rx_sweep > MaxContinuousOne Then
                    MaxContinuousOne = dqs0rx_sweep
                    dqs0rx_sweep = 0
                End If
            End If
        Next i
        '/*** if the Combined_dspwave.Element(END) = 1 ***/
        If dqs0rx_sweep < MaxContinuousOne Then
                dqs0rx_sweep = MaxContinuousOne
        End If
        
        
    Next site
    
    'TheExec.Flow.TestLimit resultVal:=dqs0rx_sweep, Tname:="Number_of_First_Continuous_One_DQS0RX", ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit resultVal:=dqs0rx_sweep, Tname:="Number_of_Max_Continuous_One_DQS0RX", ForceResults:=tlForceNone 'eng_forceflow_transfer
    
    'dqs0rx_sweep * step_size > mdll0 / 2
    
    ReportVal = dqs0rx_sweep.Multiply(Step_Size)
    
    TheExec.Datalog.WriteComment " *** DQS0RX_Sweep x Step_Size ( " & Step_Size & " ) ***"
    
    For Each site In TheExec.sites
        LoVal = mdll0
        
        If ReportVal = 0 Then ReportVal = -1        ' update by Kaino on 2017/09/20
        
        'Report_TestLimit_by_CZ_Format resultVal:=ReportVal, lowVal:=Str(LoVal), MeasType:="C", UserVar5:="EYEDQS0", scaletype:=scaleNoScaling
        TestNameInput = Report_TName_From_Instance(CalcC, "X", "EYEDQS0", 0, , , , , tlForceNone) 'eng_forceflow_transfer
        TheExec.Flow.TestLimit resultVal:=ReportVal, lowVal:=str(LoVal), Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
    Next site
    
    
    
    '/*** --------------------------------------------- ***/
    '/*** ------------------- MDLL1 ------------------- ***/
    '/*** --------------------------------------------- ***/
    
    InputKey = LCase(argv(5))
    LSW_dspwave = GetStoredCaptureData(InputKey)
    InputKey = LCase(argv(6))
    MSW_dspwave = GetStoredCaptureData(InputKey)
    
   'SampleSize = LSW_SampleSize + MSW_SampleSize
    
'    Call rundsp.CombineDSPWave(LSW_dspwave, MSW_dspwave, LSW_SampleSize, MSW_SampleSize, Combined_dspwave)
'    '/* ------------------ update on 2017/09/20 ------------------ */
'
'    '/* --- separate 64 bits data to 8 x 8 bits --- */
'    Call rundsp.ConvertToLongAndSerialToParrel(Combined_dspwave, 8, mdll_8x8)
    
    
    '/* ----- update on 2018//04/17 make one rundsp of " CombineDSPWave and ConvertToLongAndSerialToParrel "--------*/
    Call rundsp.CombineDSPWave_and_ConvertToLongAndSerialToParrel(LSW_dspwave, MSW_dspwave, LSW_SampleSize, MSW_SampleSize, Combined_dspwave, 8, mdll_8x8)
    
    
    '/* --- Calculate average of  8 x 8 bits --- */
    For Each site In TheExec.sites
        mdll1 = mdll_8x8(site).CalcMean
    Next site
    
    '/* ------------------ update on 2017/09/20 ------------------ */
    
    
    InputKey = LCase(argv(7))
    LSW_dspwave = GetStoredCaptureData(InputKey)
    InputKey = LCase(argv(8))
    MSW_dspwave = GetStoredCaptureData(InputKey)
    
    ''SampleSize = LSW_SampleSize + MSW_SampleSize
    
    Call rundsp.CombineDSPWave(LSW_dspwave, MSW_dspwave, LSW_SampleSize, MSW_SampleSize, Combined_dspwave)
    
    dqs1rx_sweep = 0
    MaxContinuousOne = 0
    For Each site In TheExec.sites
        For i = 0 To SampleSize - 1
            If Combined_dspwave(site).Element(i) = 1 Then
                dqs1rx_sweep = dqs1rx_sweep + 1
            Else
                '/*** Count the number of the first continuous '1' ***/
                'If dqs1rx_sweep > 0 Then
                '    Exit For
                'End If
                
                '/*** Count the number of the Max continuous '1' ***/
                If dqs1rx_sweep > MaxContinuousOne Then
                    MaxContinuousOne = dqs1rx_sweep
                    dqs1rx_sweep = 0
                End If
            End If
        Next i
        
        '/*** if the Combined_dspwave.Element(END) = 1 ***/
        If dqs1rx_sweep < MaxContinuousOne Then
                dqs1rx_sweep = MaxContinuousOne
        End If
        
    Next site
    
    'TheExec.Flow.TestLimit resultVal:=dqs1rx_sweep, Tname:="Number_of_First_Continuous_One_DQS1RX", ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit resultVal:=dqs1rx_sweep, Tname:="Number_of_Max_Continuous_One_DQS1RX", ForceResults:=tlForceNone 'eng_forceflow_transfer
    
    
    
    'dqs1rx_sweep * step_size > mdll1 / 2
    
    ReportVal = dqs1rx_sweep.Multiply(Step_Size)
    
    TheExec.Datalog.WriteComment " *** DQS1RX_Sweep x Step_Size ( " & Step_Size & " ) ***"
    
    For Each site In TheExec.sites
        LoVal = mdll1
        
        If ReportVal = 0 Then ReportVal = -1        ' update by Kaino on 2017/09/20
        
        'Report_TestLimit_by_CZ_Format resultVal:=ReportVal, lowVal:=Str(LoVal), MeasType:="C", UserVar5:="EYEDQS1", scaletype:=scaleNoScaling
        TestNameInput = Report_TName_From_Instance(CalcC, "X", "EYEDQS1", 0, , , , , tlForceNone) 'eng_forceflow_transfer
        TheExec.Flow.TestLimit resultVal:=ReportVal, lowVal:=str(LoVal), Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
   
    Next site
    
End Function

Public Function Calc_AverageDSP(argc As Integer, argv() As String) As Long

    Dim Val_SerialDSP_1 As New DSPWave
    Dim Val_SerialDSP_2 As New DSPWave
'''    Dim Val_ParallelDSP_1 As New DSPWave
'''    Dim Val_ParallelDSP_2 As New DSPWave
    Dim temp As New SiteDouble
    Dim outwave As New DSPWave
    
'''    Dim SampleSize1 As Long
'''    Dim SampleSize2 As Long
'''    Dim Site As Variant

'    Val_SerialDSP_1.CreateConstant 0, 11, DspLong
'    Val_SerialDSP_2.CreateConstant 0, 11, DspLong
'    Val_ParallelDSP_1.CreateConstant 0, 1, DspLong
'    Val_ParallelDSP_2.CreateConstant 0, 1, DspLong
    
    Val_SerialDSP_1 = GetStoredCaptureData(argv(0))
    Val_SerialDSP_2 = GetStoredCaptureData(argv(1))
    
'''    For Each Site In TheExec.sites
'''        SampleSize1 = Val_SerialDSP_1(Site).SampleSize
'''        SampleSize2 = Val_SerialDSP_2(Site).SampleSize
'''        Exit For
'''    Next Site
'    For Each Site In TheExec.sites
'        SampleSize1 = Val_SerialDSP_1.SampleSize
'        SampleSize2 = Val_SerialDSP_2.SampleSize
'    Next Site

'''    Call rundsp.ConvertToLongAndSerialToParrel(Val_SerialDSP_1, SampleSize1, Val_ParallelDSP_1)
'''    Call rundsp.ConvertToLongAndSerialToParrel(Val_SerialDSP_2, SampleSize2, Val_ParallelDSP_2)
'''    Call rundsp.DSP_Add(Val_ParallelDSP_1, Val_ParallelDSP_2)
    Call rundsp.Calc_Average_DSP_Porcedure(Val_SerialDSP_1, Val_SerialDSP_2, outwave, temp)
    
'    Temp = Val_ParallelDSP_1.Element(0)
'    Temp = Temp.Divide(2)
        
    'Report_TestLimit_by_CZ_Format resultVal:=Temp, ForceResults:=tlForceFlow, MeasType:="C"
    Dim TestNameInput As String
    TestNameInput = Report_TName_From_Instance(CalcC, "X", , 0)
    TheExec.Flow.TestLimit resultVal:=temp, Tname:=TestNameInput, ForceResults:=tlForceFlow

End Function

'markchen

'ADC Calculate final efuse trim code after 85C trimming
'CDNS   => REFERENCE_CTRL_DIG = round(0.25*REFERENCE_CTRL_DIG_25 + 0.75*REFERENCE_CTRL_DIG_85)
'Sicily => ADC0_VREF_85C = round(0.25*ADC0_VREF_25C + 0.75*ADC0_VREF_85C_IM)

Public Function Calc_Dict_Store(argc As Integer, argv() As String) As Long

'Dim Dict_Store_DIG_25C As New DSPWave

 
'Dict_Store_DIG_25C = argv(0)
'Dim Dict_Store_DIG_85C As New DSPWave
'Dict_Store_DIG_85C = argv(1)


'Dict_Store_DIG_25C As String, Dict_Store_DIG_85C As String

Dim DSPWave_Dict_DIG_25C As New DSPWave
Dim DSPWave_Dict_DIG_85C As New DSPWave
Dim ADC_Trim_Code_DIG_25C As New DSPWave
Dim ADC_Trim_Code_DIG_85C As New DSPWave
Dim ADC_Trim_Code_DIG_sum As New DSPWave
Dim ADC_Trim_Code_DIG_final As New DSPWave
Dim eFuse_CTRL_DIG As New DSPWave

Dim Fuse_REFERENCE_CTRL_DIG_Name As String: Fuse_REFERENCE_CTRL_DIG_Name = argv(2)

Dim site As Variant


ADC_Trim_Code_DIG_25C.CreateConstant 0, 1, DspLong
ADC_Trim_Code_DIG_85C.CreateConstant 0, 1, DspLong
ADC_Trim_Code_DIG_sum.CreateConstant 0, 1, DspLong

DSPWave_Dict_DIG_25C = GetStoredCaptureData(argv(0))
DSPWave_Dict_DIG_85C = GetStoredCaptureData(argv(1))


Call HardIP_Bin2Dec(ADC_Trim_Code_DIG_25C, DSPWave_Dict_DIG_25C)
Call HardIP_Bin2Dec(ADC_Trim_Code_DIG_85C, DSPWave_Dict_DIG_85C)

For Each site In TheExec.sites.Active
    ADC_Trim_Code_DIG_sum(site).Element(0) = FormatNumber(ADC_Trim_Code_DIG_25C(site).Element(0) * 0.25 + ADC_Trim_Code_DIG_85C(site).Element(0) * 0.75, 0)
'        Call HardIP_Dec2Bin(ADC_Trim_Code_DIG_final, ADC_Trim_Code_DIG_sum, 8)
        
        If InStr(UCase(argv(0)), UCase("ADC0")) <> 0 Then
            TheExec.Datalog.WriteComment "site " & site & " ADC0_Trim_Code_25C :" & ADC_Trim_Code_DIG_25C(site).Element(0)
            TheExec.Datalog.WriteComment "site " & site & " ADC0_Trim_Code_85C :" & ADC_Trim_Code_DIG_85C(site).Element(0)
            TheExec.Datalog.WriteComment "site " & site & " ADC0_Trim_Code_sum :" & ADC_Trim_Code_DIG_sum(site).Element(0)
            
         ElseIf InStr(UCase(argv(0)), UCase("ADC1")) <> 0 Then
            TheExec.Datalog.WriteComment "site " & site & " ADC1_Trim_Code_25C :" & ADC_Trim_Code_DIG_25C(site).Element(0)
            TheExec.Datalog.WriteComment "site " & site & " ADC1_Trim_Code_85C :" & ADC_Trim_Code_DIG_85C(site).Element(0)
            TheExec.Datalog.WriteComment "site " & site & " ADC1_Trim_Code_sum :" & ADC_Trim_Code_DIG_sum(site).Element(0)
        
         ElseIf InStr(UCase(argv(0)), UCase("ADC2")) <> 0 Then
            TheExec.Datalog.WriteComment "site " & site & " ADC2_Trim_Code_25C :" & ADC_Trim_Code_DIG_25C(site).Element(0)
            TheExec.Datalog.WriteComment "site " & site & " ADC2_Trim_Code_85C :" & ADC_Trim_Code_DIG_85C(site).Element(0)
            TheExec.Datalog.WriteComment "site " & site & " ADC2_Trim_Code_sum :" & ADC_Trim_Code_DIG_sum(site).Element(0)
        
        End If
    
Next site
Call HardIP_Dec2Bin(ADC_Trim_Code_DIG_final, ADC_Trim_Code_DIG_sum, 8)

' Dim Data_Temp As String
Dim final_Bin2_Str1(7) As String
Dim final_Bin2_Str As String
Dim efuse_REFERENCE_CTRL_DIG_Str1(7) As String
Dim efuse_REFERENCE_CTRL_DIG_Str As String
Dim i As Integer
For Each site In TheExec.sites.Active
        For i = 0 To 7
           ' Data_Temp = Data_Temp & (ADC_Trim_Code_DIG_final(site).Element(i))
             final_Bin2_Str1(i) = CStr(ADC_Trim_Code_DIG_final(site).Element(i))
                                             
        Next i
        final_Bin2_Str = Join(final_Bin2_Str1, vbNullString)
        
        If InStr(UCase(argv(0)), UCase("ADC0")) <> 0 Then
          TheExec.Datalog.WriteComment "site " & site & " ADC0_Trim_Code_final :" & final_Bin2_Str
        ElseIf InStr(UCase(argv(0)), UCase("ADC1")) <> 0 Then
          TheExec.Datalog.WriteComment "site " & site & " ADC1_Trim_Code_final :" & final_Bin2_Str
        ElseIf InStr(UCase(argv(0)), UCase("ADC2")) <> 0 Then
          TheExec.Datalog.WriteComment "site " & site & " ADC2_Trim_Code_final :" & final_Bin2_Str
        End If
          
        final_Bin2_Str = vbNullString
       ' Data_Temp = ""
Next site

Call AddStoredCaptureData(Fuse_REFERENCE_CTRL_DIG_Name, ADC_Trim_Code_DIG_final)
TheExec.Datalog.WriteComment ("DigCap data store in dictionary " & "<<" & Fuse_REFERENCE_CTRL_DIG_Name & ">>")

eFuse_CTRL_DIG = GetStoredCaptureData(Fuse_REFERENCE_CTRL_DIG_Name)

For Each site In TheExec.sites.Active
        For i = 0 To 7
          efuse_REFERENCE_CTRL_DIG_Str1(i) = CStr(eFuse_CTRL_DIG(site).Element(i))
                                             
        Next i
        efuse_REFERENCE_CTRL_DIG_Str = Join(efuse_REFERENCE_CTRL_DIG_Str1, vbNullString)
        
        If InStr(UCase(argv(0)), UCase("ADC0")) <> 0 Then
          TheExec.Datalog.WriteComment "site " & site & " Fuse ADC0_VREF_85C :" & efuse_REFERENCE_CTRL_DIG_Str
        ElseIf InStr(UCase(argv(0)), UCase("ADC1")) <> 0 Then
          TheExec.Datalog.WriteComment "site " & site & " Fuse ADC1_VREF_85C :" & efuse_REFERENCE_CTRL_DIG_Str
        ElseIf InStr(UCase(argv(0)), UCase("ADC2")) <> 0 Then
          TheExec.Datalog.WriteComment "site " & site & " Fuse ADC2_VREF_85C :" & efuse_REFERENCE_CTRL_DIG_Str
        End If
          
        final_Bin2_Str = vbNullString
       ' Data_Temp = ""
Next site


'
'For Each site In theexec.sites.Active
'    theexec.Datalog.WriteComment "site " & site & " ADC_Trim_Code_DIG_final :" & Data_Temp
''    theexec.Datalog.WriteComment "site " & site & " ADC_Trim_Code_DIG_final :" & ADC_Trim_Code_DIG_final(site).Element(0) & ADC_Trim_Code_DIG_final(site).Element(1) _
''    & ADC_Trim_Code_DIG_final(site).Element(2) & ADC_Trim_Code_DIG_final(site).Element(3) & ADC_Trim_Code_DIG_final(site).Element(4) _
''    & ADC_Trim_Code_DIG_final(site).Element(5) & ADC_Trim_Code_DIG_final(site).Element(6) & ADC_Trim_Code_DIG_final(site).Element(7)
'Next site


End Function

'
Public Function Calc_TMPS_Coeff(argc As Integer, argv() As String) As Long

    Dim site As Variant
    
    Dim Coeff_A0_Sensor1 As New DSPWave, Coeff_A1_Sensor1 As New DSPWave, Coeff_A2_Sensor1 As New DSPWave, Coeff_A3_Sensor1 As New DSPWave, Coeff_A4_Sensor1 As New DSPWave
    Dim Coeff_A0_Sensor1_Dict As New DSPWave, Coeff_A1_Sensor1_Dict As New DSPWave, Coeff_A2_Sensor1_Dict As New DSPWave, Coeff_A3_Sensor1_Dict As New DSPWave, Coeff_A4_Sensor1_Dict As New DSPWave
    Dim DataOut_85C_Sensor1 As New DSPWave, DataOut_25C_Sensor1 As New DSPWave, DSPWave_Dict As New DSPWave
    
    Coeff_A0_Sensor1.CreateConstant 0, 1, DspLong
    Coeff_A1_Sensor1.CreateConstant 0, 1, DspLong
    Coeff_A2_Sensor1.CreateConstant 0, 1, DspLong
    Coeff_A3_Sensor1.CreateConstant 0, 1, DspLong
    Coeff_A4_Sensor1.CreateConstant 0, 1, DspLong
    DataOut_25C_Sensor1.CreateConstant 0, 1, DspLong
    DataOut_85C_Sensor1.CreateConstant 0, 1, DspLong
    
    On Error GoTo errHandler
    
    If TheExec.TesterMode = testModeOffline Then
        Set DataOut_25C_Sensor1 = Nothing
        DataOut_25C_Sensor1.CreateConstant 0, 4
    Else
        'DataOut_25C_Sensor1 = GetStoredCaptureData(argv(0))
        Call HardIP_Bin2Dec(DataOut_25C_Sensor1, GetStoredCaptureData(argv(0))) ' for Turks
    End If
    
    Call HardIP_Bin2Dec(DataOut_85C_Sensor1, GetStoredCaptureData(argv(1)))

    Call TMPS_Coeff_Calculation(Coeff_A0_Sensor1, Coeff_A1_Sensor1, Coeff_A2_Sensor1, Coeff_A3_Sensor1, Coeff_A4_Sensor1, DataOut_85C_Sensor1, DataOut_25C_Sensor1)

    Call HardIP_Dec2Bin(Coeff_A0_Sensor1_Dict, Coeff_A0_Sensor1, 15)
    Call HardIP_Dec2Bin(Coeff_A1_Sensor1_Dict, Coeff_A1_Sensor1, 14)
    Call HardIP_Dec2Bin(Coeff_A2_Sensor1_Dict, Coeff_A2_Sensor1, 12)
    Call HardIP_Dec2Bin(Coeff_A3_Sensor1_Dict, Coeff_A3_Sensor1, 10)
    Call HardIP_Dec2Bin(Coeff_A4_Sensor1_Dict, Coeff_A4_Sensor1, 11)

    Call AddStoredCaptureData(argv(2), Coeff_A0_Sensor1_Dict)
    Call AddStoredCaptureData(argv(3), Coeff_A1_Sensor1_Dict)
    Call AddStoredCaptureData(argv(4), Coeff_A2_Sensor1_Dict)
    Call AddStoredCaptureData(argv(5), Coeff_A3_Sensor1_Dict)
    Call AddStoredCaptureData(argv(6), Coeff_A4_Sensor1_Dict)
        
    Exit Function
errHandler:
        TheExec.Datalog.WriteComment "TMPS Calc Temp VBT function is error "
        TheExec.Datalog.WriteComment ("Error #: " & str(err.number) & " " & err.Description)
        If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Calc_TMPS_Coeff_1point(argc As Integer, argv() As String) As Long

    Dim site As Variant
    
    Dim Coeff_A0_Sensor1 As New DSPWave, Coeff_A1_Sensor1 As New DSPWave, Coeff_A2_Sensor1 As New DSPWave, Coeff_A3_Sensor1 As New DSPWave, Coeff_A4_Sensor1 As New DSPWave
    Dim Coeff_A0_Sensor1_Dict As New DSPWave, Coeff_A1_Sensor1_Dict As New DSPWave, Coeff_A2_Sensor1_Dict As New DSPWave, Coeff_A3_Sensor1_Dict As New DSPWave, Coeff_A4_Sensor1_Dict As New DSPWave
    Dim DataOut_85C_Sensor1 As New DSPWave, DataOut_25C_Sensor1 As New DSPWave, DSPWave_Dict As New DSPWave
    
    Coeff_A0_Sensor1.CreateConstant 0, 1, DspLong
    Coeff_A1_Sensor1.CreateConstant 0, 1, DspLong
    Coeff_A2_Sensor1.CreateConstant 0, 1, DspLong
    Coeff_A3_Sensor1.CreateConstant 0, 1, DspLong
    Coeff_A4_Sensor1.CreateConstant 0, 1, DspLong
    DataOut_25C_Sensor1.CreateConstant 0, 1, DspLong
    DataOut_85C_Sensor1.CreateConstant 0, 1, DspLong
    
    On Error GoTo errHandler
    
   '' DataOut_25C_Sensor1 = GetStoredCaptureData(argv(0))

    Call HardIP_Bin2Dec(DataOut_25C_Sensor1, GetStoredCaptureData(argv(0)))

    Call TMPS_Coeff_Calculation_1point(Coeff_A0_Sensor1, Coeff_A1_Sensor1, Coeff_A2_Sensor1, Coeff_A3_Sensor1, Coeff_A4_Sensor1, DataOut_25C_Sensor1)

    Call HardIP_Dec2Bin(Coeff_A0_Sensor1_Dict, Coeff_A0_Sensor1, 15)
    Call HardIP_Dec2Bin(Coeff_A1_Sensor1_Dict, Coeff_A1_Sensor1, 14)
    Call HardIP_Dec2Bin(Coeff_A2_Sensor1_Dict, Coeff_A2_Sensor1, 12)
    Call HardIP_Dec2Bin(Coeff_A3_Sensor1_Dict, Coeff_A3_Sensor1, 10)
    Call HardIP_Dec2Bin(Coeff_A4_Sensor1_Dict, Coeff_A4_Sensor1, 11)

    Call AddStoredCaptureData(argv(1), Coeff_A0_Sensor1_Dict)
    Call AddStoredCaptureData(argv(2), Coeff_A1_Sensor1_Dict)
    Call AddStoredCaptureData(argv(3), Coeff_A2_Sensor1_Dict)
    Call AddStoredCaptureData(argv(4), Coeff_A3_Sensor1_Dict)
    Call AddStoredCaptureData(argv(5), Coeff_A4_Sensor1_Dict)
        
    Exit Function
errHandler:
        TheExec.Datalog.WriteComment "TMPS Calc Temp VBT function is error "
        TheExec.Datalog.WriteComment ("Error #: " & str(err.number) & " " & err.Description)
        If AbortTest Then Exit Function Else Resume Next
End Function



Public Function ADDRIO_TrimCodeAverage(argc As Integer, argv() As String) As Long

    Dim i As Long, j As Long
    Dim site As Variant
    Dim DSPWave_Bin() As New DSPWave
    Dim DSPWave_Dec() As New DSPWave
    ReDim DSPWave_Bin(argc - 2) As New DSPWave
    ReDim DSPWave_Dec(argc - 2) As New DSPWave
    Dim DSPWave_AverageDec As New DSPWave
    DSPWave_AverageDec.CreateConstant 0, 1
    For i = 0 To argc - 2
        DSPWave_Bin(i) = GetStoredCaptureData(argv(i))
        Call rundsp.BinToDec(DSPWave_Bin(i), DSPWave_Dec(i))
        Call rundsp.DSP_Add(DSPWave_AverageDec, DSPWave_Dec(i))
    Next i
    Call rundsp.DSP_DivideConstant(DSPWave_AverageDec, argc - 1)
''    Call rundsp.DSP_ConvertDataTypeToLong(DSPWave_AverageDec)
    For Each site In TheExec.sites
        ''20170210-Rounding
        DSPWave_AverageDec(site).Element(0) = Int(DSPWave_AverageDec(site).Element(0) + 0.5)
    Next site

    Call AddStoredCaptureData(argv(argc - 1), DSPWave_AverageDec)
    TheExec.Flow.TestLimit resultVal:=DSPWave_AverageDec.Element(0), Tname:="ADDRIO_AverageTrimCode", ForceResults:=tlForceNone 'eng_forceflow_transfer
End Function
Public Function Calc_MDLL_Monotonicity(argc As Integer, argv() As String) As Long

    Dim i As Long, j As Long
    Dim site As Variant
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    
''    Call CreateSimulateMDLL_Data(argc, argv)
    
    Dim DSPWaveBin() As New DSPWave
    ReDim DSPWaveBin(argc - 1) As New DSPWave
    Dim DSPWaveDec() As New DSPWave
    ReDim DSPWaveDec(argc - 1) As New DSPWave
    Dim TestName As String
    TestName = argv(0) & "_"
    For i = 1 To argc - 1
        DSPWaveBin(i) = GetStoredCaptureData(argv(i))
        Call rundsp.BinToDec(DSPWaveBin(i), DSPWaveDec(i))
    Next i
    Dim dataStr As String
    For Each site In TheExec.sites
        dataStr = vbNullString
        For i = 1 To argc - 1
            If i = 1 Then
                dataStr = argv(i) & " = " & DSPWaveDec(i)(site).Element(0) & ", "
            Else
                dataStr = dataStr & argv(i) & " = " & DSPWaveDec(i)(site).Element(0) & ", "
            End If
        Next i
       If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site " & site & " " & dataStr)
    Next site
    
    Dim MDLL_CurrentVal As New SiteLong
    Dim MDLL_PreviousVal  As New SiteLong
    Dim b_MDLL_DecreaseDirection As New SiteBoolean
    Dim b_MDLL_DecreaseAddIndex As New SiteBoolean
    Dim MDLL_DecreaseResultPass As New SiteLong
    Dim b_MDLL_TestResultFail As New SiteBoolean
    Dim MDLL_Index As New SiteLong
    b_MDLL_DecreaseDirection = False
    
    MDLL_DecreaseResultPass = 1
    b_MDLL_TestResultFail = False
    MDLL_Index = 1
    Dim StepSize As Long
    For Each site In TheExec.sites
        For i = 1 To argc - 1
            If i = 1 Then
                MDLL_CurrentVal(site) = DSPWaveDec(i)(site).Element(0)
                MDLL_PreviousVal(site) = MDLL_CurrentVal(site)
            Else
                MDLL_CurrentVal(site) = DSPWaveDec(i)(site).Element(0)
                b_MDLL_DecreaseDirection(site) = MDLL_CurrentVal.Subtract(MDLL_PreviousVal).compare(LessThanOrEqualTo, 0)
                
                If b_MDLL_DecreaseDirection(site) = False Then
                    MDLL_DecreaseResultPass(site) = 0
''                    b_MDLL_TestResultFail(Site) = True
                    Exit For
                End If
                
                b_MDLL_DecreaseAddIndex(site) = MDLL_CurrentVal.Subtract(MDLL_PreviousVal).compare(LessThan, 0)
                
                If b_MDLL_DecreaseAddIndex(site) = True Then
                    MDLL_Index(site) = MDLL_Index(site) + 1
                End If
''                If MDLL_Index(Site) > 1 Then
''''                    b_MDLL_TestResultFail(Site) = True
''                    Exit For
''                End If
                
                MDLL_PreviousVal(site) = MDLL_CurrentVal(site)
            End If
        Next i
    Next site
    

    TestNameInput = Report_TName_From_Instance(CalcC, "X", "MDLLDecrease", 0)
    
    TheExec.Flow.TestLimit resultVal:=MDLL_DecreaseResultPass, lowVal:=1, hiVal:=1, Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
    For Each site In TheExec.sites
        If MDLL_DecreaseResultPass.bitwiseand(1) Then
        Else
            MDLL_Index(site) = -99
        End If
    Next site
    
    TestNameInput = Report_TName_From_Instance(CalcC, "X", "MDLLUnique", 1)
    TheExec.Flow.TestLimit resultVal:=MDLL_Index, lowVal:=1, hiVal:=2, Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
End Function
Public Function Calc_MDLL_Monotonicity_Analyze(argc As Integer, argv() As String) As Long
    Dim site As Variant
    Dim i, j, k As Integer
    Dim TestName As String
    Dim TestNameInput As String
    Dim Max_Dec_Val As New SiteLong
    Dim DSP_Decimal() As New DSPWave
    Dim DSP_Captured() As New DSPWave
    Dim OutputTname_format() As String
    
    Dim MaxDiffRank As New SiteLong
    Dim DecreaseRank As New SiteLong
    Dim Uni_DLL_Indicator As New SiteLong
    
    ReDim DSP_Decimal((argc - 1))
    ReDim DSP_Captured((argc - 1))
    
    For i = 0 To argc - 1
        DSP_Captured(i) = GetStoredCaptureData(argv(i))
        For Each site In TheExec.sites.Active
            DSP_Decimal(i) = DSP_Captured(i).ConvertStreamTo(tldspParallel, DSP_Captured(i).SampleSize, 0, Bit0IsMsb)
        Next site
    Next i
    

    For Each site In TheExec.sites.Active
        MaxDiffRank(site) = 1
        DecreaseRank(site) = 1
        Uni_DLL_Indicator(site) = 1
        MaxDiffRank(site) = DSP_Decimal(0).Element(0)                                   ' Setting compare base
        For i = 0 To UBound(DSP_Decimal) - 1
            If i <> UBound(DSP_Decimal) - 1 Then
                If DecreaseRank(site) <> 0 Then
                    If DSP_Decimal(i).Element(0) < DSP_Decimal(i + 1).Element(0) Then   ' RuleCheck1:oct0>=oct1>=oct2>=oct3>=oct4>=oct5>=oct6>=oct7
                        DecreaseRank(site) = 0
                    End If
                End If
                If MaxDiffRank(site) > DSP_Decimal(i + 1).Element(0) Then               ' Record Minimum for RuleCheck3
                    MaxDiffRank(site) = DSP_Decimal(i + 1).Element(0)
                End If
            End If
            If Uni_DLL_Indicator(site) = 1 Then                                         ' RuleCheck2:The TypeNum must be less than two type
                If DSP_Decimal(i).Element(0) = DSP_Decimal(i + 1).Element(0) Then
                    Uni_DLL_Indicator(site) = 1
                ElseIf DSP_Decimal(i).Element(0) = DSP_Decimal(i + 1).Element(0) + 1 Then
                    Uni_DLL_Indicator(site) = 2                                         ' When OTC0 > OTC1 +1, then Uni_DLL_Indicator = 2
                Else
                    Uni_DLL_Indicator(site) = -2                                        ' delta(OTC(i) - OTC(i+1) )> 1 , Uni_DLL_Indicator = -2
                End If
            ElseIf Uni_DLL_Indicator(site) = 2 Then
                If DSP_Decimal(i).Element(0) = DSP_Decimal(i + 1).Element(0) Then
                    Uni_DLL_Indicator(site) = 2
                Else
                    Uni_DLL_Indicator(site) = -1                                         ' When Uni_DLL_Indicator = 2 means there have third kind of OTC value
                End If
            End If
        Next i
        MaxDiffRank(site) = DSP_Decimal(0).Element(0) - MaxDiffRank(site)               ' RuleCheck3:Maxmun & Minimum delta must be equal one
    Next site
    
  
    Call GetFlowTName
    If gl_UseStandardTestName_Flag = True Then
        gl_Tname_Alg_Index = CStr(TheExec.Flow.TestLimitIndex)
        TestNameInput = Report_TName_From_Instance("N", "x", left(argv(0), InStr(1, argv(0), "_")) & "Decrease", CInt(gl_Tname_Alg_Index), , "qq")
        TheExec.Flow.TestLimit resultVal:=DecreaseRank, lowVal:=1, hiVal:=1, Tname:=TestNameInput, ForceResults:=tlForceFlow
        gl_Tname_Alg_Index = CStr(TheExec.Flow.TestLimitIndex)
        TestNameInput = Report_TName_From_Instance("N", "x", left(argv(0), InStr(1, argv(0), "_")) & "Unique", CInt(gl_Tname_Alg_Index))
        TheExec.Flow.TestLimit resultVal:=Uni_DLL_Indicator, lowVal:=1, hiVal:=2, Tname:=TestNameInput, ForceResults:=tlForceFlow
        gl_Tname_Alg_Index = CStr(TheExec.Flow.TestLimitIndex)
        TestNameInput = Report_TName_From_Instance("N", "x", left(argv(0), InStr(1, argv(0), "_")) & "MaxDiff", CInt(gl_Tname_Alg_Index))
        TheExec.Flow.TestLimit resultVal:=MaxDiffRank, lowVal:=0, hiVal:=1, Tname:=TestNameInput, ForceResults:=tlForceFlow
    End If
  
   
End Function

Public Function Calc_LPDPTX_FXCode(argc As Integer, argv() As String) As Long
    Dim Dict_FXCode As String
    Dim Dict_Margin_5Bit As String
    Dim Dict_Margin_1Bit As String
    Dim DSP_FXCode_Bin As New DSPWave
    Dim DSP_FXCode_Dec As New DSPWave
    Dim DSP_Margin_5Bit_Dec As New DSPWave
    Dim DSP_Margin_5Bit_Bin As New DSPWave
    Dim DSP_Margin_1Bit_Dec As New DSPWave
    Dim DSP_Margin_1Bit_Bin As New DSPWave
    Dim site As Variant
    '' ----Added to truncate FXcode 20170426---
    Dim Dict_FXCode_5Bit As String
    Dim DSP_FXCode_5Bit_Bin As New DSPWave
    ''----------------------------------------
    
    ''----Added Post_Bin and Pre_Bin Procedure----
    Dim Dict_Post_Bin As String
    Dim Dict_Post_2R As String
    Dim Dict_Pre_Bin As String
    Dim Dict_Pre_2R As String
    Dim DSP_Post_Dec As New DSPWave
    Dim DSP_Post_Bin As New DSPWave
    Dim DSP_Pre_Dec As New DSPWave
    Dim DSP_Pre_Bin As New DSPWave
    Dim DSP_Post_2R_Dec As New DSPWave
    Dim DSP_Post_2R_Bin As New DSPWave
    Dim DSP_Pre_2R_Dec As New DSPWave
    Dim DSP_Pre_2R_Bin As New DSPWave
    ''-----------------------------------------------------------
    Dict_FXCode = argv(0)
    ''Dict_Margin_5Bit = argv(1)
    ''Dict_Margin_1Bit = argv(2)
    Dict_FXCode_5Bit = argv(1)
    Dict_Post_Bin = argv(2)
    Dict_Post_2R = argv(3)
    Dict_Pre_Bin = argv(4)
    Dict_Pre_2R = argv(5)
    
    
    DSP_FXCode_Bin = GetStoredCaptureData(Dict_FXCode)
    Call rundsp.BinToDec(DSP_FXCode_Bin, DSP_FXCode_Dec)
     
'     ''Simulation
'    DSP_FXCode_Dec(0).Element(0) = 12
'    DSP_FXCode_Dec(1).Element(0) = 15
    
    ''Truncate FXCode to 5 bit
    Call rundsp.DSPWaveDecToBinary(DSP_FXCode_Dec, 5, DSP_FXCode_5Bit_Bin)
    Call AddStoredCaptureData(Dict_FXCode_5Bit, DSP_FXCode_5Bit_Bin)
    
 
    DSP_Margin_5Bit_Dec.CreateConstant 0, 1, DspDouble
    DSP_Post_Dec.CreateConstant 0, 1, DspDouble
    DSP_Pre_Dec.CreateConstant 0, 1, DspDouble
    DSP_Post_2R_Dec.CreateConstant 0, 1, DspDouble
    DSP_Pre_2R_Dec.CreateConstant 0, 1, DspDouble
    
    
    DSP_Margin_5Bit_Bin.CreateConstant 0, 5, DspLong
    DSP_Margin_1Bit_Bin.CreateConstant 0, 1, DspLong
    DSP_Post_Bin.CreateConstant 0, 4, DspLong
    DSP_Pre_Bin.CreateConstant 0, 4, DspLong
    DSP_Post_2R_Bin.CreateConstant 0, 1, DspLong
    DSP_Pre_2R_Bin.CreateConstant 0, 1, DspLong
    
    For Each site In TheExec.sites
        DSP_Margin_5Bit_Dec(site).Element(0) = (DSP_FXCode_Dec(site).Element(0) + 18) / 2
        DSP_Margin_5Bit_Dec(site).Element(0) = DSP_Margin_5Bit_Dec(site).Element(0) - DSP_FXCode_Dec(site).Element(0) ''=> Rest of Margin
        
        
        If DSP_Margin_5Bit_Dec(site).Element(0) > 6 Then
           DSP_Post_Dec(site).Element(0) = Fix(DSP_Margin_5Bit_Dec.Element(0)) ''=>Integer of Rest of Margin
           DSP_Pre_Dec(site).Element(0) = 0
           DSP_Pre_2R_Dec(site).Element(0) = 0
           
            If DSP_Margin_5Bit_Dec(site).Element(0) - Int(DSP_Margin_5Bit_Dec(site).Element(0)) = 0 Then
                DSP_Post_2R_Dec.Element(0) = 0
            Else
                DSP_Post_2R_Dec.Element(0) = 1
            End If
        Else
           DSP_Pre_Dec(site).Element(0) = Fix(DSP_Margin_5Bit_Dec.Element(0))
           DSP_Post_Dec(site).Element(0) = 0
           DSP_Post_2R_Dec(site).Element(0) = 0
            
            If DSP_Margin_5Bit_Dec(site).Element(0) - Int(DSP_Margin_5Bit_Dec(site).Element(0)) = 0 Then
                DSP_Pre_2R_Dec.Element(0) = 0
            Else
                DSP_Pre_2R_Dec.Element(0) = 1
            End If
        End If
        
'        If DSP_Margin_5Bit_Dec(Site).Element(0) - Int(DSP_Margin_5Bit_Dec(Site).Element(0)) = 0 Then
'            DSP_Margin_1Bit_Bin.Element(0) = 0
'
'        Else
'            DSP_Margin_1Bit_Bin.Element(0) = 1
'            DSP_Margin_5Bit_Dec.Element(0) = Fix(DSP_Margin_5Bit_Dec.Element(0))
'        End If
    Next site
    
    ''Call AddStoredCaptureData(Dict_Margin_1Bit, DSP_Margin_1Bit_Bin)
    
    For Each site In TheExec.sites
       '' DSP_Margin_5Bit_Dec(Site) = DSP_Margin_5Bit_Dec(Site).ConvertDataTypeTo(DspLong)
        DSP_Post_Dec(site) = DSP_Post_Dec(site).ConvertDataTypeTo(DspLong)
        DSP_Pre_Dec(site) = DSP_Pre_Dec(site).ConvertDataTypeTo(DspLong)
        DSP_Post_2R_Dec(site) = DSP_Post_2R_Dec(site).ConvertDataTypeTo(DspLong)
        DSP_Pre_2R_Dec(site) = DSP_Pre_2R_Dec(site).ConvertDataTypeTo(DspLong)
    Next site
    
    ''Call rundsp.DSPWaveDecToBinary(DSP_Margin_5Bit_Dec, 5, DSP_Margin_5Bit_Bin)
    Call rundsp.DSPWaveDecToBinary(DSP_Post_Dec, 4, DSP_Post_Bin)
    Call rundsp.DSPWaveDecToBinary(DSP_Pre_Dec, 4, DSP_Pre_Bin)
    Call rundsp.DSPWaveDecToBinary(DSP_Post_2R_Dec, 1, DSP_Post_2R_Bin)
    Call rundsp.DSPWaveDecToBinary(DSP_Pre_2R_Dec, 1, DSP_Pre_2R_Bin)


    ''Call AddStoredCaptureData(Dict_Margin_5Bit, DSP_Margin_5Bit_Bin)
    Call AddStoredCaptureData(Dict_Post_Bin, DSP_Post_Bin)
    Call AddStoredCaptureData(Dict_Pre_Bin, DSP_Pre_Bin)
    Call AddStoredCaptureData(Dict_Post_2R, DSP_Post_2R_Bin)
    Call AddStoredCaptureData(Dict_Pre_2R, DSP_Pre_2R_Bin)
End Function

Public Function Calc_ADCPLL_fuse(argc As Integer, argv() As String) As Long

    Dim i As Long, j As Long
    Dim site As Variant
    Dim DSPWave_Dict As New DSPWave
    Dim fuse_name As String
    Dim Data_Temp As String
    Dim fuse_value As New SiteLong
    Dim Dict_Name As String

''    For i = 0 To argc - 2 Step 2  'arg(0)=DSPWaveA, arg(1)=Fuse_nameA, arg(2)=DSPWaveB, arg(3)=Fuse_nameB......
    Dict_Name = argv(0)
    DSPWave_Dict = GetStoredCaptureData(Dict_Name)
    Data_Temp = vbNullString
    
    For Each site In TheExec.sites
        For j = 0 To (DSPWave_Dict(site).SampleSize - 1)
            Data_Temp = Data_Temp & (DSPWave_Dict(site).Element(j))
        Next j
        fuse_value(site) = Bin2Dec_rev(Data_Temp)
        Data_Temp = vbNullString
    Next site

    fuse_name = UCase(argv(1))
    ''Call HIP_eFuse_Write("ECID", fuse_name, Fuse_Value)
    fuse_name = vbNullString
''    Next i

End Function
Public Function Calc_GrayCodeToBin(argc As Integer, argv() As String) As Long

    Dim i As Long, j As Long
    Dim site As Variant
    Dim Dict_DSP_Bin() As New DSPWave
    ReDim Dict_DSP_Bin(argc - 1) As New DSPWave
    Dim GrayCode_DSP_Bin() As New DSPWave
    ReDim GrayCode_DSP_Bin(argc - 1) As New DSPWave
    Dim GrayCode_DSP_Dec() As New DSPWave
    ReDim GrayCode_DSP_Dec(argc - 1) As New DSPWave
    Dim b_IsUnSigned As Boolean ''New SiteBoolean
    Dim OutputTname_format() As String
    Dim TestNameInput As String
    
    b_IsUnSigned = argv(0)
    
    For i = 1 To argc - 1
        Dict_DSP_Bin(i) = GetStoredCaptureData(UCase(argv(i)))
        'Call rundsp.DSP_GrayCode2Bin(b_IsUnSigned, Dict_DSP_Bin(i), GrayCode_DSP_Bin(i), GrayCode_DSP_Dec(i))
        Call GrayCode2Bin_TTR(b_IsUnSigned, Dict_DSP_Bin(i), GrayCode_DSP_Bin(i), GrayCode_DSP_Dec(i))
        TestNameInput = Report_TName_From_Instance(CalcC, "X", vbNullString, CInt(i))
        
        TheExec.Flow.TestLimit resultVal:=GrayCode_DSP_Dec(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
        
    Next i
End Function

Public Function CalcDutyDelay(argc As Integer, argv() As String) As Long

    Dim CalcDutyVal() As New PinListData
    ReDim CalcDutyVal(argc - 1) As New PinListData
    Dim DeltaDelayVal() As New PinListData
    ReDim DeltaDelayVal(argc - 1) As New PinListData
    
    Dim i As Long, j As Long, p As Long
    Dim site As Variant
    Dim PinName As String
    Dim b_FirstTime As Boolean
    b_FirstTime = True
    Dim b_DivideZeroError As New SiteBoolean
    b_DivideZeroError = False
    
    Dim TestNameInput As String
    Dim Freq_TestName_Input As String
    Dim Voltage_Name() As String
    Voltage_Name = Split(TheExec.DataManager.instancename, "_")
    Freq_TestName_Input = argv(argc - 1)
    
    Dim MaxNumOfDuty As Long
    Dim StartNumOfDuty As Long
    StartNumOfDuty = 1
    MaxNumOfDuty = 113
    Dim OutputTname_format() As String
    
    For i = StartNumOfDuty To MaxNumOfDuty
        CalcDutyVal(i) = GetStoredMeasurement(argv(i))
        If TheExec.TesterMode = testModeOffline Then
            For j = 0 To CalcDutyVal(i).Pins.Count - 1
                CalcDutyVal(i).Pins(j) = 1000000 - 1000 * j - i * 2000
            Next j
        End If
        For j = 1 To CalcDutyVal(i).Pins.Count - 1
            If InStr(UCase(CalcDutyVal(i).Pins(j)), "_P") <> 0 Then
                PinName = CalcDutyVal(i).Pins(j)
                TestNameInput = Replace(LCase(PinName), "ddr", "ch")
                TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
                TestNameInput = TestNameInput & "_" & Freq_TestName_Input
                
                For Each site In TheExec.sites
                    If CalcDutyVal(i).Pins(j).value(site) = 0 Then
                        b_DivideZeroError(site) = True
                        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site " & site & " Freq Meas 0 Hz , No CalcDutyDelay ")
                        CalcDutyVal(i).Pins(j).value = 1
                    End If
                Next site
            
                CalcDutyVal(i).Pins(j).value = CalcDutyVal(i).Pins(j).Multiply(2).Invert
                
                For Each site In TheExec.sites
                    If b_DivideZeroError(site) = True Then
                        CalcDutyVal(i).Pins(j).value = -999
    ''                TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="Delay" & CStr(i - 1) & "_" & TestNameInput, ForceResults:=tlForceFlow
                    End If
                Next site
                TestNameInput = Report_TName_From_Instance(CalcF, CalcDutyVal(i).Pins(j), vbNullString, 0)
                TheExec.Flow.TestLimit resultVal:=CalcDutyVal(i).Pins(j), scaletype:=scalePico, Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
            End If
        Next j
    Next i
    
    '' 20170228 - Add test method for octal
    Dim Freq_Dll_Str As String
    Freq_Dll_Str = argv(0)
    Dim TCycle_Val As Double
    Dim LSB_Val As Double
    Dim Oct_Ideal_Val As Double
    Select Case UCase(Freq_Dll_Str)
        Case "DDR_F0"
            TCycle_Val = 1 / (2133.3333 * MHz)
        Case "DDR_F1"
            TCycle_Val = 1 / (1466.6667 * MHz)
        Case "DDR_F2"
            TCycle_Val = 1 / (712 * MHz)
        Case "DDR_F1M9"
            TCycle_Val = 1 / (1200 * MHz)
        Case "DDR_F2M9"
            TCycle_Val = 1 / (600 * MHz)
    End Select
    
    LSB_Val = TCycle_Val / 128
    Oct_Ideal_Val = TCycle_Val / 8
    
    Dim OctantIndex As Long
    Dim OctantMaxNum As Long
    OctantIndex = 0
    OctantMaxNum = 7
    Dim Octant_Val() As New PinListData
    ReDim Octant_Val(OctantMaxNum) As New PinListData
    
    For i = StartNumOfDuty To MaxNumOfDuty Step 16
        If OctantIndex = 7 Then
            Octant_Val(OctantIndex) = CalcDutyVal(1).Math.Subtract(CalcDutyVal(i)).Add(TCycle_Val)
        Else
            Octant_Val(OctantIndex) = CalcDutyVal(i + 16).Math.Subtract(CalcDutyVal(i))
        End If
        For j = 1 To Octant_Val(OctantIndex).Pins.Count - 1
             If InStr(UCase(Octant_Val(OctantIndex).Pins(j)), "_P") <> 0 Then
                PinName = Octant_Val(OctantIndex).Pins(j)
                TestNameInput = Replace(LCase(PinName), "ddr", "ch")
                TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
                TestNameInput = TestNameInput & "_" & Freq_TestName_Input
                
                For Each site In TheExec.sites
                    If b_DivideZeroError(site) = True Then
                        Octant_Val(OctantIndex).Pins(j).value = -999
''                        TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="Oct" & CStr(OctantIndex) & "_" & TestNameInput, ForceResults:=tlForceFlow
                    End If
                Next site
                TestNameInput = Report_TName_From_Instance(CalcF, Octant_Val(OctantIndex).Pins(j), , 0)
                TheExec.Flow.TestLimit resultVal:=Octant_Val(OctantIndex).Pins(j), scaletype:=scalePico, Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
            End If
        Next j
        OctantIndex = OctantIndex + 1
    Next i
    Dim OctPhaseError() As New PinListData
    ReDim OctPhaseError(OctantMaxNum) As New PinListData
    Dim OctPhaseError_Max As New PinListData
    Dim OctPhaseError_Min As New PinListData
    
    For i = 0 To OctantMaxNum
        OctPhaseError(i) = Octant_Val(i).Math.Subtract(Oct_Ideal_Val)
        If i = 0 Then
            OctPhaseError_Max = OctPhaseError(i)
            OctPhaseError_Min = OctPhaseError(i)
        End If
        For j = 1 To OctPhaseError(i).Pins.Count - 1
            If InStr(UCase(OctPhaseError(i).Pins(j)), "_P") <> 0 Then
                PinName = OctPhaseError(i).Pins(j)
                TestNameInput = Replace(LCase(PinName), "ddr", "ch")
                TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
                TestNameInput = TestNameInput & "_" & Freq_TestName_Input
                
                For Each site In TheExec.sites
                    If OctPhaseError(i).Pins(j).value > OctPhaseError_Max.Pins(j).value Then
                        OctPhaseError_Max.Pins(j).value = OctPhaseError(i).Pins(j).value
                    End If
                    If OctPhaseError(i).Pins(j).value < OctPhaseError_Min.Pins(j).value Then
                        OctPhaseError_Min.Pins(j).value = OctPhaseError(i).Pins(j).value
                    End If
                    If b_DivideZeroError(site) = True Then
''                        TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="PE" & CStr(i) & "_" & TestNameInput, ForceResults:=tlForceFlow
                        OctPhaseError(i).Pins(j).value = -999
                    End If
''                    Else
''                        TheExec.Flow.TestLimit resultVal:=OctPhaseError(i).Pins(j).Value, ScaleType:=scalePico, Tname:="PE" & CStr(i) & "_" & TestNameInput, ForceResults:=tlForceFlow
''                    End If
''                    If i = OctantMaxNum Then
''                        If b_DivideZeroError(Site) = True Then
''                            TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="PE_MAX" & "_" & TestNameInput, ForceResults:=tlForceFlow
''                            TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="PE_MIN" & "_" & TestNameInput, ForceResults:=tlForceFlow
''                        Else
''                            TheExec.Flow.TestLimit resultVal:=OctPhaseError_Max.Pins(j).Value, ScaleType:=scalePico, Tname:="PE_MAX" & "_" & TestNameInput, ForceResults:=tlForceFlow
''                            TheExec.Flow.TestLimit resultVal:=OctPhaseError_Min.Pins(j).Value, ScaleType:=scalePico, Tname:="PE_MIN" & "_" & TestNameInput, ForceResults:=tlForceFlow
''                        End If
''                    End If
                Next site
                TestNameInput = Report_TName_From_Instance(CalcF, OctPhaseError(i).Pins(j), , 0)
                TheExec.Flow.TestLimit resultVal:=OctPhaseError(i).Pins(j), scaletype:=scalePico, Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
''                If i = OctantMaxNum Then
''                    For Each Site In TheExec.sites
''                        If b_DivideZeroError(Site) = True Then
''                            TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="PE_MAX" & "_" & TestNameInput, ForceResults:=tlForceFlow
''                            TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="PE_MIN" & "_" & TestNameInput, ForceResults:=tlForceFlow
''                        End If
''                    Next Site
''                    TheExec.Flow.TestLimit resultVal:=OctPhaseError_Max.Pins(j), ScaleType:=scalePico, Tname:="PE_MAX" & "_" & TestNameInput, ForceResults:=tlForceFlow
''                    TheExec.Flow.TestLimit resultVal:=OctPhaseError_Min.Pins(j), ScaleType:=scalePico, Tname:="PE_MIN" & "_" & TestNameInput, ForceResults:=tlForceFlow
''
''                End If
            End If
        Next j
    Next i

    For j = 1 To OctPhaseError_Max.Pins.Count - 1
        If InStr(UCase(OctPhaseError_Max.Pins(j)), "_P") <> 0 Then
            PinName = OctPhaseError_Max.Pins(j)
            TestNameInput = Replace(LCase(PinName), "ddr", "ch")
            TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
            TestNameInput = TestNameInput & "_" & Freq_TestName_Input
            
            For Each site In TheExec.sites
                If b_DivideZeroError(site) = True Then
                    OctPhaseError_Max.Pins(j).value = -999
                    OctPhaseError_Min.Pins(j).value = -999
                End If
            Next site
            
            TestNameInput = Report_TName_From_Instance(CalcF, OctPhaseError_Max.Pins(j), vbNullString, 0)
            TheExec.Flow.TestLimit resultVal:=OctPhaseError_Max.Pins(j), scaletype:=scalePico, Tname:="PE_MAX" & "_" & TestNameInput & "_" & Voltage_Name(UBound(Voltage_Name)), ForceResults:=tlForceNone 'eng_forceflow_transfer
            
            TestNameInput = Report_TName_From_Instance(CalcF, OctPhaseError_Min.Pins(j), vbNullString, 0)
            TheExec.Flow.TestLimit resultVal:=OctPhaseError_Min.Pins(j), scaletype:=scalePico, Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
        End If
    Next j

    
    For i = StartNumOfDuty To MaxNumOfDuty
        If i = 1 Then
        Else

            DeltaDelayVal(i) = CalcDutyVal(i).Math.Subtract(CalcDutyVal(i - 1))
            For j = 1 To DeltaDelayVal(i).Pins.Count - 1
                If InStr(UCase(DeltaDelayVal(i).Pins(j)), "_P") <> 0 Then
                    PinName = DeltaDelayVal(i).Pins(j)
                    TestNameInput = Replace(LCase(PinName), "ddr", "ch")
                    TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
                    TestNameInput = TestNameInput & "_" & Freq_TestName_Input
                    
                    For Each site In TheExec.sites
                        If b_DivideZeroError(site) = True Then
''                            TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="Delta_Delay_" & CStr(i - 2) & "_" & TestNameInput, ForceResults:=tlForceFlow
                            DeltaDelayVal(i).Pins(j).value = -999
                        End If
                    Next site
                    TestNameInput = Report_TName_From_Instance(CalcF, DeltaDelayVal(i).Pins(j), vbNullString, 0)
                    TheExec.Flow.TestLimit resultVal:=DeltaDelayVal(i).Pins(j), scaletype:=scalePico, Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
                End If
            Next j
        End If
    Next i
    

    Dim DNL_Val() As New PinListData
    ReDim DNL_Val(argc - 1) As New PinListData
    Dim AryShiftNum As Long
    AryShiftNum = 2
    Dim b_Linearity_Fail As Boolean
    b_Linearity_Fail = False
    Dim DNL_Val_Max As New PinListData
    Dim DNL_Val_Min As New PinListData
    Dim No_Of_Valid_Delta_Delay As Long
    
    No_Of_Valid_Delta_Delay = 111
    
    ''20170818-Sum of DNL to be INL
    ''20170901
    Dim INL() As New PinListData
    ReDim INL(argc - 1) As New PinListData
    '' Assign pins to INL and initial value to 0
''    INL = DNL_Val(0)
''    INL = 0
    
    For i = 0 + AryShiftNum To No_Of_Valid_Delta_Delay + AryShiftNum
        DNL_Val(i) = DeltaDelayVal(i).Math.divide(LSB_Val).Subtract(1)
        
        If i = 0 + AryShiftNum Then
            DNL_Val_Max = DNL_Val(i)
            DNL_Val_Min = DNL_Val(i)
            ''20170818 -  initial INL value to 0
        End If
            INL(i) = DNL_Val(i)
            INL(i) = 0
        
        For j = 1 To DeltaDelayVal(i).Pins.Count - 1
            
            If InStr(UCase(DeltaDelayVal(i).Pins(j)), "_P") <> 0 Then
                PinName = DeltaDelayVal(i).Pins(j)
                TestNameInput = Replace(LCase(PinName), "ddr", "ch")
                TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
                TestNameInput = TestNameInput & "_" & Freq_TestName_Input
                
                For Each site In TheExec.sites
    
                   If DNL_Val(i).Pins(j).value > DNL_Val_Max.Pins(j).value Then
                       DNL_Val_Max.Pins(j).value = DNL_Val(i).Pins(j).value
                   End If
                   If DNL_Val(i).Pins(j).value < DNL_Val_Min.Pins(j).value Then
                       DNL_Val_Min.Pins(j).value = DNL_Val(i).Pins(j).value
                   End If
                   
                    If b_DivideZeroError(site) = True Then
                        DNL_Val(i).Pins(j).value = -999
                    End If
                    
                       Select Case UCase(Freq_Dll_Str)
                           Case "DDR_F0"
                               If b_DivideZeroError(site) = True Then
                                    TestNameInput = Report_TName_From_Instance(CalcF, DNL_Val(i).Pins(j), vbNullString, 0)
                                    TheExec.Flow.TestLimit resultVal:=-999, lowVal:=-1, hiVal:=1, scaletype:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
                               Else
                                    TestNameInput = Report_TName_From_Instance(CalcF, DNL_Val(i).Pins(j), vbNullString, 0)
                                    TheExec.Flow.TestLimit resultVal:=DNL_Val(i).Pins(j).value, lowVal:=-1, hiVal:=1, scaletype:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
                               End If
                               If DNL_Val(i).Pins(j).value > 1 Or DNL_Val(i).Pins(j).value < -1 Then
                                   b_Linearity_Fail = True
                               End If
                           Case "DDR_F1", "DDR_F1M9"
                               If b_DivideZeroError(site) = True Then
                                   TestNameInput = Report_TName_From_Instance(CalcF, DNL_Val(i).Pins(j), vbNullString, 0)
                                   TheExec.Flow.TestLimit resultVal:=-999, lowVal:=-1, hiVal:=1, scaletype:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
                               Else
                                   TestNameInput = Report_TName_From_Instance(CalcF, DNL_Val(i).Pins(j), vbNullString, 0)
                                   TheExec.Flow.TestLimit resultVal:=DNL_Val(i).Pins(j).value, lowVal:=-1, hiVal:=1, scaletype:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
                               End If
                               If DNL_Val(i).Pins(j).value > 1 Or DNL_Val(i).Pins(j).value < -1 Then
                                   b_Linearity_Fail = True
                               End If
                           Case "DDR_F2", "DDR_F2M9"
                               If b_DivideZeroError(site) = True Then
                                   TestNameInput = Report_TName_From_Instance(CalcF, DNL_Val(i).Pins(j), vbNullString, 0)
                                   TheExec.Flow.TestLimit resultVal:=-999, lowVal:=-1, hiVal:=1.5, scaletype:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
                               Else
                                   TestNameInput = Report_TName_From_Instance(CalcF, DNL_Val(i).Pins(j), vbNullString, 0)
                                   TheExec.Flow.TestLimit resultVal:=DNL_Val(i).Pins(j).value, lowVal:=-1, hiVal:=1.5, scaletype:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
                               End If
                               If DNL_Val(i).Pins(j).value > 1.5 Or DNL_Val(i).Pins(j).value < -1 Then
                                   b_Linearity_Fail = True
                               End If
                       End Select
                Next site
                
                ''20170818-Sum of DNL to be INL
                 If i = 0 + AryShiftNum Then
                    INL(i).Pins(j) = INL(i).Pins(j).Add(DNL_Val(i).Pins(j))
                Else
                    INL(i).Pins(j) = INL(i).Pins(j).Add(DNL_Val(i).Pins(j)).Add(INL(i - 1).Pins(j))
                End If

                ''20170830 - Bypass
'                TheExec.Flow.TestLimit resultVal:=DNL_Val(i).Pins(j), ScaleType:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", Tname:="DNL" & CStr(i - 2) & "_" & TestNameInput, ForceResults:=tlForceFlow
            End If
            
        Next j
    Next i
    
''    If i = No_Of_Valid_Delta_Delay + AryShiftNum Then
    For j = 1 To DNL_Val_Max.Pins.Count - 1
            
        If InStr(UCase(DNL_Val_Max.Pins(j)), "_P") <> 0 Then
            PinName = DNL_Val_Max.Pins(j)
            TestNameInput = Replace(LCase(PinName), "ddr", "ch")
            TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
            TestNameInput = TestNameInput & "_" & Freq_TestName_Input
            
            For Each site In TheExec.sites
                If b_DivideZeroError(site) = True Then
                    DNL_Val_Max.Pins(j).value = -999
                    DNL_Val_Min.Pins(j).value = -999
                End If
            Next site
            

            TestNameInput = Report_TName_From_Instance(CalcF, DNL_Val_Max.Pins(j), vbNullString, 0)

            TheExec.Flow.TestLimit resultVal:=DNL_Val_Max.Pins(j), Tname:=TestNameInput, scaletype:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", ForceResults:=tlForceNone 'eng_forceflow_transfer
            
            TestNameInput = Report_TName_From_Instance(CalcF, DNL_Val_Min.Pins(j), vbNullString, 0)
            TheExec.Flow.TestLimit resultVal:=DNL_Val_Min.Pins(j), Tname:=TestNameInput, scaletype:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", ForceResults:=tlForceNone 'eng_forceflow_transfer
        End If
    Next j
''    End If

    Dim INL_Val_Max As New PinListData
    Dim INL_Val_Min As New PinListData
    For i = 0 + AryShiftNum To No_Of_Valid_Delta_Delay + AryShiftNum
        If i = 0 + AryShiftNum Then
            INL_Val_Max = INL(i)
            INL_Val_Min = INL(i)
        End If
       
        For j = 1 To DeltaDelayVal(i).Pins.Count - 1
            If InStr(UCase(DeltaDelayVal(i).Pins(j)), "_P") <> 0 Then
                PinName = DeltaDelayVal(i).Pins(j)
                TestNameInput = Replace(LCase(PinName), "ddr", "ch")
                TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
                TestNameInput = TestNameInput & "_" & Freq_TestName_Input
                
                For Each site In TheExec.sites
                   If INL(i).Pins(j).value > INL_Val_Max.Pins(j).value Then
                       INL_Val_Max.Pins(j).value = INL(i).Pins(j).value
                   End If
                   If INL(i).Pins(j).value < INL_Val_Min.Pins(j).value Then
                       INL_Val_Min.Pins(j).value = INL(i).Pins(j).value
                   End If
                                     
                    TestNameInput = Report_TName_From_Instance(CalcF, INL_Val_Min.Pins(j), vbNullString, 0)
                    TheExec.Flow.TestLimit resultVal:=INL(i).Pins(j).value, scaletype:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
               Next site
             End If
             
        Next j
    Next i
    
    For j = 1 To INL_Val_Max.Pins.Count - 1
            
        If InStr(UCase(INL_Val_Max.Pins(j)), "_P") <> 0 Then
            PinName = INL_Val_Max.Pins(j)
            TestNameInput = Replace(LCase(PinName), "ddr", "ch")
            TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
            TestNameInput = TestNameInput & "_" & Freq_TestName_Input
            
            For Each site In TheExec.sites
                If b_DivideZeroError(site) = True Then
                    DNL_Val_Max.Pins(j).value = -999
                    DNL_Val_Min.Pins(j).value = -999
                End If
            Next site
            
            TestNameInput = Report_TName_From_Instance(CalcF, DNL_Val_Max.Pins(j), vbNullString, 0)
                    
            TheExec.Flow.TestLimit resultVal:=INL_Val_Max.Pins(j), Tname:=TestNameInput, scaletype:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", ForceResults:=tlForceNone 'eng_forceflow_transfer
            
            TestNameInput = Report_TName_From_Instance(CalcF, DNL_Val_Min.Pins(j), vbNullString, 0)
                    
            TheExec.Flow.TestLimit resultVal:=INL_Val_Min.Pins(j), Tname:=TestNameInput, scaletype:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", ForceResults:=tlForceNone 'eng_forceflow_transfer
        End If
    Next j
    
    '' 20170818 - Test limit for INL
''    Dim INL_Val_Max As New SiteDouble
''    Dim INL_Val_Min As New SiteDouble
''    Dim Counter As Long
''
''    For i = 0 + AryShiftNum To No_Of_Valid_Delta_Delay + AryShiftNum
''        For j = 1 To INL.Pins.Count - 1
''            If InStr(UCase(INL.Pins(j)), "_P") <> 0 Then
''                PinName = INL.Pins(j)
''                TestNameInput = Replace(LCase(PinName), "ddr", "ch")
''                TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
''                TestNameInput = TestNameInput & "_" & Freq_TestName_Input
''
''                If Counter = 0 Then
''                    INL_Val_Max = INL.Pins(j)
''                    INL_Val_Min = INL.Pins(j)
''                End If
''
''                For Each Site In TheExec.sites
''                    If INL.Pins(j).Value(Site) > INL_Val_Max(Site) Then
''                        INL_Val_Max(Site) = INL.Pins(j).Value(Site)
''                    End If
''                    If INL.Pins(j).Value(Site) < INL_Val_Min(Site) Then
''                        INL_Val_Min(Site) = INL.Pins(j).Value(Site)
''                    End If
''
''                    If b_DivideZeroError(Site) = True Then
''                        INL.Pins(j).Value = -999
''                        INL.Pins(j).Value = -999
''                    End If
''                Next Site
''
''                TheExec.Flow.TestLimit resultVal:=INL.Pins(j), Tname:="INL" & "_" & TestNameInput, ScaleType:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", ForceResults:=tlForceFlow
''                Counter = Counter + 1
''            End If
''        Next j
''    Next i
''    TheExec.Flow.TestLimit resultVal:=INL_Val_Max, Tname:="INL_MAX" & "_" & Freq_TestName_Input, ScaleType:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", ForceResults:=tlForceFlow
''    TheExec.Flow.TestLimit resultVal:=INL_Val_Min, Tname:="INL_MIN" & "_" & Freq_TestName_Input, ScaleType:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", ForceResults:=tlForceFlow

'''''    For Each Site In TheExec.Sites
'''''         If b_DivideZeroError(Site) = True Then
'''''            TheExec.Flow.TestLimit resultVal:=-999, lowVal:=False, hiVal:=False, Tname:="Linearity_Pass" & "_" & TestNameInput, ForceResults:=tlForceFlow
'''''         Else
'''''            TheExec.Flow.TestLimit resultVal:=b_Linearity_Fail, lowVal:=False, hiVal:=False, Tname:="Linearity_Pass" & "_" & TestNameInput, ForceResults:=tlForceFlow
'''''        End If
'''''    Next Site
    
End Function

Public Function CalcDutyDelay_Delta(argc As Integer, argv() As String) As Long

    Dim CalcDutyVal() As New PinListData
    ReDim CalcDutyVal(argc - 1) As New PinListData
    Dim DeltaDelayVal() As New PinListData
    ReDim DeltaDelayVal(argc - 1) As New PinListData
    Dim OutputTname_format() As String
    Dim TestNameInput As String
    Dim i As Long, j As Long
    Dim site As Variant
    Dim PinName As String
    Dim b_FirstTime As Boolean
    b_FirstTime = True
    Dim b_DivideZeroError As New SiteBoolean
    b_DivideZeroError = False
    For i = 1 To argc - 1
        CalcDutyVal(i) = GetStoredMeasurement(argv(i))
        If TheExec.TesterMode = testModeOffline Then
            For j = 0 To CalcDutyVal(i).Pins.Count - 1
                CalcDutyVal(i).Pins(j) = 1000000 - 1000 * j - i * 2000
            Next j
        End If
        'For j = 1 To CalcDutyVal(i).Pins.Count - 1 Step 2
        For j = 0 To CalcDutyVal(i).Pins.Count - 1 Step 1           'Modify 20170908
            If j Mod 4 = 2 Or j Mod 4 = 3 Then                                  'Modify 20170908
                
                PinName = CalcDutyVal(i).Pins(j)
                For Each site In TheExec.sites
    
                    If CalcDutyVal(i).Pins(j).value(site) = 0 Then
                        b_DivideZeroError(site) = True
                       If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site " & site & " Freq Meas 0 Hz , No CalcDutyDelay ")
                        CalcDutyVal(i).Pins(j).value = 1
                    End If
                    
                    CalcDutyVal(i).Pins(j).value = CalcDutyVal(i).Pins(j).Multiply(2).Invert
                    
                    If b_DivideZeroError(site) = True Then
                        TestNameInput = Report_TName_From_Instance(CalcF, CalcDutyVal(i).Pins(j), vbNullString, 0, i)
                        TheExec.Flow.TestLimit resultVal:=-999, scaletype:=scalePico, PinName:=PinName, Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
                    Else
                        TestNameInput = Report_TName_From_Instance(CalcF, CalcDutyVal(i).Pins(j), vbNullString, 0, i)
                        TheExec.Flow.TestLimit resultVal:=CalcDutyVal(i).Pins(j).value, scaletype:=scalePico, PinName:=PinName, Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
                    End If
                    
                Next site
            End If                                                          'Modify 20170908
        Next j
    Next i
    
    For i = 1 To argc - 1
        If i = 1 Then
        Else

            DeltaDelayVal(i) = CalcDutyVal(i).Math.Subtract(CalcDutyVal(i - 1))
            'For j = 1 To DeltaDelayVal(i).Pins.Count - 1 Step 2
            For j = 0 To DeltaDelayVal(i).Pins.Count - 1 Step 1     'Modify 20170908
                If j Mod 4 = 2 Or j Mod 4 = 3 Then                                  'Modify 20170908
                
                    PinName = DeltaDelayVal(i).Pins(j)
                    For Each site In TheExec.sites
                        If b_DivideZeroError(site) = True Then
                            TestNameInput = Report_TName_From_Instance(CalcF, DeltaDelayVal(i).Pins(j), vbNullString, 0, i)
                            TheExec.Flow.TestLimit resultVal:=-999, scaletype:=scalePico, PinName:=PinName, Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
                        Else
                            TestNameInput = Report_TName_From_Instance(CalcF, DeltaDelayVal(i).Pins(j), vbNullString, 0, i)
                            TheExec.Flow.TestLimit resultVal:=DeltaDelayVal(i).Pins(j).value, scaletype:=scalePico, PinName:=PinName, Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
                        End If
                    Next site
                    
                End If                      'Modify 20170908
            Next j
        End If
    Next i
    
End Function

Public Function CalcDelayJitter(argc As Integer, argv() As String) As Long
    
    Dim CalcDutyVal() As New PinListData
    ReDim CalcDutyVal(argc - 1) As New PinListData
    Dim DeltaDelayVal() As New PinListData
    ReDim DeltaDelayVal(argc - 1) As New PinListData
    
    Dim i As Long, j As Long
    Dim site As Variant

    Dim b_DivideZeroError As New SiteBoolean
    b_DivideZeroError = False
    
    Dim TestNameInput As String
    Dim TestNameFromPara As String
    Dim TestNameFreq As String
    Dim OutputTname_format() As String
    
    TestNameFromPara = argv(0)
    TestNameFromPara = LCase(left(argv(0), 3))
    If InStr(argv(0), "712") Then
        TestNameFreq = LCase(right(argv(0), 3))
    Else
        TestNameFreq = LCase(right(argv(0), 4))
    End If
    
    Dim Voltage_Name() As String
    Voltage_Name = Split(TheExec.DataManager.instancename, "_")
    
    Dim MaxNumOfDuty As Long
    Dim StartNumOfDuty As Long
    StartNumOfDuty = 1
    MaxNumOfDuty = 1
    Dim PinName As String
    For i = StartNumOfDuty To MaxNumOfDuty
        CalcDutyVal(i) = GetStoredMeasurement(argv(i))
        If TheExec.TesterMode = testModeOffline Then
            For j = 0 To CalcDutyVal(i).Pins.Count - 1
                CalcDutyVal(i).Pins(j) = 1000000 - 1000 * j - i * 2000
            Next j
        End If
        For j = 1 To CalcDutyVal(i).Pins.Count - 1
            If InStr(UCase(CalcDutyVal(i).Pins(j)), "_P") <> 0 Then
                PinName = CalcDutyVal(i).Pins(j)
                TestNameInput = Replace(LCase(PinName), "ddr", "ch")
                TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
                
                For Each site In TheExec.sites
                    If CalcDutyVal(i).Pins(j).value(site) = 0 Then
                        b_DivideZeroError(site) = True
                        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site " & site & " Freq Meas 0 Hz , No CalcDutyDelay ")
                        CalcDutyVal(i).Pins(j).value = 1
                    End If
                Next site
                    
                CalcDutyVal(i).Pins(j).value = CalcDutyVal(i).Pins(j).Multiply(2).Invert
                    
                For Each site In TheExec.sites
                    If b_DivideZeroError(site) = True Then
''                        TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="Delay" & "_" & TestNameInput, ForceResults:=tlForceFlow
                        CalcDutyVal(i).Pins(j).value = -999
                    End If
                Next site
                TestNameInput = Report_TName_From_Instance(CalcF, CalcDutyVal(i).Pins(j), vbNullString, 0)
                TheExec.Flow.TestLimit resultVal:=CalcDutyVal(i).Pins(j), scaletype:=scalePico, Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
            End If
        Next j
    Next i
End Function

Public Function CalcJitter(argc As Integer, argv() As String) As Long
    
    Dim Dict_CalcDutyVal_1 As String
    Dim Dict_CalcDutyVal_2 As String
    Dim CalcDutyVal_1 As New PinListData
    Dim CalcDutyVal_2 As New PinListData
    Dim CalcDuty_Diff As New PinListData
    Dim i As Long, j As Long
    Dim site As Variant

    Dim b_DivideZeroError As New SiteBoolean
    b_DivideZeroError = False
    
    Dim TestNameInput As String
    Dim FreqTestName As String
    Dim TestNameFromPara As String
    Dim OutputTname_format() As String
    
    TestNameInput = argv(0)
    TestNameFromPara = LCase(left(argv(0), 3))
    If InStr(TestNameInput, "712") Then
        FreqTestName = right(TestNameInput, 3)
    Else
        FreqTestName = right(TestNameInput, 4)
    End If
    
    Dim Voltage_Name() As String
    Voltage_Name = Split(TheExec.DataManager.instancename, "_")
    
    Dict_CalcDutyVal_1 = argv(1)
    Dict_CalcDutyVal_2 = argv(2)
    
    CalcDutyVal_1 = GetStoredMeasurement(Dict_CalcDutyVal_1)
    CalcDutyVal_2 = GetStoredMeasurement(Dict_CalcDutyVal_2)
    Dim PinName As String
    
    For j = 1 To CalcDutyVal_1.Pins.Count - 1
        If InStr(UCase(CalcDutyVal_1.Pins(j)), "_P") <> 0 Then
            PinName = CalcDutyVal_1.Pins(j)
            TestNameInput = Replace(LCase(PinName), "ddr", "ch")
            TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
            
            For Each site In TheExec.sites
            
                If CalcDutyVal_1.Pins(j).value(site) = 0 Then
                    b_DivideZeroError(site) = True
                    If gl_Disable_HIP_debug_log = False Then If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site " & site & " Freq Meas 0 Hz , No CalcDutyDelay ")
                    CalcDutyVal_1.Pins(j).value = 1
                End If
                If CalcDutyVal_2.Pins(j).value(site) = 0 Then
                    b_DivideZeroError(site) = True
                    TheExec.Datalog.WriteComment ("Site " & site & " Freq Meas 0 Hz , No CalcDutyDelay ")
                    CalcDutyVal_2.Pins(j).value = 1
                End If
            Next site
            
            CalcDutyVal_1.Pins(j).value = CalcDutyVal_1.Pins(j).Multiply(2).Invert
            CalcDutyVal_2.Pins(j).value = CalcDutyVal_2.Pins(j).Multiply(2).Invert
        End If
    Next j
    
    CalcDuty_Diff = CalcDutyVal_1.Math.Subtract(CalcDutyVal_2)
    
    For j = 1 To CalcDuty_Diff.Pins.Count - 1
        If InStr(UCase(CalcDuty_Diff.Pins(j)), "_P") <> 0 Then
            PinName = CalcDuty_Diff.Pins(j)
            TestNameInput = Replace(LCase(PinName), "ddr", "ch")
            TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
            For Each site In TheExec.sites
                If b_DivideZeroError(site) = True Then
''                    TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="Jitter" & "_" & TestNameInput, ForceResults:=tlForceFlow
                    CalcDuty_Diff.Pins(j).value = -999
                End If
            Next site

            TestNameInput = Report_TName_From_Instance(CalcF, CalcDuty_Diff.Pins(j), vbNullString, 0)
            TheExec.Flow.TestLimit resultVal:=CalcDuty_Diff.Pins(j), scaletype:=scalePico, Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
        End If
    Next j

End Function

Public Function Calc_2S_Complement_To_SignDec(argc As Integer, argv() As String) As Long

    Dim i As Long, j As Long
    Dim site As Variant
    Dim SplitByAt() As String
    Dim DictKey_2S_BIN As String
    Dim DictKey_SIGN_DEC As String
    
    Dim DSP_DictKey_2S_BIN As New DSPWave
    Dim DSP_DictKey_SIGN_DEC() As New DSPWave

    ReDim DSP_DictKey_SIGN_DEC(argc - 1) As New DSPWave
    
    Dim TestName As String
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    
    Dim SL_BitWidth As New SiteLong
    '' Format: Dict_2S_Com_A@Dict_SignDec_A@TestName_A,Dict_2S_Com_B@Dict_SignDec_B@TestName_B
    For i = 0 To argc - 1
        SplitByAt = Split(argv(i), "@")
        DictKey_2S_BIN = SplitByAt(0)
        DictKey_SIGN_DEC = SplitByAt(1)
        TestName = SplitByAt(2)
        
        DSP_DictKey_2S_BIN = GetStoredCaptureData(DictKey_2S_BIN)
        
''        Set DSP_DictKey_DEC = Nothing
''        DSP_DictKey_DEC.CreateConstant 0, 1, DspDouble
''        Call rundsp.BinToDec(DSP_DictKey_BIN, DSP_DictKey_DEC)
        
        For Each site In TheExec.sites
            SL_BitWidth(site) = DSP_DictKey_2S_BIN(site).SampleSize
''            DSP_DictKey_DEC(0).Element(0) = 255
''            DSP_DictKey_DEC(1).Element(0) = 254
        Next site
        
        Set DSP_DictKey_SIGN_DEC(i) = Nothing
        DSP_DictKey_SIGN_DEC(i).CreateConstant 0, 1, DspLong
        
        Call rundsp.DSP_2S_Complement_To_SignDec(DSP_DictKey_2S_BIN, SL_BitWidth, DSP_DictKey_SIGN_DEC(i))
        
        Call AddStoredCaptureData(DictKey_SIGN_DEC, DSP_DictKey_SIGN_DEC(i))
        
''        TheExec.Flow.TestLimit resultVal:=DSP_DictKey_DEC.Element(0), Tname:="DEC_" & i, ForceResults:=tlForceFlow
        
        TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(i))

        TheExec.Flow.TestLimit resultVal:=DSP_DictKey_SIGN_DEC(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
        
    Next i
End Function

Public Function Calc_TMPS_Code2Temperature(argc As Integer, argv() As String) As Long
Dim site As Variant
Dim DataOut_TemperatureCode As New DSPWave
Dim DataOut_Temperature As New SiteDouble
Dim TestNameInput As String
Dim OutputTname_format() As String
Dim Temp_Pass As New SiteLong
'Dim code(165) As Long
'Dim Temperature(165) As Long
'Dim i As Integer
DataOut_TemperatureCode.CreateConstant 0, 1, DspLong

'For i = 0 To 165
'    code(i) = Worksheets("TMPS_Table").Cells(i + 2, 2).Value
'    Temperature(i) = Worksheets("TMPS_Table").Cells(i + 2, 1).Value
'Next i

Call HardIP_Bin2Dec(DataOut_TemperatureCode, GetStoredCaptureData(argv(0)))

For Each site In TheExec.sites
    DataOut_Temperature(site) = 53.2 - 0.08942 * (DataOut_TemperatureCode(site).Element(0) - 2400) - 0.0000142 * (DataOut_TemperatureCode(site).Element(0) - 2400) ^ 2 - 0.00000000231 * (DataOut_TemperatureCode(site).Element(0) - 2400) ^ 3 - 0.000000000000416 * (DataOut_TemperatureCode(site).Element(0) - 2400) ^ 4
Next site

'For Each Site In TheExec.sites
'    If DataOut_TemperatureCode(Site).Element(0) < code(165) Then
'            DataOut_Temperature(Site) = 999
'            GoTo Lable_NextSite
'    ElseIf DataOut_TemperatureCode(Site).Element(0) > code(0) Then
'            DataOut_Temperature(Site) = -999
'            GoTo Lable_NextSite
'    End If
'
'    For i = 0 To 165
'        If DataOut_TemperatureCode(Site).Element(0) < code(i) Then
'            If DataOut_TemperatureCode(Site).Element(0) > code(i + 1) Then
'                DataOut_Temperature(Site) = Temperature(i) + (Temperature(i + 1) - Temperature(i)) * (DataOut_TemperatureCode(Site).Element(0) - code(i)) / (code(i) - code(i + 1))
'                Exit For
'            End If
'        ElseIf DataOut_TemperatureCode(Site).Element(0) = code(i) Then
'                DataOut_Temperature(Site) = Temperature(i)
'                Exit For
'        End If
'    Next i
'Lable_NextSite:
'Next Site
If TheExec.DataManager.instancename Like "*BV*" Then
    TheExec.Flow.TestLimit resultVal:=DataOut_Temperature, lowVal:=15, hiVal:=35, ForceResults:=tlForceNone 'eng_forceflow_transfer
    Update_BC_PassFail_Flag
    
    If TheExec.CurrentJob = "CP1" Then
    Else: TheHdw.Wait 0.15
    End If
Else
    TestNameInput = Report_TName_From_Instance(CalcT, "X", , 0)
    TheExec.Flow.TestLimit resultVal:=DataOut_Temperature, ForceResults:=tlForceFlow, Tname:=TestNameInput
End If

Call TMPS_Temperature2iEDA(argv(0), DataOut_Temperature)


End Function

Public Function Calc_PCIE_ADC(argc As Integer, argv() As String) As Long
Dim site As Variant
Dim DataOut_ADC_Code_0 As New DSPWave
Dim DataOut_ADC_Code_1 As New DSPWave
Dim DataOut_ADC_Code_0_OffSet As New SiteLong
Dim DataOut_ADC_Code_1_OffSet As New SiteLong
Dim DataOut_ADC_Code_OffSet_Average As New SiteLong
Dim DataOut_ADC_Code_Average As New SiteLong
Dim DataOut_ADC_Code_Average_Dict As New DSPWave
Dim DataOut_ADC_Code_Final As New SiteLong
Dim DataOut_ADC_Voltage_0 As New SiteDouble
Dim DataOut_ADC_Voltage_1 As New SiteDouble
Dim DataOut_ADC_Voltage_Average As New SiteDouble
Dim DataOut_ADC_Voltage_Out As New SiteDouble
Dim Str_Split() As String
Dim i As Integer
Dim TestNameInput As String
Dim OutputTname_format() As String

DataOut_ADC_Code_0.CreateConstant 0, 1, DspLong
DataOut_ADC_Code_1.CreateConstant 0, 1, DspLong
DataOut_ADC_Code_Average_Dict.CreateConstant 0, 1, DspLong

If argv(0) Like "*adc_offset*" Then
Else
    DataOut_ADC_Code_Average_Dict = GetStoredCaptureData("ADC_OFFSET_AVERAGE_X")
End If

Call HardIP_Bin2Dec(DataOut_ADC_Code_0, GetStoredCaptureData(argv(0)))
Call HardIP_Bin2Dec(DataOut_ADC_Code_1, GetStoredCaptureData(argv(1)))

For Each site In TheExec.sites
    DataOut_ADC_Voltage_0(site) = TheHdw.DCVS.Pins("VDD12_PCIE").Voltage.value * DataOut_ADC_Code_0(site).Element(0) / 255
    DataOut_ADC_Voltage_1(site) = TheHdw.DCVS.Pins("VDD12_PCIE").Voltage.value * DataOut_ADC_Code_1(site).Element(0) / 255
    DataOut_ADC_Voltage_Average(site) = (DataOut_ADC_Voltage_0(site) + DataOut_ADC_Voltage_1(site)) / 2
    If argv(0) Like "*adc_offset_adc*" Then
        DataOut_ADC_Code_0_OffSet(site) = DataOut_ADC_Code_0(site).Element(0) - 128
        DataOut_ADC_Code_1_OffSet(site) = DataOut_ADC_Code_1(site).Element(0) - 128
        DataOut_ADC_Code_OffSet_Average(site) = (DataOut_ADC_Code_0_OffSet(site) + DataOut_ADC_Code_1_OffSet(site)) / 2
        DataOut_ADC_Code_Average_Dict(site).Element(0) = DataOut_ADC_Code_OffSet_Average(site)
    Else
    DataOut_ADC_Code_Average(site) = (DataOut_ADC_Code_0(site).Element(0) + DataOut_ADC_Code_1(site).Element(0)) / 2
    DataOut_ADC_Code_Final(site) = DataOut_ADC_Code_Average(site) - DataOut_ADC_Code_Average_Dict(site).Element(0)
    DataOut_ADC_Voltage_Out(site) = 0.25 * TheHdw.DCVS.Pins("VDD12_PCIE").Voltage.value + DataOut_ADC_Code_Final(site) * TheHdw.DCVS.Pins("VDD12_PCIE").Voltage.value * 0.5 / 256
    End If
Next site

If argv(0) Like "*adc_offset_adc*" Then
    Call AddStoredCaptureData("ADC_OFFSET_AVERAGE_X", DataOut_ADC_Code_Average_Dict)
End If

Str_Split = Split(argv(0), "_")

TheExec.Flow.TestLimit resultVal:=DataOut_ADC_Voltage_0, Tname:="Voltage_" & argv(0), ForceResults:=tlForceFlow
TheExec.Flow.TestLimit resultVal:=DataOut_ADC_Voltage_1, Tname:="Voltage_" & argv(1), ForceResults:=tlForceFlow
TheExec.Flow.TestLimit resultVal:=DataOut_ADC_Voltage_Average, Tname:="Average_Voltage_" & Str_Split(1) & "_" & Str_Split(2) & "_adc", ForceResults:=tlForceFlow

If argv(0) Like "*adc_offset_adc*" Then
    TheExec.Flow.TestLimit resultVal:=DataOut_ADC_Code_0_OffSet, Tname:="OffSet_" & argv(0), ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit resultVal:=DataOut_ADC_Code_1_OffSet, Tname:="OffSet_" & argv(1), ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit resultVal:=DataOut_ADC_Code_OffSet_Average, Tname:="Average_OffSet_" & Str_Split(1) & "_" & Str_Split(2) & "_adc", ForceResults:=tlForceFlow
Else
    TheExec.Flow.TestLimit resultVal:=DataOut_ADC_Code_Average, Tname:="Average_" & Str_Split(1) & "_" & Str_Split(2) & "_adc", ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit resultVal:=DataOut_ADC_Code_Final, Tname:="Final_" & Str_Split(1) & "_" & Str_Split(2) & "_adc", ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit resultVal:=DataOut_ADC_Voltage_Out, Tname:="Voltage_Out_" & Str_Split(1) & "_" & Str_Split(2) & "_adc", ForceResults:=tlForceFlow
End If

End Function

Public Function Calc_LPDPRX_Bin2Hex(argc As Integer, argv() As String) As Long
Dim i As Integer
Dim Data_Temp As String
Dim DSPWave_Dict As New DSPWave: DSPWave_Dict = GetStoredCaptureData(argv(0))
Dim hex_string As String
Dim site As Variant
    For Each site In TheExec.sites
    i = DSPWave_Dict(site).SampleSize - 1
        Do While (i >= 0)
            Data_Temp = Data_Temp & (DSPWave_Dict(site).Element(i))
            i = i - 1
        Loop
        hex_string = right(BinStr2HexStr(Data_Temp, DSPWave_Dict(site).SampleSize), 8)
        
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("<@Hexadecimal Code : " & UCase(argv(0)) & "|" & site & "|" & hex_string & ">")
        
        Data_Temp = vbNullString
    Next site
End Function


Public Function TX_EQXXXXXXX(argc As Integer, argv() As String) As Long

    Dim site As Variant
    Dim i As Long
    Dim j As Long
    Dim TX_Va_EQ As New PinListData
    Dim TX_Vb_EQ As New PinListData
    Dim TX_PM_EQ As New PinListData
    Dim OutputTname_format() As String
    Dim TestNameInput As String
            
    TX_Va_EQ = GetStoredMeasurement(argv(0))
    TX_Vb_EQ = GetStoredMeasurement(argv(1))
     TX_PM_EQ = TX_Va_EQ
'    TX_Va_EQ.AddPin ("Hello")
'    TX_Vb_EQ.AddPin ("Hi")
    If TheExec.TesterMode = testModeOffline Then
    
    For i = 0 To TX_PM_EQ.Pins.Count - 1
        For Each site In TheExec.sites.Active

            TX_Va_EQ.Pins(i).value(site) = 10
            TX_Vb_EQ.Pins(i).value(site) = 10
            
        Next site
'
    Next i
    
    End If
    
   
    
        
    For i = 0 To TX_PM_EQ.Pins.Count - 1
        For Each site In TheExec.sites.Active

            TX_PM_EQ.Pins(i).value(site) = 20 * log(TX_Va_EQ.Pins(i).value(site) / TX_Vb_EQ.Pins(i).value(site))

        Next site
'
    Next i

    For j = 0 To TX_PM_EQ.Pins.Count - 1
        For Each site In TheExec.sites.Active
                TestNameInput = Report_TName_From_Instance(CalcV, TX_PM_EQ.Pins(j), vbNullString, CInt(j))
                TheExec.Flow.TestLimit resultVal:=TX_PM_EQ.Pins(j).value, Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
        Next site
    Next j




End Function

'CMRR and PSSR func modified for metrology 20170711
Public Function Calc_2S_Complement_To_SignDec_Modified(argc As Integer, argv() As String) As Long

    Dim i As Long, j As Long
    Dim site As Variant
    Dim SplitByAt() As String
    Dim DictKey_2S_BIN As String
    Dim DictKey_SIGN_DEC As String
    
    Dim DSP_DictKey_2S_BIN As New DSPWave
    Dim DSP_DictKey_SIGN_DEC() As New DSPWave
Dim DSP_CMRR_CALC() As New DSPWave
Dim DSP_PSRR_CALC() As New DSPWave
    ReDim DSP_DictKey_SIGN_DEC(argc - 1) As New DSPWave
    ReDim DSP_CMRR_CALC(argc - 1) As New DSPWave
    ReDim DSP_PSRR_CALC(argc - 1) As New DSPWave
    Dim TestName As String
    
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    Dim StepIndex_Val As Long

    
    Dim SL_BitWidth As New SiteLong
    '' Format: Dict_2S_Com_A@Dict_SignDec_A@TestName_A,Dict_2S_Com_B@Dict_SignDec_B@TestName_B
    For i = 0 To argc - 1
        
        If InStr(TheExec.DataManager.instancename, "T3") Then
        
            SplitByAt = Split(argv(i), "@")
            DictKey_2S_BIN = SplitByAt(0)
            
            DictKey_SIGN_DEC = SplitByAt(1)
            TestName = SplitByAt(UBound(SplitByAt))
     
            DSP_DictKey_2S_BIN = GetStoredCaptureData(DictKey_2S_BIN)
        
        Else
        
        
            DictKey_2S_BIN = argv(0)
            DictKey_SIGN_DEC = DictKey_2S_BIN
            TestName = DictKey_2S_BIN
            DSP_DictKey_2S_BIN = GetStoredCaptureData(DictKey_2S_BIN)
        
        End If

        
        For Each site In TheExec.sites
            SL_BitWidth(site) = DSP_DictKey_2S_BIN(site).SampleSize

        Next site
        
        Set DSP_DictKey_SIGN_DEC(i) = Nothing
        DSP_DictKey_SIGN_DEC(i).CreateConstant 0, 1, DspLong
        
        Call rundsp.DSP_2S_Complement_To_SignDec(DSP_DictKey_2S_BIN, SL_BitWidth, DSP_DictKey_SIGN_DEC(i))
        
        
         Call AddStoredCaptureData(DictKey_SIGN_DEC, DSP_DictKey_SIGN_DEC(i))
        

        If Not ByPassTestLimit Then
            
            TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(i))
            TheExec.Flow.TestLimit resultVal:=DSP_DictKey_SIGN_DEC(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
        End If

        If InStr(TheExec.DataManager.instancename, "T2P6") <> 0 Then
            
                Set DSP_CMRR_CALC(i) = Nothing
                DSP_CMRR_CALC(i).CreateConstant 0, 1, DspDouble

                Dim CMRR_VIN_Calc As Double
                CMRR_VIN_Calc = CDbl(Replace(Split(DictKey_2S_BIN, "_")(2), "p", "."))
                For Each site In TheExec.sites
                    DSP_CMRR_CALC(i)(site).Element(0) = (DSP_DictKey_SIGN_DEC(i)(site).Element(0) / 131072) * 1.25
                    DSP_CMRR_CALC(i)(site).Element(0) = DSP_CMRR_CALC(i)(site).Element(0) / CMRR_VIN_Calc
                Next site
                Call AddStoredCaptureData(DictKey_2S_BIN, DSP_CMRR_CALC(i))
                If Not ByPassTestLimit Then
                    TestNameInput = Report_TName_From_Instance(CalcC, "X", "CMRR", CInt(i))
                    TheExec.Flow.TestLimit resultVal:=DSP_CMRR_CALC(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
                End If
            
        End If
        ''=CMRR calculation END=
        ''=Osprey Metrology T2P7 PSRR calculation 20170605=
        If InStr(TheExec.DataManager.instancename, "T2P7") <> 0 Then
            
                
                Set DSP_PSRR_CALC(i) = Nothing
                DSP_PSRR_CALC(i).CreateConstant 0, 1, DspDouble
                
                For Each site In TheExec.sites
                    If DSP_DictKey_SIGN_DEC(i)(site).Element(0) = 0 Then
                        DSP_DictKey_SIGN_DEC(i)(site).Element(0) = 1
                    End If
                    DSP_PSRR_CALC(i)(site).Element(0) = Abs((DSP_DictKey_SIGN_DEC(i)(site).Element(0) / 131072) * 1.25)
                    DSP_PSRR_CALC(i)(site).Element(0) = 20 * Log10(0.2 / DSP_PSRR_CALC(i)(site).Element(0))
                     ''Osprey Metrology T2P7 PSRR avergae store 20170606
                Next site
    
                Call AddStoredCaptureData(DictKey_2S_BIN, DSP_PSRR_CALC(i))
                If Not ByPassTestLimit Then
                    TestNameInput = Report_TName_From_Instance(CalcC, "X", "PSRR", CInt(i))
                 ' Call AddStoredCaptureData(SplitByAt(2), DSP_PSRR_CALC(i))
                    If Not ByPassTestLimit Then
                            TheExec.Flow.TestLimit resultVal:=DSP_PSRR_CALC(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
                    End If
                End If
        
        ''=PSRR calculation END=
        End If
    Next i

End Function

'CMRR and PSSR func modified for metrology 20170711
Public Function Calc_2S_Complement_To_SignDec_Modified_Nolimit(argc As Integer, argv() As String) As Long

    Dim i As Long, j As Long
    Dim site As Variant
    Dim SplitByAt() As String
    Dim DictKey_2S_BIN As String
    Dim DictKey_SIGN_DEC As String
    
    Dim DSP_DictKey_2S_BIN As New DSPWave
    Dim DSP_DictKey_SIGN_DEC() As New DSPWave
Dim DSP_CMRR_CALC() As New DSPWave
Dim DSP_PSRR_CALC() As New DSPWave
    ReDim DSP_DictKey_SIGN_DEC(argc - 1) As New DSPWave
    ReDim DSP_CMRR_CALC(argc - 1) As New DSPWave
    ReDim DSP_PSRR_CALC(argc - 1) As New DSPWave
    Dim TestName As String
    
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    Dim StepIndex_Val As Long

    
    Dim SL_BitWidth As New SiteLong
    '' Format: Dict_2S_Com_A@Dict_SignDec_A@TestName_A,Dict_2S_Com_B@Dict_SignDec_B@TestName_B
    For i = 0 To argc - 1
        
        If InStr(TheExec.DataManager.instancename, "T3") Then
        
            SplitByAt = Split(argv(i), "@")
            DictKey_2S_BIN = SplitByAt(0)
            
            DictKey_SIGN_DEC = SplitByAt(1)
            TestName = SplitByAt(UBound(SplitByAt))
     
            DSP_DictKey_2S_BIN = GetStoredCaptureData(DictKey_2S_BIN)
        
        Else
        
        
            DictKey_2S_BIN = argv(0)
    
            DictKey_SIGN_DEC = DictKey_2S_BIN
            
            TestName = DictKey_2S_BIN
     
            DSP_DictKey_2S_BIN = GetStoredCaptureData(DictKey_2S_BIN)
        
        End If

        
        For Each site In TheExec.sites
            SL_BitWidth(site) = DSP_DictKey_2S_BIN(site).SampleSize

        Next site
        
        Set DSP_DictKey_SIGN_DEC(i) = Nothing
        DSP_DictKey_SIGN_DEC(i).CreateConstant 0, 1, DspLong
        
        Call rundsp.DSP_2S_Complement_To_SignDec(DSP_DictKey_2S_BIN, SL_BitWidth, DSP_DictKey_SIGN_DEC(i))
        
        
         Call AddStoredCaptureData(DictKey_SIGN_DEC, DSP_DictKey_SIGN_DEC(i))
        

    Next i

End Function


Public Function Calc_MDLL_Monotonicity_DevideBlock(argc As Integer, argv() As String) As Long

    Dim i As Long, j As Long
    Dim site As Variant
    
''    Call CreateSimulateMDLL_Data(argc, argv)
    
''    Dim DSPWaveBin() As New DSPWave
''    ReDim DSPWaveBin(argc - 1) As New DSPWave
    Dim DSPWaveDec() As New DSPWave
    ReDim DSPWaveDec((argc - 1) * 2 - 1) As New DSPWave
    Dim TestName As String
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    
    TestName = argv(0) & "_"
    
    Dim DDR_MonoWithblock() As Type_MonoWithBlock
    ReDim DDR_MonoWithblock((argc - 1) * 2 - 1) As Type_MonoWithBlock
    Dim DSP_Input As New DSPWave
    Dim DSP_Input_UpperBIN As New DSPWave
    Dim DSP_Input_BelowBIN As New DSPWave
    Dim DSP_Input_UpperDEC As New DSPWave
    Dim DSP_Input_BelowDEC As New DSPWave
    Dim InputKey As String
    For i = 0 To argc - 2
        InputKey = LCase(argv(i + 1))
        DSP_Input = GetStoredCaptureData(InputKey)
        
        Call rundsp.SeprateDSP(DSP_Input, DSP_Input_UpperBIN, DSP_Input_BelowBIN)
        Call rundsp.BinToDec(DSP_Input_UpperBIN, DSP_Input_UpperDEC)
        Call rundsp.BinToDec(DSP_Input_BelowBIN, DSP_Input_BelowDEC)
        
        If InStr(InputKey, LCase("dll_l_1")) <> 0 Then
            DDR_MonoWithblock(i * 2).Block = 4
            DDR_MonoWithblock(i * 2).DSP_Bin = DSP_Input_UpperBIN
            DDR_MonoWithblock(i * 2).DSP_Dec = DSP_Input_UpperDEC
            DDR_MonoWithblock(i * 2 + 1).Block = 0
            DDR_MonoWithblock(i * 2 + 1).DSP_Bin = DSP_Input_BelowBIN
            DDR_MonoWithblock(i * 2 + 1).DSP_Dec = DSP_Input_BelowDEC
        ElseIf InStr(InputKey, LCase("dll_l_2")) <> 0 Then
            DDR_MonoWithblock(i * 2).Block = 6
            DDR_MonoWithblock(i * 2).DSP_Bin = DSP_Input_UpperBIN
            DDR_MonoWithblock(i * 2).DSP_Dec = DSP_Input_UpperDEC
            DDR_MonoWithblock(i * 2 + 1).Block = 1
            DDR_MonoWithblock(i * 2 + 1).DSP_Bin = DSP_Input_BelowBIN
            DDR_MonoWithblock(i * 2 + 1).DSP_Dec = DSP_Input_BelowDEC
        ElseIf InStr(InputKey, LCase("dll_m_1")) <> 0 Then
            DDR_MonoWithblock(i * 2).Block = 3
            DDR_MonoWithblock(i * 2).DSP_Bin = DSP_Input_UpperBIN
            DDR_MonoWithblock(i * 2).DSP_Dec = DSP_Input_UpperDEC
            DDR_MonoWithblock(i * 2 + 1).Block = 7
            DDR_MonoWithblock(i * 2 + 1).DSP_Bin = DSP_Input_BelowBIN
            DDR_MonoWithblock(i * 2 + 1).DSP_Dec = DSP_Input_BelowDEC
        ElseIf InStr(InputKey, LCase("dll_m_2")) <> 0 Then
            DDR_MonoWithblock(i * 2).Block = 2
            DDR_MonoWithblock(i * 2).DSP_Bin = DSP_Input_UpperBIN
            DDR_MonoWithblock(i * 2).DSP_Dec = DSP_Input_UpperDEC
            DDR_MonoWithblock(i * 2 + 1).Block = 5
            DDR_MonoWithblock(i * 2 + 1).DSP_Bin = DSP_Input_BelowBIN
            DDR_MonoWithblock(i * 2 + 1).DSP_Dec = DSP_Input_BelowDEC
        End If
    Next i
    
    Dim dataStr As String
    For Each site In TheExec.sites
        For i = 0 To UBound(DDR_MonoWithblock)
            dataStr = vbNullString
            For j = 0 To DDR_MonoWithblock(i).DSP_Bin.SampleSize - 1
                If j = 0 Then
                    dataStr = DDR_MonoWithblock(i).DSP_Bin(site).Element(j)
                Else
                    dataStr = dataStr & DDR_MonoWithblock(i).DSP_Bin(site).Element(j)
                End If
            Next j
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " , Block = " & DDR_MonoWithblock(i).Block & " , Binary = " & dataStr & " , Decimal = " & DDR_MonoWithblock(i).DSP_Dec.Element(0))
        Next i
    Next site
    
    '' 20170713 - Sorting DDR_MonoWithblock by block
    Dim TempBlock As Long
    Dim sd_TempDSP_BIN As New DSPWave
    Dim sd_TempDSP_DEC As New DSPWave
    For i = 0 To UBound(DDR_MonoWithblock)
        For j = i To UBound(DDR_MonoWithblock)
            If DDR_MonoWithblock(i).Block > DDR_MonoWithblock(j).Block Then
                TempBlock = DDR_MonoWithblock(i).Block
                DDR_MonoWithblock(i).Block = DDR_MonoWithblock(j).Block
                DDR_MonoWithblock(j).Block = TempBlock
                
                sd_TempDSP_BIN = DDR_MonoWithblock(i).DSP_Bin
                DDR_MonoWithblock(i).DSP_Bin = DDR_MonoWithblock(j).DSP_Bin
                DDR_MonoWithblock(j).DSP_Bin = sd_TempDSP_BIN

                sd_TempDSP_DEC = DDR_MonoWithblock(i).DSP_Dec
                DDR_MonoWithblock(i).DSP_Dec = DDR_MonoWithblock(j).DSP_Dec
                DDR_MonoWithblock(j).DSP_Dec = sd_TempDSP_DEC
            End If
        Next j
    Next i
    
    '' Print info after sorting
    If gl_Disable_HIP_debug_log = False Then
        TheExec.Datalog.WriteComment ("Print info after sorting")
        For Each site In TheExec.sites
            For i = 0 To UBound(DDR_MonoWithblock)
                dataStr = vbNullString
                For j = 0 To DDR_MonoWithblock(i).DSP_Bin.SampleSize - 1
                    If j = 0 Then
                        dataStr = DDR_MonoWithblock(i).DSP_Bin(site).Element(j)
                    Else
                        dataStr = dataStr & DDR_MonoWithblock(i).DSP_Bin(site).Element(j)
                    End If
                Next j
                TheExec.Datalog.WriteComment ("Site_" & site & " , Block = " & DDR_MonoWithblock(i).Block & " , Binary = " & dataStr & " , Decimal = " & DDR_MonoWithblock(i).DSP_Dec.Element(0))
            Next i
        Next site
    End If
    For i = 0 To UBound(DDR_MonoWithblock)
        DSPWaveDec(i) = DDR_MonoWithblock(i).DSP_Dec
    Next i
    
    For Each site In TheExec.sites
        For i = 0 To UBound(DDR_MonoWithblock)  'NEW 20170730
             'NEW 20170730
            'TestNameInput = Report_ALG_TName_From_Instance(OutputTname_format, "C", "X", "LockCodeRange", CInt(i))
            'TheExec.Flow.TestLimit resultVal:=DSPWaveDec(i)(Site).Element(0), lowVal:=0, hiVal:=119, Tname:=TestNameInput, ForceResults:=tlForceFlow
            TestNameInput = Report_TName_From_Instance(CalcC, "X", "LockCodeRange", CInt(i))
            TheExec.Flow.TestLimit resultVal:=DSPWaveDec(i)(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
        Next i
    Next site
    TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
    
    
    Dim MDLL_CurrentVal As New SiteLong
    Dim MDLL_PreviousVal  As New SiteLong
    Dim b_MDLL_DecreaseDirection As New SiteBoolean
    Dim b_MDLL_DecreaseAddIndex As New SiteBoolean
    Dim MDLL_DecreaseResultPass As New SiteLong
    Dim b_MDLL_TestResultFail As New SiteBoolean
    Dim MDLL_Index As New SiteLong
    b_MDLL_DecreaseDirection = False
    
    MDLL_DecreaseResultPass = 1
    b_MDLL_TestResultFail = False
    MDLL_Index = 1
    Dim StepSize As Long
    Dim StoreDecreaseVal As New SiteVariant
    Dim StoreDecreaseIndex As Long
    StoreDecreaseIndex = 0
    For Each site In TheExec.sites
'       For i = 1 To argc - 1
        For i = 0 To UBound(DDR_MonoWithblock)  'NEW 20170730
            If i = 0 Then
                MDLL_CurrentVal(site) = DSPWaveDec(i)(site).Element(0)
                MDLL_PreviousVal(site) = MDLL_CurrentVal(site)
                
                StoreDecreaseVal(site) = CStr(MDLL_CurrentVal(site))
                StoreDecreaseIndex = StoreDecreaseIndex + 1
            Else
                MDLL_CurrentVal(site) = DSPWaveDec(i)(site).Element(0)
                b_MDLL_DecreaseDirection(site) = MDLL_CurrentVal.Subtract(MDLL_PreviousVal).compare(LessThanOrEqualTo, 0)
                
                '' Fail  as below
                If b_MDLL_DecreaseDirection(site) = False Then
                    MDLL_DecreaseResultPass(site) = 0
''                    b_MDLL_TestResultFail(Site) = True
''                    Exit For
                End If
                
''                b_MDLL_DecreaseAddIndex(Site) = MDLL_CurrentVal.Subtract(MDLL_PreviousVal).compare(LessThan, 0)
''
''                If b_MDLL_DecreaseAddIndex(Site) = True Then
''                    MDLL_Index(Site) = MDLL_Index(Site) + 1
''
                StoreDecreaseVal(site) = StoreDecreaseVal(site) & "," & MDLL_CurrentVal(site)
                StoreDecreaseIndex = StoreDecreaseIndex + 1
''                End If
''                If MDLL_Index(Site) > 1 Then
''''                    b_MDLL_TestResultFail(Site) = True
''                    Exit For
''                End If
                
                MDLL_PreviousVal(site) = MDLL_CurrentVal(site)
            End If
        Next i
    Next site
    
    Dim OriginalVal() As String
    Dim TempVal As Double
    Dim SortedVal() As Double
    
    Dim DiffVal_Num As New SiteLong
''    Dim DiffVal_Judge As New SiteBoolean
    Dim DiffVal_MaxSubMin As New SiteLong
    DiffVal_Num = 1
    
    For Each site In TheExec.sites
        OriginalVal = Split(StoreDecreaseVal(site), ",")
        ReDim SortedVal(UBound(OriginalVal)) As Double
        For i = 0 To UBound(OriginalVal)
            SortedVal(i) = CDbl(OriginalVal(i))
        Next i
''        SortedVal = CDbl(OriginalVal)
        For i = 0 To UBound(SortedVal)
            For j = i To UBound(SortedVal)
                If SortedVal(i) > SortedVal(j) Then
                    TempVal = SortedVal(i)
                    SortedVal(i) = SortedVal(j)
                    SortedVal(j) = TempVal
                End If
            Next j
        Next i
        For i = 0 To UBound(SortedVal) - 1
            If SortedVal(i + 1) - SortedVal(i) > 0 Then
                DiffVal_Num(site) = DiffVal_Num(site) + 1
            End If
        Next i
        DiffVal_MaxSubMin(site) = SortedVal(UBound(SortedVal)) - SortedVal(0)
    Next site

    TestNameInput = Report_TName_From_Instance(CalcC, "X", "Decrease", 0)

    TheExec.Flow.TestLimit resultVal:=MDLL_DecreaseResultPass, lowVal:=1, hiVal:=1, Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
''    For Each Site In TheExec.sites
''        If MDLL_DecreaseResultPass.BitwiseAnd(1) Then
''        Else
''            MDLL_Index(Site) = -99
''        End If
''    Next Site
    TestNameInput = Report_TName_From_Instance(CalcC, "X", "Unique", 0)
    
    TheExec.Flow.TestLimit resultVal:=DiffVal_Num, lowVal:=1, hiVal:=2, Tname:=TestName & "Unique", ForceResults:=tlForceNone 'eng_forceflow_transfer
    
    TestNameInput = Report_TName_From_Instance(CalcC, "X", "MaxDiff", 0)

    TheExec.Flow.TestLimit resultVal:=DiffVal_MaxSubMin, lowVal:=0, hiVal:=1, Tname:=TestName & "Max_Diff", ForceResults:=tlForceNone 'eng_forceflow_transfer
End Function
Public Function Calc_Metrology_GainError(argc As Integer, argv() As String) As Long
    Dim Dict_ReturnKey As String
    Dim Dict_InputKey As String
    Dim InputVal As New PinListData
    Dim CalcVal As New PinListData
    
    Dict_ReturnKey = argv(0)
    Dict_InputKey = argv(1)
    InputVal = GetStoredMeasurement(Dict_InputKey)
    
    CalcVal.AddPin (InputVal.Pins(0))
    CalcVal = InputVal.Pins(0).Subtract(0.4).divide(0.7975).Subtract(1)
    Call AddStoredMeasurement(Dict_ReturnKey, CalcVal)
End Function

Public Function Calc_MIPI_CodeTolerance(argc As Integer, argv() As String) As Long
        
    Dim i As Long, j As Long
    Dim x As Integer
    Dim site As Variant
    Dim InputDSPWave_BIN As New DSPWave
    Dim InputDSPWave_DEC As New DSPWave
    Dim MIPI_threshold_Code_value1(7) As New SiteDouble
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    
    For i = 0 To UBound(argv)
        InputDSPWave_BIN = GetStoredCaptureData(argv(i))
        Call rundsp.BinToDec(InputDSPWave_BIN, InputDSPWave_DEC)
        For Each site In TheExec.sites
            MIPI_threshold_Code_value1(i)(site) = InputDSPWave_DEC(site).Element(0)
        Next site
    Next i

    Dim MIPI_threshold_lower(0) As New SiteVariant
    Dim MIPI_threshold_high(0) As New SiteVariant
    Dim MIPI_threshold_found As New SiteBoolean
    Dim MIPI_trans_mapping As Variant
    
    Dim threshold_temp As Integer
    Dim threshold_flag1 As Boolean
    Dim p  As Long
    MIPI_trans_mapping = Array(-0.2, -0.15, -0.1, -0.05, 0.05, 0.1, 0.15, 0.2)

    x = 0

    For Each site In TheExec.sites
        threshold_temp = 0
        threshold_flag1 = False
        MIPI_threshold_found(site) = False
        
        For p = 0 To UBound(argv)

''            MIPI_threshold_Code_value1(p)(Site) = DigCapVal_DSSC_Out(0, p * 2)(Site) + 256 * DigCapVal_DSSC_Out(0, p * 2 + 1)(Site)

            If MIPI_threshold_Code_value1(p)(site) = 0 Then
                If threshold_flag1 = False Then
                    MIPI_threshold_lower(0)(site) = p
                    threshold_flag1 = True
                    MIPI_threshold_found(site) = True
                End If
                If threshold_flag1 = True Then
                    MIPI_threshold_high(0)(site) = p
                End If
            End If
            If MIPI_threshold_Code_value1(p)(site) > 0 Then
                threshold_temp = threshold_temp + 1
            End If
        Next p
        
        If threshold_temp = 0 Then
            MIPI_threshold_found = False
        End If
        
        If MIPI_threshold_lower(0)(site) <> "" Then
            MIPI_threshold_lower(0)(site) = MIPI_trans_mapping(MIPI_threshold_lower(0)(site))
        Else
            MIPI_threshold_lower(0)(site) = 999
        End If

         If MIPI_threshold_high(0)(site) <> "" Then
            MIPI_threshold_high(0)(site) = MIPI_trans_mapping(MIPI_threshold_high(0)(site))
        Else
            MIPI_threshold_high(0)(site) = 999
        End If

    Next site

    For p = 0 To 7
        TestNameInput = Report_TName_From_Instance(CalcC, "code1_" & p + 1, , CInt(x))
        
        TheExec.Flow.TestLimit MIPI_threshold_Code_value1(p), 0, 2 ^ 10 - 1, PinName:="code1_" & p + 1, ForceResults:=tlForceFlow
    Next p

    TestNameInput = Report_TName_From_Instance(CalcC, "MIPI_Tolerance1_1", , CInt(x))

    TheExec.Flow.TestLimit MIPI_threshold_lower(0), scaletype:=scaleNone, PinName:="MIPI_Tolerance1_1", ForceResults:=tlForceFlow
    'TheExec.Flow.TestLimit MIPI_threshold_lower(0), ScaleType:=None, PinName:="MIPI_Tolerance1_1", ForceResults:=tlForceFlow ''OscarLi_Compile,20190629
    TestNameInput = Report_TName_From_Instance(CalcC, "MIPI_Tolerance1_2", , CInt(x))
      
    TheExec.Flow.TestLimit MIPI_threshold_high(0), scaletype:=scaleNone, PinName:="MIPI_Tolerance1_2", ForceResults:=tlForceFlow
    'TheExec.Flow.TestLimit MIPI_threshold_high(0), ScaleType:=None, PinName:="MIPI_Tolerance1_1", ForceResults:=tlForceFlow ''OscarLi_Compile,20190629
    TestNameInput = Report_TName_From_Instance(CalcC, "MIPI_threshold_found", , CInt(x))
    
    TheExec.Flow.TestLimit MIPI_threshold_found, True, True, PinName:="MIPI_threshold_found", ForceResults:=tlForceFlow

End Function
Public Function Calc_Metrology_GainErrorOffset(argc As Integer, argv() As String) As Long

    Dim site As Variant
    Dim Dict_tfe_vol_1 As String
    Dim Dict_tfe_vol_0 As String

    Dim CapturedCode1 As String
    Dim CapturedCode2 As String
    Dim CapturedCode3 As String
    Dim CapturedCode4 As String
    Dim TestNameInput As String
    Dim OutputTname_format() As String

    Dim DSP_tfe_vol_1_in_decimal As New DSPWave
    Dim DSP_tfe_vol_1_in_binary As New DSPWave


    Dim SL_BitWidth As New SiteLong
    
    Dim x As Long

    Dict_tfe_vol_1 = argv(0)
    CapturedCode1 = argv(1)
    CapturedCode2 = argv(2)
    CapturedCode3 = argv(3)
    CapturedCode4 = argv(4)
    Dict_tfe_vol_0 = argv(5)


    Dim DSP_tfe_vol_0_in_2S_binary As New DSPWave
    Dim DSP_tfe_vol_0_in_decimal As New DSPWave

    Dim DSP_gainErrorOffset1 As New DSPWave
    Dim DSP_gainErrorOffset2 As New DSPWave
    Dim DSP_gainErrorOffset3 As New DSPWave
    Dim DSP_gainErrorOffset4 As New DSPWave

    Dim DSP_gainErrorOffset1_decimal As New DSPWave
    Dim DSP_gainErrorOffset2_decimal As New DSPWave
    Dim DSP_gainErrorOffset3_decimal As New DSPWave
    Dim DSP_gainErrorOffset4_decimal As New DSPWave

    x = 0

    DSP_gainErrorOffset1 = GetStoredCaptureData(CapturedCode1)
    DSP_gainErrorOffset2 = GetStoredCaptureData(CapturedCode2)
    DSP_gainErrorOffset3 = GetStoredCaptureData(CapturedCode3)
    DSP_gainErrorOffset4 = GetStoredCaptureData(CapturedCode4)

    DSP_tfe_vol_0_in_2S_binary = GetStoredCaptureData(Dict_tfe_vol_0)
    For Each site In TheExec.sites
            SL_BitWidth(site) = DSP_tfe_vol_0_in_2S_binary(site).SampleSize
            
            'Test Run
            
'            '111111111111111000
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(0) = 0
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(1) = 0
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(2) = 0
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(3) = 1
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(4) = 1
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(5) = 1
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(6) = 1
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(7) = 1
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(8) = 1
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(9) = 1
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(10) = 1
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(11) = 1
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(12) = 1
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(13) = 1
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(14) = 1
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(15) = 1
'             DSP_tfe_vol_0_in_2S_binary(Site).Element(16) = 1
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(17) = 1
'
'
'            '001000000000001110
'            DSP_gainErrorOffset1(Site).Element(0) = 0
'            DSP_gainErrorOffset1(Site).Element(1) = 1
'            DSP_gainErrorOffset1(Site).Element(2) = 1
'            DSP_gainErrorOffset1(Site).Element(3) = 1
'            DSP_gainErrorOffset1(Site).Element(4) = 0
'            DSP_gainErrorOffset1(Site).Element(5) = 0
'            DSP_gainErrorOffset1(Site).Element(6) = 0
'            DSP_gainErrorOffset1(Site).Element(7) = 0
'            DSP_gainErrorOffset1(Site).Element(8) = 0
'            DSP_gainErrorOffset1(Site).Element(9) = 0
'            DSP_gainErrorOffset1(Site).Element(10) = 0
'            DSP_gainErrorOffset1(Site).Element(11) = 0
'            DSP_gainErrorOffset1(Site).Element(12) = 0
'            DSP_gainErrorOffset1(Site).Element(13) = 0
'            DSP_gainErrorOffset1(Site).Element(14) = 0
'            DSP_gainErrorOffset1(Site).Element(15) = 1
'             DSP_gainErrorOffset1(Site).Element(16) = 0
'            DSP_gainErrorOffset1(Site).Element(17) = 0
'
'
'            '000111111111101011
'            DSP_gainErrorOffset2(Site).Element(0) = 1
'            DSP_gainErrorOffset2(Site).Element(1) = 1
'            DSP_gainErrorOffset2(Site).Element(2) = 0
'            DSP_gainErrorOffset2(Site).Element(3) = 1
'            DSP_gainErrorOffset2(Site).Element(4) = 0
'            DSP_gainErrorOffset2(Site).Element(5) = 1
'            DSP_gainErrorOffset2(Site).Element(6) = 1
'            DSP_gainErrorOffset2(Site).Element(7) = 1
'            DSP_gainErrorOffset2(Site).Element(8) = 1
'            DSP_gainErrorOffset2(Site).Element(9) = 1
'            DSP_gainErrorOffset2(Site).Element(10) = 1
'            DSP_gainErrorOffset2(Site).Element(11) = 1
'            DSP_gainErrorOffset2(Site).Element(12) = 1
'            DSP_gainErrorOffset2(Site).Element(13) = 1
'            DSP_gainErrorOffset2(Site).Element(14) = 1
'            DSP_gainErrorOffset2(Site).Element(15) = 0
'             DSP_gainErrorOffset2(Site).Element(16) = 0
'            DSP_gainErrorOffset2(Site).Element(17) = 0
'
'            '000111111111100110
'            DSP_gainErrorOffset3(Site).Element(0) = 0
'            DSP_gainErrorOffset3(Site).Element(1) = 1
'            DSP_gainErrorOffset3(Site).Element(2) = 1
'            DSP_gainErrorOffset3(Site).Element(3) = 0
'            DSP_gainErrorOffset3(Site).Element(4) = 0
'            DSP_gainErrorOffset3(Site).Element(5) = 1
'            DSP_gainErrorOffset3(Site).Element(6) = 1
'            DSP_gainErrorOffset3(Site).Element(7) = 1
'            DSP_gainErrorOffset3(Site).Element(8) = 1
'            DSP_gainErrorOffset3(Site).Element(9) = 1
'            DSP_gainErrorOffset3(Site).Element(10) = 1
'            DSP_gainErrorOffset3(Site).Element(11) = 1
'            DSP_gainErrorOffset3(Site).Element(12) = 1
'            DSP_gainErrorOffset3(Site).Element(13) = 1
'            DSP_gainErrorOffset3(Site).Element(14) = 1
'            DSP_gainErrorOffset3(Site).Element(15) = 0
'             DSP_gainErrorOffset3(Site).Element(16) = 0
'            DSP_gainErrorOffset3(Site).Element(17) = 0
'
'
'            '001000000000011010
'            DSP_gainErrorOffset4(Site).Element(0) = 0
'            DSP_gainErrorOffset4(Site).Element(1) = 1
'            DSP_gainErrorOffset4(Site).Element(2) = 0
'            DSP_gainErrorOffset4(Site).Element(3) = 1
'            DSP_gainErrorOffset4(Site).Element(4) = 1
'            DSP_gainErrorOffset4(Site).Element(5) = 0
'            DSP_gainErrorOffset4(Site).Element(6) = 0
'            DSP_gainErrorOffset4(Site).Element(7) = 0
'            DSP_gainErrorOffset4(Site).Element(8) = 0
'            DSP_gainErrorOffset4(Site).Element(9) = 0
'            DSP_gainErrorOffset4(Site).Element(10) = 0
'            DSP_gainErrorOffset4(Site).Element(11) = 0
'            DSP_gainErrorOffset4(Site).Element(12) = 0
'            DSP_gainErrorOffset4(Site).Element(13) = 0
'            DSP_gainErrorOffset4(Site).Element(14) = 0
'            DSP_gainErrorOffset4(Site).Element(15) = 1
'             DSP_gainErrorOffset4(Site).Element(16) = 0
'            DSP_gainErrorOffset4(Site).Element(17) = 0
            
            
            

    Next site
    DSP_tfe_vol_0_in_decimal.CreateConstant 0, 1, DspLong
    
    

    Call rundsp.DSP_2S_Complement_To_SignDec(DSP_tfe_vol_0_in_2S_binary, SL_BitWidth, DSP_tfe_vol_0_in_decimal)

    TestNameInput = Report_TName_From_Instance(CalcC, "X", "tfe_vol_0", CInt(x))

    TheExec.Flow.TestLimit resultVal:=DSP_tfe_vol_0_in_decimal.Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer

    Call rundsp.BinToDec(DSP_gainErrorOffset1, DSP_gainErrorOffset1_decimal)
    Call rundsp.BinToDec(DSP_gainErrorOffset2, DSP_gainErrorOffset2_decimal)
    Call rundsp.BinToDec(DSP_gainErrorOffset3, DSP_gainErrorOffset3_decimal)
    Call rundsp.BinToDec(DSP_gainErrorOffset4, DSP_gainErrorOffset4_decimal)

    DSP_tfe_vol_1_in_decimal.CreateConstant 0, 1, DspLong

    TestNameInput = Report_TName_From_Instance(CalcC, "X", "CapCode1_Dec", CInt(x))

    TheExec.Flow.TestLimit resultVal:=DSP_gainErrorOffset1_decimal.Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
    
    TestNameInput = Report_TName_From_Instance(CalcC, "X", "CapCode2_Dec", CInt(x))
    
    TheExec.Flow.TestLimit resultVal:=DSP_gainErrorOffset2_decimal.Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
    
    TestNameInput = Report_TName_From_Instance(CalcC, "X", "CapCode3_Dec", CInt(x))
    
    TheExec.Flow.TestLimit resultVal:=DSP_gainErrorOffset3_decimal.Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
    
    TestNameInput = Report_TName_From_Instance(CalcC, "X", "CapCode4_Dec", CInt(x))
    
    TheExec.Flow.TestLimit resultVal:=DSP_gainErrorOffset4_decimal.Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer

    
    TestNameInput = Report_TName_From_Instance(CalcC, "X", "Gain_vol_1_LimitExceeded_Dec", CInt(x))
    
    For Each site In TheExec.sites


        DSP_tfe_vol_1_in_decimal(site).Element(0) = DSP_gainErrorOffset1_decimal(site).Element(0) + DSP_gainErrorOffset2_decimal(site).Element(0) + DSP_gainErrorOffset3_decimal(site).Element(0) + DSP_gainErrorOffset4_decimal(site).Element(0) - 4 * DSP_tfe_vol_0_in_decimal(site).Element(0)
        If (DSP_tfe_vol_1_in_decimal(site).Element(0) > 262143) Then
            TheExec.Datalog.WriteComment ("Site:" + CStr(site) + "  Gain_vol_1_LimitExceeded_Dec = " + CStr(DSP_tfe_vol_1_in_decimal(site).Element(0)) + ", Force tfe_vol_1_dec = 174762")
            DSP_tfe_vol_1_in_decimal(site).Element(0) = 174762
        End If

    Next site
    
    TestNameInput = Report_TName_From_Instance(CalcC, "X", "tfe_vol_1_dec", CInt(x))
    
    TheExec.Flow.TestLimit resultVal:=DSP_tfe_vol_1_in_decimal.Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
    

    Call rundsp.DSPWf_Dec2Binary(DSP_tfe_vol_1_in_decimal, 18, DSP_tfe_vol_1_in_binary)

    Call AddStoredCaptureData(Dict_tfe_vol_1, DSP_tfe_vol_1_in_binary)
    
    Dim tfe_vol_1_bin_str As String
    Dim i As Long
   
    For Each site In TheExec.sites

            tfe_vol_1_bin_str = vbNullString
         For i = DSP_tfe_vol_1_in_binary(site).SampleSize - 1 To 0 Step -1
         
                tfe_vol_1_bin_str = tfe_vol_1_bin_str + CStr(DSP_tfe_vol_1_in_binary(site).Element(i))
            
         Next i
      
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site:" + CStr(site) + "  tfe_vol_1_binary_fuse_Code  " + tfe_vol_1_bin_str)
    Next site
     
   ' TheExec.Flow.TestLimit resultVal:=DSP_tfe_vol_1_in_binary.Element(0), Tname:="tfe_vol_1", ForceResults:=tlForceFlow


End Function
Public Function Calc_Metrology_EncodeActualTemp(argc As Integer, argv() As String) As Long

Dim actual_temp As Double
Dim Dict_actual_temp As String
Dim site As Variant

actual_temp = CDbl(argv(0))
Dict_actual_temp = argv(1)


Dim actual_temp_cal1 As Double
actual_temp_cal1 = (actual_temp - 25) * 64

Dim conv_temp_rounded As Long

conv_temp_rounded = FormatNumber(actual_temp_cal1)


Dim DSP_conv_temp_rounded As New DSPWave
Dim DSP_conv_temp_rounded_binary As New DSPWave

DSP_conv_temp_rounded.CreateConstant 0, 1, DspLong
DSP_conv_temp_rounded_binary.CreateConstant 0, 1, DspLong


For Each site In TheExec.sites

DSP_conv_temp_rounded(site).Element(0) = conv_temp_rounded

Next site

Call rundsp.DSPWf_Dec2Binary(DSP_conv_temp_rounded, 10, DSP_conv_temp_rounded_binary)

Call AddStoredCaptureData(Dict_actual_temp, DSP_conv_temp_rounded_binary)


End Function

Public Function Calc_Metrology_DecodeActualTemp(argc As Integer, argv() As String) As Long

Dim Dict_decoded_temp As String
Dim Dict_encoded_temp As String
Dim site As Variant
Dim SL_BitWidth As New SiteLong

Dim Dict_encoded_temp_in_2S_binary As New DSPWave
Dim Dict_encoded_temp_in_Decimal As New DSPWave



Dim Dict_decoded_temp_in_Decimal As New DSPWave


Dict_encoded_temp = argv(0)
Dict_decoded_temp = argv(1)

Dict_encoded_temp_in_2S_binary = GetStoredCaptureData(Dict_encoded_temp)


''  Test Data for y0 25C

'For Each Site In TheExec.sites

  '  Dict_encoded_temp_in_2S_binary(Site).Element(0) = 1
   ' Dict_encoded_temp_in_2S_binary(Site).Element(1) = 0

   ' Dict_encoded_temp_in_2S_binary(Site).Element(2) = 1
   ' Dict_encoded_temp_in_2S_binary(Site).Element(3) = 1
   ' Dict_encoded_temp_in_2S_binary(Site).Element(4) = 0
   ' Dict_encoded_temp_in_2S_binary(Site).Element(5) = 1
    'Dict_encoded_temp_in_2S_binary(Site).Element(6) = 1
    'Dict_encoded_temp_in_2S_binary(Site).Element(7) = 0
    'Dict_encoded_temp_in_2S_binary(Site).Element(8) = 1
    'Dict_encoded_temp_in_2S_binary(Site).Element(9) = 1


'Next Site


''

For Each site In TheExec.sites
            SL_BitWidth(site) = Dict_encoded_temp_in_2S_binary(site).SampleSize
Next site

Dict_encoded_temp_in_Decimal.CreateConstant 0, 1, DspLong


Call rundsp.DSP_2S_Complement_To_SignDec(Dict_encoded_temp_in_2S_binary, SL_BitWidth, Dict_encoded_temp_in_Decimal)

Dict_decoded_temp_in_Decimal.CreateConstant 0, 1, DspDouble







For Each site In TheExec.sites

Dict_decoded_temp_in_Decimal(site).Element(0) = (CDbl(Dict_encoded_temp_in_Decimal(site).Element(0)) / 64) + 25

Next site



Call AddStoredCaptureData(Dict_decoded_temp, Dict_decoded_temp_in_Decimal)


End Function

Public Function Calc_Metrology_adc_tfe_temp_fuses(argc As Integer, argv() As String) As Long

    Dim site As Variant

    Dim Dict_name_tfe_vol_x1 As String
    Dim cal_tfe_vol_y1 As String

    Dim fuse_read_tfe_vol_0 As String
    Dim fuse_read_tfe_vol_1 As String
    Dim fuse_read_tfe_x0 As String
    Dim fuse_read_tfe_y0 As String

    Dim fuse_write_tfe_temp_0 As String
    Dim fuse_write_tfe_temp_1 As String

    Dim tfe_y0_decimal As String
    Dim tfe_y1_decimal As String


    Dim actual_Temp_CP2 As Double
    
    Dim x As Long
    
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    

    Dict_name_tfe_vol_x1 = argv(0)
    cal_tfe_vol_y1 = argv(1)
    fuse_read_tfe_vol_0 = argv(2)
    fuse_read_tfe_vol_1 = argv(3)
    fuse_read_tfe_x0 = argv(4)
    fuse_read_tfe_y0 = argv(5)

    fuse_write_tfe_temp_0 = argv(6)
    fuse_write_tfe_temp_1 = argv(7)




    'Get Cap data for t5p2 at 85C and Fuse Data for offset,gain and x0 at 25C
    Dim DSP_tfe_vol_x1_binary As New DSPWave
    Dim DSP_fuse_read_tfe_vol_0_2S_binary As New DSPWave
    Dim DSP_fuse_read_tfe_vol_1_binary As New DSPWave
    Dim DSP_fuse_read_tfe_x0_binary As New DSPWave





    DSP_tfe_vol_x1_binary = GetStoredCaptureData(Dict_name_tfe_vol_x1)
    DSP_fuse_read_tfe_vol_0_2S_binary = GetStoredCaptureData(fuse_read_tfe_vol_0)
    DSP_fuse_read_tfe_vol_1_binary = GetStoredCaptureData(fuse_read_tfe_vol_1)
    DSP_fuse_read_tfe_x0_binary = GetStoredCaptureData(fuse_read_tfe_x0)



    ' Test Inputs
    
'    For Each Site In TheExec.sites
'
'    'test Run
'
'            '000010000101001000
'            DSP_fuse_read_tfe_x0_binary(Site).Element(0) = 0
'            DSP_fuse_read_tfe_x0_binary(Site).Element(1) = 1
'            DSP_fuse_read_tfe_x0_binary(Site).Element(2) = 1
'            DSP_fuse_read_tfe_x0_binary(Site).Element(3) = 0
'            DSP_fuse_read_tfe_x0_binary(Site).Element(4) = 0
'            DSP_fuse_read_tfe_x0_binary(Site).Element(5) = 0
'            DSP_fuse_read_tfe_x0_binary(Site).Element(6) = 0
'            DSP_fuse_read_tfe_x0_binary(Site).Element(7) = 1
'            DSP_fuse_read_tfe_x0_binary(Site).Element(8) = 0
'            DSP_fuse_read_tfe_x0_binary(Site).Element(9) = 0
'            DSP_fuse_read_tfe_x0_binary(Site).Element(10) = 0
'            DSP_fuse_read_tfe_x0_binary(Site).Element(11) = 0
'            DSP_fuse_read_tfe_x0_binary(Site).Element(12) = 0
'            DSP_fuse_read_tfe_x0_binary(Site).Element(13) = 1
'            DSP_fuse_read_tfe_x0_binary(Site).Element(14) = 0
'            DSP_fuse_read_tfe_x0_binary(Site).Element(15) = 0
'             DSP_fuse_read_tfe_x0_binary(Site).Element(16) = 0
'            DSP_fuse_read_tfe_x0_binary(Site).Element(17) = 0
'
'
'            '000010011100101001
''            DSP_tfe_vol_x1_binary(Site).Element(0) = 1
''            DSP_tfe_vol_x1_binary(Site).Element(1) = 0
''            DSP_tfe_vol_x1_binary(Site).Element(2) = 0
''            DSP_tfe_vol_x1_binary(Site).Element(3) = 1
''            DSP_tfe_vol_x1_binary(Site).Element(4) = 0
''            DSP_tfe_vol_x1_binary(Site).Element(5) = 1
''            DSP_tfe_vol_x1_binary(Site).Element(6) = 0
''            DSP_tfe_vol_x1_binary(Site).Element(7) = 0
''            DSP_tfe_vol_x1_binary(Site).Element(8) = 1
''            DSP_tfe_vol_x1_binary(Site).Element(9) = 1
''            DSP_tfe_vol_x1_binary(Site).Element(10) = 1
''            DSP_tfe_vol_x1_binary(Site).Element(11) = 0
''            DSP_tfe_vol_x1_binary(Site).Element(12) = 0
''            DSP_tfe_vol_x1_binary(Site).Element(13) = 1
''            DSP_tfe_vol_x1_binary(Site).Element(14) = 0
''            DSP_tfe_vol_x1_binary(Site).Element(15) = 0
''             DSP_tfe_vol_x1_binary(Site).Element(16) = 0
''            DSP_tfe_vol_x1_binary(Site).Element(17) = 0
''
'
'            '111111111111111000
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(0) = 0
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(1) = 0
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(2) = 0
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(3) = 1
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(4) = 0
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(5) = 1
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(6) = 1
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(7) = 1
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(8) = 1
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(9) = 1
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(10) = 1
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(11) = 1
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(12) = 1
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(13) = 1
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(14) = 1
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(15) = 1
'             DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(16) = 1
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(17) = 1
'
'
'
'
'            '100000000000011000
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(0) = 1
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(1) = 0
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(2) = 0
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(3) = 1
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(4) = 1
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(5) = 1
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(6) = 0
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(7) = 0
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(8) = 0
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(9) = 0
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(10) = 0
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(11) = 0
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(12) = 0
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(13) = 0
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(14) = 0
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(15) = 0
'             DSP_fuse_read_tfe_vol_1_binary(Site).Element(16) = 0
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(17) = 1
'
'
'
'    Next Site
'
    
    ''


    ' y0 in decimal for 25C
    Dim DSP_fuse_read_tfe_y0_in_double As New DSPWave
    Dim decoded_Dic_tfe_y0_in_double As String
    decoded_Dic_tfe_y0_in_double = "decoded_Dic_tfe_y0_in_double"

    Dim call_decode_argv(2) As String
    call_decode_argv(0) = fuse_read_tfe_y0
    call_decode_argv(1) = decoded_Dic_tfe_y0_in_double
    Dim call_decodeActualTemp As Long
    call_decodeActualTemp = Calc_Metrology_DecodeActualTemp(1, call_decode_argv)

    DSP_fuse_read_tfe_y0_in_double = GetStoredCaptureData(decoded_Dic_tfe_y0_in_double)



    ' y1 in decimal for 85C .. for now..will be changed in future
    If cal_tfe_vol_y1 Like "CP2" Then

        actual_Temp_CP2 = 85

    End If

    Dim DSP_tfe_y1_in_double As New DSPWave

    DSP_tfe_y1_in_double.CreateConstant 0, 1, DspDouble

    For Each site In TheExec.sites

    DSP_tfe_y1_in_double(site).Element(0) = actual_Temp_CP2

    Next site

'    'Check for Encode Logic ..Can comment it
'
'    Dim encoded_tfe_y1_in_2S_binary As String
'    Dim DSP_tfe_y1_in_2S_binary As New DSPWave
'    encoded_tfe_y1_in_2S_binary = "encoded_tfe_y1_in_2S_binary"
'        Dim call_encode_argv(2) As String
'    call_encode_argv(0) = CStr(actual_Temp_CP2)
'    call_encode_argv(1) = encoded_tfe_y1_in_2S_binary
'
'    Dim call_encodeActualTemp As Long
'    call_encodeActualTemp = Calc_Metrology_EncodeActualTemp(1, call_encode_argv)
'
'    DSP_tfe_y1_in_2S_binary = GetStoredCaptureData(encoded_tfe_y1_in_2S_binary)
'
'    'Check End


    'Start the algo


    'Define Constants

    Dim C0 As Double
    Dim c1 As Double
    Dim C2 As Double
    Dim C3 As Double

    'Values for Constants

    C0 = CDbl("-21.5822184999726")
    c1 = CDbl("428.0092266096283") 'truncated one digit
    C2 = CDbl("-133.4543109228228") 'truncated one digit
    C3 = CDbl("19.0485545665615")
    



    'Convert x1 to decimal

    Dim DSP_tfe_vol_x1_in_decimal As New DSPWave

    Call rundsp.BinToDec(DSP_tfe_vol_x1_binary, DSP_tfe_vol_x1_in_decimal)



    'Convert x0 to decimal

    Dim DSP_fuse_read_tfe_x0_in_decimal As New DSPWave

    Call rundsp.BinToDec(DSP_fuse_read_tfe_x0_binary, DSP_fuse_read_tfe_x0_in_decimal)



    'Convert vol_0 2S to decimal
     Dim DSP_fuse_read_tfe_vol_0_in_decimal As New DSPWave
     Dim SL_BitWidth As New SiteLong
     For Each site In TheExec.sites
            SL_BitWidth(site) = 18

    Next site

    DSP_fuse_read_tfe_vol_0_in_decimal.CreateConstant 0, 1, DspLong



    Call rundsp.DSP_2S_Complement_To_SignDec(DSP_fuse_read_tfe_vol_0_2S_binary, SL_BitWidth, DSP_fuse_read_tfe_vol_0_in_decimal)



    'Convert vol_1 to Decimal
    Dim DSP_fuse_read_tfe_vol_1_in_decimal As New DSPWave

    Call rundsp.BinToDec(DSP_fuse_read_tfe_vol_1_binary, DSP_fuse_read_tfe_vol_1_in_decimal)




    Dim x0 As New SiteDouble
    Dim x1 As New SiteDouble

    Dim Y0 As New SiteDouble
    Dim Y1 As New SiteDouble

    Dim c1_cal As New SiteDouble
    Dim C0_CAL As New SiteDouble

    Dim tfe_temp0_double As New SiteDouble
    Dim tfe_temp1_double As New SiteDouble

    Dim tfe_temp0_long As New SiteLong
    Dim tfe_temp1_long As New SiteLong

    Dim Dsp_tfe_temp0_in_decimal As New DSPWave
    Dim Dsp_tfe_temp1_in_decimal As New DSPWave
    
    Dsp_tfe_temp0_in_decimal.CreateConstant 0, 1, DspDouble
     Dsp_tfe_temp1_in_decimal.CreateConstant 0, 1, DspDouble
    
'    'Test Data
'    For Each Site In TheExec.sites
'
'    DSP_fuse_read_tfe_x0_in_decimal(Site).Element(0) = 8520
'    DSP_fuse_read_tfe_y0_in_double(Site).Element(0) = 22.7031
'
'    DSP_fuse_read_tfe_vol_0_in_decimal(Site).Element(0) = -8
'    DSP_fuse_read_tfe_vol_1_in_decimal(Site).Element(0) = 131097
'
'
'    DSP_tfe_vol_x1_in_decimal(Site).Element(0) = 10025
'    DSP_tfe_y1_in_double(Site).Element(0) = 86.1
'
'
'    Next Site
'
'    ''Test data end

    For Each site In TheExec.sites
                x = 0
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "tfex0", CInt(x))
                 
                TheExec.Flow.TestLimit resultVal:=DSP_fuse_read_tfe_x0_in_decimal(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "tfey0", CInt(x))
                  
                TheExec.Flow.TestLimit resultVal:=DSP_fuse_read_tfe_y0_in_double(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "tfex1", CInt(x))
                    
                TheExec.Flow.TestLimit resultVal:=DSP_tfe_vol_x1_in_decimal(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "tfey1", CInt(x))
                    
                TheExec.Flow.TestLimit resultVal:=DSP_tfe_y1_in_double(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
                
                TestNameInput = Report_ALG_TName_From_Instance(OutputTname_format, "C", "X", "tfevol0", CInt(x))
                
                TheExec.Flow.TestLimit resultVal:=DSP_fuse_read_tfe_vol_0_in_decimal(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "tfevol1", CInt(x))
                
                TheExec.Flow.TestLimit resultVal:=DSP_fuse_read_tfe_vol_1_in_decimal(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
                
               
                
                If (DSP_fuse_read_tfe_x0_in_decimal(site).Element(0) = DSP_tfe_vol_x1_in_decimal(site).Element(0)) Or (DSP_fuse_read_tfe_vol_1_in_decimal(site).Element(0) = 0) Then
                         
                         tfe_temp0_double(site) = 178956970

                        tfe_temp1_double(site) = 178956970
                        
                            Dsp_tfe_temp0_in_decimal(site).Element(0) = FormatNumber(tfe_temp0_double(site))

                        Dsp_tfe_temp1_in_decimal(site).Element(0) = FormatNumber(tfe_temp1_double(site))
                    TestNameInput = Report_TName_From_Instance(CalcC, "X", "Error_code_temp_0", CInt(x))
                    
                    TheExec.Flow.TestLimit resultVal:=Dsp_tfe_temp0_in_decimal(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
                    TestNameInput = Report_TName_From_Instance(CalcC, "X", "Error_code_temp_1", CInt(x))
                        
                    TheExec.Flow.TestLimit resultVal:=Dsp_tfe_temp1_in_decimal(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
                Else
                

                    x0(site) = ((DSP_fuse_read_tfe_x0_in_decimal(site).Element(0) - DSP_fuse_read_tfe_vol_0_in_decimal(site).Element(0)) / CDbl(DSP_fuse_read_tfe_vol_1_in_decimal(site).Element(0))) * 16
            
        x1(site) = ((DSP_tfe_vol_x1_in_decimal(site).Element(0) - DSP_fuse_read_tfe_vol_0_in_decimal(site).Element(0)) / CDbl(DSP_fuse_read_tfe_vol_1_in_decimal(site).Element(0))) * 16


        Y0(site) = 273.15 + DSP_fuse_read_tfe_y0_in_double(site).Element(0) - C2 * x0(site) * x0(site) - C3 * x0(site) * x0(site) * x0(site)


        Y1(site) = 273.15 + DSP_tfe_y1_in_double(site).Element(0) - C2 * x1(site) * x1(site) - C3 * x1(site) * x1(site) * x1(site)


        c1_cal(site) = (Y1(site) - Y0(site)) / (x1(site) - x0(site))

        C0_CAL(site) = (x1(site) * Y0(site) - x0(site) * Y1(site)) / (x1(site) - x0(site))

        tfe_temp0_double(site) = (C0_CAL(site) - C0) * (2 ^ 13)

        tfe_temp1_double(site) = (c1_cal(site) - c1) * (2 ^ 13)

    'tfe_temp0_long(Site) = FormatNumber(tfe_temp0_double(Site))

    'tfe_temp1_long(Site) = FormatNumber(tfe_temp1_double(Site))
                
                
                    If (tfe_temp0_double(site) > 134217727) Or (tfe_temp0_double(site) < -134217728) Then
                    

                        TestNameInput = Report_TName_From_Instance(CalcC, "X", "UpperLimit_Reached_temp_0", CInt(x))
                            
                        TheExec.Flow.TestLimit resultVal:=tfe_temp0_double(site), Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
 
                                               
                        tfe_temp0_double(site) = 178956970
                                           
                        
                    
                    End If
                    If (tfe_temp1_double(site) > 134217727) Or (tfe_temp1_double(site) < -134217728) Then
                        TestNameInput = Report_TName_From_Instance(CalcC, "X", "UpperLimit_Reached_temp_1", CInt(x))
                            
                        TheExec.Flow.TestLimit resultVal:=tfe_temp1_double(site), Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
                    
                      tfe_temp1_double(site) = 178956970
                    End If
                
                
                    Dsp_tfe_temp0_in_decimal(site).Element(0) = FormatNumber(tfe_temp0_double(site))

                    Dsp_tfe_temp1_in_decimal(site).Element(0) = FormatNumber(tfe_temp1_double(site))
        
        
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "X0", CInt(x))
                    
                TheExec.Flow.TestLimit resultVal:=x0(site), Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "X1", CInt(x))
                   
                TheExec.Flow.TestLimit resultVal:=x1(site), Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer

                TestNameInput = Report_TName_From_Instance(CalcC, "X", "Y0", CInt(x))
                    
                TheExec.Flow.TestLimit resultVal:=Y0(site), Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "Y1", CInt(x))
                
                TheExec.Flow.TestLimit resultVal:=Y1(site), Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "c1_calc", CInt(x))
                    
                TheExec.Flow.TestLimit resultVal:=c1_cal(site), Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "c0_calc", CInt(x))
                    
                TheExec.Flow.TestLimit resultVal:=C0_CAL(site), Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "temp_0", CInt(x))
                
                TheExec.Flow.TestLimit resultVal:=Dsp_tfe_temp0_in_decimal(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "temp_1", CInt(x))
                    
                TheExec.Flow.TestLimit resultVal:=Dsp_tfe_temp1_in_decimal(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
                
            End If

    Next site



    Dim Dsp_tfe_temp0_in_binary As New DSPWave
    Dim Dsp_tfe_temp1_in_binary As New DSPWave


    Call rundsp.DSPWf_Dec2Binary(Dsp_tfe_temp0_in_decimal, 28, Dsp_tfe_temp0_in_binary)

    Call rundsp.DSPWf_Dec2Binary(Dsp_tfe_temp1_in_decimal, 28, Dsp_tfe_temp1_in_binary)

    'Algo end

    'Store Data
    
    
    ''test dspWave
    
'    Dim test_dspWave As New DSPWave
'    test_dspWave.CreateConstant 0, 5, DspLong
'
'    For Each Site In TheExec.sites
'    test_dspWave(Site).Element(0) = 2
'    test_dspWave(Site).Element(1) = -2
'    test_dspWave(Site).Element(2) = 3
'    test_dspWave(Site).Element(3) = 4
'    test_dspWave(Site).Element(4) = 14
'    Next Site
'
'
'    Dim test_dspWave_inBinary As New DSPWave
  '  Call rundsp.DSPWf_Dec2Binary(test_dspWave, 4, test_dspWave_inBinary)
    
    ''end test

    Call AddStoredCaptureData(fuse_write_tfe_temp_0, Dsp_tfe_temp0_in_binary)
    Call AddStoredCaptureData(fuse_write_tfe_temp_1, Dsp_tfe_temp1_in_binary)



End Function



Public Function Calc_MTR_REL_Freq_Diff_Percentage(argc As Integer, argv() As String) As Long

    Dim site As Variant
    Dim freq_Dut As String
    Dim freq_ref As String

    Dim fdiff_percent As String


    Dim DSP_fdiff_percent As New DSPWave


    DSP_fdiff_percent.CreateConstant 0, 1, DspDouble


    Dim DSP_freq_Dut As New DSPWave
    Dim DSP_freq_ref As New DSPWave



    freq_Dut = argv(0)
    freq_ref = argv(1)
    fdiff_percent = argv(2)

    Dim TestName As String
    TestName = "f_diff_" + freq_Dut
    DSP_freq_Dut = GetStoredCaptureData(freq_Dut)
    DSP_freq_ref = GetStoredCaptureData(freq_ref)

   
    For Each site In TheExec.sites

    If DSP_freq_Dut(site).Element(0) <> 0 Then
        DSP_fdiff_percent(site).Element(0) = ((DSP_freq_Dut(site).Element(0) - DSP_freq_ref(site).Element(0)) / DSP_freq_Dut(site).Element(0)) * 100
     
    

    Else
        DSP_fdiff_percent(site).Element(0) = 99999
         If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site:" + CStr(site) + "  freq_of_Dut " + freq_Dut + " is 0")

    End If

                TheExec.Flow.TestLimit resultVal:=DSP_fdiff_percent(site).Element(0), Tname:=TestName, ForceResults:=tlForceNone 'eng_forceflow_transfer
    Next site


    Call AddStoredCaptureData(fdiff_percent, DSP_fdiff_percent)



End Function

Public Function Calc_MIPID_VCMTX(argc As Integer, argv() As String) As Long

    Dim i As Long, j As Long
    Dim TestNameInput As String
    Dim OutputTname_format() As String

    Dim DSPWave_Binary() As New DSPWave
    ReDim DSPWave_Binary(argc - 1) As New DSPWave
    
    Dim DSPWave_Combine As New DSPWave
    DSPWave_Combine.CreateConstant 0, 10, DspLong
    
'    Dim DSPWave_Combine_verify As New DSPWave
'    DSPWave_Combine_verify.CreateConstant 0, 10, DspLong

    
    Dim DSPWave_Combine_Dec As New DSPWave
    DSPWave_Combine_Dec.CreateConstant 0, 1, DspLong
    
    Dim TestName As String
    Dim site As Variant
    
    For i = 0 To argc - 1 '20190523 CWCIOU
        DSPWave_Binary(i) = GetStoredCaptureData(argv(i))
    Next i
    
    TestName = argv(argc - 1)
    
    For j = 0 To DSPWave_Combine.SampleSize - 1
        For Each site In TheExec.sites
            If j < 8 Then
                DSPWave_Combine.Element(j) = DSPWave_Binary(0).Element(j)
            Else
                DSPWave_Combine.Element(j) = DSPWave_Binary(1).Element(j - 8)
            End If
        Next site
    Next j

    Call rundsp.ConvertToLongAndSerialToParrel(DSPWave_Combine, 10, DSPWave_Combine_Dec)
    
    Dim VCMTX As New DSPWave
    VCMTX.CreateConstant 0, 1, DspDouble
    Dim VDD18_MIPID_value As Double
    VDD18_MIPID_value = TheHdw.DCVS.Pins("VDD18_MIPID").Voltage.Main.value '20190523 CWCIOU
    If VDD18_MIPID_value = 0 Then
            VDD18_MIPID_value = 999
            TheExec.Datalog.WriteComment ("Error! Apply VDD18_MIPID=0 V  ")
    End If
    
    For Each site In TheExec.sites
        VCMTX(site).Element(0) = DSPWave_Combine_Dec(site).Element(0) / 1024 * VDD18_MIPID_value
    Next site
'    Call rundsp.DSPWaveDecToBinary(DSPWave_Combine_Dec, 10, DSPWave_Combine_verify)
    
    TestNameInput = Report_TName_From_Instance(CalcC, "X", , 0)
    'TestNameInput = Report_TName_From_Instance("V", "X", , 0)

    TheExec.Flow.TestLimit resultVal:=VCMTX.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
End Function

Public Function Calc_DigCapCombine(argc As Integer, argv() As String) As Long

    Dim i As Long, j As Long

    Dim DSPWave_Binary() As New DSPWave
    ReDim DSPWave_Binary(argc - 1) As New DSPWave
    
    Dim DSPWave_Combine As New DSPWave
    DSPWave_Combine.CreateConstant 0, 10, DspLong
    
'    Dim DSPWave_Combine_verify As New DSPWave
'    DSPWave_Combine_verify.CreateConstant 0, 10, DspLong

    
    Dim DSPWave_Combine_Dec As New DSPWave
    DSPWave_Combine_Dec.CreateConstant 0, 1, DspLong
    
    Dim TestNameInput As String
    Dim site As Variant
    
    For i = 0 To argc - 1
        DSPWave_Binary(i) = GetStoredCaptureData(argv(i))
    Next i
    

    For j = 0 To DSPWave_Combine.SampleSize - 1
        For Each site In TheExec.sites
            If j < 8 Then
                DSPWave_Combine.Element(j) = DSPWave_Binary(0).Element(j)
            Else
                DSPWave_Combine.Element(j) = DSPWave_Binary(1).Element(j - 8)
            End If
        Next site
    Next j

    Call rundsp.ConvertToLongAndSerialToParrel(DSPWave_Combine, 10, DSPWave_Combine_Dec)
'    Call rundsp.DSPWaveDecToBinary(DSPWave_Combine_Dec, 10, DSPWave_Combine_verify)
    
    
    Dim OutputTname_format() As String

    TestNameInput = Report_TName_From_Instance(CalcC, "X", "DEC" & i, CInt(i))
    
    TheExec.Flow.TestLimit resultVal:=DSPWave_Combine_Dec.Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
End Function
Public Function Calc_VDiff_t6p1_metrologyGR(argc As Integer, argv() As String) As Long
    Dim Dict_V2 As String
    Dim Dict_V1 As String
    Dim TestName As String
    Dim Input_V1 As New PinListData
    Dim Input_V2 As New PinListData
    Dim result As New DSPWave
    Dim CalcVal As New PinListData
    Dim DummyPinListData As New PinListData
    Dim site As Variant
    Dim x As Integer
    Dim OutputTname_format() As String
    Dim TestNameInput As String
    
    
    x = 0
   
    
    result.CreateConstant 0, 1, DspDouble

 
    
    Dict_V1 = argv(0)
    Dict_V2 = argv(1)
    TestName = argv(2)
    Input_V1 = GetStoredMeasurement(Dict_V1)
      Input_V2 = GetStoredMeasurement(Dict_V2)
      
      
    DummyPinListData.AddPin (Input_V1.Pins(0))
      DummyPinListData = Input_V1.Pins(0).Subtract(Input_V2.Pins(0)).Abs
      
      
      
      
      For Each site In TheExec.sites
        result(site).Element(0) = DummyPinListData.Pins(0).value
      Next site
      


'    CalcVal.AddPin (InputVal.Pins(0))
'    CalcVal = InputVal.Pins(0).Subtract(0.4).Divide(0.7975).Subtract(1)
    
         If Not ByPassTestLimit Then
            TestNameInput = Report_TName_From_Instance(CalcV, "X", , CInt(x))
            TheExec.Flow.TestLimit resultVal:=result.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
        End If
'    Call AddStoredMeasurement(Dict_ReturnKey, CalcVal)
End Function
Public Function Calc_DigCapAvg(argc As Integer, argv() As String) As Long

    Dim i As Long, j As Long

    Dim DSPWave_Binary() As New DSPWave
    ReDim DSPWave_Binary(argc - 4) As New DSPWave
    
    Dim DSPWave_Dec() As New DSPWave
    ReDim DSPWave_Dec(argc - 4) As New DSPWave
    
    Dim DSPWave_Avg_Dec As New DSPWave
    DSPWave_Avg_Dec.CreateConstant 0, 1, DspLong
    'ReDim DSPWave_Avg(argc - 1) As New DSPWave
    
    Dim DSPWave_Avg_Bin As New DSPWave
    'ReDim DSPWave_Avg_Bin(argc - 3) As New DSPWave
    
    Dim TestName As String
    Dim site As Variant
    Dim Dict As String
    Dim bitwidth As Long
    
    For i = 0 To 1
        DSPWave_Binary(i) = GetStoredCaptureData(argv(i))
        Call rundsp.BinToDec(DSPWave_Binary(i), DSPWave_Dec(i))
    Next i
    
    TestName = argv(argc - 1)
    bitwidth = argv(argc - 2)
    Dict = argv(argc - 3)
    
    For Each site In TheExec.sites
            DSPWave_Avg_Dec.Element(0) = Int(((DSPWave_Dec(0).Element(0) + DSPWave_Dec(1).Element(0)) / 2) + 0.5) ''Example 1). 78.4=>78  2). 78.5=79
    Next site
    Call rundsp.DSPWaveDecToBinary(DSPWave_Avg_Dec, bitwidth, DSPWave_Avg_Bin)
    Call AddStoredCaptureData(Dict, DSPWave_Avg_Bin)
    Dim TestNameInput As String
    Dim OutputTname_format() As String

    TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(i))
    
    TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
    TheExec.Flow.TestLimit resultVal:=DSPWave_Avg_Dec.Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
    
End Function
Public Function Calc_CalR_FVMI_IO(argc As Integer, argv() As String) As Long

    Dim StoredCurrent As New PinListData
    Dim CalR As New PinListData
    Dim ForceVoltVal As Double
    Dim PowerPinName As String
    
    Dim i, p As Long
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    Dim site As Variant
    Dim pin  As Variant
    Dim Lowlimitval_temp As Double
    Dim Hilimitval_temp As Double
        
    PowerPinName = argv(1)
    StoredCurrent = GetStoredMeasurement(argv(0))
    ForceVoltVal = argv(2)
    
    For Each pin In StoredCurrent.Pins
        For Each site In TheExec.sites
            If StoredCurrent.Pins(pin).value(site) = 0 Then
                StoredCurrent.Pins(pin).value(site) = 0.000000000001
            End If
        Next site
    Next pin

    
    CalR = StoredCurrent.Math.Invert.Multiply(ForceVoltVal).Abs
          
    '===============RAK read
    Dim GetRakVal As New PinListData
    GetRakVal = CurrentJob_Card_RAK
       
            For Each site In TheExec.sites
                GetRakVal = CurrentJob_Card_RAK.Pins(PowerPinName).value(site)
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment PowerPinName & " = " & CalR.Pins.item(PowerPinName).value(site) & ", RAK val = " & GetRakVal.Pins(PowerPinName).value
                CalR.Pins.item(PowerPinName).value(site) = CalR.Pins.item(PowerPinName).value(site) - GetRakVal.Pins(PowerPinName).value
            Next site
    
        For p = 0 To CalR.Pins.Count - 1
            If LCase(CalR.Pins.item(p).name) Like LCase((PowerPinName)) Then
                    TestNameInput = Report_TName_From_Instance("R", CalR.Pins(p), , CInt(p))
                    Hilimitval_temp = 96
                    Lowlimitval_temp = 64
                    TheExec.Flow.TestLimit CalR.Pins(p), Lowlimitval_temp, Hilimitval_temp, , , , unitCustom, , TestNameInput, , , , , " ohm", , ForceResults:=tlForceFlow
            End If
        Next p
    
'
End Function
Public Function Calc_CalZ_FVMI(argc As Integer, argv() As String) As Long

    Dim StoredCurrent As New PinListData
    Dim StoredCurrent_I2 As New PinListData
    Dim StoredCurrent_I1 As New PinListData
    Dim CalR As New PinListData
    Dim ForceVoltVal As Double
    Dim PowerPinName As String
    
    Dim i, p As Long
    Dim TestNameInput As String
    Dim site As Variant
    Dim pin  As Variant

        
        'argv() :V2,V1,I2,I1
        

    StoredCurrent_I1 = GetStoredMeasurement(argv(3))
    StoredCurrent_I2 = GetStoredMeasurement(argv(2))
  
    ForceVoltVal = argv(0) - argv(1)
    

    StoredCurrent = StoredCurrent_I2.Math.Subtract(StoredCurrent_I1)
    
        For Each pin In StoredCurrent.Pins ' To prevent i=0
            For Each site In TheExec.sites
                If StoredCurrent.Pins(pin).value(site) = 0 Then
                    StoredCurrent.Pins(pin).value(site) = 0.000000000001
                End If
            Next site
        Next pin
    
        CalR = StoredCurrent.Math.Invert.Multiply(ForceVoltVal).Abs

        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment CalR.Pins(p) & " = " & CalR.Pins(p).value(site)
        TestNameInput = Report_TName_From_Instance(CalcR, "X", vbNullString, 0)
        TheExec.Flow.TestLimit CalR, , , , , , unitCustom, , TestNameInput, , , , , " ohm", , ForceResults:=tlForceFlow


End Function


Public Function Sub_MDLL(DSP_Input_UpperBIN_1 As DSPWave, DSP_Input_UpperBIN_2 As DSPWave, DSP_Input_BelowBIN_1 As DSPWave, DSP_Input_BelowBIN_2 As DSPWave, DSP_Input_UpperDEC_1 As DSPWave, DSP_Input_UpperDEC_2 As DSPWave, DSP_Input_BelowDEC_1 As DSPWave, DSP_Input_BelowDEC_2 As DSPWave, _
ByRef Temp_DSP_Input_UpperBIN_1 As DSPWave, ByRef Temp_DSP_Input_UpperBIN_2 As DSPWave, ByRef Temp_DSP_Input_BelowBIN_1 As DSPWave, ByRef Temp_DSP_Input_BelowBIN_2 As DSPWave, ByRef Temp_DSP_Input_UpperDEC_1 As DSPWave, ByRef Temp_DSP_Input_UpperDEC_2 As DSPWave, ByRef Temp_DSP_Input_BelowDEC_1 As DSPWave, ByRef Temp_DSP_Input_BelowDEC_2 As DSPWave, _
Binary_Start As Long, Binary_End As Long, Dec_data As Long) As Long

    Dim i As Long: i = 0
    Dim j  As Long: j = 0
    Dim site As Variant
    For Each site In TheExec.sites
        Temp_DSP_Input_UpperDEC_1(site).Element(0) = DSP_Input_UpperDEC_1(site).Element(Dec_data)
        Temp_DSP_Input_UpperDEC_2(site).Element(0) = DSP_Input_UpperDEC_2(site).Element(Dec_data)
        Temp_DSP_Input_BelowDEC_1(site).Element(0) = DSP_Input_BelowDEC_1(site).Element(Dec_data)
        Temp_DSP_Input_BelowDEC_2(site).Element(0) = DSP_Input_BelowDEC_2(site).Element(Dec_data)
        
        For i = Binary_Start To Binary_End
            Temp_DSP_Input_UpperBIN_1(site).Element(j) = DSP_Input_UpperBIN_1(site).Element(i)
            Temp_DSP_Input_UpperBIN_2(site).Element(j) = DSP_Input_UpperBIN_2(site).Element(i)
            Temp_DSP_Input_BelowBIN_1(site).Element(j) = DSP_Input_BelowBIN_1(site).Element(i)
            Temp_DSP_Input_BelowBIN_2(site).Element(j) = DSP_Input_BelowBIN_2(site).Element(i)
            j = j + 1
        Next i
        j = 0
    Next site
    
End Function

Public Function Calc_FromLoad_MTR_SE_CAL_Coeff(SensorTempName_rot As String, SensorTempName_rov As String, ByVal Temperature As Long, ByVal FuseSize_1 As Long, ByVal FuseSize_2 As Long, ByRef DSPWave_Coeff_1 As DSPWave, ByRef DSPWave_Coeff_2 As DSPWave, ByRef OutDspWaveToFuse_1 As DSPWave, ByRef OutDspWaveToFuse_2 As DSPWave, MTR_CAL_Sheet As Long) As Long

    Dim site As Variant
    Dim piU1(3, 7) As Double
    Dim piU2(2, 7) As Double
    Dim piU3(3, 7) As Double
    Dim piU4(2, 7) As Double
    Dim a1 As New DSPWave
    Dim a2 As New DSPWave
    Dim a3 As New DSPWave
    Dim a4 As New DSPWave
    a1.CreateConstant 0, 4, DspDouble
    a2.CreateConstant 0, 3, DspDouble
    a3.CreateConstant 0, 4, DspDouble
    a4.CreateConstant 0, 3, DspDouble
    Dim a1_max(3) As Double
    Dim a2_max(2) As Double
    Dim a3_max(3) As Double
    Dim a4_max(2) As Double
    Dim a1_min(3) As Double
    Dim a2_min(2) As Double
    Dim a3_min(3) As Double
    Dim a4_min(2) As Double
    
    Dim Row As Long
    Dim Col As Long


    
    Dim MTRMatricesSheet As Worksheet
    If MTR_CAL_Sheet = 0 Then
        Set MTRMatricesSheet = Sheets("MTR_CAL_matrices_Group1")
    For Row = 2 To 5
            For Col = 1 To 7
            piU1(Row - 2, Col - 1) = MTRMatricesSheet.Cells(Row, Col)
        Next Col
    Next Row
    
    For Row = 7 To 9
            For Col = 1 To 7
            piU2(Row - 7, Col - 1) = MTRMatricesSheet.Cells(Row, Col)
        Next Col
    Next Row
    
    For Row = 11 To 14
            For Col = 1 To 7
            piU3(Row - 11, Col - 1) = MTRMatricesSheet.Cells(Row, Col)
        Next Col
    Next Row
    
    For Row = 16 To 18
            For Col = 1 To 7
            piU4(Row - 16, Col - 1) = MTRMatricesSheet.Cells(Row, Col)
        Next Col
    Next Row
    Else
        Set MTRMatricesSheet = Sheets("MTR_CAL_matrices_Group2")
        For Row = 2 To 5
            For Col = 1 To 5
                piU1(Row - 2, Col - 1) = MTRMatricesSheet.Cells(Row, Col)
            Next Col
        Next Row
    
        For Row = 7 To 9
            For Col = 1 To 5
                piU2(Row - 7, Col - 1) = MTRMatricesSheet.Cells(Row, Col)
            Next Col
        Next Row
    
        For Row = 11 To 14
            For Col = 1 To 5
                piU3(Row - 11, Col - 1) = MTRMatricesSheet.Cells(Row, Col)
            Next Col
        Next Row
    
        For Row = 16 To 18
            For Col = 1 To 5
                piU4(Row - 16, Col - 1) = MTRMatricesSheet.Cells(Row, Col)
            Next Col
        Next Row
    End If
    
    For Col = 2 To 5
        a1_max(Col - 2) = MTRMatricesSheet.Cells(20, Col)
    Next Col
        For Col = 2 To 5
        a1_min(Col - 2) = MTRMatricesSheet.Cells(21, Col)
    Next Col
        For Col = 2 To 4
        a2_max(Col - 2) = MTRMatricesSheet.Cells(22, Col)
    Next Col
        For Col = 2 To 4
        a2_min(Col - 2) = MTRMatricesSheet.Cells(23, Col)
    Next Col
    For Col = 2 To 5
        a3_max(Col - 2) = MTRMatricesSheet.Cells(24, Col)
    Next Col
    For Col = 2 To 5
        a3_min(Col - 2) = MTRMatricesSheet.Cells(25, Col)
    Next Col
    For Col = 2 To 4
        a4_max(Col - 2) = MTRMatricesSheet.Cells(26, Col)
    Next Col
    For Col = 2 To 4
        a4_min(Col - 2) = MTRMatricesSheet.Cells(27, Col)
    Next Col
    
    
    
    
    
    
    
    Dim temp_rowVal_a1 As New SiteDouble
    Dim temp_rowVal_a2 As New SiteDouble
    Dim temp_rowVal_a3 As New SiteDouble
    Dim temp_rowVal_a4 As New SiteDouble
    Dim TestName As String
    Dim currBinaryStr As String
    Dim totalBinaryStr As String
    Dim currElementDspWave As Long
'    Dim OutDspWaveToFuse As New DSPWave
'    OutDspWaveToFuse.CreateConstant 0, FuseSize, DspLong


    Dim decimalPlaces As Long
    decimalPlaces = 8
    
    Dim DSPWave_Matrix_rot As New DSPWave
    Dim DSPWave_Matrix_rov As New DSPWave
    DSPWave_Matrix_rot = GetStoredCaptureData(SensorTempName_rot)
    DSPWave_Matrix_rov = GetStoredCaptureData(SensorTempName_rov)
    
    If Temperature = 25 Then

        For Each site In TheExec.sites
            totalBinaryStr = vbNullString
            For Row = 0 To 3
                currBinaryStr = vbNullString
                temp_rowVal_a1(site) = 0

                If MTR_CAL_Sheet = 0 Then
                    For Col = 0 To 6
                    temp_rowVal_a1(site) = temp_rowVal_a1(site) + piU1(Row, Col) * DSPWave_Matrix_rot(site).Element(Col)
                Next Col
                Else
                    For Col = 0 To 4
                        temp_rowVal_a1(site) = temp_rowVal_a1(site) + piU1(Row, Col) * DSPWave_Matrix_rot(site).Element(Col)
                    Next Col
                End If

                a1(site).Element(Row) = (temp_rowVal_a1(site) - a1_min(Row)) / (a1_max(Row) - a1_min(Row))
                temp_rowVal_a1 = 0
                If (Row = 0) Then
                    Call MTR_Cal_DecimalToBinary(a1(site).Element(Row), 15, decimalPlaces, currBinaryStr)
                Else
                    Call MTR_Cal_DecimalToBinary(a1(site).Element(Row), 14, decimalPlaces, currBinaryStr)
                End If
                totalBinaryStr = totalBinaryStr + currBinaryStr

                'Added on 20180131 To Force Error
                If (a1(site).Element(Row) = 0) Then
                    a1(site).Element(Row) = -0.000001
                ElseIf (a1(site).Element(Row) = 1) Then
                    a1(site).Element(Row) = 1.000001
                End If
            Next Row
            currElementDspWave = 0
            TheExec.Datalog.WriteComment ("Fuse Binary Str  a1 for Site:" + CStr(site) + " is " + totalBinaryStr)
            totalBinaryStr = StrReverse(totalBinaryStr)
            If Len(totalBinaryStr) = OutDspWaveToFuse_1.SampleSize Then
                Do While currElementDspWave < FuseSize_1
                    OutDspWaveToFuse_1(site).Element(currElementDspWave) = CInt(mid(totalBinaryStr, currElementDspWave + 1, 1))
                    currElementDspWave = currElementDspWave + 1
                Loop
            End If
            
            
            totalBinaryStr = vbNullString
            For Row = 0 To 2
                currBinaryStr = vbNullString
                temp_rowVal_a2(site) = 0
                    
                If MTR_CAL_Sheet = 0 Then
                    For Col = 0 To 6
                    temp_rowVal_a2(site) = temp_rowVal_a2(site) + piU2(Row, Col) * DSPWave_Matrix_rov(site).Element(Col)
                Next Col
                Else
                    For Col = 0 To 4
                        temp_rowVal_a2(site) = temp_rowVal_a2(site) + piU2(Row, Col) * DSPWave_Matrix_rov(site).Element(Col)
                    Next Col
                End If

                a2(site).Element(Row) = (temp_rowVal_a2(site) - a2_min(Row)) / (a2_max(Row) - a2_min(Row))
                temp_rowVal_a2 = 0
                If (Row = 0) Then
                    Call MTR_Cal_DecimalToBinary(a2(site).Element(Row), 15, decimalPlaces, currBinaryStr)
                Else
                    Call MTR_Cal_DecimalToBinary(a2(site).Element(Row), 14, decimalPlaces, currBinaryStr)
                End If
                totalBinaryStr = totalBinaryStr + currBinaryStr

                'Added on 20180131 To Force Error
                If (a2(site).Element(Row) = 0) Then
                    a2(site).Element(Row) = -0.000001
                ElseIf (a2(site).Element(Row) = 1) Then
                    a2(site).Element(Row) = 1.000001
                End If
            Next Row
            currElementDspWave = 0
            TheExec.Datalog.WriteComment ("Fuse Binary Str  a2 for Site:" + CStr(site) + " is " + totalBinaryStr)
            totalBinaryStr = StrReverse(totalBinaryStr)
            If Len(totalBinaryStr) = OutDspWaveToFuse_2.SampleSize Then

                Do While currElementDspWave < FuseSize_2
                    OutDspWaveToFuse_2(site).Element(currElementDspWave) = CInt(mid(totalBinaryStr, currElementDspWave + 1, 1))
                    currElementDspWave = currElementDspWave + 1
                Loop
            End If
        Next site
            For Row = 0 To 3
                TestName = "a1_row_" + SensorTempName_rot + "_" + CStr(Row + 1) + ":"
                TheExec.Flow.TestLimit resultVal:=a1.Element(Row), Tname:=TestName, ForceResults:=tlForceFlow
            Next Row
            Set DSPWave_Coeff_1 = a1
            For Row = 0 To 2
                TestName = "a2_row_" + SensorTempName_rov + "_" + CStr(Row + 1) + ":"
                TheExec.Flow.TestLimit resultVal:=a2.Element(Row), Tname:=TestName, ForceResults:=tlForceFlow
            Next Row
            Set DSPWave_Coeff_2 = a2
            
    ElseIf Temperature = 85 Then

        For Each site In TheExec.sites
            totalBinaryStr = vbNullString
            For Row = 0 To 3
                temp_rowVal_a3(site) = 0
                currBinaryStr = vbNullString


                If MTR_CAL_Sheet = 0 Then
                    For Col = 0 To 6
                    temp_rowVal_a3(site) = temp_rowVal_a3(site) + piU3(Row, Col) * DSPWave_Matrix_rot(site).Element(Col)
                Next Col
                Else
                    For Col = 0 To 4
                        temp_rowVal_a3(site) = temp_rowVal_a3(site) + piU3(Row, Col) * DSPWave_Matrix_rot(site).Element(Col)
                    Next Col
                End If

                a3(site).Element(Row) = (temp_rowVal_a3(site) - a3_min(Row)) / (a3_max(Row) - a3_min(Row))
                temp_rowVal_a3 = 0
                If (Row = 0) Then
                    Call MTR_Cal_DecimalToBinary(a3(site).Element(Row), 15, decimalPlaces, currBinaryStr)
                Else
                    Call MTR_Cal_DecimalToBinary(a3(site).Element(Row), 14, decimalPlaces, currBinaryStr)
                End If
                totalBinaryStr = totalBinaryStr + currBinaryStr
                'Added on 20180131 To Force Error
                If (a3(site).Element(Row) = 0) Then
                    a3(site).Element(Row) = -0.000001
                ElseIf (a3(site).Element(Row) = 1) Then
                    a3(site).Element(Row) = 1.000001
                End If
            Next Row
            currElementDspWave = 0
            TheExec.Datalog.WriteComment ("Fuse Binary Str  a3 for Site:" + CStr(site) + " is " + totalBinaryStr)
            totalBinaryStr = StrReverse(totalBinaryStr)
            If Len(totalBinaryStr) = OutDspWaveToFuse_1.SampleSize Then
                
                Do While currElementDspWave < FuseSize_1
                    OutDspWaveToFuse_1(site).Element(currElementDspWave) = CInt(mid(totalBinaryStr, currElementDspWave + 1, 1))
                    currElementDspWave = currElementDspWave + 1
                Loop
            End If
            totalBinaryStr = vbNullString
            For Row = 0 To 2
                temp_rowVal_a4(site) = 0
                currBinaryStr = vbNullString

                    
                If MTR_CAL_Sheet = 0 Then
                    For Col = 0 To 6
                    temp_rowVal_a4(site) = temp_rowVal_a4(site) + piU4(Row, Col) * DSPWave_Matrix_rov(site).Element(Col)
                Next Col
                Else
                    For Col = 0 To 4
                        temp_rowVal_a4(site) = temp_rowVal_a4(site) + piU4(Row, Col) * DSPWave_Matrix_rov(site).Element(Col)
                    Next Col
                End If
                    
                a4(site).Element(Row) = (temp_rowVal_a4(site) - a4_min(Row)) / (a4_max(Row) - a4_min(Row))
                temp_rowVal_a4 = 0
                If (Row = 0) Then
                    Call MTR_Cal_DecimalToBinary(a4(site).Element(Row), 15, decimalPlaces, currBinaryStr)
                Else
                    Call MTR_Cal_DecimalToBinary(a4(site).Element(Row), 14, decimalPlaces, currBinaryStr)
                End If
                totalBinaryStr = totalBinaryStr + currBinaryStr

                'Added on 20180131 To Force Error
                If (a4(site).Element(Row) = 0) Then
                    a4(site).Element(Row) = -0.000001
                ElseIf (a4(site).Element(Row) = 1) Then
                    a4(site).Element(Row) = 1.000001
                End If
            Next Row
            currElementDspWave = 0
            TheExec.Datalog.WriteComment ("Fuse Binary Str  a4 for Site:" + CStr(site) + " is " + totalBinaryStr)
            totalBinaryStr = StrReverse(totalBinaryStr)
            If Len(totalBinaryStr) = OutDspWaveToFuse_2.SampleSize Then

                Do While currElementDspWave < FuseSize_2
                    OutDspWaveToFuse_2(site).Element(currElementDspWave) = CInt(mid(totalBinaryStr, currElementDspWave + 1, 1))
                    currElementDspWave = currElementDspWave + 1
                Loop
            End If
        Next site
        For Row = 0 To 3
            TestName = "a3_row_" + SensorTempName_rot + "_" + CStr(Row + 1) + ":"
            TheExec.Flow.TestLimit resultVal:=a3.Element(Row), Tname:=TestName, ForceResults:=tlForceFlow
        Next Row
        Set DSPWave_Coeff_1 = a3
        For Row = 0 To 2
            TestName = "a4_row_" + SensorTempName_rov + "_" + CStr(Row + 1) + ":"
            TheExec.Flow.TestLimit resultVal:=a4.Element(Row), Tname:=TestName, ForceResults:=tlForceFlow
        Next Row
        Set DSPWave_Coeff_2 = a4
    End If
End Function

Public Function MTR_Cal_DecimalToBinary(ByVal inputDecimal As Double, ByVal bitSize As Long, ByVal placesAfterDecimal As Long, ByRef outBinaryStr As String) As Long
    
    Dim i As Long
    Dim fractional As Double
    Dim integral  As Long
    Dim currIntegral As Long
    Dim decimalFract As Double
    Dim binaryStr As String
    Dim currDecimal As Double
    Dim theDecimal As Double
    Dim currCount As Long
 
    theDecimal = FormatNumber(inputDecimal, placesAfterDecimal)
       
     
    integral = Int(theDecimal)
    
    fractional = theDecimal - integral
    
    If (theDecimal > 0) And (theDecimal < 1) Then
    
        
        
        currCount = 0
        Do While currCount < bitSize
        
            currDecimal = fractional * 2
            currIntegral = Int(currDecimal)
            decimalFract = decimalFract + CStr(currIntegral) * (2 ^ (bitSize - currCount))
            binaryStr = binaryStr + CStr(currIntegral)
            fractional = currDecimal - currIntegral
        
            
            currCount = currCount + 1
            
        Loop
        outBinaryStr = binaryStr
        

    Else
        currCount = 0
        binaryStr = vbNullString
        Do While currCount < bitSize
            binaryStr = binaryStr + "1"
            currCount = currCount + 1
        
        Loop
        outBinaryStr = binaryStr
    End If


End Function


Public Function TX_Low_Level(argc As Integer, argv() As String) As Long

    Dim DictKey_V1 As String, DictKey_V2 As String
    Dim pld_V1 As New PinListData, pld_V2 As New PinListData
    Dim pld_upd_V1 As New PinListData, pld_upd_V2 As New PinListData
    Dim Pin_Name_1 As String, Pin_Name_2 As String
    'Dim Rak_Pin_Name_1() As Double
    'Dim Rak_Pin_Name_2() As Double
    Dim GetRakVal As Double
    Dim OutputTname_format() As String
    Dim TestNameInput As String

    DictKey_V1 = argv(0)
    DictKey_V2 = argv(1)
    Pin_Name_1 = argv(2)
    Pin_Name_2 = argv(3)
    Dim site As Variant
    pld_V1 = GetStoredMeasurement(DictKey_V1)
    pld_V2 = GetStoredMeasurement(DictKey_V2)
    
    pld_upd_V1.AddPin (Pin_Name_1)
    pld_upd_V2.AddPin (Pin_Name_2)
    
    For Each site In TheExec.sites
        'Rak_Pin_Name_1 = TheHdw.PPMU.ReadRakValuesByPinnames(Pin_Name_1, site)
        'Rak_Pin_Name_2 = TheHdw.PPMU.ReadRakValuesByPinnames(Pin_Name_2, site)
        GetRakVal = (CurrentJob_Card_RAK.Pins(Pin_Name_1).value(site) + CurrentJob_Card_RAK.Pins(Pin_Name_2).value(site)) / 2
        pld_upd_V1.Pins(Pin_Name_1).value(site) = pld_V1.Pins(Pin_Name_1).Multiply(45).divide(45 + 45 + GetRakVal).value(site)
        pld_upd_V2.Pins(Pin_Name_2).value(site) = pld_V2.Pins(Pin_Name_2).Multiply(45).divide(45 + 45 + GetRakVal).value(site)
    Next site
    
    TestNameInput = Report_TName_From_Instance(CalcV, "X", vbNullString, 0)
    TheExec.Flow.TestLimit resultVal:=pld_upd_V1, ForceResults:=tlForceFlow, Tname:=TestNameInput
    
    TestNameInput = Report_TName_From_Instance(CalcV, "X", vbNullString, 0)
    TheExec.Flow.TestLimit resultVal:=pld_upd_V2, ForceResults:=tlForceFlow, Tname:=TestNameInput
    
    
End Function

Public Function Calc_MIPI_Tolerance(argc As Integer, argv() As String) As Long



    Dim site As Variant
    Dim i, j As Long
    Dim DSPWave_First As New DSPWave
    Dim DSPWave_Second As New DSPWave
    Dim DSPWave_Combine() As New DSPWave
    Dim TestNameInput As String
    Dim SplitByAdd() As String
    Dim First_StartElement As Long
    Dim First_EndElement As Long
    Dim Second_StartElement As Long
    Dim Second_EndElement As Long
    
    Dim DictKey_DSPWave_Combine As String
    
    Dim DataString_First As String
    Dim DataString_Second As String
    Dim DataString_Combine As String
    
    ReDim DSPWave_Combine(argc - 1) As New DSPWave
    Dim DSPWave_Combine_Dec As New DSPWave
    Dim OutputTname_format() As String
'    Dim TestNameInput As String
    
    For i = 0 To argc - 1
        'TestNameInput = "ConcatenateDSP_"
        SplitByAdd = Split(argv(i), "+")
        DSPWave_First = GetStoredCaptureData(SplitByAdd(0))
        First_StartElement = 0
        First_EndElement = 7
        DSPWave_Second = GetStoredCaptureData(SplitByAdd(1))
        Second_StartElement = 0
        Second_EndElement = 1
        

        Call ConcatenateDSP_TTR(DSPWave_First, First_StartElement, First_EndElement, DSPWave_Second, Second_StartElement, Second_EndElement, DSPWave_Combine(i))
        
        ''20170718 - Store Concatenate DSP to Dict.
'        If UBound(SplitByAt) = 6 Then
'            DictKey_DSPWave_Combine = SplitByAt(6)
'            Call AddStoredCaptureData(DictKey_DSPWave_Combine, DSPWave_Combine(i))
'        End If
        
        For Each site In TheExec.sites
            DataString_First = vbNullString
            DataString_Second = vbNullString
            DataString_Combine = vbNullString
            For j = 0 To DSPWave_First.SampleSize - 1
                DataString_First = DataString_First & DSPWave_First(site).Element(j)
            Next j
            For j = 0 To DSPWave_Second.SampleSize - 1
                DataString_Second = DataString_Second & DSPWave_Second(site).Element(j)
            Next j
            For j = 0 To DSPWave_Combine(i).SampleSize - 1
                DataString_Combine = DataString_Combine & DSPWave_Combine(i)(site).Element(j)
            Next j
            
           If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site " & site & " Dictionary " & SplitByAdd(0) & " Output Bits = " & DataString_First & " Extract Bits [" & First_StartElement & "-" & First_EndElement & "]" & _
                                                           " ,Dictionary " & SplitByAdd(1) & " Output Bits = " & DataString_Second & " Extract Bits [" & Second_StartElement & "-" & Second_EndElement & "]" & _
                                                           " ,Dictionary " & DictKey_DSPWave_Combine & " Output Bits = " & DataString_Combine)
        
        
        
        
        
        
        DSPWave_Combine(i)(site) = DSPWave_Combine(i)(site).ConvertDataTypeTo(DspLong)
        DSPWave_Combine_Dec(site) = DSPWave_Combine(i)(site).ConvertStreamTo(tldspParallel, DSPWave_Combine(i)(site).SampleSize, 0, Bit0IsMsb)
        
        
        Next site
        'Call rundsp.BinToDec(DSPWave_Combine(i), DSPWave_Combine_Dec)
                
        If gl_Disable_HIP_debug_log = False Then

            TestNameInput = Report_TName_From_Instance(CalcC, "X", "ConcatenateDSP", 0)
               
            TheExec.Flow.TestLimit resultVal:=DSPWave_Combine_Dec.Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
      
        End If
        
      
        Dim MIPI_threshold_Code_value_P(7) As New SiteDouble
        Dim MIPI_threshold_Code_value_N(7) As New SiteDouble
        
        For Each site In TheExec.sites
        
        
        'split code p and code n
        If i < 8 Then
            MIPI_threshold_Code_value_P(i)(site) = DSPWave_Combine_Dec(site).Element(0)
        Else
            i = i - 8
            MIPI_threshold_Code_value_N(i)(site) = DSPWave_Combine_Dec(site).Element(0)
            i = i + 8
        End If
        
        Next site

        
    Next i


    Dim MIPI_threshold_lower_p(0) As New SiteVariant
    Dim MIPI_threshold_high_p(0) As New SiteVariant
    'Dim MIPI_threshold_found_p As New SiteBoolean   'Change to SiteLong, due to SiteBoolean True = -1
    Dim MIPI_threshold_lower_n(0) As New SiteVariant
    Dim MIPI_threshold_high_n(0) As New SiteVariant
    'Dim MIPI_threshold_found_n As New SiteBoolean  'Change to SiteLong, due to SiteBoolean True = -1
    Dim MIPI_trans_mapping As Variant
    Dim MIPI_threshold_found_p_value As New SiteLong
    Dim MIPI_threshold_found_n_value As New SiteLong
    
    Dim threshold_temp_p As Integer
    Dim threshold_flag_p As Boolean
    Dim threshold_temp_n As Integer
    Dim threshold_flag_n As Boolean
    Dim p  As Long
    Dim n  As Long
    MIPI_trans_mapping = Array(-0.2, -0.15, -0.1, -0.05, 0.05, 0.1, 0.15, 0.2)

    For Each site In TheExec.sites
    
    'code p process
        threshold_temp_p = 0
        threshold_flag_p = False
        'MIPI_threshold_found_p(Site) = False
        MIPI_threshold_found_p_value(site) = -1  'Clear = -1
        
        
        For p = 0 To 7

            If MIPI_threshold_Code_value_P(p)(site) = 0 Then
                If threshold_flag_p = False Then
                    MIPI_threshold_lower_p(0)(site) = p
                    threshold_flag_p = True
                    MIPI_threshold_found_p_value(site) = 1 'True = 1
                End If
                If threshold_flag_p = True Then
                    MIPI_threshold_high_p(0)(site) = p
                End If
            End If
            If MIPI_threshold_Code_value_P(p)(site) > 0 Then
                threshold_temp_p = threshold_temp_p + 1
            End If
        Next p
        
        If threshold_temp_p = 0 Then
            MIPI_threshold_found_p_value(site) = 0 'Flase = 0
        End If
        
        If MIPI_threshold_lower_p(0)(site) <> "" Then
            MIPI_threshold_lower_p(0)(site) = MIPI_trans_mapping(MIPI_threshold_lower_p(0)(site))
        Else
            MIPI_threshold_lower_p(0)(site) = 999
        End If

         If MIPI_threshold_high_p(0)(site) <> "" Then
            MIPI_threshold_high_p(0)(site) = MIPI_trans_mapping(MIPI_threshold_high_p(0)(site))
        Else
            MIPI_threshold_high_p(0)(site) = 999
        End If

       'code n process
        threshold_temp_n = 0
        threshold_flag_n = False
        'MIPI_threshold_found_n(Site) = False
        MIPI_threshold_found_n_value(site) = -1 'Clear = -1
        
        
        For n = 0 To 7

            If MIPI_threshold_Code_value_N(n)(site) = 0 Then
                If threshold_flag_n = False Then
                    MIPI_threshold_lower_n(0)(site) = n
                    threshold_flag_n = True
                    MIPI_threshold_found_n_value(site) = 1 'True = 1
                End If
                If threshold_flag_n = True Then
                    MIPI_threshold_high_n(0)(site) = n
                End If
            End If
            If MIPI_threshold_Code_value_N(n)(site) > 0 Then
                threshold_temp_n = threshold_temp_n + 1
            End If
        Next n
        
        If threshold_temp_n = 0 Then
            MIPI_threshold_found_n_value(site) = 0 'Flase = 0
        End If
        
        If MIPI_threshold_lower_n(0)(site) <> "" Then
            MIPI_threshold_lower_n(0)(site) = MIPI_trans_mapping(MIPI_threshold_lower_n(0)(site))
        Else
            MIPI_threshold_lower_n(0)(site) = 999
        End If

         If MIPI_threshold_high_n(0)(site) <> "" Then
            MIPI_threshold_high_n(0)(site) = MIPI_trans_mapping(MIPI_threshold_high_n(0)(site))
        Else
            MIPI_threshold_high_n(0)(site) = 999
        End If


    Next site
    If gl_Disable_HIP_debug_log = False Then
  ' print datdlog
    For p = 0 To 7
        TestNameInput = Report_TName_From_Instance(CalcC, "code_P_" & p, vbNullString, CLng(p))
        TheExec.Flow.TestLimit MIPI_threshold_Code_value_P(p), 0, 2 ^ 10 - 1, PinName:="code_P_" & p, ForceResults:=tlForceNone, Tname:=TestNameInput 'eng_forceflow_transfer
    Next p
    End If
    TestNameInput = Report_TName_From_Instance(CalcC, "DATA0_Term_Tol1", vbNullString, 0)
    TheExec.Flow.TestLimit MIPI_threshold_lower_p(0), scaletype:=scaleNone, PinName:="DATA0_Term_Tol1", ForceResults:=tlForceFlow, Tname:=TestNameInput
        
    TestNameInput = Report_TName_From_Instance(CalcC, "DATA0_Term_Tol2", vbNullString, 0)
    TheExec.Flow.TestLimit MIPI_threshold_high_p(0), scaletype:=scaleNone, PinName:="DATA0_Term_Tol2", ForceResults:=tlForceFlow, Tname:=TestNameInput
    
    TestNameInput = Report_TName_From_Instance(CalcC, "DATA0_Found_Thresh", vbNullString, 0)
    TheExec.Flow.TestLimit MIPI_threshold_found_p_value, 1, 1, PinName:="DATA0_Found_Thresh", ForceResults:=tlForceFlow, Tname:=TestNameInput
    
    If gl_Disable_HIP_debug_log = False Then
    
    For n = 0 To 7
        TestNameInput = Report_TName_From_Instance(CalcC, "code_N_" & n, vbNullString, CLng(n))
        TheExec.Flow.TestLimit MIPI_threshold_Code_value_N(n), 0, 2 ^ 10 - 1, PinName:="code_N_" & n, ForceResults:=tlForceNone, Tname:=TestNameInput 'eng_forceflow_transfer
        
    Next n
    End If

    TestNameInput = Report_TName_From_Instance(CalcC, "DATA1_Term_Tol1", vbNullString, 0)
    TheExec.Flow.TestLimit MIPI_threshold_lower_n(0), scaletype:=scaleNone, PinName:="DATA1_Term_Tol1", ForceResults:=tlForceFlow, Tname:=TestNameInput
    
    TestNameInput = Report_TName_From_Instance(CalcC, "DATA1_Term_Tol2", vbNullString, 0)
    TheExec.Flow.TestLimit MIPI_threshold_high_n(0), scaletype:=scaleNone, PinName:="DATA1_Term_Tol2", ForceResults:=tlForceFlow, Tname:=TestNameInput
        
    TestNameInput = Report_TName_From_Instance(CalcC, "DATA1_Found_Thresh", vbNullString, 0)
    TheExec.Flow.TestLimit MIPI_threshold_found_n_value, 1, 1, PinName:="DATA1_Found_Thresh", ForceResults:=tlForceFlow, Tname:=TestNameInput
    
End Function


Public Function Calc_ADC_Error_code(argc As Integer, argv() As String) As Long

Dim site As Variant
Dim ADC_Trim_Code As New DSPWave: ADC_Trim_Code.CreateConstant 0, 1, DspLong
Dim Error_Code As New DSPWave: Error_Code.CreateConstant 0, 1, DspLong
Dim ERROR_CODE_Dict As New DSPWave
Dim ADC_Error_Code_Str As String
Dim ADC_Trim_Code_Str As String
Dim REFERENCE_CTRL As Long
Dim ADC_Error_Code_Str_25 As String
Dim ADC_Error_Code_Str_85 As String
Dim Error_Code_25C_Dec As New DSPWave: Error_Code_25C_Dec.CreateConstant 0, 1, DspLong
Dim Error_Code_85C_Dec As New DSPWave: Error_Code_85C_Dec.CreateConstant 0, 1, DspLong
Dim Error_Code_25C As New DSPWave
Dim Error_Code_85C As New DSPWave
Dim SL_BitWidth As New SiteLong
Dim ADC_Final_RefCtrl_Str As String
Dim ADC_Final_RefCtrl As New DSPWave: ADC_Final_RefCtrl.CreateConstant 0, 1, DspLong
Dim ADC_Final_RefCtrl_Dict As New DSPWave
Dim OutputTname_format() As String
Dim TestNameInput As String



    ADC_Trim_Code_Str = argv(0)
    ADC_Error_Code_Str = argv(1)
    REFERENCE_CTRL = argv(2)

    Call HardIP_Bin2Dec(ADC_Trim_Code, GetStoredCaptureData(ADC_Trim_Code_Str))
    For Each site In TheExec.sites
        Error_Code(site).Element(0) = ADC_Trim_Code(site).Element(0) - REFERENCE_CTRL
    Next site
    
    TestNameInput = Report_TName_From_Instance(CalcC, ADC_Error_Code_Str, vbNullString)
    TheExec.Flow.TestLimit resultVal:=Error_Code.Element(0), lowVal:=-127, hiVal:=127, ForceResults:=tlForceNone, Tname:=TestNameInput 'eng_forceflow_transfer
        
    For Each site In TheExec.sites
        If Error_Code(site).Element(0) < -128 Then
            Error_Code(site).Element(0) = 128
        ElseIf Error_Code(site).Element(0) < 0 Then
            Error_Code(site).Element(0) = 2 ^ 8 + FormatNumber(Error_Code(site).Element(0))
        ElseIf Error_Code(site).Element(0) > 127 Then
            Error_Code(site).Element(0) = 127
        End If
    Next site
    Call HardIP_Dec2Bin(ERROR_CODE_Dict, Error_Code, 8)
    Call AddStoredCaptureData(ADC_Error_Code_Str, ERROR_CODE_Dict)
    
    
    If argc >= 4 Then
        ADC_Error_Code_Str_25 = argv(3)
        ADC_Error_Code_Str_85 = argv(1)
        ADC_Final_RefCtrl_Str = argv(4)
        
        Error_Code_25C = GetStoredCaptureData(ADC_Error_Code_Str_25)
        Error_Code_85C = GetStoredCaptureData(ADC_Error_Code_Str_85)
    
        For Each site In TheExec.sites
            SL_BitWidth(site) = Error_Code_25C(site).SampleSize
        Next site
        
        Call rundsp.DSP_2S_Complement_To_SignDec(Error_Code_25C, SL_BitWidth, Error_Code_25C_Dec)
        Call rundsp.DSP_2S_Complement_To_SignDec(Error_Code_85C, SL_BitWidth, Error_Code_85C_Dec)
    
        For Each site In TheExec.sites
            ADC_Final_RefCtrl(site).Element(0) = REFERENCE_CTRL + (Error_Code_25C_Dec(site).Element(0) + Error_Code_85C_Dec(site).Element(0)) / 2
        Next site
        TestNameInput = Report_TName_From_Instance(CalcC, "FinalReferenceControlCode", vbNullString)
        TheExec.Flow.TestLimit resultVal:=ADC_Final_RefCtrl.Element(0), ForceResults:=tlForceNone, Tname:=TestNameInput 'eng_forceflow_transfer
        
        Call HardIP_Dec2Bin(ADC_Final_RefCtrl_Dict, ADC_Final_RefCtrl, 8)
        Call AddStoredCaptureData(ADC_Final_RefCtrl_Str, ADC_Final_RefCtrl_Dict)
    
    
    End If
    
    
End Function

Public Function ADC_code_toV(argc As Integer, argv() As String) As Long '----------------add by CSHO 20171227

Dim ADCcapcode As String
Dim USBvoltages As String
Dim USBvoltages2 As String
Dim devideV As Long
Dim ADC_voltages As String
Dim InputKey As String
Dim DSP_Input As New DSPWave
Dim LSB As Double
Dim bitprint As String
Dim i As Integer
Dim ADC_voltages_final As Double
Dim site As Variant

InputKey = argv(0)
USBvoltages = argv(1)
devideV = argv(2)

Set DSP_Input = Nothing
DSP_Input = GetStoredCaptureData(InputKey)
 For Each site In TheExec.sites
 For i = 0 To DSP_Input.SampleSize - 1
    If i = 0 Then
      bitprint = DSP_Input.Element(0)
    
    Else
      bitprint = bitprint & DSP_Input.Element(i)
    
    End If
  Next i


ADC_voltages = Bin2Dec(bitprint)

USBvoltages2 = TheExec.Specs.DC.item(USBvoltages).ContextValue

LSB = CDbl(USBvoltages2) / devideV

ADC_voltages_final = ADC_voltages * LSB

If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "site" & site & "Convert ADC codes to voltages" & " V: " & ADC_voltages_final
Next site


End Function



Public Function Calc_MTR_REL_Freq_Diff_AVG(argc As Integer, argv() As String) As Long

    Dim site As Variant
    Dim freq_Dut As String
    Dim freq_ref As String
    
    Dim dut As String
    Dim ref As String

    Dim fdiff_percent As String
    Dim TestName As String
    Dim Efuse_Dict_Name As String
'    Dim DSP_fdiff_percent As String

'    Dim DSP_freq_Dut As String
'    Dim DSP_freq_ref As String
    Dim index_name As String
    Dim Index_count As Long
    Dim i, k As Integer

    Dim freq_Dut_dsp As New DSPWave: freq_Dut_dsp = Nothing
    Dim freq_ref_dsp As New DSPWave: freq_ref_dsp = Nothing
    Dim fdiff_percent_dsp As New DSPWave: fdiff_percent_dsp = Nothing
    
    Dim freq_Dut_wav As New DSPWave: freq_Dut_wav = Nothing
    Dim freq_ref_wav As New DSPWave: freq_ref_wav = Nothing
    Dim fdiff_percent_wav As New DSPWave: fdiff_percent_wav = Nothing
    
    Dim freq_Dut_mean As New SiteDouble
    Dim freq_ref_mean As New SiteDouble
    Dim fdiff_percent_mean As New SiteDouble
    
    Dim freq_Dut_std As New SiteDouble
    Dim freq_ref_std As New SiteDouble
    Dim fdiff_percent_std As New SiteDouble
    
    Dim RSD_DUT As New SiteDouble
    Dim RSD_REF As New SiteDouble
    Dim R_Ref As New SiteDouble
    
    Dim DSP_fdiff_percent As New DSPWave: DSP_fdiff_percent.CreateConstant 0, 1, DspDouble
    Dim DSP_freq_Dut As New DSPWave: DSP_freq_Dut = Nothing
    Dim DSP_freq_ref As New DSPWave: DSP_freq_ref = Nothing
    
    Dim dut_array() As String
    Dim ref_array() As String
    Dim freq_Dut_array() As String
    Dim freq_ref_array() As String
    Dim fdiff_percent_array() As String
    Dim Check_Freq As New SiteBoolean
    Dim Check_STD As New SiteBoolean
    Dim Check_Ratio As New SiteBoolean
    
    Dim Freq_HiLimit As Double: Freq_HiLimit = 1150000000
    Dim Freq_LoLimit As Double: Freq_LoLimit = 650000000
    Dim STD_HiLimit As Double: STD_HiLimit = 0.2
    Dim STD_LoLimit As Double: STD_LoLimit = 0
    Dim F_Ratio_HiLimit As Double: F_Ratio_HiLimit = 106
    Dim F_Ratio_LoLimit As Double: F_Ratio_LoLimit = 94
    Dim Fuse_Code() As New DSPWave
    Dim Final_Fuse_Code As New DSPWave: Final_Fuse_Code = Nothing
    Dim Final_Fuse_Code_DEC As New DSPWave: Final_Fuse_Code_DEC = Nothing
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    

    Dim xxx As New DSPWave
    Dim yyy As New DSPWave
    Final_Fuse_Code_DEC.CreateConstant 0, 1, DspDouble
'    yyy.CreateConstant 0, 8, DspLong
'    xxx.CreateConstant 0, 8, DspLong
'    xxx(0).Element(0) = 1
'
'    xxx = xxx.ConvertStreamTo(tldspParallel, 8, 0, Bit0IsMsb)
'
'    xxx = xxx.ConvertDataTypeTo(DspLong)
'    xxx(0).Element(0) = 128
'    yyy = xxx.ConvertStreamTo(tldspSerial, 8, 0, Bit0IsMsb)




    dut = argv(0)
    ref = argv(1)
    freq_Dut = argv(2)
    freq_ref = argv(3)
    fdiff_percent = argv(4)
    
    index_name = argv(5)
    Index_count = argv(6)
    Efuse_Dict_Name = argv(7)
    
    
    dut_array = Split(dut, "@")
    ref_array = Split(ref, "@")
    freq_Dut_array = Split(freq_Dut, "@")
    freq_ref_array = Split(freq_ref, "@")
    fdiff_percent_array = Split(fdiff_percent, "@")
    
    
    For k = 0 To UBound(dut_array)
    
        TestName = "f_diff_" + Replace(freq_Dut_array(k), index_name, TheExec.Flow.var(index_name).value)
        DSP_freq_Dut = GetStoredCaptureData(dut_array(k))
        DSP_freq_ref = GetStoredCaptureData(ref_array(k))
        For Each site In TheExec.sites.Active
            DSP_freq_Dut = DSP_freq_Dut.ConvertStreamTo(tldspParallel, 16, 0, Bit0IsMsb)
            DSP_freq_ref = DSP_freq_ref.ConvertStreamTo(tldspParallel, 16, 0, Bit0IsMsb)
            DSP_freq_Dut = DSP_freq_Dut.Multiply(93750)
            DSP_freq_ref = DSP_freq_ref.Multiply(93750)
        Next site
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "======================================================================Start Calc Freq======================================================================"
        For Each site In TheExec.sites
            If DSP_freq_Dut.Element(0) <> 0 Then
                DSP_fdiff_percent.Element(0) = ((DSP_freq_Dut.Element(0) - DSP_freq_ref.Element(0)) / DSP_freq_Dut.Element(0)) * 100
            Else
                DSP_fdiff_percent.Element(0) = 99999
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site:" + CStr(site) + "  freq_of_Dut " + freq_Dut_array(k) + " is 0")
            End If
        Next site
        
        TestNameInput = Report_TName_From_Instance(CalcF, "X", Replace(freq_Dut_array(k), index_name, vbNullString), CInt(TheExec.Flow.var(index_name).value))
        
        TheExec.Flow.TestLimit resultVal:=DSP_freq_Dut.Element(0), Tname:=Replace(freq_Dut_array(k), index_name, TheExec.Flow.var(index_name).value), ForceResults:=tlForceNone, scaletype:=scaleMega 'eng_forceflow_transfer
        
        TestNameInput = Report_TName_From_Instance(CalcF, "X", Replace(freq_ref_array(k), index_name, vbNullString), CInt(TheExec.Flow.var(index_name).value))
        
        TheExec.Flow.TestLimit resultVal:=DSP_freq_ref.Element(0), Tname:=Replace(freq_ref_array(k), index_name, TheExec.Flow.var(index_name).value), ForceResults:=tlForceNone, scaletype:=scaleMega 'eng_forceflow_transfer
        
        TestNameInput = Report_TName_From_Instance(CalcF, "X", "Percent", CInt(TheExec.Flow.var(index_name).value))
        
        TheExec.Flow.TestLimit resultVal:=DSP_fdiff_percent.Element(0), Tname:=TestName, ForceResults:=tlForceNone 'eng_forceflow_transfer
            
        
    
        
        
        Call AddStoredCaptureData(Replace(freq_Dut_array(k), index_name, TheExec.Flow.var(index_name).value), DSP_freq_Dut)
        Call AddStoredCaptureData(Replace(freq_ref_array(k), index_name, TheExec.Flow.var(index_name).value), DSP_freq_ref)
    
        Call AddStoredCaptureData(Replace(fdiff_percent_array(k), index_name, TheExec.Flow.var(index_name).value), DSP_fdiff_percent)
        
        
        
        
        If (TheExec.Flow.var(index_name).value + 1 = Index_count) Then
           If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "======================================================================Start Calc Mean,SD,Ratio======================================================================"
            Set freq_Dut_dsp = New DSPWave
            freq_Dut_dsp.CreateConstant 0, Index_count
            Set freq_ref_dsp = New DSPWave
            freq_ref_dsp.CreateConstant 0, Index_count
            Set fdiff_percent_dsp = New DSPWave
            fdiff_percent_dsp.CreateConstant 0, Index_count
            
            
            For Each site In TheExec.sites.Active
                Check_Freq = True
                Check_STD = True
                Check_Ratio = True
            Next site
            
            For i = 0 To Index_count - 1
                freq_Dut_wav = GetStoredCaptureData(Replace(freq_Dut_array(k), index_name, CStr(i)))
                freq_ref_wav = GetStoredCaptureData(Replace(freq_ref_array(k), index_name, CStr(i)))
                fdiff_percent_wav = GetStoredCaptureData(Replace(fdiff_percent_array(k), index_name, CStr(i)))
                
                For Each site In TheExec.sites.Active
                    freq_Dut_dsp.Element(i) = freq_Dut_wav.Element(0)
                    If (freq_Dut_wav.Element(0) < Freq_HiLimit And freq_Dut_wav.Element(0) > Freq_LoLimit) Then
                        Check_Freq = Check_Freq And True
                    Else
                        Check_Freq = False
                    End If
                    
                    freq_ref_dsp.Element(i) = freq_ref_wav.Element(0)
                        If (freq_ref_wav.Element(0) < Freq_HiLimit And freq_ref_wav.Element(0) > Freq_LoLimit) Then
                        Check_Freq = Check_Freq And True
                    Else
                        Check_Freq = False
                    End If
                    fdiff_percent_dsp.Element(i) = fdiff_percent_wav.Element(0)
                Next site
            Next i
            
            Dim freq_Dut_std_dbl As Double
            Dim freq_ref_std_dbl As Double
            
            
            
            For Each site In TheExec.sites.Active
                freq_Dut_mean = freq_Dut_dsp.CalcMeanWithStdDev(freq_Dut_std_dbl)
                freq_ref_mean = freq_ref_dsp.CalcMeanWithStdDev(freq_ref_std_dbl)
                fdiff_percent_mean = fdiff_percent_dsp.CalcMean
                
'                freq_Dut_dsp.CalcMeanWithStdDev (freq_Dut_std)
'                freq_ref_dsp.CalcMeanWithStdDev (freq_ref_std)
'                fdiff_percent_dsp.CalcMeanWithStdDev (fdiff_percent_std)
                If freq_Dut_mean = 0 Then
                    RSD_DUT = 0
                Else
                    RSD_DUT = 3 * freq_Dut_std_dbl / freq_Dut_mean * 100
                End If
                If (RSD_DUT < STD_HiLimit And RSD_DUT > STD_LoLimit) Then
                    Check_STD = Check_STD And True
                Else
                    Check_STD = False
                End If
                
                If freq_ref_mean = 0 Then
                    RSD_REF = 0
                Else
                    RSD_REF = 3 * freq_ref_std_dbl / freq_ref_mean * 100
                End If
                If (RSD_REF < STD_HiLimit And RSD_REF > STD_LoLimit) Then
                    Check_STD = Check_STD And True
                Else
                    Check_STD = False
                End If
                If freq_Dut_mean = 0 Then
                    R_Ref = 0
                Else
                    R_Ref = freq_ref_mean / freq_Dut_mean * 100
                End If
                If (R_Ref < F_Ratio_HiLimit And R_Ref > F_Ratio_LoLimit) Then
                    Check_Ratio = Check_Ratio And True
                Else
                    Check_Ratio = False
                End If
                
            Next site

            TheExec.Flow.TestLimit resultVal:=freq_Dut_mean, Tname:="freq_Dut_mean", ForceResults:=tlForceFlow
            TheExec.Flow.TestLimit resultVal:=freq_ref_mean, Tname:="freq_ref_mean", ForceResults:=tlForceFlow
            TheExec.Flow.TestLimit resultVal:=RSD_DUT, Tname:="RSD_DUT", ForceResults:=tlForceFlow
            TheExec.Flow.TestLimit resultVal:=RSD_REF, Tname:="RSD_REF", ForceResults:=tlForceFlow
            TheExec.Flow.TestLimit resultVal:=R_Ref, Tname:="R_Ref", ForceResults:=tlForceFlow
            TheExec.Flow.TestLimit resultVal:=fdiff_percent_mean, Tname:="Avg_R0t0_E3", ForceResults:=tlForceFlow

            ReDim Fuse_Code(UBound(dut_array)) As New DSPWave
            Dim fuse_code_dec As New DSPWave
            
            Set Fuse_Code(k) = New DSPWave
            Fuse_Code(k).CreateConstant 0, 16, DspLong
            Set fuse_code_dec = New DSPWave
            fuse_code_dec.CreateConstant 0, 1, DspLong
            
            For Each site In TheExec.sites.Active
                If Check_Freq = False Then
                    fuse_code_dec.Element(0) = 65533    '0xFFFD
                ElseIf Check_STD = False Then
                    fuse_code_dec.Element(0) = 65534    '0xFFFE
                ElseIf Check_Ratio = False Then
                    fuse_code_dec.Element(0) = 65532    '0xFFFC
                Else
                    If (fdiff_percent_mean >= 0) Then
                        fuse_code_dec.Element(0) = Abs(fdiff_percent_mean) * 1000
                    Else
                        fuse_code_dec.Element(0) = Abs(fdiff_percent_mean) * 1000 + 32768
                    End If
                End If
                fuse_code_dec = fuse_code_dec.ConvertDataTypeTo(DspLong)
                Fuse_Code(k) = fuse_code_dec.ConvertStreamTo(tldspSerial, 16, 0, Bit0IsMsb)
               If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Efuse Write,Site:" + CStr(site) + " Value: " + CStr(fuse_code_dec.Element(0))
                Final_Fuse_Code = Final_Fuse_Code.Concatenate(Fuse_Code(k))
                Final_Fuse_Code_DEC.Element(0) = Final_Fuse_Code_DEC.Element(0) * (2 ^ 16) * k + fuse_code_dec.Element(0)
            Next site

        End If
    Next k
    If (TheExec.Flow.var(index_name).value + 1 = Index_count) Then
    
        If gl_Disable_HIP_debug_log = False Then
            For Each site In TheExec.sites.Active
                TheExec.Datalog.WriteComment "Final Efuse Write Value , Site:" + CStr(site) + " Value: " + CStr(Final_Fuse_Code_DEC.Element(0))
            Next site
        End If
        
        Call AddStoredCaptureData(Efuse_Dict_Name, Final_Fuse_Code_DEC)
    End If
End Function

Public Function Calc_MTR_AVG(argc As Integer, argv() As String) As Long

'    Dim index_name As String
'    Dim Sweep_Dictionary As String
    Dim Loop_count As Long
    Dim Loop_Index As Long
    
    Dim DSP_Capture As New DSPWave
    Dim i, j As Long

    Dim Sweep_index As Long
    Dim Sweep_Info() As Power_Sweep
    Dim Sweep_Count As Long
    Dim dict_key As String
    Dim DSP_Result As New DSPWave
    Dim site As Variant
    Dim Sweep_Mean As New SiteDouble
    Dim OutputTname_format() As String
    Dim TestNameInput As String
    
    Loop_Index = argv(argc - 1)

    ReDim Sweep_Info(argc - 2) As Power_Sweep
    
    
    For i = 0 To argc - 2
        Sweep_Info(i).PinName = Split(argv(i), "@")(1)
        Sweep_Info(i).from = Split(argv(i), "@")(3)
        Sweep_Info(i).stop = Split(argv(i), "@")(4)
        Sweep_Info(i).step = Split(argv(i), "@")(5)
        If (CDbl(Sweep_Info(i).stop) < CDbl(Sweep_Info(i).from)) Then Sweep_Info(i).step = "-" & Sweep_Info(i).step
        Sweep_Info(i).Loop_Index_Name = Split(argv(i), "@")(6)
        Sweep_Info(i).Loop_count = Split(argv(i), "@")(7)
        Sweep_Info(i).key = Split(argv(i), "@")(8)
    Next i
    
    Sweep_Count = CLng(Abs((Sweep_Info(0).stop - Sweep_Info(0).from) / Sweep_Info(0).step)) + 1
    Loop_count = CLng(Sweep_Info(0).Loop_count)
    
    If (TheExec.Flow.var(Sweep_Info(0).Loop_Index_Name).value = Loop_count - 1 And Loop_Index = Sweep_Count - 1) Then
        
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "====================================Start Calc Mean===================================="

        
        For j = 0 To Sweep_Count - 1
            dict_key = vbNullString
            For Sweep_index = 0 To UBound(Sweep_Info)
                If dict_key = "" Then
                    dict_key = Replace(CStr(CDbl(Sweep_Info(Sweep_index).from) + CDbl(Sweep_Info(Sweep_index).step) * j), ".", "p")
                Else
                    dict_key = dict_key & "_" & Replace(CStr(CDbl(Sweep_Info(Sweep_index).from) + CDbl(Sweep_Info(Sweep_index).step) * j), ".", "p")
                End If
            Next Sweep_index
            
            dict_key = Sweep_Info(0).key & "_" & dict_key
            For Each site In TheExec.sites.Active
                DSP_Result.CreateConstant 0, Loop_count, DspDouble
            Next site
            For i = 0 To Loop_count - 1
                
                DSP_Capture = GetStoredCaptureData(dict_key & "_" & CStr(i))
                For Each site In TheExec.sites.Active
                    DSP_Capture = DSP_Capture.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspParallel, DSP_Capture.SampleSize, 0, Bit0IsMsb)
                'For Each site In TheExec.sites.Active
                    DSP_Result.Element(i) = DSP_Capture.Element(0)
                Next site
            Next i
            For Each site In TheExec.sites.Active
                Sweep_Mean = DSP_Result.CalcMean
            Next site
            
            TestNameInput = Report_TName_From_Instance(CalcC, "X", dict_key & "Mean", CInt(j))
                        
            TheExec.Flow.TestLimit resultVal:=Sweep_Mean, Tname:=TestNameInput, ForceResults:=tlForceFlow
        Next j
    End If
    
'    index_name = argv(0)
'    Sweep_Dictionary = argv(1)
'    Loop_Count = argv(2)
'    Loop_Index = theexec.Flow.var(index_name).Value
'
'
'    If (Loop_Index = Loop_Count - 1) Then
'        DSP_Result.CreateConstant 0, Loop_Count, DspLong
'        Dictionary_Key = Split(Sweep_Dictionary, ":")
'
'        For Each key In Dictionary_Key
'            DSP_Capture = GetStoredCaptureData(CStr(key))
'
'        Next key
'
'
'    End If
    
    

End Function



Public Function USB3_ADC(argc As Integer, argv() As String) As Long '----------------add by CSHO 20171227

Dim ADCcapcode As String
Dim USBvoltages As String
Dim USBvoltages2 As String
Dim devideV As Long
Dim ADC_voltages As String
Dim InputKey As String
Dim DSP_Input As New DSPWave
Dim DSP_Input_2 As New SiteDouble
Dim ADC_Output As New SiteDouble
Dim LSB As Double
Dim bitprint As String
Dim i As Integer
Dim ADC_voltages_final As Double
Dim MinusValue As Double
Dim OutputTname_format() As String
Dim TestNameInput As String


USBvoltages = argv(0)
MinusValue = ProcessEvaluateDCSpec(USBvoltages)

devideV = argv(1)
DSP_Input.CreateConstant 0, 1, DspDouble

For i = 2 To argc - 1
    InputKey = argv(i)
    Set DSP_Input = Nothing
    DSP_Input_2 = GetStoredData(InputKey & "_para")
    'DSP_Input = GetStoredCaptureData(InputKey & "_para")
    ADC_Output = ADC_Output.Add(DSP_Input_2)
    ADC_Output = ADC_Output.Multiply(MinusValue).divide(devideV)
    TestNameInput = Report_TName_From_Instance(CalcC, InputKey, "_ADC", CInt(i - 2))
    TheExec.Flow.TestLimit resultVal:=ADC_Output, ForceResults:=tlForceFlow, Tname:=TestNameInput
Next i

End Function

Public Function Print_Shmoo_Voltage(argc As Integer, argv() As String) As Long

Dim i As Long
Dim z As Long

Dim Input_Pins() As String
Dim num_pins As Long
Dim Voltage_Value As Double

For i = 0 To argc - 1
    TheExec.DataManager.DecomposePinList argv(i), Input_Pins(), num_pins
    For z = 0 To UBound(Input_Pins)
        Voltage_Value = TheHdw.DCVS.Pins(Input_Pins(z)).Voltage.Main
        TheExec.Datalog.WriteComment Input_Pins(z) & " Vmain=" & Voltage_Value
        Voltage_Value = TheHdw.DCVS.Pins(Input_Pins(z)).Voltage.Alt
        TheExec.Datalog.WriteComment Input_Pins(z) & " Valt=" & Voltage_Value
    Next z

Next i


End Function
Public Function Calc_memcheck(argc As Integer, argv() As String) As Long
    
    Dim temp_dsp As New DSPWave
    Dim dataWave As New DSPWave
    Dim hexWave As New DSPWave
    Dim i As Long
    Dim CurSite As Variant
    Dim HexStr As String
    Dim DataFormat As String: DataFormat = "Hex"
    Dim cap_dec_data As New SiteLong
    Dim dc_read As New SiteLong: dc_read = 1
    Dim j As Integer
    Dim first_flag As New SiteBoolean
    Dim second_flag As New SiteBoolean
    Dim Dec_Str_All(3) As New DSPWave

        
        first_flag = False
        second_flag = False
        For i = 0 To argc - 1
            Dec_Str_All(i).CreateConstant 0, 4, DspLong
        Next i
    For i = 0 To argc - 1
        temp_dsp = GetStoredCaptureData(argv(i))
        For Each CurSite In TheExec.sites
            HexStr = vbNullString
            ' convert bits to hex formatted stream
            Dim bin_str As String
            bin_str = vbNullString
               For j = 0 To temp_dsp.SampleSize - 1
                    bin_str = bin_str & temp_dsp.Element(j)
            Next j
            bin_str = StrReverse(bin_str)
            TheExec.Datalog.WriteComment "(MSB -> LSB)"
            TheExec.Datalog.WriteComment bin_str

            hexWave = temp_dsp.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspParallel, 4, 0, Bit0IsMsb)

            For j = (hexWave.SampleSize - 1) To 0 Step -1
                    HexStr = HexStr + Hex(hexWave.Element(j))
            Next j

               cap_dec_data(CurSite) = CLng("&H" & CStr(HexStr))
               dc_read(CurSite) = CLng("&H" & CStr(Hex(temp_dsp.Element(temp_dsp.SampleSize - 2)))) * dc_read(CurSite)  ' If two cycle HSC_READ are "1" then read=1
               TheExec.Datalog.WriteComment " Hex:  0x " & HexStr
               Dec_Str_All(i)(CurSite).Element(0) = cap_dec_data 'store data
        Next CurSite
        TheExec.Flow.TestLimit resultVal:=cap_dec_data, ForceResults:=tlForceFlow, unit:=unitCustom, customUnit:=vbNullString   ', Tname:="FailBitCount", Unit:=unitNone, ScaleType:=scaleNone
    Next i


'/////////////// Judgement passing flag///////////////////
    
    Dim temp_HexStr As String
        For Each CurSite In TheExec.sites
            temp_HexStr = vbNullString
            HexStr = vbNullString
            For i = 0 To argc - 1
             temp_dsp = GetStoredCaptureData(argv(i))
                HexStr = CStr(Hex(Dec_Str_All(i)(CurSite).Element(0)))
                If first_flag = False Then
                    If HexStr = "E910" And temp_dsp(CurSite).Element(temp_dsp.SampleSize - 1) = 1 Then
                        first_flag = True
                        temp_HexStr = HexStr
                    ElseIf HexStr = "91E" And temp_dsp(CurSite).Element(temp_dsp.SampleSize - 1) = 0 Then
                        first_flag = True
                        temp_HexStr = HexStr
                    Else
                        first_flag = False
                    End If
                 Else
                    If HexStr <> temp_HexStr Then
                        Select Case HexStr
                            Case "E910":
                                If temp_dsp(CurSite).Element(temp_dsp.SampleSize - 1) = 1 Then second_flag = True
                            Case "91E":
                                If temp_dsp(CurSite).Element(temp_dsp.SampleSize - 1) = 0 Then second_flag = True
                        End Select
                    End If
                 End If
             Next i
        Next CurSite
'////////////////////////////////////////////////////////////
    TheExec.Flow.TestLimit resultVal:=dc_read, ForceResults:=tlForceFlow, unit:=unitCustom, customUnit:=vbNullString
    TheExec.Flow.TestLimit resultVal:=second_flag, ForceResults:=tlForceFlow, unit:=unitCustom, customUnit:=vbNullString

End Function

Public Function LP5_LB_PI(argc As Integer, argv() As String) As Long

   'New LP5 eye model 20190417
   
   Dim i As Long, j As Long, k As Long, L As Long
   Dim site As Variant
   Dim SplitByAt() As String
   Dim DSP_Captured() As New DSPWave
   Dim DSP_EYE() As New DSPWave
   Dim tmp_element As Long
   Dim tmp_name As String
   Dim EYE_arr() As Long
   Dim DSP_INV() As New DSPWave
   Dim DSP_CK() As New DSPWave
   Dim DSP_CKTemp() As New DSPWave
   Dim DSP_INVTemp() As New DSPWave
   ReDim DSP_INV(CStr(argc) - 1)
   ReDim DSP_CK(CStr(argc) - 1)
   ReDim DSP_INVTemp(CStr(argc) - 1)
   ReDim DSP_CKTemp(CStr(argc) - 1)
   
   Dim tmp_max_eye As Long
   Dim Eye_str As String
   Dim Eye_str_result() As New SiteVariant
   Dim Eye_str_long As New DSPWave
   Eye_str_long.CreateConstant 0, CLng(argc)
   ReDim Eye_str_result(CStr(argc))
   'ReDim Eye_str_long(CStr(argc)) As String
   Dim TestNameInput As String
   Dim DSP_Record() As New SiteVariant
   
   'argv(0) = "WCK0Sweep_2@WCK0Sweep_3@INVDQ0Sweep_0@INVDQ0Sweep_1"
   'argv(1) = "WCK1Sweep_2@WCK1Sweep_3@INVDQ1Sweep_0@INVDQ1Sweep_1"

   '' Split DSPWave captured to number of components of sweep
   
   
For Each site In TheExec.sites

   For i = 0 To argc - 1
      SplitByAt = Split(argv(i), "@") ' list of sweep names in order of concatination should be performed and INV if reverse is required
       ReDim Preserve DSP_Record((UBound(SplitByAt) + 1) * CStr(argc) - 1)
      ' Resize capture and final EYE DSPWaves to
      ReDim DSP_Captured(UBound(SplitByAt))
      ReDim DSP_EYE(UBound(SplitByAt))
      ReDim EYE_arr(UBound(SplitByAt))
      ReDim DSP_INV(UBound(SplitByAt))
      'ReDim Preserve DSP_INV(UBound(SplitByAt))

      Set DSP_EYE(i) = DSP_EYE(i).ConvertDataTypeTo(DspLong)
      Set DSP_INV(i) = DSP_INV(i).ConvertDataTypeTo(DspLong)
      Set DSP_CK(i) = DSP_CK(i).ConvertDataTypeTo(DspLong)
      ' ============== Prepare data capture for calculation ==============
        For j = 0 To UBound(SplitByAt)
        ' ======= INV Data order MSB -> LSB require inversion =======
            If SplitByAt(j) Like "INV*" Then
                tmp_name = mid(SplitByAt(j), 4) ' remove INV from the beginning
                'tmp_name = SplitByAt(j)
                DSP_Captured(j) = GetStoredCaptureData(tmp_name)
                
                DSP_INVTemp(i).CreateConstant 0, DSP_Captured(j).SampleSize, DspLong
                
'                For k = DSP_Captured(j).SampleSize To 1 Step -1
                For k = 0 To DSP_Captured(j).SampleSize - 1
                                  
                    DSP_INVTemp(i).Element(k) = DSP_Captured(j).Element(UBound(DSP_Captured(j).data) - k)
                    DSP_Record(i * (UBound(SplitByAt) + 1) + j) = CStr(DSP_Record(i * (UBound(SplitByAt) + 1) + j)) & CStr(DSP_INVTemp(i).Element(k))
                   
                Next k
                 
 
                Set DSP_INV(i) = DSP_INV(i).Concatenate(DSP_INVTemp(i)) 'Merge all need flipped bit into one DSP
                
      
            Else
                
                DSP_Captured(j) = GetStoredCaptureData(SplitByAt(j))
                DSP_CKTemp(i).CreateConstant 0, DSP_Captured(j).SampleSize, DspLong
                
                For k = 0 To DSP_Captured(j).SampleSize - 1
                    DSP_CKTemp(i).Element(k) = DSP_Captured(j).Element(UBound(DSP_Captured(j).data) - k)
                    DSP_Record(i * (UBound(SplitByAt) + 1) + j) = DSP_Record(i * (UBound(SplitByAt) + 1) + j) & CStr(DSP_Captured(j).Element(UBound(DSP_Captured(j).data) - k))
                Next k
               Set DSP_CK(i) = DSP_CK(i).Concatenate(DSP_CKTemp(i))
            
   
            End If
    
            If UCase(SplitByAt(j)) Like "*WCK*" Then
                DSP_Record(i * (UBound(SplitByAt) + 1) + j) = "WCK:" & DSP_Record(i * (UBound(SplitByAt) + 1) + j)
            ElseIf UCase(SplitByAt(j)) Like "*CK*" Then
               DSP_Record(i * (UBound(SplitByAt) + 1) + j) = "CK:" & DSP_Record(i * (UBound(SplitByAt) + 1) + j)
            Else
               DSP_Record(i * (UBound(SplitByAt) + 1) + j) = "INV:" & DSP_Record(i * (UBound(SplitByAt) + 1) + j)
            End If
            
        Next j
        
        '/////// INV_Cap reverse ////////////

        Dim iMidPt As Long
        Dim iUpper As Long
          iUpper = UBound(DSP_INV(i).data)
          iMidPt = (UBound(DSP_INV(i).data) - LBound(DSP_INV(i).data)) \ 2 + LBound(DSP_INV(i).data)
          For k = LBound(DSP_INV(i).data) To iMidPt
              tmp_element = DSP_INV(i).Element(iUpper)
              DSP_INV(i).Element(iUpper) = DSP_INV(i).Element(k)
              DSP_INV(i).Element(k) = tmp_element
              iUpper = iUpper - 1
          Next k
        '////////////////////////////////////
        
        
      ' next sweep register sw0, sw1, ...

       ' ====== Concat EYE data ============

        Set DSP_EYE(i) = DSP_CK(i).Concatenate(DSP_INV(i))  ' Concatenate UnFlip code + INV Flip code
         'Set DSP_EYE(i) = DSP_INV(i).Concatenate(DSP_CK(i))
      '==========================================================================

      'theexec.Datalog.WriteComment "EYE " & i
      For k = 0 To DSP_EYE(i).SampleSize - 1
          'Debug.Print DSP_EYE(i).Element(k);
          If k = 0 Then
            Eye_str = DSP_EYE(i).Element(k)
            Else
            Eye_str = Eye_str & DSP_EYE(i).Element(k)
            End If

      Next k

      Eye_str_result(i) = Eye_str

      'Debug.Print
      'theexec.Datalog.WriteComment Eye_str
      '====== Calculate number of 'ones' in the EYE ===========
      tmp_max_eye = 0 ' reset tmp_max_eye
      For k = 0 To DSP_EYE(i).SampleSize - 1
         If DSP_EYE(i).Element(k) = 1 Then
            EYE_arr(i) = EYE_arr(i) + 1
         Else
            If tmp_max_eye < EYE_arr(i) Then
                tmp_max_eye = EYE_arr(i) ' update max eye width
            End If
            EYE_arr(i) = 0
         End If
      Next k

             If tmp_max_eye < EYE_arr(i) Then
                 tmp_max_eye = EYE_arr(i) ' update max eye width
             End If

      Eye_str_long(site).Element(i) = tmp_max_eye
      
      
'''      'T_Name Edit
'''      '**************************************
'''      TestNameInput = "EYEDDR" & CStr(i)
'''      TestNameInput = Report_TName_From_Instance("X", "x", TestNameInput, CInt(theexec.Flow.TestLimitIndex), 0)
'''      theexec.Flow.TestLimit resultVal:=Eye_str_long(i), FormatStr:="%i", TName:=TestNameInput, ForceResults:=tlForceFlow
'''      '**************************************

'      theexec.Datalog.WriteComment "EYE " & i & " width " & tmp_max_eye
     Next i ' next DDR bus : DQ0, DQ1, CA0, CA1 ...
    '=============================================================================
Next site

   'T_Name Edit
      '**************************************
    Dim TnumRecord As Long
    
    For i = 0 To argc - 1
        
            TnumRecord = TheExec.sites.item(site).TestNumber
            TestNameInput = "EYEDDR" & CStr(i)
            TestNameInput = Report_TName_From_Instance("X", "x", TestNameInput, CInt(TheExec.Flow.TestLimitIndex), 0)
        
            For Each site In TheExec.sites
                TheExec.Flow.TestLimit resultVal:=Eye_str_long.Element(i), formatStr:="%i", Tname:=TestNameInput, ForceResults:=tlForceFlow, TNum:=TnumRecord, scaletype:=scaleNoScaling
        '            TheExec.Flow.TestLimit lowVal:=mdll_low(i)(Site), resultVal:=Eye_str_long(Site).Element(i) * 4, FormatStr:="%i", TName:=TestNameInput, ForceResults:=tlForceFlow, TNum:=TnumRecord
                TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
            Next site
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
            'TheExec.sites.item(Site).TestNumber = TheExec.sites.item(Site).TestNumber + 1
    Next i
   '**************************************
    
    
For Each site In TheExec.sites
   TheExec.Datalog.WriteComment "/////////" & "Site: " & site & "/////////"
      Count = 0
    For L = 0 To argc - 1

        SplitByAt = Split(argv(L), "@")

        For i = 0 To UBound(SplitByAt)
           If SplitByAt(i) Like "INV*" Then
           TheExec.Datalog.WriteComment mid(SplitByAt(i), 4)
           Else
           TheExec.Datalog.WriteComment SplitByAt(i)
           End If
           
           TheExec.Datalog.WriteComment DSP_Record(Count)
           Count = Count + 1
        Next i
        '****************************************
        TheExec.Datalog.WriteComment "EYE " & L
        TheExec.Datalog.WriteComment Eye_str_result(L)
        TheExec.Datalog.WriteComment "EYE " & L & " width " & Eye_str_long(site).Element(L)
    Next L

   
Next site

    TheExec.Datalog.WriteComment " ------------------ End ---"
    TheExec.Datalog.WriteComment "                           "




End Function
Public Function LP5_LB_DLL(argc As Integer, argv() As String) As Long

   'New LP5 eye model 20190417
   
    Dim i As Long, j As Long, k As Long, L As Long, z As Long
    Dim site As Variant
    Dim SplitByAt() As String
    Dim DictKey As String
    Dim DSP_Captured() As New DSPWave
    Dim DSP_EYE() As New DSPWave
    Dim tmp_element As Long
    Dim tmp_name As String
    Dim EYE_arr() As Long
    Dim DSP_INV() As New DSPWave
    Dim DSP_CK() As New DSPWave
    Dim DSP_CKTemp() As New DSPWave
    Dim DSP_INVTemp() As New DSPWave
    ReDim DSP_INV(CStr(argc) - 1)
    ReDim DSP_CK(CStr(argc) - 1)
    ReDim DSP_INVTemp(CStr(argc) - 1)
    ReDim DSP_CKTemp(CStr(argc) - 1)
    Dim tmp_max_eye As Long
    Dim Eye_str As String
    'Dim Eye_str_result() As String
    Dim Eye_str_long As New DSPWave
    Eye_str_long.CreateConstant 0, CLng(argc - 1)
    Dim Eye_str_result() As New SiteVariant
    ReDim Eye_str_result(CStr(argc))
    'ReDim Eye_str_long(CStr(argc)) As String
    Dim TestNameInput As String
    Dim OutputTname_formatQQ() As String
    Dim Mdll_value() As String
    Dim Mdll_ChannelInfo() As String
    
    Dim mdll_12x8 As New DSPWave
    Dim mdll As New SiteDouble
    Dim mdll_low() As New SiteDouble
    ReDim mdll_low(CLng(argc - 2))
    Dim mdll_high() As New SiteDouble
    ReDim mdll_high(CLng(argc - 2))
    Dim Mdll_width As Long
   
   
    Dim DSP_Record() As New SiteVariant
   
   
   'argv(0) = "WCK0Sweep_2@WCK0Sweep_3@INVDQ0Sweep_0@INVDQ0Sweep_1"
   'argv(1) = "WCK1Sweep_2@WCK1Sweep_3@INVDQ1Sweep_0@INVDQ1Sweep_1"
   'argv(2) = "ch0_mdll_w210|ch0_mdll_w543|ch0_mdll_w76|ch1_mdll_w210|ch1_mdll_w543|ch1_mdll_w76"    'for mdll high low clac
   '' Split DSPWave captured to number of components of sweep
   
    For Each site In TheExec.sites

        For i = 0 To argc - 2

            SplitByAt = Split(argv(i), "@") ' list of sweep names in order of concatination should be performed and INV if reverse is required
            ReDim Preserve DSP_Record((UBound(SplitByAt) + 1) * CStr(argc) - 2)

          ' Resize capture and final EYE DSPWaves to
            ReDim DSP_Captured(UBound(SplitByAt))
            ReDim DSP_EYE(UBound(SplitByAt))
            ReDim EYE_arr(UBound(SplitByAt))
            ReDim DSP_INV(UBound(SplitByAt))
            Set DSP_EYE(i) = DSP_EYE(i).ConvertDataTypeTo(DspLong)
            Set DSP_INV(i) = DSP_INV(i).ConvertDataTypeTo(DspLong)
            Set DSP_CK(i) = DSP_CK(i).ConvertDataTypeTo(DspLong)
     
      ' ============== Prepare data capture for calculation ==============
            For j = 0 To UBound(SplitByAt)
        ' ======= INV Data order MSB -> LSB require inversion =======
                If SplitByAt(j) Like "INV*" Then
                    tmp_name = mid(SplitByAt(j), 4) ' remove INV from the beginning
                    'tmp_name = SplitByAt(j)
                    DSP_Captured(j) = GetStoredCaptureData(tmp_name)
                    'DSP_Captured(j).CreateRandom 0, 1, 10, 1, DspLong '<- should be replaced by previous Line
                    'Set DSP_INV(i) = DSP_INV(i).Concatenate(DSP_Captured(j)) 'Merge all need flipped bit into one DSP
                    DSP_INVTemp(i).CreateConstant 0, DSP_Captured(j).SampleSize, DspLong
                    For k = 0 To DSP_Captured(j).SampleSize - 1
                        DSP_INVTemp(i).Element(k) = DSP_Captured(j).Element(UBound(DSP_Captured(j).data) - k)
                        DSP_Record(i * (UBound(SplitByAt) + 1) + j) = CStr(DSP_Record(i * (UBound(SplitByAt) + 1) + j)) & CStr(DSP_INVTemp(i).Element(k))
                    Next k
                    Set DSP_INV(i) = DSP_INV(i).Concatenate(DSP_INVTemp(i)) 'Merge all need flipped bit into one D
                Else
                    DSP_Captured(j) = GetStoredCaptureData(SplitByAt(j))
                    DSP_CKTemp(i).CreateConstant 0, DSP_Captured(j).SampleSize, DspLong
                    For k = 0 To DSP_Captured(j).SampleSize - 1
                        DSP_CKTemp(i).Element(k) = DSP_Captured(j).Element(UBound(DSP_Captured(j).data) - k)
                        DSP_Record(i * (UBound(SplitByAt) + 1) + j) = DSP_Record(i * (UBound(SplitByAt) + 1) + j) & CStr(DSP_Captured(j).Element(UBound(DSP_Captured(j).data) - k))
                    Next k
                    Set DSP_CK(i) = DSP_CK(i).Concatenate(DSP_CKTemp(i))
                End If
            

                If UCase(SplitByAt(j)) Like "*WCK*" Then
                    DSP_Record(i * (UBound(SplitByAt) + 1) + j) = "WCK:" & DSP_Record(i * (UBound(SplitByAt) + 1) + j)
                ElseIf UCase(SplitByAt(j)) Like "*CK*" Then
                    DSP_Record(i * (UBound(SplitByAt) + 1) + j) = "CK:" & DSP_Record(i * (UBound(SplitByAt) + 1) + j)
                Else
                    DSP_Record(i * (UBound(SplitByAt) + 1) + j) = "INV:" & DSP_Record(i * (UBound(SplitByAt) + 1) + j)
                End If
                'ReDim Preserve DSP_Record(UBound(DSP_Record) + 1)
            Next j
     
        '/////// INV_Cap reverse ////////////

            Dim iMidPt As Long
            Dim iUpper As Long
            iUpper = UBound(DSP_INV(i).data)
            iMidPt = (UBound(DSP_INV(i).data) - LBound(DSP_INV(i).data)) \ 2 + LBound(DSP_INV(i).data)
            For k = LBound(DSP_INV(i).data) To iMidPt
                tmp_element = DSP_INV(i).Element(iUpper)
                DSP_INV(i).Element(iUpper) = DSP_INV(i).Element(k)
                DSP_INV(i).Element(k) = tmp_element
                iUpper = iUpper - 1
            Next k
        '////////////////////////////////////

      ' next sweep register sw0, sw1, ...
       ' ====== Concat EYE data ============
            Set DSP_EYE(i) = DSP_CK(i).Concatenate(DSP_INV(i))  ' Concatenate UnFlip code + INV Flip code
            'Set DSP_EYE(i) = DSP_INV(i).Concatenate(DSP_CK(i))
      '==========================================================================
     
            'theexec.Datalog.WriteComment "EYE " & i
            For k = 0 To DSP_EYE(i).SampleSize - 1
                'Debug.Print DSP_EYE(i).Element(k);
                If k = 0 Then
                    Eye_str = DSP_EYE(i).Element(k)
                Else
                    Eye_str = Eye_str & DSP_EYE(i).Element(k)
                End If
            Next k
            Eye_str_result(i) = Eye_str

            'Debug.Print
            'theexec.Datalog.WriteComment Eye_str
            '====== Calculate number of 'ones' in the EYE ===========
            tmp_max_eye = 0 ' reset tmp_max_eye
            For k = 0 To DSP_EYE(i).SampleSize - 1
                If DSP_EYE(i).Element(k) = 1 Then
                    EYE_arr(i) = EYE_arr(i) + 1
                Else
                    If tmp_max_eye < EYE_arr(i) Then
                        tmp_max_eye = EYE_arr(i) ' update max eye width
                    End If
                    EYE_arr(i) = 0
                End If
                
                
                    If tmp_max_eye < EYE_arr(i) Then
                        tmp_max_eye = EYE_arr(i) ' update max eye width
                    End If
               
                
            Next k
            'Eye_str_long(i) = tmp_max_eye
            'Eye_str_long.CreateConstant 0, CLng(argc - 1)
            Eye_str_long(site).Element(i) = tmp_max_eye
        Next i ' next DDR bus : DQ0, DQ1, CA0, CA1 ...
    '=============================================================================
    Next site
    
    
    Dim DSP_Mdll_Temp As New DSPWave
    Dim DSP_Mdll_All() As New DSPWave
    Dim DSP_Mdll_Capture() As New DSPWave
    ReDim DSP_Mdll_All(argc - 1) As New DSPWave
    ReDim DSP_Mdll_Capture(argc - 1) As New DSPWave
    DSP_Mdll_Temp.CreateConstant 0, 1, DspLong
    For i = 0 To argc - 2
    '************Only for CACK read Mdll DSSCOUT and get Hi/Low limit****************
        DSP_Mdll_All(i).CreateConstant 0, 1, DspLong
        Mdll_ChannelInfo = Split(argv(UBound(argv)), "&")
        Mdll_value = Split(Mdll_ChannelInfo(i), "|")
        For z = 0 To UBound(Mdll_value)
            DSP_Mdll_Capture(i) = GetStoredCaptureData(Mdll_value(z))
'            rundsp.ConvertToLongAndSerialToParrel DSP_Mdll_Capture(i), DSP_Mdll_Capture(i).SampleSize, DSP_Mdll_Temp
            For Each site In TheExec.sites
                DSP_Mdll_Temp = DSP_Mdll_Capture(i).ConvertStreamTo(tldspParallel, DSP_Mdll_Capture(i).SampleSize, 0, Bit0IsMsb)
                DSP_Mdll_All(i).Element(0) = DSP_Mdll_All(i).Element(0) + DSP_Mdll_Temp.Element(0)
            Next site
        Next z
        For Each site In TheExec.sites
            TheExec.Datalog.WriteComment "Site: " & site & "   ,octants code sum : " & DSP_Mdll_All(i).Element(0) & ",for Argc Number " & i + 1
            mdll_low(i)(site) = DSP_Mdll_All(i).Element(0) / 2     'fix 20190715
            mdll_high(i)(site) = DSP_Mdll_All(i).Element(0) * 2    ' fix 20190601
        Next site
    '*******************************************************************************
    Next i
    
    
    Dim TnumRecord As Long
    
    For i = 0 To argc - 2
        'For Each Site In TheExec.sites
            SplitByAt = Split(argv(i), "@")
            
            TnumRecord = TheExec.sites.item(site).TestNumber
            
            TestNameInput = left(SplitByAt(0), InStr(1, SplitByAt(0), "_")) & "EYEDDR" & CStr(i)
            
            TestNameInput = Report_TName_From_Instance("X", "x", TestNameInput, CInt(TheExec.Flow.TestLimitIndex), 0)
            
        'Next Site
           For Each site In TheExec.sites
''''''''''            TheExec.Flow.TestLimit LowVal:=mdll_low(i), HiVal:=mdll_high(i), resultVal:=Eye_str_long.Element(i) * 8, FormatStr:="%i", TName:=TestNameInput, ForceResults:=tlForceFlow, TNum:=TnumRecord, ScaleType:=scaleNoScaling
                 TheExec.Flow.TestLimit lowVal:=mdll_low(i), hiVal:=mdll_high(i), resultVal:=Eye_str_long.Element(i) * 8, formatStr:="%i", Tname:=TestNameInput, ForceResults:=tlForceFlow, TNum:=TnumRecord, scaletype:=scaleNoScaling
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
            Next site
            
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
           ' TheExec.sites.item(Site).TestNumber = TheExec.sites.item(Site).TestNumber + 1
        'Next Site
    Next i
    
    For Each site In TheExec.sites
        TheExec.Datalog.WriteComment "/////////" & "Site: " & site & "/////////"
        Count = 0
        For L = 0 To argc - 2
            SplitByAt = Split(argv(L), "@")
            For i = 0 To UBound(SplitByAt)
                If SplitByAt(i) Like "INV*" Then
                    TheExec.Datalog.WriteComment mid(SplitByAt(i), 4)
                Else
                    TheExec.Datalog.WriteComment SplitByAt(i)
                End If
                TheExec.Datalog.WriteComment DSP_Record(Count)
                Count = Count + 1
            Next i
        '****************************************
        TheExec.Datalog.WriteComment "EYE " & L
        TheExec.Datalog.WriteComment Eye_str_result(L)
        TheExec.Datalog.WriteComment "EYE " & L & " width " & Eye_str_long(site).Element(L)
        Next L
    Next site
    TheExec.Datalog.WriteComment " ------------------ End ------------------"
    TheExec.Datalog.WriteComment "                           "

End Function


Public Function LP5_LB_RDDLL(argc As Integer, argv() As String) As Long

   'New LP5 eye model 20190417
   
    Dim i As Long, j As Long, k As Long, L As Long, z As Long
    Dim site As Variant
    Dim SplitByAt() As String
    Dim DictKey As String
    Dim DSP_Captured() As New DSPWave
    Dim DSP_EYE() As New DSPWave
    Dim tmp_element As Long
    Dim tmp_name As String
    Dim EYE_arr() As Long
    Dim DSP_INV() As New DSPWave
    Dim DSP_CK() As New DSPWave
    Dim DSP_CKTemp() As New DSPWave
    Dim DSP_INVTemp() As New DSPWave
    ReDim DSP_INV(CStr(argc) - 1)
    ReDim DSP_CK(CStr(argc) - 1)
    ReDim DSP_INVTemp(CStr(argc) - 1)
    ReDim DSP_CKTemp(CStr(argc) - 1)
    Dim tmp_max_eye As Long
    Dim Eye_str As String
    'Dim Eye_str_result() As String
    Dim Eye_str_long As New DSPWave
    Eye_str_long.CreateConstant 0, CLng(argc - 1)
    Dim Eye_str_result() As New SiteVariant
    ReDim Eye_str_result(CStr(argc))
    'ReDim Eye_str_long(CStr(argc)) As String
    Dim TestNameInput As String
    Dim OutputTname_formatQQ() As String
    Dim Mdll_value() As String
    Dim Mdll_ChannelInfo() As String
    
    Dim mdll_12x8 As New DSPWave
    Dim mdll As New SiteDouble
    Dim mdll_low() As New SiteDouble
    ReDim mdll_low(CLng(argc - 2))
    Dim mdll_high() As New SiteDouble
    ReDim mdll_high(CLng(argc - 2))
    Dim Mdll_width As Long
   
   
    Dim DSP_Record() As New SiteVariant
   
   
   'argv(0) = "WCK0Sweep_2@WCK0Sweep_3@INVDQ0Sweep_0@INVDQ0Sweep_1"
   'argv(1) = "WCK1Sweep_2@WCK1Sweep_3@INVDQ1Sweep_0@INVDQ1Sweep_1"
   'argv(2) = "ch0_mdll_w210|ch0_mdll_w543|ch0_mdll_w76|ch1_mdll_w210|ch1_mdll_w543|ch1_mdll_w76"    'for mdll high low clac
   '' Split DSPWave captured to number of components of sweep
   
    For Each site In TheExec.sites

        For i = 0 To argc - 2

            SplitByAt = Split(argv(i), "@") ' list of sweep names in order of concatination should be performed and INV if reverse is required
            ReDim Preserve DSP_Record((UBound(SplitByAt) + 1) * CStr(argc) - 2)

          ' Resize capture and final EYE DSPWaves to
            ReDim DSP_Captured(UBound(SplitByAt))
            ReDim DSP_EYE(UBound(SplitByAt))
            ReDim EYE_arr(UBound(SplitByAt))
            ReDim DSP_INV(UBound(SplitByAt))
            Set DSP_EYE(i) = DSP_EYE(i).ConvertDataTypeTo(DspLong)
            Set DSP_INV(i) = DSP_INV(i).ConvertDataTypeTo(DspLong)
            Set DSP_CK(i) = DSP_CK(i).ConvertDataTypeTo(DspLong)
     
      ' ============== Prepare data capture for calculation ==============
            For j = 0 To UBound(SplitByAt)
        ' ======= INV Data order MSB -> LSB require inversion =======
                If SplitByAt(j) Like "INV*" Then
                    tmp_name = mid(SplitByAt(j), 4) ' remove INV from the beginning
                    'tmp_name = SplitByAt(j)
                    DSP_Captured(j) = GetStoredCaptureData(tmp_name)
                    'DSP_Captured(j).CreateRandom 0, 1, 10, 1, DspLong '<- should be replaced by previous Line
                    'Set DSP_INV(i) = DSP_INV(i).Concatenate(DSP_Captured(j)) 'Merge all need flipped bit into one DSP
                    DSP_INVTemp(i).CreateConstant 0, DSP_Captured(j).SampleSize, DspLong
                    For k = 0 To DSP_Captured(j).SampleSize - 1
                        DSP_INVTemp(i).Element(k) = DSP_Captured(j).Element(UBound(DSP_Captured(j).data) - k)
                        DSP_Record(i * (UBound(SplitByAt) + 1) + j) = CStr(DSP_Record(i * (UBound(SplitByAt) + 1) + j)) & CStr(DSP_INVTemp(i).Element(k))
                    Next k
                    Set DSP_INV(i) = DSP_INV(i).Concatenate(DSP_INVTemp(i)) 'Merge all need flipped bit into one D
                Else
                    DSP_Captured(j) = GetStoredCaptureData(SplitByAt(j))
                    DSP_CKTemp(i).CreateConstant 0, DSP_Captured(j).SampleSize, DspLong
                    For k = 0 To DSP_Captured(j).SampleSize - 1
                        DSP_CKTemp(i).Element(k) = DSP_Captured(j).Element(UBound(DSP_Captured(j).data) - k)
                        DSP_Record(i * (UBound(SplitByAt) + 1) + j) = DSP_Record(i * (UBound(SplitByAt) + 1) + j) & CStr(DSP_Captured(j).Element(UBound(DSP_Captured(j).data) - k))
                    Next k
                    Set DSP_CK(i) = DSP_CK(i).Concatenate(DSP_CKTemp(i))
                End If
            

                If UCase(SplitByAt(j)) Like "*WCK*" Then
                    DSP_Record(i * (UBound(SplitByAt) + 1) + j) = "WCK:" & DSP_Record(i * (UBound(SplitByAt) + 1) + j)
                ElseIf UCase(SplitByAt(j)) Like "*CK*" Then
                    DSP_Record(i * (UBound(SplitByAt) + 1) + j) = "CK:" & DSP_Record(i * (UBound(SplitByAt) + 1) + j)
                ElseIf UCase(SplitByAt(j)) Like "*RDQS*" Then
                    DSP_Record(i * (UBound(SplitByAt) + 1) + j) = "RDQS:" & DSP_Record(i * (UBound(SplitByAt) + 1) + j)
                Else
                    DSP_Record(i * (UBound(SplitByAt) + 1) + j) = "INV:" & DSP_Record(i * (UBound(SplitByAt) + 1) + j)
                End If
                'ReDim Preserve DSP_Record(UBound(DSP_Record) + 1)
            Next j
     
        '/////// INV_Cap reverse ////////////

            Dim iMidPt As Long
            Dim iUpper As Long
            iUpper = UBound(DSP_INV(i).data)
            iMidPt = (UBound(DSP_INV(i).data) - LBound(DSP_INV(i).data)) \ 2 + LBound(DSP_INV(i).data)
            For k = LBound(DSP_INV(i).data) To iMidPt
                tmp_element = DSP_INV(i).Element(iUpper)
                DSP_INV(i).Element(iUpper) = DSP_INV(i).Element(k)
                DSP_INV(i).Element(k) = tmp_element
                iUpper = iUpper - 1
            Next k
        '////////////////////////////////////

      ' next sweep register sw0, sw1, ...
       ' ====== Concat EYE data ============
            Set DSP_EYE(i) = DSP_CK(i).Concatenate(DSP_INV(i))  ' Concatenate UnFlip code + INV Flip code
             'Set DSP_EYE(i) = DSP_INV(i).Concatenate(DSP_CK(i))
      '==========================================================================
     
            'theexec.Datalog.WriteComment "EYE " & i
            For k = 0 To DSP_EYE(i).SampleSize - 1
                'Debug.Print DSP_EYE(i).Element(k);
                If k = 0 Then
                    Eye_str = DSP_EYE(i).Element(k)
                Else
                    Eye_str = Eye_str & DSP_EYE(i).Element(k)
                End If
            Next k
            Eye_str_result(i) = Eye_str

            'Debug.Print
            'theexec.Datalog.WriteComment Eye_str
            '====== Calculate number of 'ones' in the EYE ===========
            tmp_max_eye = 0 ' reset tmp_max_eye
            For k = 0 To DSP_EYE(i).SampleSize - 1
                If DSP_EYE(i).Element(k) = 1 Then
                    EYE_arr(i) = EYE_arr(i) + 1
                Else
                    If tmp_max_eye < EYE_arr(i) Then
                        tmp_max_eye = EYE_arr(i) ' update max eye width
                    End If
                    EYE_arr(i) = 0
                End If
                
                
                    If tmp_max_eye < EYE_arr(i) Then
                        tmp_max_eye = EYE_arr(i) ' update max eye width
                    End If
              
            
            Next k
            'Eye_str_long(i) = tmp_max_eye
            'Eye_str_long.CreateConstant 0, CLng(argc - 1)
            Eye_str_long(site).Element(i) = tmp_max_eye
        Next i ' next DDR bus : DQ0, DQ1, CA0, CA1 ...
    '=============================================================================
    Next site
    
    
    Dim DSP_Mdll_Temp As New DSPWave
    Dim DSP_Mdll_All() As New DSPWave
    Dim DSP_Mdll_Capture() As New DSPWave
    ReDim DSP_Mdll_All(argc - 1) As New DSPWave
    ReDim DSP_Mdll_Capture(argc - 1) As New DSPWave
    DSP_Mdll_Temp.CreateConstant 0, 1, DspLong
    For i = 0 To argc - 2
    '************Only for CACK read Mdll DSSCOUT and get Hi/Low limit****************
        DSP_Mdll_All(i).CreateConstant 0, 1, DspLong
        Mdll_ChannelInfo = Split(argv(UBound(argv)), "&")
        Mdll_value = Split(Mdll_ChannelInfo(i), "|")
        For z = 0 To UBound(Mdll_value)
            DSP_Mdll_Capture(i) = GetStoredCaptureData(Mdll_value(z))
'            rundsp.ConvertToLongAndSerialToParrel DSP_Mdll_Capture(i), DSP_Mdll_Capture(i).SampleSize, DSP_Mdll_Temp
            For Each site In TheExec.sites
                DSP_Mdll_Temp = DSP_Mdll_Capture(i).ConvertStreamTo(tldspParallel, DSP_Mdll_Capture(i).SampleSize, 0, Bit0IsMsb)
                DSP_Mdll_All(i).Element(0) = DSP_Mdll_All(i).Element(0) + DSP_Mdll_Temp.Element(0)
            Next site
        Next z
        For Each site In TheExec.sites
        TheExec.Datalog.WriteComment "Site: " & site & "   ,octants code sum : " & DSP_Mdll_All(i).Element(0) & ",for Argc Number " & i + 1
            mdll_low(i)(site) = DSP_Mdll_All(i).Element(0) / 8
            mdll_high(i)(site) = DSP_Mdll_All(i).Element(0)
        Next site
    '*******************************************************************************
    Next i
    
    
    Dim TnumRecord As Long
    
    For i = 0 To argc - 2
        
            SplitByAt = Split(argv(i), "@")
            
            TnumRecord = TheExec.sites.item(site).TestNumber
            
            'TestNameInput = Left(SplitByAt(0), InStr(1, SplitByAt(0), "_")) & "EYEDDR" & CStr(i)
             TestNameInput = "EYE" & left(SplitByAt(0), 7)
            
            TestNameInput = Report_TName_From_Instance("X", "x", TestNameInput, CInt(TheExec.Flow.TestLimitIndex), 0)
       
            For Each site In TheExec.sites
                'TheExec.Flow.TestLimit LowVal:=mdll_low(i)(Site), HiVal:=mdll_high(i)(Site), resultVal:=Eye_str_long(Site).Element(i) * 4, FormatStr:="%i", TName:=TestNameInput, ForceResults:=tlForceFlow, TNum:=TnumRecord, ScaleType:=scaleNoScaling
                TheExec.Flow.TestLimit lowVal:=mdll_low(i), hiVal:=mdll_high(i), resultVal:=Eye_str_long.Element(i), formatStr:="%i", Tname:=TestNameInput, ForceResults:=tlForceFlow, TNum:=TnumRecord, scaletype:=scaleNoScaling
        '           TheExec.Flow.TestLimit lowVal:=mdll_low(i)(Site), resultVal:=Eye_str_long(Site).Element(i) * 4, FormatStr:="%i", TName:=TestNameInput, ForceResults:=tlForceFlow, TNum:=TnumRecord
                TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
            Next site
                TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
            'TheExec.sites.item(Site).TestNumber = TheExec.sites.item(Site).TestNumber + 1
    Next i
    
    For Each site In TheExec.sites
        TheExec.Datalog.WriteComment "/////////" & "Site: " & site & "/////////"
        Count = 0
        For L = 0 To argc - 2
            SplitByAt = Split(argv(L), "@")
            For i = 0 To UBound(SplitByAt)
                If SplitByAt(i) Like "INV*" Then
                    TheExec.Datalog.WriteComment mid(SplitByAt(i), 4)
                Else
                    TheExec.Datalog.WriteComment SplitByAt(i)
                End If
                TheExec.Datalog.WriteComment DSP_Record(Count)
                Count = Count + 1
            Next i
        '****************************************
        TheExec.Datalog.WriteComment "EYE " & L
        TheExec.Datalog.WriteComment Eye_str_result(L)
        TheExec.Datalog.WriteComment "EYE " & L & " width " & Eye_str_long(site).Element(L)
        Next L
    Next site
    TheExec.Datalog.WriteComment " ------------------ End ------------------"
    TheExec.Datalog.WriteComment "                           "

End Function
Public Function Calc_MTR_BinStr2HexStr(ByVal binstr As String, ByVal HexBit As Long) As String

    Dim i As Integer, j As Integer
    Dim BinStrLen As Long
    Dim HexMOD As Integer
    Dim HexStr As String
    Dim HexVal As String
    Dim HexLen As Long

    HexStr = vbNullString
    
    BinStrLen = Len(binstr)
    If (BinStrLen Mod (4)) > 0 Then
        HexLen = (BinStrLen \ 4) + 1
    Else
        HexLen = BinStrLen \ 4
    End If
    
    If HexBit > HexLen Then
        HexLen = HexBit
    End If

    HexMOD = HexLen * 4 - BinStrLen
    
    If HexMOD > 0 Then
        For i = 0 To HexMOD - 1
            binstr = "0" & binstr
        Next i
    End If

    For i = 0 To HexLen - 1
        If mid(binstr, i * 4 + 1, 4) = "0000" Then
            HexVal = "0"
        ElseIf mid(binstr, i * 4 + 1, 4) = "0001" Then
            HexVal = "1"
        ElseIf mid(binstr, i * 4 + 1, 4) = "0010" Then
            HexVal = "2"
        ElseIf mid(binstr, i * 4 + 1, 4) = "0011" Then
            HexVal = "3"
        ElseIf mid(binstr, i * 4 + 1, 4) = "0100" Then
            HexVal = "4"
        ElseIf mid(binstr, i * 4 + 1, 4) = "0101" Then
            HexVal = "5"
        ElseIf mid(binstr, i * 4 + 1, 4) = "0110" Then
            HexVal = "6"
        ElseIf mid(binstr, i * 4 + 1, 4) = "0111" Then
            HexVal = "7"
        ElseIf mid(binstr, i * 4 + 1, 4) = "1000" Then
            HexVal = "8"
        ElseIf mid(binstr, i * 4 + 1, 4) = "1001" Then
            HexVal = "9"
        ElseIf mid(binstr, i * 4 + 1, 4) = "1010" Then
            HexVal = "A"
        ElseIf mid(binstr, i * 4 + 1, 4) = "1011" Then
            HexVal = "B"
        ElseIf mid(binstr, i * 4 + 1, 4) = "1100" Then
            HexVal = "C"
        ElseIf mid(binstr, i * 4 + 1, 4) = "1101" Then
            HexVal = "D"
        ElseIf mid(binstr, i * 4 + 1, 4) = "1110" Then
            HexVal = "E"
        ElseIf mid(binstr, i * 4 + 1, 4) = "1111" Then
            HexVal = "F"
        Else
            HexVal = "X"
        End If

        HexStr = HexStr & HexVal
    Next i

    Calc_MTR_BinStr2HexStr = HexStr

End Function

Public Function Calc_Voff_t6p2_MetrologyGR(argc As Integer, argv() As String) As Long
    Dim Dict_V0 As String
    Dim Dict_V1 As String
    Dim Dict_V2 As String
    Dim Fuse_BitCount As Double
    Dim Fuse_Voff_Round As String
    Dim Dict_Ratio_off_Per As String
    
    Dim Input_V0 As New PinListData
    Dim Input_V1 As New PinListData
    Dim Input_V2 As New PinListData
    
    Dim Voff_PinListData As New PinListData
    Dim Voff_PinListData_Round As New PinListData
    
    Dim UnSinged_Voff_Round As New DSPWave
    UnSinged_Voff_Round.CreateConstant 0, 1, DspDouble
    
    Dim Ratio_off_PinListData As New PinListData
    Dim Ratio_off_Per_DSP As New DSPWave
    Ratio_off_Per_DSP.CreateConstant 0, 1, DspDouble
    Dim site As Variant
    
    Dict_V0 = argv(0)
    Dict_V1 = argv(1)
    Dict_V2 = argv(2)
    Fuse_BitCount = argv(3)
    Fuse_Voff_Round = argv(4)
    Dict_Ratio_off_Per = argv(5)
    
    Input_V0 = GetStoredMeasurement(Dict_V0)
    Input_V1 = GetStoredMeasurement(Dict_V1)
    Input_V2 = GetStoredMeasurement(Dict_V2)
    
    Voff_PinListData.AddPin (Input_V1.Pins(0))
    Voff_PinListData = Input_V1.Pins(0).Subtract(Input_V0.Pins(0))
    Voff_PinListData = Voff_PinListData.Math.divide(0.001)
    
    Voff_PinListData_Round.AddPin (Input_V1.Pins(0))
    Voff_PinListData_Round = Voff_PinListData.Pins(0).divide(0.5)
    For Each site In TheExec.sites
        Voff_PinListData_Round.Pins(0).value(site) = CDbl(FormatNumber(Voff_PinListData_Round.Pins(0).value(site), 0))
    Next site
    
    Ratio_off_PinListData.AddPin (Input_V1.Pins(0))
    Ratio_off_PinListData = Input_V1.Pins(0).divide(Input_V2.Pins(0)).divide(2).Subtract(1)
    For Each site In TheExec.sites
            Ratio_off_Per_DSP(site).Element(0) = Ratio_off_PinListData.Pins(0).value(site)
    Next site
    
    Call AddStoredCaptureData(Dict_Ratio_off_Per, Ratio_off_Per_DSP)
                                       
    TheExec.Flow.TestLimit resultVal:=Voff_PinListData.Pins(0), ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit resultVal:=Voff_PinListData_Round.Pins(0), ForceResults:=tlForceFlow
    For Each site In TheExec.sites
        If Voff_PinListData_Round.Pins(0).value(site) < 0 Then
            UnSinged_Voff_Round(site).Element(0) = Voff_PinListData_Round.Pins(0).value(site) + (2 ^ Fuse_BitCount)
        Else
            UnSinged_Voff_Round(site).Element(0) = Voff_PinListData_Round.Pins(0).value(site)
        End If
    Next site
    
    Call AddStoredCaptureData(Fuse_Voff_Round, UnSinged_Voff_Round)
    
    TheExec.Flow.TestLimit resultVal:=Ratio_off_PinListData.Pins(0), ForceResults:=tlForceFlow
    
End Function

Public Function Calc_Ratio_off_average_t6p2_MetrologyGR(argc As Integer, argv() As String) As Long
    Dim i As Long
    Dim DSPWave_Ratio_off_per() As New DSPWave
    ReDim DSPWave_Ratio_off_per(argc - 3) As New DSPWave
    Dim DSPWave_Average As New DSPWave
    DSPWave_Average.CreateConstant 0, 1
    Dim DSPWave_Round_Average As New DSPWave
    DSPWave_Round_Average.CreateConstant 0, 1
    Dim DSPWave_Unsinged_Round_Average As New DSPWave
    DSPWave_Unsinged_Round_Average.CreateConstant 0, 1
    Dim site As Variant
    Dim Sweep_Count As Double
    Dim Fuse_BitCount As Double
    Sweep_Count = argc - 2
    Fuse_BitCount = argv(argc - 1)
    Dim Dict_Unsinged_Round_Avg As String
    Dict_Unsinged_Round_Avg = argv(argc - 2)
    
    For i = 0 To argc - 3
        DSPWave_Ratio_off_per(i) = GetStoredCaptureData(argv(i))
        Call rundsp.DSP_Add(DSPWave_Average, DSPWave_Ratio_off_per(i))
    Next i
    Call rundsp.DSP_DivideConstant(DSPWave_Average, Sweep_Count)
    
    If TheExec.TesterMode = testModeOffline Then            'for offline run
        For Each site In TheExec.sites
            DSPWave_Average(site).Element(0) = -0.00060171
        Next site
    End If
    
    TheExec.Flow.TestLimit resultVal:=DSPWave_Average.Element(0), Tname:="Ratio_off_per_avg", ForceResults:=tlForceNone 'eng_forceflow_transfer
    
    For Each site In TheExec.sites
        DSPWave_Round_Average(site).Element(0) = FormatNumber(DSPWave_Average(site).Element(0) * 1600, 0)
    Next site
    
    TheExec.Flow.TestLimit resultVal:=DSPWave_Round_Average.Element(0), Tname:="Round_Ratio_off_per_avg", ForceResults:=tlForceNone 'eng_forceflow_transfer
    
    For Each site In TheExec.sites
        If DSPWave_Round_Average(site).Element(0) < 0 Then
            DSPWave_Unsinged_Round_Average(site).Element(0) = DSPWave_Round_Average(site).Element(0) + (2 ^ Fuse_BitCount)
        Else
            DSPWave_Unsinged_Round_Average(site).Element(0) = DSPWave_Round_Average(site).Element(0)
        End If
    Next site
    
    Call AddStoredCaptureData(Dict_Unsinged_Round_Avg, DSPWave_Unsinged_Round_Average)
    
End Function

Public Function Calc_2S_Complement_To_SignDec_DivConst(argc As Integer, argv() As String) As Long

    Dim i As Long, j As Long
    Dim site As Variant
    Dim SplitByAt() As String
    Dim DictKey_2S_BIN As String
    Dim DictKey_SIGN_DEC As String
    
    Dim DSP_DictKey_2S_BIN As New DSPWave
    Dim DSP_DictKey_SIGN_DEC() As New DSPWave

    ReDim DSP_DictKey_SIGN_DEC(argc - 1) As New DSPWave
    
    Dim TestName As String
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    
    Dim DivConst As Double
    
    Dim SL_BitWidth As New SiteLong
    '' Format: Dict_2S_Com_A@Dict_SignDec_A@TestName_A,Dict_2S_Com_B@Dict_SignDec_B@TestName_B
    For i = 0 To argc - 1
        SplitByAt = Split(argv(i), "@")
        DictKey_2S_BIN = SplitByAt(0)
        DictKey_SIGN_DEC = SplitByAt(1)
        TestName = SplitByAt(2)
        DivConst = SplitByAt(3)
        
        DSP_DictKey_2S_BIN = GetStoredCaptureData(DictKey_2S_BIN)
        
''        Set DSP_DictKey_DEC = Nothing
''        DSP_DictKey_DEC.CreateConstant 0, 1, DspDouble
''        Call rundsp.BinToDec(DSP_DictKey_BIN, DSP_DictKey_DEC)
        
        For Each site In TheExec.sites
            SL_BitWidth(site) = DSP_DictKey_2S_BIN(site).SampleSize
''            DSP_DictKey_DEC(0).Element(0) = 255
''            DSP_DictKey_DEC(1).Element(0) = 254
        Next site
        
        Set DSP_DictKey_SIGN_DEC(i) = Nothing
        'DSP_DictKey_SIGN_DEC(i).CreateConstant 0, 1, DspLong
        DSP_DictKey_SIGN_DEC(i).CreateConstant 0, 1
        
        Call rundsp.DSP_2S_Complement_To_SignDec(DSP_DictKey_2S_BIN, SL_BitWidth, DSP_DictKey_SIGN_DEC(i))
        
        Call AddStoredCaptureData(DictKey_SIGN_DEC, DSP_DictKey_SIGN_DEC(i))

        TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(i))
        
        Call rundsp.DSP_DivideConstant(DSP_DictKey_SIGN_DEC(i), DivConst)
        
        TheExec.Flow.TestLimit resultVal:=DSP_DictKey_SIGN_DEC(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
        
    Next i
End Function

Public Function Calc_Metrology_Trim_Vdiff(argc As Integer, argv() As String) As Long

    Dim Dict_V1 As String
    Dim Dict_V2 As String
    Dim Dict_Vdiff As String
    
    Dim PinList_V1 As New PinListData
    Dim PinList_V2 As New PinListData
    Dim PinList_Vdiff As New PinListData
    
    Dict_V1 = argv(0)
    Dict_V2 = argv(1)
    Dict_Vdiff = argv(2)
    
    PinList_V1 = GetStoredMeasurement(Dict_V1)
    PinList_V2 = GetStoredMeasurement(Dict_V2)
    
    PinList_Vdiff.AddPin (PinList_V1.Pins(0))
    PinList_Vdiff.Pins(0) = PinList_V1.Math.Subtract(PinList_V2)
    
    Call AddStoredMeasurement(Dict_Vdiff, PinList_Vdiff)
    
    TheExec.Datalog.WriteComment ("Voltage Difference Calculation")
    
End Function



Public Function Calc_MetrologyGR_t5p5(argc As Integer, argv() As String) As Long
    Dim Vsrp As New SiteDouble
    Dim Vsrn As New SiteDouble
    Dim Vdiff As New SiteDouble
    Dim TestNameInput As String
    Dim site As Variant
    Dim OutputTname_format() As String
    
    'GetStoredMeasurement
    For Each site In TheExec.sites
        Vsrp = GetStoredMeasurement(argv(0))
        Vsrn = GetStoredMeasurement(argv(1))
        Vdiff = Vsrn.Subtract(Vsrp).divide(0.000005)
    Next site
    
    TestNameInput = Report_TName_From_Instance(CalcV, "X", "Rpsr")
    OutputTname_format = Split(TestNameInput, "_")
    OutputTname_format(6) = "Rpsr"
    TestNameInput = Merge_TName(OutputTname_format)
    TheExec.Flow.TestLimit resultVal:=Vdiff, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
End Function

Public Function Calc_MetrologyGR_t1p0(argc As Integer, argv() As String) As Long
    Dim Vsrp As New SiteDouble
    Dim Vsrn As New SiteDouble
    Dim Vdiff As New SiteDouble
    Dim TestNameInput As String
    Dim site As Variant
    Dim OutputTname_format() As String
    
    'GetStoredMeasurement
    For Each site In TheExec.sites
        Vsrp = GetStoredMeasurement(argv(0))
        Vsrn = GetStoredMeasurement(argv(1))
        Vdiff = Vsrp.Subtract(Vsrn)
    Next site
    
    TestNameInput = Report_TName_From_Instance(CalcV, "X", "Rpsr")
    OutputTname_format = Split(TestNameInput, "_")
    OutputTname_format(6) = "Vdiff"
    OutputTname_format(7) = CStr(TheExec.Flow.var("SrcCodeIndx").value)
    TestNameInput = Merge_TName(OutputTname_format)
    TheExec.Flow.TestLimit resultVal:=Vdiff, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
End Function



Public Function Calc_PCIE_RXTERM(argc As Integer, argv() As String) As Long
    Dim DSP_RCAL_TX_DIV4_CODE As New DSPWave: DSP_RCAL_TX_DIV4_CODE = GetStoredCaptureData(argv(0))
    Dim DSP_RXTERM_CODE_Binary As New DSPWave
    Dim DSP_RXTERM_CODE_Dec As New DSPWave
    Dim TestNameInput As String
    For Each site In TheExec.sites.Active
        DSP_RXTERM_CODE_Binary = DSP_RCAL_TX_DIV4_CODE.Select(0, 1, DSP_RCAL_TX_DIV4_CODE.SampleSize - 1).Copy
        DSP_RXTERM_CODE_Dec = DSP_RXTERM_CODE_Binary.ConvertStreamTo(tldspParallel, DSP_RXTERM_CODE_Binary.SampleSize, 0, Bit0IsMsb)
    Next site
    TestNameInput = Report_TName_From_Instance(CalcC, vbNullString)
    TheExec.Flow.TestLimit resultVal:=DSP_RXTERM_CODE_Dec.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Call AddStoredCaptureData(argv(1), DSP_RXTERM_CODE_Binary)
End Function

Public Function Auth_Debug(argc As Integer, argv() As String) As Long

    Dim noncedsp As New DSPWave
    Dim eciddsp As New DSPWave
    Dim chipiddsp As New DSPWave
    Dim briddsp As New DSPWave
    Dim prod_moddsp As New DSPWave
    Dim secure_moddsp As New DSPWave
    Dim secure_domaindsp As New DSPWave
    
    
    Dim site_number As Long
    site_number = TheExec.sites.Existing.Count
    
    Dim nonce() As String
    ReDim nonce(site_number) As String
    Dim ecid() As String
    ReDim ecid(site_number) As String
    Dim chipid() As String
    ReDim chipid(site_number) As String
    Dim brid() As String
    ReDim brid(site_number) As String
    Dim prod_mod() As String
    ReDim prod_mod(site_number) As String
    Dim secure_mod() As String
    ReDim secure_mod(site_number) As String
    Dim secure_domain() As String
    ReDim secure_domain(site_number) As String
    
    
    Dim OutDspWave As New DSPWave
    
    Dim nonce_64() As String
    ReDim nonce_64(site_number) As String
    Dim ecid_dec() As String
    ReDim ecid_dec(site_number) As String
    Dim chipid_dec() As String
    ReDim chipid_dec(site_number) As String
    Dim brid_int() As String
    ReDim brid_int(site_number) As String
    Dim prod_mod_boo() As String
    ReDim prod_mod_boo(site_number) As String
    Dim secure_mod_boo() As String
    ReDim secure_mod_boo(site_number) As String
    Dim secure_domain_int() As String
    ReDim secure_domain_int(site_number) As String
    
    Dim sitev As Variant
    Dim iv As Long
    Dim i As Long
    
    Dim fetch_data() As String
    ReDim fetch_data(site_number) As String
    
    Dim get_store_name As String
    Dim source_dsp As New DSPWave
    
    source_dsp.CreateConstant 0, argv(1)
    
    
    get_store_name = argv(0)
    
    OutDspWave = GetStoredCaptureData(get_store_name)
    
    
    For Each sitev In TheExec.sites
    
        ecid(sitev) = vbNullString
    
        noncedsp.ConvertDataTypeTo (DspLong)
        eciddsp.ConvertDataTypeTo (DspLong)
        chipiddsp.ConvertDataTypeTo (DspLong)
        noncedsp = OutDspWave.Select(32, 1, 128).ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspParallel, 4, 0, Bit0IsMsb)
        eciddsp = OutDspWave.Select(320, 1, 64)
        chipiddsp = OutDspWave.Select(224, 1, 32)
        briddsp = OutDspWave.Select(160, 1, 32)
        prod_moddsp = OutDspWave.Select(256, 1, 32)
        secure_moddsp = OutDspWave.Select(288, 1, 32)
        secure_domaindsp = OutDspWave.Select(384, 1, 32)
'                eciddsp = OutDspWave.Select(320, 1, 64).ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspParallel, 4, 0, Bit0IsMsb)
'                chipiddsp = OutDspWave.Select(224, 1, 32).ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspParallel, 4, 0, Bit0IsMsb)
        
        For i = 0 To noncedsp.SampleSize - 1
            'nonce = nonce & CStr(Hex(noncedsp.Element(i)))
            nonce(sitev) = CStr(Hex(noncedsp.Element(i))) & nonce(sitev)
        Next
        For i = 0 To eciddsp.SampleSize - 1
            ecid(sitev) = ecid(sitev) & eciddsp.Element(i)
            'Ecid = CStr(Hex(eciddsp.Element(i))) & Ecid
        Next
        For i = 0 To chipiddsp.SampleSize - 1
            chipid(sitev) = chipid(sitev) & chipiddsp.Element(i)
            'chipid = CStr(Hex(chipiddsp.Element(i))) & chipid
        Next
        For i = 0 To briddsp.SampleSize - 1
            brid(sitev) = brid(sitev) & briddsp.Element(i)
            'chipid = CStr(Hex(chipiddsp.Element(i))) & chipid
        Next
        'theexec.Datalog.WriteComment "Site" + CStr(sitev) + " Ecid = " + Ecid
         For i = 0 To prod_moddsp.SampleSize - 1
            prod_mod(sitev) = prod_mod(sitev) & prod_moddsp.Element(i)
            'chipid = CStr(Hex(chipiddsp.Element(i))) & chipid
        Next
         For i = 0 To secure_moddsp.SampleSize - 1
            secure_mod(sitev) = secure_mod(sitev) & secure_moddsp.Element(i)
            'chipid = CStr(Hex(chipiddsp.Element(i))) & chipid
        Next
         For i = 0 To secure_domaindsp.SampleSize - 1
            secure_domain(sitev) = secure_domain(sitev) & secure_domaindsp.Element(i)
            'chipid = CStr(Hex(chipiddsp.Element(i))) & chipid
        Next
    
    
        TheExec.Datalog.WriteComment "Site = " & CStr(sitev) & ",  Nonce Hex = " & nonce(sitev)
        
        
        nonce_64(sitev) = Hex2Base64(nonce(sitev))
        ecid_dec(sitev) = binToDecStr(ecid(sitev))
        chipid_dec(sitev) = binToDecStr(chipid(sitev))
        brid_int(sitev) = binToDecStr(brid(sitev))
        prod_mod_boo(sitev) = binToDecStr(prod_mod(sitev))
        If prod_mod_boo(sitev) = "0" Then prod_mod_boo(sitev) = "false"
        If prod_mod_boo(sitev) = "1" Then prod_mod_boo(sitev) = "true"
        secure_mod_boo(sitev) = binToDecStr(secure_mod(sitev))
        If secure_mod_boo(sitev) = "0" Then secure_mod_boo(sitev) = "false"
        If secure_mod_boo(sitev) = "1" Then secure_mod_boo(sitev) = "true"
        secure_domain_int(sitev) = binToDecStr(secure_domain(sitev))
        'fetch_data(sitev) = Server_Connection(nonce_64(sitev), ecid_dec(sitev), chipid_dec(sitev), brid_int(sitev), prod_mod_boo(sitev), secure_mod_boo(sitev), secure_domain_int(sitev))
        fetch_data(sitev) = Server_Connection(nonce_64(sitev), ecid_dec(sitev), chipid_dec(sitev), brid_int(sitev), prod_mod_boo(sitev), secure_mod_boo(sitev), secure_domain_int(sitev), sitev)
  
        
        TheExec.Datalog.WriteComment "Site = " & CStr(sitev) & ", Ecid = " & ecid_dec(sitev)
        TheExec.Datalog.WriteComment "Site = " & CStr(sitev) & ",  Data from server (Hex)= " & fetch_data(sitev)

   
            Dim i2 As Long
            Dim j2 As Long
            Dim binstr2() As String
            ReDim binstr2(site_number) As String
            Dim binstr3() As String
            ReDim binstr3(site_number) As String
            Dim counterbit As Integer: counterbit = 1
            Dim xortemp As Long
            
            For i2 = 1 To Len(fetch_data(sitev)) - 1 Step 2
                binstr2(sitev) = binstr2(sitev) & StrReverse(auto_Hex2BinStr(mid(fetch_data(sitev), i2 + 1, 1)))
                binstr2(sitev) = binstr2(sitev) & StrReverse(auto_Hex2BinStr(mid(fetch_data(sitev), i2, 1)))
            Next
            For i2 = (8192 - Len(binstr2(sitev)) - 1) To 0 Step -1
                binstr2(sitev) = binstr2(sitev) + "0"
            Next
            For i2 = 1 To Len(binstr2(sitev))
                binstr3(sitev) = binstr3(sitev) & mid(binstr2(sitev), i2, 1)
                If i2 Mod 32 = 1 Then
                    xortemp = CLng(mid(binstr2(sitev), i2, 1))
                Else
                    xortemp = xortemp Xor CLng(mid(binstr2(sitev), i2, 1))
                End If
                If i2 Mod 32 = 0 Then binstr3(sitev) = binstr3(sitev) & CStr(xortemp)
            Next
            
            
            For i2 = 0 To Len(binstr3(sitev)) - 1
                source_dsp.Element(i2) = CLng(mid(binstr3(sitev), i2 + 1, 1))
            Next
            
            'Call AddStoredCaptureData(argv(2), source_dsp)
'            Dim DigSrc_Equation As String
'            Dim DigSrc_Assignment As String
'            DigSrc_Equation = ""
'            DigSrc_Assignment = ""
'
'            For i2 = 0 To 255
'                DigSrc_Equation = DigSrc_Equation + "input_" + CStr(i2) + "+"
'                DigSrc_Assignment = DigSrc_Assignment + "input_" + CStr(i2) + "=" + Mid(binstr3(sitev), (i2) * 33 + 1, 33) + ";"
'            Next
'            DigSrc_Equation = Left(DigSrc_Equation, Len(DigSrc_Equation) - 1)
'            DigSrc_Assignment = Left(DigSrc_Assignment, Len(DigSrc_Assignment) - 1)
 
     Next
     Call AddStoredCaptureData(argv(2), source_dsp)

End Function

Public Function Calc_DigCap_MeanWithVariance(argc As Integer, argv() As String) As Long

'Arguments: SegmentSize_8, DictKey1, DictKey2, DictKey3.....


    Dim i As Long, j As Long
    Dim site As Variant
    Dim SegmentSize As Long
    Dim SegmentCount As Long
    Dim IndexOffset As Long
    Dim GroupCount As Long
    Dim str_temp As String
    
    Dim mean As Double
    Dim STDEV As Double
    Dim Variance As Double
    
    Dim TestNameInput As String
    
'    Dim Output_Mean() As New DSPWave
'    Dim Output_STDEV() As New DSPWave
'    Dim Output_Variance() As New DSPWave
    Dim CalcResult() As New PinListData
    
'    Dim Mean() As New SiteDouble
'    Dim STDEV() As New SiteDouble
'    Dim Variance() As New SiteDouble
    
    Dim DSPwave_temp As New DSPWave
    Dim DSPWave_UnitSegment() As New DSPWave
    Dim DSPWave_MergedSegment() As New DSPWave

    
    SegmentSize = CLng(Split(argv(0), "_")(1))
    GroupCount = argc - 1 ''Argv(0) define segment size, the others are group1, group2 ....
    ReDim DSPWave_UnitSegment(GroupCount - 1)
    ReDim DSPWave_MergedSegment(GroupCount - 1)
    
    ReDim CalcResult(GroupCount - 1)
    Dim MSB_First_Flag As Boolean
    
'    ReDim Output_Mean(GroupCount - 1)
'    ReDim Output_STDEV(GroupCount - 1)
'    ReDim Output_Variance(GroupCount - 1)
    
    If UBound(Split(argv(0), "_")) = 2 Then
        If Split(argv(0), "_")(2) = "MSB" Then MSB_First_Flag = True
    End If
    
    
    For Each site In TheExec.sites.Active
        For i = 0 To GroupCount - 1
            DSPWave_UnitSegment(i) = GetStoredCaptureData(argv(i + 1))
            CalcResult(i).AddPin "Mean"
            CalcResult(i).AddPin "STDEV"
            CalcResult(i).AddPin "Variance"
'            Output_Mean(i).CreateConstant 0, 1, DspDouble
'            Output_STDEV(i).CreateConstant 0, 1, DspDouble
'            Output_Variance(i).CreateConstant 0, 1, DspDouble
        Next i
        
        SegmentCount = CLng(DSPWave_UnitSegment(0).SampleSize) \ SegmentSize
        
    
        For i = 0 To GroupCount - 1 '' For example, SegmentSize = 8; Binary 8 Bit -> Dec 1 number (DSPWave_MergedSegment)
            If MSB_First_Flag Then
                DSPWave_MergedSegment(i) = DSPWave_UnitSegment(i).ConvertStreamTo(tldspParallel, SegmentSize, 0, Bit0IsLsb)
            Else
                DSPWave_MergedSegment(i) = DSPWave_UnitSegment(i).ConvertStreamTo(tldspParallel, SegmentSize, 0, Bit0IsMsb)
            End If
            mean = DSPWave_MergedSegment(i).CalcMeanWithStdDev(STDEV)
            CalcResult(i).Pins("Mean").value(site) = mean
            CalcResult(i).Pins("STDEV").value(site) = STDEV
            CalcResult(i).Pins("Variance").value(site) = STDEV * STDEV
        Next i
        
    Next site
   
    
    TheExec.Datalog.WriteComment vbNullString
    
    For i = 0 To GroupCount - 1
        'TestNameInput = Report_TName_From_Instance("C", "X", , CInt(i))
        
        For Each site In TheExec.sites.Active
            For j = 0 To SegmentCount - 1
                TheExec.Flow.TestLimit resultVal:=DSPWave_MergedSegment(i).data(j), _
                Tname:=Report_TName_From_Instance("C", "X", argv(i + 1), Instance_Data.TestSeqNum + CInt(i * (SegmentCount + 2) + j)), ForceResults:=tlForceNone 'eng_forceflow_transfer
            Next j
                TheExec.Flow.TestLimit resultVal:=CalcResult(i).Pins("Mean").value(site), _
                Tname:=Report_TName_From_Instance("C", "X", argv(i + 1) & "Mean", Instance_Data.TestSeqNum + CInt(i * (SegmentCount + 2) + SegmentCount)), ForceResults:=tlForceNone 'eng_forceflow_transfer
                TheExec.Flow.TestLimit resultVal:=CalcResult(i).Pins("Variance").value(site), _
                Tname:=Report_TName_From_Instance("C", "X", argv(i + 1) & "Variance", Instance_Data.TestSeqNum + CInt(i * (SegmentCount + 2) + SegmentCount + 1)), ForceResults:=tlForceNone 'eng_forceflow_transfer
        Next site
        
    Next i
    
    'ReDim DSPWave_Avg_Bin(argc - 3) As New DSPWave
    
'    Dim TestName As String
'    Dim Site As Variant
'    Dim Dict As String
'    Dim BitWidth As Long
'
'    For i = 0 To 1
'        DSPWave_Binary(i) = GetStoredCaptureData(argv(i))
'        Call rundsp.BinToDec(DSPWave_Binary(i), DSPWave_Dec(i))
'    Next i
'
'    TestName = argv(argc - 1)
'    BitWidth = argv(argc - 2)
'    Dict = argv(argc - 3)
'
'    For Each Site In TheExec.sites
'            DSPWave_Avg_Dec.Element(0) = Int(((DSPWave_Dec(0).Element(0) + DSPWave_Dec(1).Element(0)) / 2) + 0.5) ''Example 1). 78.4=>78  2). 78.5=79
'    Next Site
'    Call rundsp.DSPWaveDecToBinary(DSPWave_Avg_Dec, BitWidth, DSPWave_Avg_Bin)
'    Call AddStoredCaptureData(Dict, DSPWave_Avg_Bin)
'    Dim TestNameInput As String
'    Dim OutputTname_format() As String
'
'    TestNameInput = Report_TName_From_Instance("C", "X", , CInt(i))
'
'    TheExec.Flow.TestLimit resultVal:=DSPWave_Avg_Dec.Element(0), TName:=TestNameInput, ForceResults:=tlForceFlow
    
End Function

Public Function Trim_Pll_Freq(argc As Integer, argv() As String) As Long


    Dim UseLimitTname As String
    Dim TestNameInput As String
    Dim SplitTrimFreq() As String
    Dim SplitVro() As String
    Dim SplitDCO() As String
    Dim SplitCap() As String
    Dim SplitBias() As String
    Dim i, j, k As Long
    Dim F_TrimComplete() As New SiteBoolean
    Dim F_Vro() As New SiteBoolean
    Dim DSPWaveFromDict As New DSPWave
    Dim DSPWaveDecType As New DSPWave
    Dim BiasTargetIndex() As New SiteLong
    Dim FinalBiasIndex As New SiteLong
    Dim VroTarget As New SiteDouble
    Dim FinalVroIndex As New SiteLong
    Dim VroFromDict As New SiteLong: VroFromDict = 0
    Dim VroVoltage() As New SiteDouble
    Dim CapDSPWave As New DSPWave
    Dim BiasDSPWave As New DSPWave
    Dim DCODSPWave As New DSPWave
    Dim FinalDSPWave As New DSPWave
    Dim Cap_arry() As Long
    Dim Bias_arry() As Long
    Dim DCO_arry() As Long
    'Trim_Pll_Freq@1000,Vro@300,DCO@6,Fcount_Cap@Fcount-Cap0-@Fcount-Cap1-@Fcount-Cap3-,Fcount_Bias@Bias0@Bias1@Bias2@Bias3@Bias4@Bias5@Bias6@Bias7,Storename_src_bit
    'argv(0) : Trim_Pll_Freq@1000
    'argv(1) : Vro@300
    'argv(2) : DCO@6
    'argv(3) : Fcount_Cap@Fcount-Cap0-@Fcount-Cap1-@Fcount-Cap3-
    'argv(4) : Fcount_Bias@Bias0@Bias1@Bias2@Bias3@Bias4@Bias5@Bias6@Bias7
    'argv(5) : Storename_src_bit
    
    SplitTrimFreq = Split(argv(0), "@") 'Trim_Pll_Freq@1000
    SplitVro = Split(argv(1), "@") 'Vro@300
    SplitDCO = Split(argv(2), "@") 'DCO@5
    SplitCap = Split(argv(3), "@") 'Fcount_Cap@Fcount-Cap0-@Fcount-Cap1-@Fcount-Cap3-
    SplitBias = Split(argv(4), "@") 'Fcount_Bias@Bias0@Bias1@Bias2@Bias3@Bias4@Bias5@Bias6@Bias7
    ReDim F_TrimComplete(UBound(SplitCap))
    ReDim F_Vro(UBound(SplitCap))
    ReDim VroVoltage(UBound(SplitCap))
    ReDim BiasTargetIndex(UBound(SplitCap))
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'Follow Customer Instruction to create the bit space
    ReDim Bias_arry(2)  'Bit0-2
    ReDim Cap_arry(1)   'Bit3-4
    ReDim DCO_arry(2)   'Bit5-7
    CapDSPWave.CreateConstant 0, (UBound(Cap_arry) + 1)
    BiasDSPWave.CreateConstant 0, (UBound(Bias_arry) + 1)
    DCODSPWave.CreateConstant 0, (UBound(DCO_arry) + 1)
    FinalDSPWave.CreateConstant 0, ((CapDSPWave.SampleSize) + (BiasDSPWave.SampleSize) + (DCODSPWave.SampleSize))
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
    For i = 0 To UBound(SplitCap)
        VroVoltage(i) = 0
        F_TrimComplete(i) = False
        F_Vro(i) = False
    Next i
        
    For i = 1 To UBound(SplitBias)
        For j = 1 To UBound(SplitCap)
            For Each site In TheExec.sites.Active
                If Not F_TrimComplete(j) Then
                    DSPWaveFromDict = GetStoredCaptureData(LCase(SplitCap(j) & SplitBias(i)))
                    ''DSPWaveFromDict.Element(i + 4) = 1 '' Debug
                    DSPWaveDecType = DSPWaveFromDict.ConvertStreamTo(tldspParallel, 16, 0, Bit0IsMsb)
                    If DSPWaveDecType.Element(0) >= CInt(SplitTrimFreq(1)) Then
                        F_TrimComplete(j) = True
                        VroVoltage(j) = GetStoredMeasurement(CStr(LCase(SplitCap(j) & SplitBias(i) & SplitVro(0))))
                        ''VroVoltage(j) = VroVoltage(j).Add(j * 0.15) '' Debug
                        If VroVoltage(j) * 1000 >= SplitVro(1) Then
                            F_Vro(j) = True
                            VroTarget = VroVoltage(j)
                            FinalVroIndex = j
                            BiasTargetIndex(j) = i
                        Else
                            F_Vro(j) = False
                        End If
                    End If
                End If
            Next site
        Next j
    Next i
    
    For Each site In TheExec.sites.Active
        For k = 1 To UBound(SplitCap)
            If F_Vro(k) Then
                If VroVoltage(k) <= VroTarget Then
                    FinalVroIndex = k
                    VroTarget = VroVoltage(k)
                    FinalBiasIndex = BiasTargetIndex(k)
                End If
            End If
        Next k
        
        If right(SplitBias(FinalBiasIndex), 1) = "s" Then
        
            TheExec.Datalog.WriteComment ("ERROR: No parameter can reach the target")
        Else
            FinalBiasIndex = CLng(right(SplitBias(FinalBiasIndex), 1))
            FinalVroIndex = CLng(mid(SplitCap(FinalVroIndex), (Len(SplitCap(FinalVroIndex)) - 1), 1))
'            CapDSPWave.CreateConstant 0, (UBound(Cap_arry) + 1)
'            BiasDSPWave.CreateConstant 0, (UBound(Bias_arry) + 1)
'            DCODSPWave.CreateConstant 0, (UBound(DCO_arry) + 1)
'            FinalDSPWave.CreateConstant 0, ((CapDSPWave.SampleSize) + (BiasDSPWave.SampleSize) + (DCODSPWave.SampleSize))
            Call Dec2Bin(FinalBiasIndex, Bias_arry())
            For i = 0 To UBound(Bias_arry)
                FinalDSPWave.Element(i) = Bias_arry((UBound(Bias_arry)) - i)
            Next i
            Call Dec2Bin(FinalVroIndex, Cap_arry())
            For j = 0 To UBound(Cap_arry)
                FinalDSPWave.Element(j + UBound(Bias_arry) + 1) = Cap_arry((UBound(Cap_arry)) - j) 'Bit start from 3
            Next j
            
           ' If Bias_arry(0) = 1 And Bias_arry(1) = 1 And Bias_arry(2) = 1 And FinalVroIndex = 0 Then
           
            UseLimitTname = CStr(Instance_Data.Tname(TheExec.Flow.TestLimitIndex))          ' Dylan Edit by 20190616
            TestNameInput = Report_TName_From_Instance("Calc", "X", UseLimitTname, 0, 0)
            If FinalBiasIndex = 7 And FinalVroIndex = 0 Then
                TheExec.Flow.TestLimit resultVal:=1, hiVal:=0, lowVal:=0, Tname:=TestNameInput, ForceResults:=tlForceFlowFail
            Else
                TheExec.Flow.TestLimit resultVal:=0, hiVal:=0, lowVal:=0, Tname:=TestNameInput, ForceResults:=tlForceFlowPass
            End If
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1       ' Edited for avoid useLimit index messy
            
            Call Dec2Bin(CLng(SplitDCO(1)), DCO_arry())
            For k = 0 To UBound(DCO_arry)
                FinalDSPWave.Element(k + UBound(Bias_arry) + UBound(Cap_arry) + 1 + 1) = DCO_arry((UBound(DCO_arry)) - k) 'Bit start from 5
            Next k
        End If
        TheExec.Datalog.WriteComment ("Final Bias Index : " & FinalBiasIndex)
        TheExec.Datalog.WriteComment ("Final Cap Index : " & FinalVroIndex)
        TheExec.Datalog.WriteComment ("DCO : " & CLng(SplitDCO(1)))
    Next site
    Call AddStoredCaptureData(LCase(argv(5)), FinalDSPWave)
End Function
Public Function Calc_DCC_Skew_Range_DSP(argc As Integer, argv() As String) As Long
 
    ''''Demo String : CH@CH0@CH1,DQ@DQ0@DQ1,CountIN@0x1F@0x0@1Fx0,Count100@0x1F@0x0@1Fx0,SkewFactor@0.5,InputFactor@1.5, PatternBit@13
    Dim i, j, k, y As Long
    Dim SplitCH() As String
    Dim SplitDQ() As String
    Dim SplitCountIN() As String
    Dim SplitCount100() As String
    Dim DC_Skew_Input_Array() As New SiteDouble
    Dim DC_Input_CLK_UP As New SiteDouble
    Dim DC_Input_CLK_NO_DCC As New SiteDouble
    Dim DC_Input_CLK_DOWN As New SiteDouble
    Dim DC_Skew_Input_CLK_UP As New SiteDouble
    Dim DC_Skew_Input_CLK_NO_DCC As New SiteDouble
    Dim DC_Skew_Input_CLK_DOWN As New SiteDouble
    Dim DCC_RANGE_UP As New SiteDouble
    Dim DCC_RANGE_DOWN As New SiteDouble
    Dim SkewFactor As Double
    Dim InputFactor As Double
    Dim DSPWaveTemp As New DSPWave
    Dim DSPWaveTemp1 As New DSPWave
    Dim DSPWaveDec As New DSPWave
    Dim DSPWaveDec1 As New DSPWave
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    
    SplitCH = Split(argv(0), "@")   'CH@CH0@CH1
    SplitDQ = Split(argv(1), "@")   'DQ@DQ0@DQ1
    SplitCountIN = Split(argv(2), "@")   'CountIN@0x1F@0x0@1Fx0
    SplitCount100 = Split(argv(3), "@")   'Count100@0x1F@0x0@1Fx0
    SkewFactor = Split(argv(4), "@")(1) 'Factor1@0.5
    InputFactor = Split(argv(5), "@")(1) 'Factor2@1.5
    DSPWaveTemp.CreateConstant 0, Split(argv(6), "@")(1), DspLong 'PatternBit@13
    DSPWaveTemp1.CreateConstant 0, Split(argv(6), "@")(1), DspLong 'PatternBit@13
    DSPWaveDec.CreateConstant 0, 1
    DSPWaveDec1.CreateConstant 0, 1
    ReDim DC_Skew_Input_Array(UBound(SplitCountIN) - 1)
    
    For i = 0 To (UBound(SplitCH) - 1)
        For j = 0 To (UBound(SplitDQ) - 1)
            For Each site In TheExec.sites.Active
                For k = 0 To (UBound(SplitCountIN) - 1)
''''''''''                    DSPWaveTemp = GetStoredCaptureData(SplitCH(i + 1) & SplitDQ(j + 1) & "x" & SplitCountIN(k + 1) & SplitCountIN(0))
''''''''''                    DSPWaveTemp1 = GetStoredCaptureData(SplitCH(i + 1) & SplitDQ(j + 1) & "x" & SplitCount100(k + 1) & SplitCount100(0))
                    DSPWaveDec = GetStoredCaptureData("2SDEC_" & SplitCH(i + 1) & SplitDQ(j + 1) & "x" & SplitCountIN(k + 1) & SplitCountIN(0))
                    DSPWaveDec1 = GetStoredCaptureData("2SDEC_" & SplitCH(i + 1) & SplitDQ(j + 1) & "x" & SplitCount100(k + 1) & SplitCount100(0))
                    
''''''''''                    DSPWaveDec = DSPWaveTemp.ConvertStreamTo(tldspParallel, Split(argv(6), "@")(1), 0, Bit0IsMsb)
''''''''''                    DSPWaveDec1 = DSPWaveTemp1.ConvertStreamTo(tldspParallel, Split(argv(6), "@")(1), 0, Bit0IsMsb)
                    If DSPWaveDec1.Element(0) = 0 Then
                        TheExec.Datalog.WriteComment ("Can't divide by 0")
                    Else
                        DC_Skew_Input_Array(k) = (DSPWaveDec.Element(0) / DSPWaveDec1.Element(0)) * SkewFactor
                    End If
                Next k
                DC_Input_CLK_UP = DC_Skew_Input_Array(0) + 0.5
                DC_Input_CLK_NO_DCC = DC_Skew_Input_Array(1) + 0.5
                DC_Input_CLK_DOWN = DC_Skew_Input_Array(2) + 0.5
                
                DC_Skew_Input_CLK_UP = DC_Skew_Input_Array(0)
                DC_Skew_Input_CLK_NO_DCC = DC_Skew_Input_Array(1)
                DC_Skew_Input_CLK_DOWN = DC_Skew_Input_Array(2)
                DCC_RANGE_UP = DC_Skew_Input_CLK_UP - DC_Skew_Input_CLK_NO_DCC
                DCC_RANGE_DOWN = DC_Skew_Input_CLK_DOWN - DC_Skew_Input_CLK_NO_DCC
''''''''''                TheExec.Datalog.WriteComment ("Site " & Site & " : " & SplitCH(i + 1) & "_" & SplitDQ(j + 1) & "_" & "DCC_RANGE_UP" & " = " & DCC_RANGE_UP)
''''''''''                TheExec.Datalog.WriteComment ("Site " & Site & " : " & SplitCH(i + 1) & "_" & SplitDQ(j + 1) & "_" & "DCC_RANGE_DOWN" & " = " & DCC_RANGE_DOWN)
            Next site
            
            TestNameInput = Report_TName_From_Instance("Calc", SplitCH(i + 1) & SplitDQ(j + 1), "DC_Input_CLK", CInt(i), , "replace;7=UP", , , tlForceNone) 'eng_forceflow_transfer
            TheExec.Flow.TestLimit resultVal:=DC_Input_CLK_UP.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%" 'eng_forceflow_transfer
            TestNameInput = Report_TName_From_Instance("Calc", SplitCH(i + 1) & SplitDQ(j + 1), "DC_Input_CLK", CInt(i), , "replace;7=NODCC", , , tlForceNone) 'eng_forceflow_transfer
            TheExec.Flow.TestLimit resultVal:=DC_Input_CLK_NO_DCC.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%" 'eng_forceflow_transfer
            TestNameInput = Report_TName_From_Instance("Calc", SplitCH(i + 1) & SplitDQ(j + 1), "DC_Input_CLK", CInt(i), , "replace;7=DOWN", , , tlForceNone) 'eng_forceflow_transfer
            TheExec.Flow.TestLimit resultVal:=DC_Input_CLK_DOWN.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%" 'eng_forceflow_transfer
            
            TestNameInput = Report_TName_From_Instance("Calc", SplitCH(i + 1) & SplitDQ(j + 1), "DC_Skew_Input_CLK", CInt(i), , "replace;7=UP", , , tlForceNone) 'eng_forceflow_transfer
            TheExec.Flow.TestLimit resultVal:=DC_Skew_Input_CLK_UP.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%" 'eng_forceflow_transfer
            TestNameInput = Report_TName_From_Instance("Calc", SplitCH(i + 1) & SplitDQ(j + 1), "DC_Skew_Input_CLK", CInt(i), , "replace;7=NODCC", , , tlForceNone) 'eng_forceflow_transfer
            TheExec.Flow.TestLimit resultVal:=DC_Skew_Input_CLK_NO_DCC.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%" 'eng_forceflow_transfer
            TestNameInput = Report_TName_From_Instance("Calc", SplitCH(i + 1) & SplitDQ(j + 1), "DC_Skew_Input_CLK", CInt(i), , "replace;7=DOWN", , , tlForceNone) 'eng_forceflow_transfer
            TheExec.Flow.TestLimit resultVal:=DC_Skew_Input_CLK_DOWN.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%" 'eng_forceflow_transfer
            
            TestNameInput = Report_TName_From_Instance("Calc", SplitCH(i + 1) & SplitDQ(j + 1), "DCC_RANGE_UP", CInt(i), , , , , tlForceNone) 'eng_forceflow_transfer
            TheExec.Flow.TestLimit resultVal:=DCC_RANGE_UP.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%" 'eng_forceflow_transfer
            TestNameInput = Report_TName_From_Instance("Calc", SplitCH(i + 1) & SplitDQ(j + 1), "DCC_RANGE_DOWN", CInt(i), , , , , tlForceNone) 'eng_forceflow_transfer
            TheExec.Flow.TestLimit resultVal:=DCC_RANGE_DOWN.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%" 'eng_forceflow_transfer
        
        Next j
    Next i
End Function


Public Function Calc_DiCap_ParallelMode_For_IPPM(argc As Integer, argv() As String) As Long
' Edided by 20190613
'**********************************************************
' Format : DictionaryName & dc_meas @ t_meas
' Calculate Function : duty_cycle = (dc_meas_int*2) / (t_meas_int+2^17)
'**********************************************************

    Dim site As Variant
    Dim i, j, k As Integer
    Dim BitPosition As Long
    Dim RegSplit() As String
    Dim AssembleStr() As String
    Dim FormatSplit() As String
    Dim BinaryWave As New DSPWave
    Dim CalDSPWave As New DSPWave
    Dim duty_cycle() As New DSPWave
    Dim SplitDspWave() As New DSPWave
    ReDim AssembleStr(argc - 1)
    ReDim duty_cycle(argc - 1)
    
    For i = 0 To argc - 1
        'BitPosition = 0
        FormatSplit = Split(argv(i), "&")
        RegSplit = Split(FormatSplit(1), "@")
        
        ReDim SplitDspWave(UBound(RegSplit))
        duty_cycle(i).CreateConstant 0, 1, DspDouble
        BinaryWave = GetStoredCaptureData(FormatSplit(0))
        
        For Each site In TheExec.sites
            For j = 0 To BinaryWave.SampleSize - 1
                AssembleStr(i) = CStr(BinaryWave(site).Element(j)) & AssembleStr(i)
            Next j
            BitPosition = 0
            For j = 0 To UBound(RegSplit)
               
                SplitDspWave(j).CreateConstant 0, 1, DspLong
                BinaryWave = BinaryWave(site).ConvertDataTypeTo(DspLong)
                CalDSPWave = BinaryWave(site).Select(CLng(BitPosition), 1, CLng(RegSplit(j))).Copy
                SplitDspWave(j) = CalDSPWave.ConvertStreamTo(tldspParallel, CLng(RegSplit(j)), 0, Bit0IsMsb)
                BitPosition = BitPosition + CLng(RegSplit(j))
            Next j
            duty_cycle(i).Element(0) = (SplitDspWave(0).Element(0) * 2) / (SplitDspWave(1).Element(0) + 2 ^ 17)
            duty_cycle(i).Element(0) = duty_cycle(i).Element(0) * 100
        Next site
        TheExec.Datalog.WriteComment FormatSplit(0) & " Binary Value : " & AssembleStr(i)
    Next i
    For i = 0 To argc - 1
        TheExec.Flow.TestLimit resultVal:=duty_cycle(i).Element(0), Tname:=FormatSplit(0), ForceResults:=tlForceFlow
    Next
End Function

Public Function prasing_ADC(RAW_DSP As DSPWave, ADC_bits As Long, sgmt_size) As DSPWave

    Dim new_DSP As New DSPWave
    Dim i As Long, j As Long
    Dim sgmt_cnt As Long
    sgmt_cnt = RAW_DSP.SampleSize / ADC_bits / sgmt_size
    
    For i = 0 To sgmt_cnt - 1
        For j = 0 To sgmt_size - 1
            If new_DSP.SampleSize = 0 Then
                new_DSP = RAW_DSP.Select(0, sgmt_size, ADC_bits).Copy
            Else
                new_DSP = new_DSP.Concatenate(RAW_DSP.Select(i * ADC_bits * sgmt_size + j, sgmt_size, ADC_bits).Copy)
            End If
        Next j
    Next i
    
    Set prasing_ADC = new_DSP.ConvertStreamTo(tldspParallel, ADC_bits, 0, Bit0IsMsb)

End Function

Public Function Calc_12bADC_MeanWithVariance(argc As Integer, argv() As String) As Long
'Alg::Calc_DigCap_MeanWithVariance(SegmentSize_32,CFG_FIFO_SDM5|2047,CFG_FIFO_SDM6|1024,CFG_FIFO_SDM7|3072)

    Dim i As Long, j As Long, k As Long
    Dim site As Variant
    Dim SegmentSize As Long: SegmentSize = CLng(Split(argv(0), "_")(1))
    Dim GroupCount As Long: GroupCount = argc - 1
    Dim ADC_bits As Long: ADC_bits = 12
    
    Dim mean As Double
    Dim STDEV As Double
    Dim TestNameInput As String
    Dim Limit_Dev As Long: Limit_Dev = 60
    
    Dim CalcResult() As New PinListData
    Dim Limit_Val() As Long
    Dim key() As String
    ReDim CalcResult(GroupCount - 1)
    ReDim Limit_Val(GroupCount - 1)
    ReDim key(GroupCount - 1)
    
    Dim DSPWave_Ori() As New DSPWave
    ReDim DSPWave_Ori(GroupCount - 1)
    
    Dim ADC_result() As New DSPWave
    ReDim ADC_result(GroupCount - 1)

    For i = 0 To GroupCount - 1
        key(i) = Split(argv(i + 1), "|")(0)
        DSPWave_Ori(i) = GetStoredCaptureData(key(i))
        CalcResult(i).AddPin "Mean"
        CalcResult(i).AddPin "Variance"
        CalcResult(i).AddPin "MaxErr"
        Limit_Val(i) = Split(argv(i + 1), "|")(1)
        Set ADC_result(i) = Nothing
        ADC_result(i) = prasing_ADC(DSPWave_Ori(i), ADC_bits, SegmentSize)
        'Call DebugPrintRawDigCap(DSPWave_Ori(i), SegmentSize)
    Next i

    For Each site In TheExec.sites.Active
        For i = 0 To GroupCount - 1
            mean = ADC_result(i).CalcMeanWithStdDev(STDEV)
            CalcResult(i).Pins("Mean").value(site) = mean
            CalcResult(i).Pins("Variance").value(site) = STDEV * STDEV
            CalcResult(i).Pins("MaxErr").value(site) = ADC_result(i).CalcMaximumValue - ADC_result(i).CalcMinimumValue
        Next i
    Next site
    
    k = 0
    For i = 0 To GroupCount - 1
        For Each site In TheExec.sites.Active
            If True Then
                For j = 0 To ADC_result(i).SampleSize - 1
                    TheExec.Flow.TestLimit ADC_result(i).data(j), Limit_Val(i) - Limit_Dev, Limit_Val(i) + Limit_Dev, _
                    Tname:=Report_TName_From_Instance("C", "X", key(i), Instance_Data.TestSeqNum + CInt(k)), ForceResults:=tlForceNone 'eng_forceflow_transfer
                    k = k + 1
                Next j
            End If
            TheExec.Flow.TestLimit resultVal:=CalcResult(i).Pins("Mean").value(site), _
            Tname:=Report_TName_From_Instance("C", "X", key(i), Instance_Data.TestSeqNum + CInt(k + 1), , "replace;7=Mean"), ForceResults:=tlForceNone 'eng_forceflow_transfer
            TheExec.Flow.TestLimit resultVal:=CalcResult(i).Pins("Variance").value(site), _
            Tname:=Report_TName_From_Instance("C", "X", key(i), Instance_Data.TestSeqNum + CInt(k + 2), , "replace;7=Variance"), ForceResults:=tlForceNone 'eng_forceflow_transfer
            TheExec.Flow.TestLimit resultVal:=CalcResult(i).Pins("MaxErr").value(site), _
            Tname:=Report_TName_From_Instance("C", "X", key(i), Instance_Data.TestSeqNum + CInt(k + 3), , "replace;7=MaxErr"), ForceResults:=tlForceNone 'eng_forceflow_transfer
            k = k + 3
        Next site
    Next i
    
End Function

Public Function DebugPrintRawDigCap(InWave As DSPWave, sgmt_size As Long)
    
    Dim i As Long, j As Long
    Dim prtStr As String
    Dim PartialWave As New DSPWave
    
    For i = 0 To (InWave.SampleSize / sgmt_size) - 1
        prtStr = vbNullString
        Set PartialWave = Nothing
        
        PartialWave = InWave.Select(i * sgmt_size, 1, sgmt_size).Copy
        For j = 0 To sgmt_size - 1
            prtStr = prtStr & PartialWave.Element(j)
        Next j
        TheExec.Datalog.WriteComment "Line" & Space(3 - Len(CStr(i))) & i & ":" & prtStr
        
    Next i
End Function

Public Function Calc_ADC_average(argc As Integer, argv() As String) As Long

    Dim InWave As New DSPWave
    Dim i As Long, j As Long, k As Long
    Dim ADC_wave(2) As New DSPWave
    Dim Chk_wave(2) As New DSPWave
    Dim mean As Double
    Dim STDEV As Double
    Dim MaxErr As Double
    
    InWave = GetStoredCaptureData(argv(0))
    InWave = InWave.ConvertDataTypeTo(DspLong)
    
    For i = 0 To 2
        k = 0
        'Chk_wave(i) = InWave.Select(13 * (i + 1) - 1, 39, 256).Copy
        Chk_wave(i) = InWave.Select(3328 * i - 1 + 13, 13, 256).Copy
        Set ADC_wave(i) = Nothing
        ADC_wave(i) = ADC_wave(i).ConvertDataTypeTo(DspLong)
        For j = 0 To 255
'            ADC_wave(i) = ADC_wave(i).Concatenate(InWave.Select(39 * j + 13 * i, 1, 12).Copy)
             ADC_wave(i) = ADC_wave(i).Concatenate(InWave.Select(13 * j + 3328 * i, 1, 12).Copy)
        Next j
        
        ADC_wave(i) = ADC_wave(i).ConvertStreamTo(tldspParallel, 12, 0, Bit0IsMsb)
        
        For j = 0 To 255
            TheExec.Flow.TestLimit ADC_wave(i).data(j), _
            Tname:=Report_TName_From_Instance("C", "X", "SDM" & (5 + i), CInt(k)), ForceResults:=tlForceNone 'eng_forceflow_transfer
            TheExec.Flow.TestLimit Chk_wave(i).data(j), _
            Tname:=Report_TName_From_Instance("C", "X", "SDM" & (5 + i), CInt(k), , "replace;7=Chk"), ForceResults:=tlForceNone 'eng_forceflow_transfer
            k = k + 1
        Next j
        
        mean = ADC_wave(i).CalcMeanWithStdDev(STDEV)
        STDEV = STDEV * STDEV
        MaxErr = ADC_wave(i).CalcMaximumValue - ADC_wave(i).CalcMinimumValue
                    
        TheExec.Flow.TestLimit resultVal:=mean, Tname:=Report_TName_From_Instance("C", "X", "SDM" & (5 + i), CInt(k), , "replace;7=Mean"), ForceResults:=tlForceNone 'eng_forceflow_transfer
        TheExec.Flow.TestLimit resultVal:=STDEV, Tname:=Report_TName_From_Instance("C", "X", "SDM" & (5 + i), CInt(k + 1), , "replace;7=Variance"), ForceResults:=tlForceNone 'eng_forceflow_transfer
        TheExec.Flow.TestLimit resultVal:=MaxErr, Tname:=Report_TName_From_Instance("C", "X", "SDM" & (5 + i), CInt(k + 2), , "replace;7=MaxErr"), ForceResults:=tlForceNone 'eng_forceflow_transfer
    Next i
    

End Function
Public Function Calc_D2D_MAX_MIN_V2(argc As Integer, argv() As String) As Long

    Dim site As Variant
    Dim i, j, k As Integer
    Dim FirstLoop As String
    Dim SecondLoop As String
    Dim SplitCalStr() As String
    Dim TestNameInput As String
    Dim ValueMax As New SiteLong
    Dim ValueMin As New SiteLong
    Dim MaxCKDNL As New SiteLong
    Dim MinCKDNL As New SiteLong
    
    Dim DCKMAX As New SiteLong
    Dim DCKMin As New SiteLong
    Dim Idsvalue As New SiteDouble
    Dim DCKIdsvalue As New SiteDouble
    
    Dim IndexOfMinimumValue As Long
    Dim IndexOfMaximumValue As Long
    
    Dim SaveDCK_DSPWave As New DSPWave
    Dim DNLValue_DSPWave As New DSPWave
    Dim PreValue_DSPWave As New DSPWave
    Dim DeltaDCK_DSPWave As New DSPWave
    Dim DictValue_DSPWave As New DSPWave
    Dim DNLDCKValue_DSPWave As New DSPWave
    Dim SaveMeasNum_DSPWave As New DSPWave
    Dim SaveDeltaValue_DSPWave As New DSPWave
    
    
    
    FirstLoop = CStr(TheExec.Flow.var("SrcCodeIndx").value)
    SecondLoop = CStr(TheExec.Flow.var("SrcCodeIndx1").value)
''''''''''    Debug.Print "SrcCodeIndx Value : " & FirstLoop
''''''''''    Debug.Print "SrcCodeIndx1 Value : " & SecondLoop
    
    
''''''''''    LoopNum = SecondLoop = SrcCodeIndx1
''''''''''    LoopNum1 = FirstLoop = SrcCodeIndx
    
    For i = 0 To argc - 1
        SplitCalStr = Split(argv(i), "@")
        DictValue_DSPWave = GetStoredCaptureData(SplitCalStr(0))
        For Each site In TheExec.sites.Active
            DictValue_DSPWave = DictValue_DSPWave.ConvertDataTypeTo(DspLong)
            DictValue_DSPWave = DictValue_DSPWave.ConvertStreamTo(tldspParallel, DictValue_DSPWave.SampleSize, 0, Bit0IsMsb)
        Next site
        Call AddStoredCaptureData(SplitCalStr(0) & "_" & SecondLoop, DictValue_DSPWave)
               
        If SecondLoop = 0 Then
            SaveDeltaValue_DSPWave.CreateConstant 0, CLng(SplitCalStr(1)), DspLong
            SaveMeasNum_DSPWave.CreateConstant 0, CLng(SplitCalStr(1)) + 1, DspLong
            
            If FirstLoop = 0 Then
                SaveDCK_DSPWave.CreateConstant 0, CLng(SplitCalStr(1)) + 1, DspLong
            Else
                SaveDCK_DSPWave = GetStoredCaptureData("SaveDCK_DSPWaveData")
            End If
            For Each site In TheExec.sites.Active
                SaveDCK_DSPWave.Element(FirstLoop) = DictValue_DSPWave.Element(0)
                SaveMeasNum_DSPWave.Element(SecondLoop) = DictValue_DSPWave.Element(0)
            Next site
            Call AddStoredCaptureData("SaveDCK_DSPWaveData", SaveDCK_DSPWave)
            Call AddStoredCaptureData("SaveMeasNum_DSPWaveData", SaveMeasNum_DSPWave)
            Call AddStoredCaptureData("SaveDeltaValue_DSPWaveData", SaveDeltaValue_DSPWave)
            
        ElseIf SecondLoop <> 64 Then
            SaveMeasNum_DSPWave = GetStoredCaptureData("SaveMeasNum_DSPWaveData")
            SaveDeltaValue_DSPWave = GetStoredCaptureData("SaveDeltaValue_DSPWaveData")
            PreValue_DSPWave = GetStoredCaptureData(SplitCalStr(0) & "_" & SecondLoop - 1)
            
            For Each site In TheExec.sites.Active
                SaveDeltaValue_DSPWave.Element(SecondLoop - 1) = Abs(DictValue_DSPWave.Element(0) - PreValue_DSPWave.Element(0))
                SaveMeasNum_DSPWave.Element(SecondLoop) = DictValue_DSPWave.Element(0)
            Next site
            Call AddStoredCaptureData("SaveMeasNum_DSPWaveData", SaveMeasNum_DSPWave)
            Call AddStoredCaptureData("SaveDeltaValue_DSPWaveData", SaveDeltaValue_DSPWave)
            
        ElseIf SecondLoop = 64 Then
            SaveMeasNum_DSPWave = GetStoredCaptureData("SaveMeasNum_DSPWaveData")
            SaveDeltaValue_DSPWave = GetStoredCaptureData("SaveDeltaValue_DSPWaveData")
            
            For Each site In TheExec.sites.Active
                ValueMax = SaveMeasNum_DSPWave.CalcMaximumValue(IndexOfMaximumValue)
                ValueMin = SaveMeasNum_DSPWave.CalcMinimumValue(IndexOfMinimumValue)
                Idsvalue = (ValueMax - ValueMin) / (CLng(SplitCalStr(1)) + 1)
''''''''''                TheExec.Datalog.WriteComment "MAX_Value:  " & SaveMeasNum_DSPWave.CalcMaximumValue(IndexOfMaximumValue)
''''''''''                TheExec.Datalog.WriteComment "MIN_Value:  " & SaveMeasNum_DSPWave.CalcMinimumValue(IndexOfMinimumValue)
''''''''''                TheExec.Datalog.WriteComment "FinalIds value : " & Idsvalue
            Next site
            TestNameInput = Report_TName_From_Instance("C", "X", "CK" & Format(FirstLoop, "00") & "-START", CInt(i))
            TheExec.Flow.TestLimit resultVal:=ValueMax, Tname:=TestNameInput, ForceResults:=tlForceFlow
            TestNameInput = Report_TName_From_Instance("C", "X", "CK" & Format(FirstLoop, "00") & "-END", CInt(i))
            TheExec.Flow.TestLimit resultVal:=ValueMin, Tname:=TestNameInput, ForceResults:=tlForceFlow
            TestNameInput = Report_TName_From_Instance("C", "X", "DCK" & Format(FirstLoop, "00") & "-IDEAL", CInt(i))
            TheExec.Flow.TestLimit resultVal:=Idsvalue, Tname:=TestNameInput, ForceResults:=tlForceFlow

            DNLValue_DSPWave.CreateConstant 0, CLng(SplitCalStr(1)), DspDouble
            For j = 0 To CLng(SplitCalStr(1)) - 1
                For Each site In TheExec.sites.Active
                    DNLValue_DSPWave.Element(j) = (SaveDeltaValue_DSPWave.Element(j) - Idsvalue) / Idsvalue
''''''''''                    TheExec.Datalog.WriteComment "CK Delta value" & j & " = " & CStr(SaveDeltaValue_DSPWave.Element(j))
''''''''''                    TheExec.Datalog.WriteComment "DNL value on point" & j & " = " & CStr(DNLValue_DSPWave.Element(j))
                Next site
                TestNameInput = Report_TName_From_Instance("C", "X", "CKDeltavalue" & Format(j, "00"), CInt(i))
                TheExec.Flow.TestLimit resultVal:=SaveDeltaValue_DSPWave.Element(j), Tname:=TestNameInput, ForceResults:=tlForceFlow
                TestNameInput = Report_TName_From_Instance("C", "X", "CKDNLvalue" & Format(j, "00"), CInt(i))
                TheExec.Flow.TestLimit resultVal:=DNLValue_DSPWave.Element(j), Tname:=TestNameInput, ForceResults:=tlForceFlow
            Next j
            
            For Each site In TheExec.sites.Active
                MaxCKDNL = DNLValue_DSPWave.CalcMaximumValue(IndexOfMaximumValue)
                MinCKDNL = DNLValue_DSPWave.CalcMinimumValue(IndexOfMinimumValue)
''''''''''                TheExec.Datalog.WriteComment "MAX DNL:" & MaxCKDNL
''''''''''                TheExec.Datalog.WriteComment "MIN DNL:" & MinCKDNL
            Next site
            TestNameInput = Report_TName_From_Instance("C", "X", "CK" & Format(FirstLoop, "00") & "-MAX-DNL", CInt(i))
            TheExec.Flow.TestLimit resultVal:=MaxCKDNL, Tname:=TestNameInput, ForceResults:=tlForceFlow
            TestNameInput = Report_TName_From_Instance("C", "X", "CK" & Format(FirstLoop, "00") & "-MIN-DNL", CInt(i))
            TheExec.Flow.TestLimit resultVal:=MinCKDNL, Tname:=TestNameInput, ForceResults:=tlForceFlow
            TestNameInput = Report_TName_From_Instance("C", "X", "CK" & Format(FirstLoop, "00") & "-MAXStepDelta", CInt(i))
            For Each site In TheExec.sites.Active
                TheExec.Flow.TestLimit resultVal:=CStr(SaveDeltaValue_DSPWave.CalcMaximumValue(IndexOfMaximumValue)), Tname:=TestNameInput, ForceResults:=tlForceFlow
                TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
            Next site
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
            
            TestNameInput = Report_TName_From_Instance("C", "X", "CK" & Format(FirstLoop, "00") & "-MINStepDelta", CInt(i))
            For Each site In TheExec.sites.Active
                TheExec.Flow.TestLimit resultVal:=CStr(SaveDeltaValue_DSPWave.CalcMinimumValue(IndexOfMinimumValue)), Tname:=TestNameInput, ForceResults:=tlForceFlow
                TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
            Next site
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
        End If
            
        If SecondLoop = 64 And FirstLoop = 64 Then
            SaveDCK_DSPWave = GetStoredCaptureData("SaveDCK_DSPWaveData")
            
            For Each site In TheExec.sites.Active
                DCKMAX = SaveDCK_DSPWave.CalcMaximumValue(IndexOfMaximumValue)
                DCKMin = SaveDCK_DSPWave.CalcMinimumValue(IndexOfMinimumValue)
                DCKIdsvalue = (DCKMAX - DCKMin) / (CLng(SplitCalStr(1)) + 1)
            Next site
            TheExec.Datalog.WriteComment "------------------------------------------------- SDLL_DCK summary--------------------------------------------------------------"
            TestNameInput = Report_TName_From_Instance("C", "X", "DCK" & Format(FirstLoop, "00") & "-START", CInt(i))
            TheExec.Flow.TestLimit resultVal:=DCKMAX, Tname:=TestNameInput, ForceResults:=tlForceFlow
            TestNameInput = Report_TName_From_Instance("C", "X", "DCK" & Format(FirstLoop, "00") & "-END", CInt(i))
            TheExec.Flow.TestLimit resultVal:=DCKMin, Tname:=TestNameInput, ForceResults:=tlForceFlow
            TestNameInput = Report_TName_From_Instance("C", "X", "DCK" & Format(FirstLoop, "00") & "-IDEAL", CInt(i))
            TheExec.Flow.TestLimit resultVal:=DCKIdsvalue, Tname:=TestNameInput, ForceResults:=tlForceFlow

            DeltaDCK_DSPWave.CreateConstant 0, CLng(SplitCalStr(1)), DspDouble
            DNLDCKValue_DSPWave.CreateConstant 0, CLng(SplitCalStr(1)), DspDouble

            For j = 1 To CLng(SplitCalStr(1))
                For Each site In TheExec.sites.Active
                    DeltaDCK_DSPWave.Element(j - 1) = SaveDCK_DSPWave.Element(j - 1) - SaveDCK_DSPWave.Element(j)
                    DNLDCKValue_DSPWave.Element(j - 1) = (DeltaDCK_DSPWave.Element(j - 1) - DCKIdsvalue) / DCKIdsvalue
                Next site
                TestNameInput = Report_TName_From_Instance("C", "X", "DCKDeltavalue" & Format(j - 1, "00"), CInt(i))
                TheExec.Flow.TestLimit resultVal:=SaveDCK_DSPWave.Element(j - 1), Tname:=TestNameInput, ForceResults:=tlForceFlow
                TestNameInput = Report_TName_From_Instance("C", "X", "DCKDNLvalue" & Format(j - 1, "00"), CInt(i))
                TheExec.Flow.TestLimit resultVal:=DNLDCKValue_DSPWave.Element(j - 1), Tname:=TestNameInput, ForceResults:=tlForceFlow
            Next j
            
            TestNameInput = Report_TName_From_Instance("C", "X", "DCK" & Format(FirstLoop, "00") & "-MAX-DNL", CInt(i))
            For Each site In TheExec.sites.Active
                TheExec.Flow.TestLimit resultVal:=DNLDCKValue_DSPWave.CalcMaximumValue(IndexOfMaximumValue), Tname:=TestNameInput, ForceResults:=tlForceFlow
                TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
            Next site
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
            
            TestNameInput = Report_TName_From_Instance("C", "X", "DCK" & Format(FirstLoop, "00") & "-MIN-DNL", CInt(i))
            For Each site In TheExec.sites.Active
                TheExec.Flow.TestLimit resultVal:=DNLDCKValue_DSPWave.CalcMinimumValue(IndexOfMinimumValue), Tname:=TestNameInput, ForceResults:=tlForceFlow
                TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
            Next site
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
                
            TestNameInput = Report_TName_From_Instance("C", "X", "DCK" & Format(FirstLoop, "00") & "-MAXStepDelta", CInt(i))
            For Each site In TheExec.sites.Active
                TheExec.Flow.TestLimit resultVal:=DeltaDCK_DSPWave.CalcMaximumValue(IndexOfMaximumValue), Tname:=TestNameInput, ForceResults:=tlForceFlow
                TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
            Next site
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
            
            TestNameInput = Report_TName_From_Instance("C", "X", "DCK" & Format(FirstLoop, "00") & "-MINStepDelta", CInt(i))
            For Each site In TheExec.sites.Active
                TheExec.Flow.TestLimit resultVal:=DeltaDCK_DSPWave.CalcMinimumValue(IndexOfMinimumValue), Tname:=TestNameInput, ForceResults:=tlForceFlow
                TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
            Next site
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
            
        End If
    Next i
End Function



Public Function P2PBundle_eye(argc As Integer, argv() As String) As Long

' Format:P2PBundle_eye([SweepLoopName,StartValue,TargetValue,DivideForEye,Site0@Site1&Site2@Site3])
' SweepLoopName : This StringName should be same with Split_Loop_DigSrc_Str(6)
' Site0@Site1&Site2@Site : Exchange data site0 & site1 , site2 & site3

    Dim site As Variant
    Dim EyeWidth As Integer
    Dim EyeWidthTemp As Integer
    Dim i, j, k, z, x As Integer
    Dim strTemp() As String
    Dim EyeDivide As String
    Dim TempString As String
    Dim SiteBundle() As String
    Dim SweepConterStr As String
    Dim DestinationSite As String
    Dim CounterByStart As String
    Dim CounterByTarget As String
    Dim CounterByWidth As String
    Dim DictCounterName As String
    Dim SiteBundleIndex() As String
    Dim Eye_information() As String
    Dim EyeStep() As String
    Dim DataSite() As New SiteVariant
    Dim Mdll_lock As New DSPWave
    Dim Mdll_lockvalue As New DSPWave
    Dim mdll_low() As New SiteDouble
    Dim mdll_high() As New SiteDouble
    ReDim mdll_low(0)
    ReDim mdll_high(0)
    Dim Eye_precent As Double
    
    Dim DSPWaveTemp As New DSPWave
    Dim New_DSPWave As New DSPWave
    Dim PrintEye() As New SiteLong
    Dim UnitCellrecord() As New SiteVariant
    Dim TestNameInput As String
    Dim TestNameInputeye As String
    Dim TnumRecord As Long
    
    Dim ConvertRXTXBundleString() As String
    Dim TempRXTXBundleName() As String
    Dim ConvertRXTXUnitCellString() As String
    Dim TempRXTXUnitCellString() As String
    
    
'----------------------------------Debug by Dylan--------------------------------------'
    Dim t As Integer
    Dim m, n As Integer
    Dim minmumValue As Integer
    Dim MaxmumValue As Integer
    Dim SplitRegister() As String
    
    Dim SweepRange() As String
    Dim BundleNum() As String
    Dim StrTempNumber() As String
    Dim RegNameSplit() As String
    Dim RegRange() As String
    Dim RegRangeValue() As String
    
    Dim LowLimitValue As New SiteDouble
    Dim HighLimitValue As New SiteDouble
    
    SplitRegister = Split("DDR2X:3-11,32-103,124-132|SDR:12-31,104-123", "|")
    ' DDR2X /8, /2
    ' SDR /4, /1
'---------------------------------------------------------------------------------------'
    

    For i = 0 To argc - 5
        Eye_information = Split(argv(0), "=")
        EyeStep = Split(Eye_information(1), "@")
        
        CounterByStart = EyeStep(0)
        CounterByTarget = EyeStep(1)
        CounterByWidth = EyeStep(2)
        EyeDivide = argv(1)
        argv(2) = Replace(argv(2), "]", vbNullString)
        strTemp = Split(argv(3), "+")
        DictCounterName = Replace(Eye_information(0), "[", vbNullString)
        
        If LCase(TheExec.CurrentChanMap) = LCase("ChannelMap_FT_4_site_2C") Then  ' 20201027 CT add for 2C FT
             argv(2) = "0@2&1@3"
        End If
        
        SiteBundleIndex = Split(argv(2), "&")
        ReDim DataSite((UBound(strTemp)))
        Public_GetStoredString (DictCounterName)
        SweepConterStr = gl_SpecialString
        New_DSPWave.CreateConstant 0, 2 * (UBound(strTemp) + 1), DspLong
        Mdll_lockvalue = GetStoredCaptureData(argv(4))
        'Mdll_lockvalue.ConvertDataTypeTo (DspLong)
        Mdll_lock.CreateConstant 0, 1, DspLong
        For Each site In TheExec.sites
            Mdll_lock(site) = Mdll_lockvalue(site).ConvertStreamTo(tldspParallel, Mdll_lockvalue.SampleSize, 0, Bit0IsMsb)
''''''''''            mdll_low(0)(Site) = Mdll_lock(Site).Element(0) / 4
''''''''''            mdll_high(0)(Site) = Mdll_lock(Site).Element(0) / 2
            mdll_low(0)(site) = Mdll_lock(site).Element(0)
            mdll_high(0)(site) = Mdll_lock(site).Element(0)
        Next site
        
        ReDim ConvertRXTXBundleString(UBound(strTemp)) As String
        ReDim ConvertRXTXUnitCellString(CounterByWidth - 1) As String
        
         ReDim PrintEye((UBound(strTemp) + 1) * CounterByWidth) As New SiteLong
         ReDim UnitCellrecord((UBound(strTemp) + 1) * CounterByWidth) As New SiteVariant
        For j = 0 To UBound(strTemp)
            DSPWaveTemp = GetStoredCaptureData(strTemp(j))
            If CLng(SweepConterStr) <> CLng(CounterByStart) Then
                DataSite(j) = GetStoredMeasurement(strTemp(j) & "_" & "AssemblyStr")
            End If
            For k = 0 To UBound(SiteBundleIndex)
                SiteBundle = Split(SiteBundleIndex(k), "@")
                For x = 0 To UBound(SiteBundle)
                    TempString = vbNullString
                    DestinationSite = SiteBundle(UBound(SiteBundle) - x)
                    
                    If TheExec.sites(DestinationSite).Active Then ' 20201027 CT add for 2C FT
                    For z = 0 To DSPWaveTemp(DestinationSite).SampleSize - 1
                        If z = 0 Then
                            TempString = CStr(DSPWaveTemp(DestinationSite).Element(0))
                        Else
                            TempString = CStr(DSPWaveTemp(DestinationSite).Element(z)) & TempString
                        End If
                    Next z
                    DataSite(j)(SiteBundle(x)) = TempString & DataSite(j)(SiteBundle(x))
                    End If
                    
                Next x
            Next k
            Call AddStoredMeasurement(strTemp(j) & "_" & "AssemblyStr", DataSite(j))
            If CLng(SweepConterStr) = CLng(CounterByTarget) Then
            
                Dim UnitCellString As String
                Dim SweepStep As Long
                SweepStep = CLng(CounterByTarget / EyeDivide)
                
                
                'ReDim PrintEye((UBound(StrTemp) + 1) * CounterByWidth) As New SiteLong
                
                For Each site In TheExec.sites
                    For z = 1 To CounterByWidth     '6= Unitcell Number
                        UnitCellString = vbNullString
                        For k = 0 To SweepStep
                        
                          If k = 0 Then
                            UnitCellString = mid(DataSite(j)(site), z + k * 6, 1)
                          Else
                            UnitCellString = UnitCellString & mid(DataSite(j)(site), z + k * 6, 1)
                    
                          End If
                        Next k
                    
                        EyeWidth = 0
                        EyeWidthTemp = 0
                    
                        For k = 0 To Len(UnitCellString)
                            If mid(UnitCellString, k + 1, 1) = "1" Then
                                EyeWidthTemp = EyeWidthTemp + 1
                            ElseIf k = Len(UnitCellString) And EyeWidthTemp > EyeWidth Then
                                    EyeWidth = EyeWidthTemp
                                    EyeWidthTemp = 0
                            Else
                                If EyeWidthTemp > EyeWidth Then
                                    EyeWidth = EyeWidthTemp
                                    EyeWidthTemp = 0
                                Else
                                    EyeWidthTemp = 0
                                End If
                            End If
                        Next k
                      
'                       Dim PrintEye() As SiteLong
'                       ReDim PrintEye(SweepStep * CounterByWidth)
                      
                      PrintEye((z - 1) + j * CounterByWidth)(site) = EyeWidth
                      UnitCellrecord((z - 1) + j * CounterByWidth)(site) = UnitCellString
                      
'                      If EyeDivide <> "" Then
'                            EyeWidthTemp = CStr(FormatNumber((EyeWidth * EyeDivide), 0))
'                            TheExec.Datalog.WriteComment "Site" & CStr(Site) & " , " & "UnitCell" & z - 1 & "_" & "EyeWidth : " & EyeWidthTemp
'                      Else
'                            TheExec.Datalog.WriteComment "Site" & CStr(Site) & " , " & "UnitCell" & z - 1 & "_" & "EyeWidth : " & EyeWidth
'                      End If
'                            TheExec.Datalog.WriteComment "Site" & CStr(Site) & " , " & "UnitCell" & z - 1 & "_" & CStr(StrTemp(j)) & " : " & UnitCellString
                    Next z
                
                Next site
           End If
        Next j
        
           
         If CLng(SweepConterStr) = CLng(CounterByTarget) Then
            TheExec.Datalog.WriteComment "=========================== Count EYE=========================== "
              For Each site In TheExec.sites
              
                   For k = 0 To UBound(strTemp)
                   'SplitByAt = Split(argv(i), "@")
                   
                   'TnumRecord = TheExec.sites.item(Site).TestNumber
                  
                       For z = 0 To CounterByWidth - 1
                        TestNameInput = "UnitCell" & z & "_" & CStr(strTemp(k)) & "EYE"
                        TestNameInputeye = "UnitCell" & z & "_" & CStr(strTemp(k)) & "EYE-percent"
                        
                        TestNameInput = Report_TName_From_Instance("X", "x", TestNameInput, x, 0)
                        
                   
                       
                        TheExec.Flow.TestLimit lowVal:=mdll_low(0), hiVal:=mdll_high(0), resultVal:=PrintEye(z + k * CounterByWidth) * EyeDivide, formatStr:="%i", Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling 'eng_forceflow_transfer
                        
                        
                        TestNameInputeye = Report_TName_From_Instance("calc", "x", TestNameInputeye, x, 0)
                        
                        Eye_precent = (PrintEye(z + k * CounterByWidth) * EyeDivide) / Mdll_lock(site).Element(0)
                        
                        TheExec.Flow.TestLimit lowVal:=25, hiVal:=50, resultVal:=Format(Eye_precent * 100, 0), formatStr:="%i", Tname:=TestNameInputeye, ForceResults:=tlForceNone, scaletype:=scaleNoScaling 'eng_forceflow_transfer
                        
                        
                        
                        
                        'TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
                          'theexec.sites.item(Site).TestNumber = theexec.sites.item(Site).TestNumber + 1
                       Next z
                   'TheExec.sites.item(Site).TestNumber = TheExec.sites.item(Site).TestNumber + 1
                   Next k
           
               Next site
               
               
               For Each site In TheExec.sites
                 For k = 0 To UBound(strTemp)
                    For z = 0 To CounterByWidth - 1
                      If EyeDivide <> "" Then
                            
                            TheExec.Datalog.WriteComment "Site" & CStr(site) & " , " & "UnitCell" & z - 1 & "_" & "EyeWidth : " & PrintEye(z + k * CounterByWidth) * EyeDivide
                      Else
                            TheExec.Datalog.WriteComment "Site" & CStr(site) & " , " & "UnitCell" & z - 1 & "_" & "EyeWidth : " & PrintEye(z + k * CounterByWidth)
                      End If
                      TheExec.Datalog.WriteComment "Site" & CStr(site) & " , " & "UnitCell" & z - 1 & "_" & CStr(strTemp(k)) & " : " & CStr(UnitCellrecord(z + k * CounterByWidth))
                    Next z
                 Next k
              Next site
               
               
               
        End If
    Next i
End Function
        

Public Function P2PBundle_Unflipeye(argc As Integer, argv() As String) As Long

' Format:P2PBundle_eye([SweepLoopName,StartValue,TargetValue,DivideForEye,Site0@Site1&Site2@Site3])
' SweepLoopName : This StringName should be same with Split_Loop_DigSrc_Str(6)
' Site0@Site1&Site2@Site : Exchange data site0 & site1 , site2 & site3

    Dim site As Variant
    Dim EyeWidth As Integer
    Dim EyeWidthTemp As Integer
    Dim i, j, k, z, x As Integer
    Dim strTemp() As String
    Dim EyeDivide As String
    Dim TempString As String
    Dim SiteBundle() As String
    Dim SweepConterStr As String
    Dim DestinationSite As String
    Dim CounterByStart As String
    Dim CounterByTarget As String
    Dim CounterByWidth As String
    Dim DictCounterName As String
    Dim SiteBundleIndex() As String
    Dim Eye_information() As String
    Dim EyeStep() As String
    Dim DataSite() As New SiteVariant
    Dim Mdll_lock As New DSPWave
    Dim Mdll_lockvalue As New DSPWave
    Dim mdll_low() As New SiteDouble
    Dim mdll_high() As New SiteDouble
    ReDim mdll_low(0)
    ReDim mdll_high(0)
    Dim Eye_precent As Double
    
    Dim DSPWaveTemp As New DSPWave
    Dim New_DSPWave As New DSPWave
    Dim PrintEye() As New SiteLong
    Dim UnitCellrecord() As New SiteVariant
    Dim TestNameInput As String
    Dim TestNameInputeye As String
    Dim TnumRecord As Long
    
    'Alg::P2PBundle_unflipeye([Loop_DigSrc=64@804@10,1,0@1],Bundle+++++,D2D_CMN__MDLL_LOCK_CODE_mdll_dcode_lock_NV)

    For i = 0 To argc - 5
        Eye_information = Split(argv(0), "=")
        EyeStep = Split(Eye_information(1), "@")
        
        CounterByStart = EyeStep(0)
        CounterByTarget = EyeStep(1)
        CounterByWidth = EyeStep(2)
        EyeDivide = argv(1)
        argv(2) = Replace(argv(2), "]", vbNullString)
        strTemp = Split(argv(3), "+")
        DictCounterName = Replace(Eye_information(0), "[", vbNullString)
        SiteBundleIndex = Split(argv(2), "&")
        ReDim DataSite((UBound(strTemp)))
        Public_GetStoredString (DictCounterName)
        SweepConterStr = gl_SpecialString
        New_DSPWave.CreateConstant 0, 2 * (UBound(strTemp) + 1), DspLong
        Mdll_lockvalue = GetStoredCaptureData(argv(4))
        'Mdll_lockvalue.ConvertDataTypeTo (DspLong)
        Mdll_lock.CreateConstant 0, 1, DspLong
        For Each site In TheExec.sites
            Mdll_lock(site) = Mdll_lockvalue(site).ConvertStreamTo(tldspParallel, Mdll_lockvalue.SampleSize, 0, Bit0IsMsb)
            mdll_low(0)(site) = Mdll_lock(site).Element(0) / 4
            mdll_high(0)(site) = Mdll_lock(site).Element(0) / 2
        Next site
        
        
        
        
         ReDim PrintEye((UBound(strTemp) + 1) * CounterByWidth) As New SiteLong
         ReDim UnitCellrecord((UBound(strTemp) + 1) * CounterByWidth) As New SiteVariant
        For j = 0 To UBound(strTemp)
            DSPWaveTemp = GetStoredCaptureData(strTemp(j))
            If CLng(SweepConterStr) <> CLng(CounterByStart) Then
                DataSite(j) = GetStoredMeasurement(strTemp(j) & "_" & "AssemblyStr")
            End If
            For k = 0 To UBound(SiteBundleIndex)
                SiteBundle = Split(SiteBundleIndex(k), "@")
               ''''' For x = 0 To UBound(SiteBundle)
                 For Each site In TheExec.sites
                    TempString = vbNullString
                    '''''DestinationSite = SiteBundle(x)  'SiteBundle(UBound(SiteBundle) - x)  change for no flip site
                    
                    
                    For z = 0 To DSPWaveTemp(site).SampleSize - 1
                        If z = 0 Then
                            TempString = CStr(DSPWaveTemp(site).Element(0))
                        Else
                            TempString = TempString & CStr(DSPWaveTemp(site).Element(z)) 'CStr(DSPWaveTemp(DestinationSite).Element(z)) & TempString  change for no unit cell flip
                        End If
                    Next z
                    DataSite(j)(site) = TempString & DataSite(j)(site)
                Next site
            Next k
            Call AddStoredMeasurement(strTemp(j) & "_" & "AssemblyStr", DataSite(j))
            If CLng(SweepConterStr) = CLng(CounterByTarget) Then
            
                Dim UnitCellString As String
                Dim SweepStep As Long
                SweepStep = CLng(CounterByTarget / EyeDivide)
                
                
                'ReDim PrintEye((UBound(StrTemp) + 1) * CounterByWidth) As New SiteLong
                
                For Each site In TheExec.sites
                    For z = 1 To CounterByWidth     '6= Unitcell Number
                        UnitCellString = vbNullString
                        For k = 0 To SweepStep
                        
                          If k = 0 Then
                            UnitCellString = mid(DataSite(j)(site), z + k * 6, 1)
                          Else
                            UnitCellString = UnitCellString & mid(DataSite(j)(site), z + k * 6, 1)
                    
                          End If
                        Next k
                    
                        EyeWidth = 0
                        EyeWidthTemp = 0
                    
                        For k = 0 To Len(UnitCellString)
                            If mid(UnitCellString, k + 1, 1) = "1" Then
                                EyeWidthTemp = EyeWidthTemp + 1
                            ElseIf k = Len(UnitCellString) And EyeWidthTemp > EyeWidth Then
                                    EyeWidth = EyeWidthTemp
                                    EyeWidthTemp = 0
                            Else
                                If EyeWidthTemp > EyeWidth Then
                                    EyeWidth = EyeWidthTemp
                                    EyeWidthTemp = 0
                                Else
                                    EyeWidthTemp = 0
                                End If
                            End If
                        Next k
                      
'                       Dim PrintEye() As SiteLong
'                       ReDim PrintEye(SweepStep * CounterByWidth)
                      
                      PrintEye((z - 1) + j * CounterByWidth)(site) = EyeWidth
                      UnitCellrecord((z - 1) + j * CounterByWidth)(site) = UnitCellString
                      
'                      If EyeDivide <> "" Then
'                            EyeWidthTemp = CStr(FormatNumber((EyeWidth * EyeDivide), 0))
'                            TheExec.Datalog.WriteComment "Site" & CStr(Site) & " , " & "UnitCell" & z - 1 & "_" & "EyeWidth : " & EyeWidthTemp
'                      Else
'                            TheExec.Datalog.WriteComment "Site" & CStr(Site) & " , " & "UnitCell" & z - 1 & "_" & "EyeWidth : " & EyeWidth
'                      End If
'                            TheExec.Datalog.WriteComment "Site" & CStr(Site) & " , " & "UnitCell" & z - 1 & "_" & CStr(StrTemp(j)) & " : " & UnitCellString
                    Next z
                
                Next site
           End If
        Next j
        
           
         If CLng(SweepConterStr) = CLng(CounterByTarget) Then
            TheExec.Datalog.WriteComment "=========================== Count EYE=========================== "
              For Each site In TheExec.sites
              
                   For k = 0 To UBound(strTemp)
                   'SplitByAt = Split(argv(i), "@")
                   
                   'TnumRecord = TheExec.sites.item(Site).TestNumber
                  
                       For z = 0 To CounterByWidth - 1
                        TestNameInput = "UnitCell" & z & "_" & CStr(strTemp(k)) & "EYE"
                        TestNameInputeye = "UnitCell" & z & "_" & CStr(strTemp(k)) & "EYE-percent"
                        
                        TestNameInput = Report_TName_From_Instance("X", "x", TestNameInput, x, 0)
                        
                   
                       
                        'TheExec.Flow.TestLimit LowVal:=mdll_low(0), HiVal:=mdll_high(0), resultVal:=PrintEye(z + k * CounterByWidth) * EyeDivide, FormatStr:="%i", TName:=TestNameInput, ForceResults:=tlForceFlow, ScaleType:=scaleNoScaling
                        TheExec.Flow.TestLimit lowVal:=mdll_low(0), hiVal:=mdll_high(0), resultVal:=PrintEye(z + k * CounterByWidth) * EyeDivide, formatStr:="%i", Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling 'eng_forceflow_transfer
                        
                        TestNameInputeye = Report_TName_From_Instance("calc", "x", TestNameInputeye, x, 0)
                        
                        Eye_precent = (PrintEye(z + k * CounterByWidth) * EyeDivide) / Mdll_lock(site).Element(0)
                        
                        TheExec.Flow.TestLimit lowVal:=25, hiVal:=50, resultVal:=Format(Eye_precent * 100, 0), formatStr:="%i", Tname:=TestNameInputeye, ForceResults:=tlForceNone, scaletype:=scaleNoScaling 'eng_forceflow_transfer
                        
                        
                        
                        
                        'TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
                          'theexec.sites.item(Site).TestNumber = theexec.sites.item(Site).TestNumber + 1
                       Next z
                   'TheExec.sites.item(Site).TestNumber = TheExec.sites.item(Site).TestNumber + 1
                   Next k
           
               Next site
               
               
               For Each site In TheExec.sites
                 For k = 0 To UBound(strTemp)
                    For z = 0 To CounterByWidth - 1
                      If EyeDivide <> "" Then
                            
                            TheExec.Datalog.WriteComment "Site" & CStr(site) & " , " & "UnitCell" & z & "_" & "EyeWidth : " & PrintEye(z + k * CounterByWidth) * EyeDivide
                      Else
                            TheExec.Datalog.WriteComment "Site" & CStr(site) & " , " & "UnitCell" & z & "_" & "EyeWidth : " & PrintEye(z + k * CounterByWidth)
                      End If
                      TheExec.Datalog.WriteComment "Site" & CStr(site) & " , " & "UnitCell" & z & "_" & CStr(strTemp(k)) & " : " & CStr(UnitCellrecord(z + k * CounterByWidth))
                    Next z
                 Next k
              Next site
               
               
               
        End If
    Next i
End Function

Public Function LB_Error_Count(argc As Integer, argv() As String) As Long
    
    Dim TestNameInput As String
    Dim site As Variant
    Dim LPK_result As New SiteDouble
    Dim ERR_CNT As New DSPWave
    Dim ParallelStream As New DSPWave
    Dim ErrorStr As String
    Dim i As Long
    
    ERR_CNT = GetStoredCaptureData(argv(0))
    For Each site In TheExec.sites
        LPK_result = ERR_CNT.Element(ERR_CNT.SampleSize - 1)
        ParallelStream = ERR_CNT.ConvertStreamTo(tldspParallel, 4, 0, Bit0IsMsb)
    Next site
    
    TestNameInput = Report_TName_From_Instance("Calc", "ERR_CNT", , 0, 0, , , , tlForceFlow)
    TheExec.Flow.TestLimit LPK_result, 1, 1, , , , , "0.0f", Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer

    For Each site In TheExec.sites
        ErrorStr = vbNullString
        For i = (ParallelStream.SampleSize - 1) To 0 Step -1
            ErrorStr = ErrorStr & Hex(ParallelStream.Element(i))
        Next i
        TheExec.Datalog.WriteComment "Site " & site & ":Error Count on (" & TheExec.DataManager.instancename & "),Hex =" & ErrorStr
    Next site
    
End Function

Public Function Calc_EQcal_DigSrc(argc As Long, argv() As String)
    
    Dim eqc_ctrl As New DSPWave
    Dim eqc_ctrl_Dec As New DSPWave
    Dim EQC As New SiteDouble
    Dim EQC_temp As New SiteLong
    Dim eqclen As New SiteLong
    Dim EQCWave As New DSPWave
    Dim site As Variant
    Dim i As Integer
    Dim counter As Integer
    Dim EQCWave_Dec As New SiteLong
    For Each site In TheExec.sites.Active
        'For i = 0 To argc - 2
            counter = 0
            eqc_ctrl = GetStoredCaptureData(argv(0))
            eqclen = eqc_ctrl.SampleSize
            For i = 0 To eqclen - 1
                If eqc_ctrl.Element(i) = 1 Then counter = counter + 1
            Next i
            eqc_ctrl_Dec.CreateConstant 0, 1, DspLong
            Call HardIP_Bin2Dec(eqc_ctrl_Dec, eqc_ctrl)
            'EQC = (10 * eqc_ctrl_Dec.Element(0) - 18) / 11
            EQC = Int((10 * counter - 18) / 11)
            If EQC > (10 * counter - 18) / 11 Then
            Else
                EQC = EQC + 1
            End If
            
        EQCWave.CreateConstant 0, 18
        EQCWave_Dec = 2 ^ (EQC) - 1
        EQC_temp = EQCWave_Dec

        For i = 0 To EQCWave.SampleSize - 1
            If EQCWave_Dec > 0 Then
                EQCWave.Element(i) = EQCWave_Dec Mod 2
                EQCWave_Dec = Int(EQCWave_Dec / 2)
            Else
                EQCWave.Element(i) = 0
            End If
        Next i
        AddStoredCaptureData argv(1), EQCWave
    
        
    Next site
    
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    Dim TestName As String

'''''    If gl_UseStandardTestName_Flag = True Then                     'Roger add
'''''        Call Report_ALG_TName_From_Instance(OutputTname_format, "C", "X", gl_Tname_Alg, 1)
'''''        'OutputTname_format(6) = OutputTname_format(6) & "MDLLUnique"
'''''        If gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex) <> "" Then
'''''                OutputTname_format(6) = gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex)
'''''                'OutputTname_format(9) = 64
'''''            End If
'''''        TestNameInput = Merge_TName(OutputTname_format)
'''''    Else
'''''        TestNameInput = TestName & "EQCalWave"
'''''    End If
'''''
    'TheExec.Flow.TestLimit resultVal:=EQC_temp, Tname:=TestNameInput, ForceResults:=tlForceFlow
   
   
   
   'i_min = GetStoredCaptureData("aus16pll_fcal_bypass_code")
   
   
    TestNameInput = Report_TName_From_Instance(CalcC, "X", argv(1) & "_EQCalWave", CInt(i))
    
    TheExec.Flow.TestLimit resultVal:=EQC_temp, Tname:=TestNameInput, ForceResults:=tlForceFlow

End Function
Public Function Calc_DCSkew_IN_CLK(argc As Integer, argv() As String) As Long
    'Calc_DCSkew_IN_CLK(CNT_100_0,CNT_INPUT)
    Dim CNT_100_0 As New DSPWave: CNT_100_0 = GetStoredCaptureData(argv(0))
    Dim CNT_INPUT As New DSPWave: CNT_INPUT = GetStoredCaptureData(argv(1))
    Dim CNT_100_0_DSP As New DSPWave
    Dim CNT_INPUT_DSP As New DSPWave
    Dim Cnt_100_Val As New SiteDouble
    Dim Cnt_IN_Val As New SiteDouble
    
    Dim MSB_sign_Init As New SiteBoolean
    Dim MSB_sign_Ref As New SiteBoolean
    Dim MSB_sign_In As New SiteBoolean
    Dim i As Variant
    Dim DC_Skew_InClk As New SiteDouble
    Dim DC_InClk As New SiteDouble
    Dim TestNameInput As String
    
    MSB_sign_Init = True
    CNT_100_0_DSP.CreateConstant 0, 1, DspDouble
    CNT_INPUT_DSP.CreateConstant 0, 1, DspDouble
    
    Call HardIP_Bin2Dec(CNT_100_0_DSP, CNT_100_0)
    Call HardIP_Bin2Dec(CNT_INPUT_DSP, CNT_INPUT)
    For Each i In TheExec.sites
        Cnt_100_Val = CNT_100_0_DSP.Element(0)
        Cnt_IN_Val = CNT_INPUT_DSP.Element(0)
    Next i
    
    MSB_sign_Ref = Cnt_100_Val.compare(GreaterThan, 4096)
    If MSB_sign_Ref.Any(True) Then
        TheExec.sites.Selected = MSB_sign_Ref
        Cnt_100_Val = Cnt_100_Val.Subtract(4096).Multiply(-1)
        TheExec.sites.Selected = MSB_sign_Init
    End If
    
    MSB_sign_In = Cnt_IN_Val.compare(GreaterThan, 4096)
    If MSB_sign_In.Any(True) Then
        TheExec.sites.Selected = MSB_sign_In
        Cnt_IN_Val = Cnt_IN_Val.Subtract(4096).Multiply(-1)
        TheExec.sites.Selected = MSB_sign_Init
    End If
    
    DC_Skew_InClk = Cnt_IN_Val.divide(Cnt_100_Val).Multiply(0.5) '.Multiply(IIf(MSB_sign, -1, 1))
    DC_InClk = DC_Skew_InClk.Add(0.5)
    
    TestNameInput = Report_TName_From_Instance(CalcC, "DC_IN_CLK")
    TheExec.Flow.TestLimit resultVal:=DC_InClk, Tname:=TestNameInput, ForceResults:=tlForceFlow

End Function

Public Function Functional_Parametric_LoopFunction(argc As Integer, argv() As String) As Long

    Dim site As Variant
    Dim IndexName As String
    Dim PassFailEachSite As New SiteLong
    Dim PassFailBoolean As New SiteBoolean
    
    IndexName = CStr(argv(0))
    PassFailEachSite = TheExec.Flow.LastFlowStepResult
    
    For Each site In TheExec.sites.Active
        If PassFailEachSite = 1 Then                     ' 1 = tlResultPass
            PassFailBoolean = True
        ElseIf PassFailEachSite = 0 Then                 ' 0 = tlResultFail
            PassFailBoolean = False
        End If
    Next site
    
    If PassFailBoolean.All(True) Then
        TheExec.Flow.var(IndexName).value = 100
        For Each site In TheExec.sites
            TheExec.sites.item(site).FlagState(argv(1)) = logicFalse
        Next site
        
    End If
    
    
End Function

Public Function Calc_RAW_DCO_DAC(argc As Integer, argv() As String) As Long
     
    Dim DAC_result() As New DSPWave
    Dim DAC_3_0() As New DSPWave
    Dim DAC_5_4() As New DSPWave
    Dim DicName_3_0 As String
    Dim DicName_5_4 As String
    Dim site As Variant
    Dim DAC_temp As New DSPWave
    Dim i As Long
    Dim TestNameInput As String
    
    ReDim DAC_result(argc - 1)
    ReDim DAC_3_0(argc - 1)
    ReDim DAC_5_4(argc - 1)
    
    For i = 0 To argc - 1
        Set DAC_temp = Nothing
        DicName_5_4 = Split(argv(i), "&")(1)
        DicName_3_0 = Split(argv(i), "&")(0)
        DAC_5_4(i) = GetStoredCaptureData(DicName_5_4)
        DAC_3_0(i) = GetStoredCaptureData(DicName_3_0)
        For Each site In TheExec.sites
            DAC_temp = DAC_3_0(i).Concatenate(DAC_5_4(i))
            DAC_result(i) = DAC_temp.ConvertStreamTo(tldspParallel, 6, 0, Bit0IsMsb)
        Next site
        
        TestNameInput = Report_TName_From_Instance(CalcC, vbNullString, , , i)
        TheExec.Flow.TestLimit resultVal:=DAC_result(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i
End Function

Public Function Calc_GRPPercentageF(argc As Integer, argv() As String) As Long
    'Update from Staten @William 211101
    'Merge Calc_GRPPercentageF1/Calc_GRPPercentageF2 function -- 20220214
    'Alg::Calc_GRPPercentageF(DDR0_CATXMDLL_CODE_HV_F2&DDR0_seqreg32_rddata,DDR1_CATXMDLL_CODE_HV_F2&DDR1_seqreg32_rddata,DDR2_CATXMDLL_CODE_HV_F2&DDR2_seqreg32_rddata,Ratio=2.1)
    
    Dim i As Integer
    Dim site As Variant
    Dim SplitStrAry() As String
    Dim DSP_LoLimit As New DSPWave
    Dim ratio As Double  ' New Add for define ratio from argument -- 20220214
    
'   Dim DSP_Temp As New DSPWave
'   Dim DSP_Summary As New DSPWave
    Dim DSP_EyeWidth As New DSPWave
    Dim DSP_Percentage As New DSPWave
    Dim TestNameInput As String
    Dim Percentage_bysite() As New SiteDouble
    
    Dim DSP_Temp() As New DSPWave
'   Dim DSP_Summary() As New DSPWave
    Dim DSP_Summary() As New SiteDouble 'Upadte variable type due to new "Calc_MDLL_Monotonicity_DevideBlock_SEGTTR" @CW 211004 by Carter
    ReDim DSP_Temp(argc - 2)            ' org : argc-1
    ReDim DSP_Summary(argc - 2)         ' org : argc-1
    ReDim Percentage_bysite(argc - 2)   ' org : argc-1
    
    'New Add Format : Ratio=2.1   -- 20220214
    If UCase(argv(argc - 1)) Like "*RATIO=*" Then
        SplitStrAry = Split(argv(argc - 1), "=")
        ratio = CDbl(SplitStrAry(1))
    Else
        'Judge run time error for Invalid Ration define
        TheExec.Datalog.WriteComment ("Error! Invalid Ratio Defined!!")
        TheExec.Flow.TestLimit 9999, 0, 1, , , , , , PinName:="X", Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
        Exit Function
    End If
    
    For i = 0 To argc - 2   ' org : argc-1
        SplitStrAry = Split(argv(i), "&")
'       DSP_Summary(i) = GetStoredCaptureData("Summary" & SplitStrAry(0))
        DSP_Summary(i) = GetStoredData("Summary" & SplitStrAry(0)) 'Upadte variable type due to new "Calc_MDLL_Monotonicity_DevideBlock_SEGTTR" @CW 211004 by Carter
        DSP_Temp(i) = GetStoredCaptureData(SplitStrAry(1))
       
        For Each site In TheExec.sites
            DSP_EyeWidth = DSP_Temp(i).ConvertStreamTo(tldspParallel, DSP_Temp(i).SampleSize, 0, Bit0IsMsb)
'           DSP_Percentage = DSP_EyeWidth.Multiply(100).Divide(DSP_Summary(i).Multiply(2.1))
'           DSP_Percentage = DSP_EyeWidth.Multiply(100).Divide(DSP_Summary(i).Multiply(2.3))
'           DSP_Percentage = DSP_EyeWidth.Multiply(100).Divide(DSP_Summary(i).Multiply(2))
            DSP_Percentage = DSP_EyeWidth.Multiply(100).divide(DSP_Summary(i).Multiply(ratio))      ' New Add for define ratio from argument -- 20220214
            Percentage_bysite(i) = DSP_Percentage.Element(0)
        Next site
        
        '@210707 CW TTR eye_width
        '--------------------------
'        TestNameInput = Report_TName_From_Instance("CalcC", "", "", i)
'        For Each site In TheExec.sites
'            'theexec.Flow.TestLimit DSP_EyeWidth.Element(0), DSP_Summary(i).Element(0) / 2, DSP_Summary(i).Element(0) * 2.1, , , , , , Tname:=TestNameInput, ForceResults:=tlForceFlow
'            'theexec.Flow.TestLimit DSP_EyeWidth.Element(0), DSP_Summary(i).Element(0) / 2, DSP_Summary(i).Element(0) * 2.3, , , , , , Tname:=TestNameInput, ForceResults:=tlForceFlow
'
'            If TheExec.Flow.EnableWord("AMPLP5_BinCut_Enable_Flag") = True Then
'                TheExec.Flow.TestLimit DSP_EyeWidth.Element(0), , , , , , , , PinName:="X", Tname:=TestNameInput, ForceResults:=tlForceFlow
'            Else
'                TheExec.Flow.TestLimit DSP_EyeWidth.Element(0), DSP_Summary(i) / 2, DSP_Summary(i) * Ratio, , , , , , PinName:="X", Tname:=TestNameInput, ForceResults:=tlForceFlow
'            End If
'
'
'            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
'        Next site
'        TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
        '------------------------------
    Next i
    
    For i = 0 To argc - 2   ' org : argc-1
        SplitStrAry = Split(argv(i), "&")
        'TestNameInput = Report_TName_From_Instance("C", SplitStrAry(1), "Percentage" & i, i, , , , , tlForceFlow)
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, vbNullString, i)
        'TheExec.Flow.TestLimit Percentage_bysite(i), 25, 120, , , , , , PinName:="X", Tname:=TestNameInput, ForceResults:=tlForceFlow
        TheExec.Flow.TestLimit Percentage_bysite(i), , , , , , , , PinName:="X", Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i
    
End Function

Public Function P2PBundle_Unflipeye_NewCal(argc As Integer, argv() As String) As Long

' Format:P2PBundle_eye([SweepLoopName,StartValue,TargetValue,DivideForEye,Site0@Site1&Site2@Site3])
' SweepLoopName : This StringName should be same with Split_Loop_DigSrc_Str(6)
' Site0@Site1&Site2@Site : Exchange data site0 & site1 , site2 & site3

    Dim site As Variant
    Dim EyeWidth As Integer
    Dim EyeWidthTemp As Integer
    Dim i, j, k, z, x As Integer
    Dim strTemp() As String
    Dim EyeDivide As String
    Dim TempString As String
    Dim SiteBundle() As String
    Dim SweepConterStr As String
    Dim DestinationSite As String
    Dim CounterByStart As String
    Dim CounterByTarget As String
    Dim CounterByWidth As String
    Dim DictCounterName As String
    Dim SiteBundleIndex() As String
    Dim Eye_information() As String
    Dim EyeStep() As String
    Dim DataSite() As New SiteVariant
    Dim Mdll_lock As New DSPWave
    Dim Mdll_lockvalue As New DSPWave
    Dim mdll_low() As New SiteDouble
    Dim mdll_high() As New SiteDouble
    ReDim mdll_low(0)
    ReDim mdll_high(0)
    Dim Eye_precent As Double
    
    Dim DSPWaveTemp As New DSPWave
    Dim New_DSPWave As New DSPWave
    Dim PrintEye() As New SiteLong
    Dim UnitCellrecord() As New SiteVariant
    Dim TestNameInput As String
    Dim TestNameInputeye As String
    Dim TnumRecord As Long
    
    Dim SplitRegister() As String
    
    
    
'----------------------------------Debug by Dylan--------------------------------------'
    Dim t As Integer
    Dim m, n As Integer
    Dim minmumValue As Integer
    Dim MaxmumValue As Integer
    
    Dim SweepRange() As String
    Dim BundleNum() As String
    Dim StrTempNumber() As String
    Dim RegNameSplit() As String
    Dim RegRange() As String
    Dim RegRangeValue() As String
    
    Dim LowLimitValue As New SiteDouble
    Dim HighLimitValue As New SiteDouble
    
    SplitRegister = Split("DDR2X:3-11,32-103,124-132|SDR:12-31,104-123", "|")
    ' SplitRegister = Split("DDR2X:3-11,32-103,124-132|SDR:12-31,104-123", "|")
    ' DDR2X /8, /2
    ' SDR /4, /1
'---------------------------------------------------------------------------------------'
    
    
    
    
    
    'Alg::P2PBundle_unflipeye([Loop_DigSrc=64@804@10,1,0@1],Bundle+++++,D2D_CMN__MDLL_LOCK_CODE_mdll_dcode_lock_NV)

    For i = 0 To argc - 5
        Eye_information = Split(argv(0), "=")
        EyeStep = Split(Eye_information(1), "@")
        
        CounterByStart = EyeStep(0)
        CounterByTarget = EyeStep(1)
        CounterByWidth = EyeStep(2)
        EyeDivide = argv(1)
        argv(2) = Replace(argv(2), "]", vbNullString)
        strTemp = Split(argv(3), "+")
        DictCounterName = Replace(Eye_information(0), "[", vbNullString)
        SiteBundleIndex = Split(argv(2), "&")
        ReDim DataSite((UBound(strTemp)))
        Public_GetStoredString (DictCounterName)
        SweepConterStr = gl_SpecialString
        New_DSPWave.CreateConstant 0, 2 * (UBound(strTemp) + 1), DspLong
        Mdll_lockvalue = GetStoredCaptureData(argv(4))
        'Mdll_lockvalue.ConvertDataTypeTo (DspLong)
        Mdll_lock.CreateConstant 0, 1, DspLong
        For Each site In TheExec.sites
            Mdll_lock(site) = Mdll_lockvalue(site).ConvertStreamTo(tldspParallel, Mdll_lockvalue.SampleSize, 0, Bit0IsMsb)
''''''''''            mdll_low(0)(site) = Mdll_lock(site).Element(0) / 8
''''''''''            mdll_high(0)(site) = Mdll_lock(site).Element(0) / 2
            mdll_low(0)(site) = Mdll_lock(site).Element(0)
            mdll_high(0)(site) = Mdll_lock(site).Element(0)
        Next site
        
        
        
        
         ReDim PrintEye((UBound(strTemp) + 1) * CounterByWidth) As New SiteLong
         ReDim UnitCellrecord((UBound(strTemp) + 1) * CounterByWidth) As New SiteVariant
        For j = 0 To UBound(strTemp)
            DSPWaveTemp = GetStoredCaptureData(strTemp(j))
            If CLng(SweepConterStr) <> CLng(CounterByStart) Then
                DataSite(j) = GetStoredMeasurement(strTemp(j) & "_" & "AssemblyStr")
            End If
            For k = 0 To UBound(SiteBundleIndex)
                SiteBundle = Split(SiteBundleIndex(k), "@")
               ''''' For x = 0 To UBound(SiteBundle)
                 For Each site In TheExec.sites
                    TempString = vbNullString
                    '''''DestinationSite = SiteBundle(x)  'SiteBundle(UBound(SiteBundle) - x)  change for no flip site
                    
                    
                    For z = 0 To DSPWaveTemp(site).SampleSize - 1
                        If z = 0 Then
                            TempString = CStr(DSPWaveTemp(site).Element(0))
                        Else
                            TempString = TempString & CStr(DSPWaveTemp(site).Element(z)) 'CStr(DSPWaveTemp(DestinationSite).Element(z)) & TempString  change for no unit cell flip
                        End If
                    Next z
                    DataSite(j)(site) = TempString & DataSite(j)(site)
                Next site
            Next k
            Call AddStoredMeasurement(strTemp(j) & "_" & "AssemblyStr", DataSite(j))
            If CLng(SweepConterStr) = CLng(CounterByTarget) Then
            
                Dim UnitCellString As String
                Dim SweepStep As Long
                SweepStep = CLng(CounterByTarget / EyeDivide)
                
                
                'ReDim PrintEye((UBound(StrTemp) + 1) * CounterByWidth) As New SiteLong
                
                For Each site In TheExec.sites
                    For z = 1 To CounterByWidth     '6= Unitcell Number
                        UnitCellString = vbNullString
                        For k = 0 To SweepStep
                        
                          If k = 0 Then
                            UnitCellString = mid(DataSite(j)(site), z + k * 6, 1)
                          Else
                            UnitCellString = UnitCellString & mid(DataSite(j)(site), z + k * 6, 1)
                    
                          End If
                        Next k
                    
                        EyeWidth = 0
                        EyeWidthTemp = 0
                    
                        For k = 0 To Len(UnitCellString)
                            If mid(UnitCellString, k + 1, 1) = "1" Then
                                EyeWidthTemp = EyeWidthTemp + 1
                            ElseIf k = Len(UnitCellString) And EyeWidthTemp > EyeWidth Then
                                    EyeWidth = EyeWidthTemp
                                    EyeWidthTemp = 0
                            Else
                                If EyeWidthTemp > EyeWidth Then
                                    EyeWidth = EyeWidthTemp
                                    EyeWidthTemp = 0
                                Else
                                    EyeWidthTemp = 0
                                End If
                            End If
                        Next k
                      
'                       Dim PrintEye() As SiteLong
'                       ReDim PrintEye(SweepStep * CounterByWidth)
                      
                      PrintEye((z - 1) + j * CounterByWidth)(site) = EyeWidth
                      UnitCellrecord((z - 1) + j * CounterByWidth)(site) = UnitCellString
                      
'                      If EyeDivide <> "" Then
'                            EyeWidthTemp = CStr(FormatNumber((EyeWidth * EyeDivide), 0))
'                            TheExec.Datalog.WriteComment "Site" & CStr(Site) & " , " & "UnitCell" & z - 1 & "_" & "EyeWidth : " & EyeWidthTemp
'                      Else
'                            TheExec.Datalog.WriteComment "Site" & CStr(Site) & " , " & "UnitCell" & z - 1 & "_" & "EyeWidth : " & EyeWidth
'                      End If
'                            TheExec.Datalog.WriteComment "Site" & CStr(Site) & " , " & "UnitCell" & z - 1 & "_" & CStr(StrTemp(j)) & " : " & UnitCellString
                    Next z
                
                Next site
           End If
        Next j
        
           
         If CLng(SweepConterStr) = CLng(CounterByTarget) Then
            TheExec.Datalog.WriteComment "=========================== Count EYE=========================== "
              For Each site In TheExec.sites
              
                   For k = 0 To UBound(strTemp)
                   'SplitByAt = Split(argv(i), "@")
                   
                   'TnumRecord = TheExec.sites.item(Site).TestNumber
                  
                       For z = 0 To CounterByWidth - 1
                        TestNameInput = "UnitCell" & z & "_" & CStr(strTemp(k)) & "EYE"
                        TestNameInputeye = "UnitCell" & z & "_" & CStr(strTemp(k)) & "EYE-percent"
                        
                        TestNameInput = Report_TName_From_Instance("X", "x", TestNameInput, , 0, ForceResult:=tlForceNone) 'eng_forceflow_transfer
                        
                   
                   
                        '----------------------------------Debug by Dylan--------------------------------------'
                        
                        For t = 0 To UBound(SplitRegister)
                            RegNameSplit = Split(SplitRegister(t), ":")
                            RegRangeValue = Split(RegNameSplit(1), ",")
                            For m = 0 To UBound(RegRangeValue)
                                minmumValue = Split(RegRangeValue(m), "-")(0)
                                MaxmumValue = Split(RegRangeValue(m), "-")(1)
                                If CInt(Split(strTemp(k), "-")(1)) >= minmumValue And CInt(Split(strTemp(k), "-")(1)) <= MaxmumValue Then
                                    If RegNameSplit(0) = "DDR2X" Then
                                        Debug.Print "Site :" & "site" & "    " & strTemp(k) & "----------------" & "DDR2X"
                                        LowLimitValue = mdll_low(0) / 8
                                        HighLimitValue = mdll_low(0) / 2
                                    ElseIf RegNameSplit(0) = "SDR" Then
                                        Debug.Print "Site :" & "site" & "    " & strTemp(k) & "----------------" & "SDR"
                                        LowLimitValue = mdll_low(0) / 4
                                        HighLimitValue = mdll_low(0) / 1
                                    End If
                                    Exit For
                                End If
                            Next m
                        Next t
                        
                        
                        
                        TheExec.Flow.TestLimit lowVal:=LowLimitValue, hiVal:=HighLimitValue, resultVal:=PrintEye(z + k * CounterByWidth) * EyeDivide, formatStr:="%i", Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling 'eng_forceflow_transfer
                        '---------------------------------------------------------------------------------------'
                   
                        'TheExec.Flow.TestLimit LowVal:=mdll_low(0), HiVal:=mdll_high(0), resultVal:=PrintEye(z + k * CounterByWidth) * EyeDivide, FormatStr:="%i", TName:=TestNameInput, ForceResults:=tlForceFlow, ScaleType:=scaleNoScaling
''''''''''                        TheExec.Flow.TestLimit lowVal:=mdll_low(0), hiVal:=mdll_high(0), resultVal:=PrintEye(z + k * CounterByWidth) * EyeDivide, formatStr:="%i", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                        
                        TestNameInputeye = Report_TName_From_Instance("calc", "x", TestNameInputeye, , 0, ForceResult:=tlForceNone) 'eng_forceflow_transfer
                        
                        If TheExec.TesterMode = testModeOffline Then
                            Eye_precent = (PrintEye(z + k * CounterByWidth) * EyeDivide) / 1  ' 20200722 add by CT to prevent "divide by 0"
                        Else
''''''''''                            Eye_precent = (PrintEye(z + k * CounterByWidth) * EyeDivide) / Mdll_lock(site).Element(0)
                            For t = 0 To UBound(SplitRegister)
                                RegNameSplit = Split(SplitRegister(t), ":")
                                RegRangeValue = Split(RegNameSplit(1), ",")
                                For m = 0 To UBound(RegRangeValue)
                                    minmumValue = Split(RegRangeValue(m), "-")(0)
                                    MaxmumValue = Split(RegRangeValue(m), "-")(1)
                                    If CInt(Split(strTemp(k), "-")(1)) >= minmumValue And CInt(Split(strTemp(k), "-")(1)) <= MaxmumValue Then
                                        
                                        If Mdll_lock(site).Element(0) = 0 Then
                                            Mdll_lock(site).Element(0) = 0.000001
                                        End If
                                    
                                        If RegNameSplit(0) = "DDR2X" Then
                                            Debug.Print "Site :" & "site" & "    " & strTemp(k) & "----------------" & "DDR2X"
                                            Eye_precent = ((PrintEye(z + k * CounterByWidth) * EyeDivide)) / (0.5 * Mdll_lock(site).Element(0))
                                        ElseIf RegNameSplit(0) = "SDR" Then
                                            Debug.Print "Site :" & "site" & "    " & strTemp(k) & "----------------" & "SDR"
                                            Eye_precent = ((PrintEye(z + k * CounterByWidth) * EyeDivide)) / (Mdll_lock(site).Element(0))
                                        End If
                                        Exit For
                                    End If
                                Next m
                            Next t
                        End If
                        
                        TheExec.Flow.TestLimit lowVal:=25, hiVal:=100, resultVal:=Format(Eye_precent * 100, 0), formatStr:="%i", Tname:=TestNameInputeye, ForceResults:=tlForceNone, scaletype:=scaleNoScaling 'eng_forceflow_transfer
                        
                        
                        
                        
                        'TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
                          'theexec.sites.item(Site).TestNumber = theexec.sites.item(Site).TestNumber + 1
                       Next z
                   'TheExec.sites.item(Site).TestNumber = TheExec.sites.item(Site).TestNumber + 1
                   Next k
           
               Next site
               
               
               For Each site In TheExec.sites
                 For k = 0 To UBound(strTemp)
                    For z = 0 To CounterByWidth - 1
                      If EyeDivide <> "" Then
                            
                            TheExec.Datalog.WriteComment "Site" & CStr(site) & " , " & "UnitCell" & z & "_" & "EyeWidth : " & PrintEye(z + k * CounterByWidth) * EyeDivide
                      Else
                            TheExec.Datalog.WriteComment "Site" & CStr(site) & " , " & "UnitCell" & z & "_" & "EyeWidth : " & PrintEye(z + k * CounterByWidth)
                      End If
                      TheExec.Datalog.WriteComment "Site" & CStr(site) & " , " & "UnitCell" & z & "_" & CStr(strTemp(k)) & " : " & CStr(UnitCellrecord(z + k * CounterByWidth))
                    Next z
                 Next k
              Next site
               
               
               
        End If
    Next i
End Function
Public Function Check_Powerrails_check(argc As Integer, argv() As String) As Long
'VDD_FIXED_PS,MON121_DEC@10,MON122_DEC@16

Dim Check_pinname As String: Check_pinname = argv(0)
Dim First_fuseName() As String
Dim Second_fuseName() As String
Dim Save_firstDSPvalue As New DSPWave
Dim Save_SecondDSPvalue As New DSPWave
Dim Check_pinvalue As New SiteDouble
Dim TestNameInput As String

First_fuseName = Split(argv(1), "@")
Second_fuseName = Split(argv(2), "@")
Save_firstDSPvalue.CreateConstant 0, 1, DspDouble
Save_SecondDSPvalue.CreateConstant 0, 1, DspDouble
'
'
'For Each site In TheExec.sites
'    Check_pinvalue = thehdw.DCVS.Pins(Check_pinname).Voltage.Main.value
'    If Check_pinvalue > 0.85 And Check_pinvalue < 0.84 Then
'
'Next site

Check_pinvalue = TheHdw.DCVS.Pins(Check_pinname).Voltage.Main.ValuePerSite
TestNameInput = Report_TName_From_Instance(CalcC, Check_pinname, "CheckPowerRail", CInt(0))
TheExec.Flow.TestLimit resultVal:=Check_pinvalue, Tname:=TestNameInput, ForceResults:=tlForceFlow

For Each site In TheExec.sites
    If Check_pinvalue > 0.805 And Check_pinvalue < 0.84 Then
        Save_firstDSPvalue.Element(0) = First_fuseName(1)
        Save_SecondDSPvalue.Element(0) = Second_fuseName(1)
    Else
        Save_firstDSPvalue.Element(0) = 0
        Save_SecondDSPvalue.Element(0) = 0

    End If
Next site
    
Call AddStoredCaptureData(CStr(First_fuseName(0)), Save_firstDSPvalue)
Call AddStoredCaptureData(CStr(Second_fuseName(0)), Save_SecondDSPvalue)



End Function

Public Function TX_EQ(argc As Integer, argv() As String) As Long

    Dim site As Variant
    Dim i As Long
    Dim j As Long
    Dim TX_Va_EQ As New PinListData
    Dim TX_Vb_EQ As New PinListData
    Dim TX_Vc_EQ As New PinListData
    
    Dim TX_PM_EQ As New PinListData
    Dim TX_PM_PRE As New PinListData
    Dim OutputTname_format() As String
    Dim TestNameInput As String
    
    TX_Va_EQ = GetStoredMeasurement(argv(0))
    TX_Vb_EQ = GetStoredMeasurement(argv(1))
    TX_PM_EQ = TX_Va_EQ
    
    If (argc = 3) Then
        TX_Vc_EQ = GetStoredMeasurement(argv(2))
        TX_PM_PRE = TX_Vc_EQ
    End If
    
    
    
    If TheExec.TesterMode = testModeOffline Then
    For i = 0 To TX_PM_EQ.Pins.Count - 1
        For Each site In TheExec.sites.Active

            TX_Va_EQ.Pins(i).value(site) = 10
            TX_Vb_EQ.Pins(i).value(site) = 10
            If (argc = 3) Then
                TX_Vc_EQ.Pins(i).value(site) = 10
            End If
        Next site
'
    Next i
    End If
    
    
    
        
    For i = 0 To TX_PM_EQ.Pins.Count - 1
        For Each site In TheExec.sites.Active

            If ((TX_Vb_EQ.Pins(i).value(site) / TX_Va_EQ.Pins(i).value(site)) > 0) Then
                TX_PM_EQ.Pins(i).value(site) = 20 * Log10((TX_Vb_EQ.Pins(i).value(site) / TX_Va_EQ.Pins(i).value(site)))
            Else
                TX_PM_EQ.Pins(i).value(site) = 0
            End If
            If (argc = 3) Then
                If (TX_Vc_EQ.Pins(i).value(site) / TX_Vb_EQ.Pins(i).value(site) > 0) Then
                    TX_PM_PRE.Pins(i).value(site) = 20 * Log10(TX_Vc_EQ.Pins(i).value(site) / TX_Vb_EQ.Pins(i).value(site))
                Else
                    TX_PM_PRE.Pins(i).value(site) = 0
                End If
            End If
        Next site
'
    Next i
    

    For j = 0 To TX_PM_EQ.Pins.Count - 1
        TestNameInput = Report_TName_From_Instance(CalcV, TX_PM_EQ.Pins(j), vbNullString, CInt(j))
        TheExec.Flow.TestLimit resultVal:=TX_PM_EQ.Pins(j), ForceResults:=tlForceFlow, Tname:=TestNameInput
    Next j

    If (argc = 3) Then
        For j = 0 To TX_PM_PRE.Pins.Count - 1
            TestNameInput = Report_TName_From_Instance(CalcV, TX_PM_PRE.Pins(j), vbNullString, CInt(j))
            TheExec.Flow.TestLimit resultVal:=TX_PM_PRE.Pins(j), ForceResults:=tlForceFlow, Tname:=TestNameInput
        Next j
    End If


End Function
Public Function Calc_PVTx(argc As Long, argv() As String) As Long

    Dim site As Variant
    Dim k As Integer
    
    Dim DigCapWave As New DSPWave
    Dim DigSrcWave As New DSPWave
    Dim DigCapKey As String
    Dim Savekeyname As String
    Dim PS As New DSPWave
    Dim SweepFrom As Integer
    Dim SweepTo As Integer
    
    DigCapKey = argv(0)
    Savekeyname = argv(1)
    
    SweepFrom = argv(2)
    SweepTo = argv(3)
    
    DigCapWave = GetStoredCaptureData(DigCapKey)

    

    k = TheExec.Flow.var("SrcCodeIndx").value


    If k = SweepFrom Then
            'Set ParallelStream = New DSPWave
            


        For Each site In TheExec.sites
            DigCapStrs = str(DigCapWave.Element(0))
        Next site
        PVTx_1to0 = -1

    ElseIf (SweepFrom < SweepTo And k <= SweepTo) Or (SweepFrom > SweepTo And k >= SweepTo) Then


        For Each site In TheExec.sites
            If right(DigCapStrs, 1) = "1" And DigCapWave.Element(0) = 0 Then
                PVTx_1to0 = k - (SweepTo - SweepFrom) / Abs(SweepTo - SweepFrom)
                

            End If
            DigCapStrs = DigCapStrs & Trim(str(DigCapWave.Element(0)))

        Next site


        If k = SweepTo Then
            Dim TestNameInput As String
            Dim gl_FlowForLoop_DigSrc_SweepCode_temp As String
            
            
            PS.CreateConstant 0, 1, DspLong
            TheExec.Datalog.WriteComment " *** PVTX-SEARCH (1->0) ***"
            If SweepTo = 63 Then
                TheExec.Datalog.WriteComment "         0         1         2         3         4         5         6"
            ElseIf SweepTo = 0 Then
                TheExec.Datalog.WriteComment "        63  6         5         4         3         2         1         0"
            End If
            For Each site In TheExec.sites
                TheExec.Datalog.WriteComment "site(" & site & "):" & DigCapStrs(site)
                
                If PVTx_1to0 >= 0 Then
                PS.Element(0) = PVTx_1to0
                Else
                    PS.Element(0) = 0
                    'Stop
                End If
                
                DigSrcWave = PS.ConvertStreamTo(tldspSerial, 6, 0, Bit0IsMsb)
            Next site

            AddStoredCaptureData Savekeyname, DigSrcWave
            
            
            
            gl_FlowForLoop_DigSrc_SweepCode_temp = gl_FlowForLoop_DigSrc_SweepCode
            gl_FlowForLoop_DigSrc_SweepCode = vbNullString
            TestNameInput = Report_TName_From_Instance(CalcC, "X")
            gl_FlowForLoop_DigSrc_SweepCode = gl_FlowForLoop_DigSrc_SweepCode_temp
            
            TheExec.Flow.TestLimit PVTx_1to0, Tname:=TestNameInput, ForceResults:=tlForceFlow

        End If
    Else

    End If

    
    
End Function


Public Function Calc_MDLL_BIST(argc As Integer, argv() As String) As Long
    'Calc_ONE_MAX_COUNT
    Dim InputKey() As String
  
    Dim site As Variant
    Dim arg As Long
    Dim i As Integer
    Dim WaveeStr As String
    
    Dim Input_Dspwave() As New DSPWave
    Dim SampleSize As Integer
    
    Dim Temp_ContinuousOne As New SiteLong
    Dim Max_ContinuousOne As New SiteLong
    Dim MDLL_Calc As New SiteDouble
    Dim EachRCapDspWave As New DSPWave
    
    '/* ------------------------------ */
    ReDim Input_Dspwave(argc)
    ReDim InputKey(argc)
    'NAND_T9BISTWRV1818_PP_TURA0_S_FULP_AN_AN01_PFF_JTG_CAL_V1818_SI_BISTWR_T9_HV
    '(0) leading
    '(1) trailing
    '(2) nis_bist_wr_bitmap
    'NAND_T10BISTRDV1818_PP_TURA0_S_FULP_AN_AN01_PFF_JTG_CAL_V1818_SI_BISTRD_T10_HV
    '(2) nis_bist_rd_pos_bitmap
    '(2) nis_bist_rd_neg_bitmap
    
    For arg = 0 To argc - 1
        InputKey(arg) = LCase(argv(arg))
        Input_Dspwave(arg) = GetStoredCaptureData(InputKey(arg))
        For Each site In TheExec.sites
        Call Wave2Str_Single(Input_Dspwave(arg), WaveeStr)
        TheExec.Datalog.WriteComment "Site(" & site & "):" & InputKey(arg) & " = " & WaveeStr
        Next site
    Next arg
    

    For arg = 1 To argc - 1 Step 3
        Call rundsp.ConvertToLongAndSerialToParrel(Input_Dspwave(arg - 1), 9, EachRCapDspWave)
            For Each site In TheExec.sites
                Temp_ContinuousOne = 0
                Max_ContinuousOne = 0
                For i = 0 To Input_Dspwave(arg + 1).SampleSize - 1
                    If Input_Dspwave(arg + 1).Element(i) = 1 Then
                        Temp_ContinuousOne = Temp_ContinuousOne + 1
                    Else
                        If Temp_ContinuousOne > Max_ContinuousOne Then
                            Max_ContinuousOne = Temp_ContinuousOne
                        End If
                        Temp_ContinuousOne = 0
                    End If
                Next i
                If Temp_ContinuousOne > Max_ContinuousOne Then
                        Max_ContinuousOne = Temp_ContinuousOne
                        Temp_ContinuousOne = 0
                End If
             TheExec.Datalog.WriteComment "Site(" & site & "):" & "Decimal of Leading" & " = " & EachRCapDspWave.Element(0)
             TheExec.Datalog.WriteComment "Site(" & site & "):" & "Max Continuous One of Bitmap" & " = " & Max_ContinuousOne
             TheExec.Datalog.WriteComment "Site(" & site & "):" & "Formula" & " = " & Max_ContinuousOne & "/" & EachRCapDspWave.Element(0) & "*" & 360
             If EachRCapDspWave.Element(0) = 0 Then
                TheExec.Datalog.WriteComment "Site(" & site & "):" & "Decimal of Leading" & " = 0, Set Decimal of Leading=99999999"
                EachRCapDspWave.Element(0) = 99999999
             End If
             MDLL_Calc = Round((Max_ContinuousOne * 360) / EachRCapDspWave.Element(0), 5)
         Next site
        'Report_TestLimit_by_CZ_Format resultVal:=Max_ContinuousOne, MeasType:="C", UserVar5:="MaxOneCount", UserVar7:=InputKey(arg + 1), scaletype:=scaleNoScaling, ForceResults:=tlForceFlow
        'Report_TestLimit_by_CZ_Format resultVal:=MDLL_Calc, MeasType:="C", UserVar7:=InputKey(arg + 1), scaletype:=scaleNoScaling, ForceResults:=tlForceFlow
   
         
        Dim TestNameInput As String
        Dim gl_FlowForLoop_DigSrc_SweepCode_temp As String
        gl_FlowForLoop_DigSrc_SweepCode_temp = gl_FlowForLoop_DigSrc_SweepCode
        
        gl_FlowForLoop_DigSrc_SweepCode = Replace(InputKey(arg + 1), "_", vbNullString)
        TestNameInput = Report_TName_From_Instance(CalcC, PinName:="MaxOneCount")
        
            
        TheExec.Flow.TestLimit resultVal:=Max_ContinuousOne, Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
        TestNameInput = Report_TName_From_Instance(CalcC, "Degrees")
        TheExec.Flow.TestLimit resultVal:=MDLL_Calc, Tname:=TestNameInput, ForceResults:=tlForceFlow
        
        gl_FlowForLoop_DigSrc_SweepCode = gl_FlowForLoop_DigSrc_SweepCode_temp
   
   Next arg
   
End Function

Public Function Calc_NAND_PHY_MDLL(argc As Integer, argv() As String) As Long

    Dim i As Long, j As Long
    Dim site As Variant
    Dim Result_Ratio As New SiteDouble
    Dim DSPWave_Binary() As New DSPWave
    ReDim DSPWave_Binary(argc - 1) As New DSPWave
    
    Dim DSPWave_Combine_Dec() As New DSPWave
    ReDim DSPWave_Combine_Dec(argc - 1) As New DSPWave
    
    Dim DSPWave_Result As New DSPWave
    DSPWave_Result.CreateConstant 0, 3, DspLong

    Dim DSPWave_Result_K As New DSPWave
    DSPWave_Result_K.CreateConstant 0, 1, DspLong
        
        
    If TheExec.TesterMode = testModeOnline Then
    
        If TheExec.Flow.enableWord("TTR") = True Then
            For i = 0 To 3
                DSPWave_Binary(i) = GetStoredCaptureData(argv(i))
            Next i
            Call rundsp.Calc_NAND_PHY_MDLL_DSP(DSPWave_Binary(0), DSPWave_Binary(1), DSPWave_Binary(2), DSPWave_Binary(3), DSPWave_Result, Result_Ratio)
        Else
        
            For i = 0 To 3
                DSPWave_Combine_Dec(i).CreateConstant 0, 1, DspLong
                DSPWave_Binary(i) = GetStoredCaptureData(argv(i))
                Call rundsp.ConvertToLongAndSerialToParrel(DSPWave_Binary(i), 9, DSPWave_Combine_Dec(i))
            Next i
            
             For Each site In TheExec.sites
                DSPWave_Result.Element(0) = DSPWave_Combine_Dec(0).Element(0) - DSPWave_Combine_Dec(1).Element(0)
                DSPWave_Result.Element(1) = DSPWave_Combine_Dec(2).Element(0) - DSPWave_Combine_Dec(3).Element(0)
        '        DSPWave_Result.Element(2) = DSPWave_Result.Element(1) - DSPWave_Result.Element(0)
                If DSPWave_Result.Element(0) = 0 Then DSPWave_Result.Element(0) = 99999999
                Result_Ratio = DSPWave_Result.Element(1) / DSPWave_Result.Element(0)
            Next site
        End If
    
    'If TheExec.TesterMode = testModeOffline Then
    Else
    'testModeOffline
        For i = 0 To 3
            DSPWave_Combine_Dec(i).CreateConstant 0, 1, DspLong
            DSPWave_Binary(i) = GetStoredCaptureData(argv(i))
            Call rundsp.ConvertToLongAndSerialToParrel(DSPWave_Binary(i), 9, DSPWave_Combine_Dec(i))
        Next i
        
         For Each site In TheExec.sites
            DSPWave_Result.Element(0) = DSPWave_Combine_Dec(0).Element(0) - DSPWave_Combine_Dec(1).Element(0)
            DSPWave_Result.Element(1) = DSPWave_Combine_Dec(2).Element(0) - DSPWave_Combine_Dec(3).Element(0)
    '        DSPWave_Result.Element(2) = DSPWave_Result.Element(1) - DSPWave_Result.Element(0)
            If DSPWave_Result.Element(0) = 0 Then DSPWave_Result.Element(0) = 99999999
            Result_Ratio = DSPWave_Result.Element(1) / DSPWave_Result.Element(0)
        Next site
    
        For Each site In TheExec.sites
            DSPWave_Result.Element(1) = DSPWave_Combine_Dec(2).Element(0) - DSPWave_Combine_Dec(3).Element(0) + 400 + (site * 10)
        Next site
    End If
    
    
    If TheExec.Flow.enableWord("CZ2_PRINT_EN") = False Then
        TheExec.Flow.TestLimit DSPWave_Result.Element(0), , , , ForceResults:=tlForceFlow  'chyehq
        TheExec.Flow.TestLimit DSPWave_Result.Element(1), , , , ForceResults:=tlForceFlow  'chyehq
        TheExec.Flow.TestLimit Result_Ratio, , , , ForceResults:=tlForceFlow  'chyehq
    Else
         
        Dim TestNameInput As String
        Dim gl_FlowForLoop_DigSrc_SweepCode_temp As String
        gl_FlowForLoop_DigSrc_SweepCode_temp = gl_FlowForLoop_DigSrc_SweepCode
        
        
        
        
    
        'Report_TestLimit_by_CZ_Format DSPWave_Result.Element(0), , , , ForceResults:=tlForceFlow, MeasType:="C"
        'Report_TestLimit_by_CZ_Format DSPWave_Result.Element(1), , , , ForceResults:=tlForceFlow, MeasType:="C"
        'Report_TestLimit_by_CZ_Format Result_Ratio, , , ForceResults:=tlForceFlow, MeasType:="C", UserVar5:="MDLL", UserVar6:="CAL", UserVar7:="RATIO"
        
        TestNameInput = Report_TName_From_Instance(CalcC, vbNullString, ForceResult:=tlForceFlow)
        TheExec.Flow.TestLimit DSPWave_Result.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
        
        TestNameInput = Report_TName_From_Instance(CalcC, vbNullString, ForceResult:=tlForceFlow)
        TheExec.Flow.TestLimit DSPWave_Result.Element(1), Tname:=TestNameInput, ForceResults:=tlForceFlow
        
        TestNameInput = Report_TName_From_Instance(CalcC, PinName:="MDLL", ForceResult:=tlForceFlow)
        'Stop
        
        TheExec.Flow.TestLimit Result_Ratio, Tname:=TestNameInput, ForceResults:=tlForceFlow
        
    End If
    
    
'     For Each Site In TheExec.sites
'         If DSPWave_Result.Element(1) < 1 Then DSPWave_Result.Element(1) = 1
'
'         If TheExec.Flow.EnableWord("CZ2_PRINT_EN") = False Then
'            TheExec.Flow.TestLimit DSPWave_Result.Element(2), 1, DSPWave_Result.Element(1), Tname:="MDLLCALDIFF", ForceResults:=tlForceFlow  'chyehq
'        Else
'            Report_TestLimit_by_CZ_Format DSPWave_Result.Element(2), 1, DSPWave_Result.Element(1), ForceResults:=tlForceFlow, MeasType:="C", UserVar5:="MDLL", UserVar6:="CAL", UserVar7:="DIFF"
'        End If
'    Next Site
    
End Function

Public Function Wave2Str_Single(InDSPWave As DSPWave, outstr As Variant, Optional SP As Integer = 0, Optional EP As Integer = 0) As Long
    Dim i As Integer
    outstr = vbNullString
    If SP + EP > 0 Then
        If EP > InDSPWave.SampleSize - 1 Then EP = InDSPWave.SampleSize - 1
        For i = SP To EP
            outstr = outstr & CStr(InDSPWave.Element(i))
        Next i
    Else
        For i = 0 To InDSPWave.SampleSize - 1
            outstr = outstr & CStr(InDSPWave.Element(i))
        Next i
    End If
End Function

Public Function Calc_PVTP(argc As Long, argv() As String) As Long
    
    Dim site As Variant
    Dim PVTPNR_int As New SiteLong
    Dim i, j, k As Integer
    Dim WaveeStr As String
    Dim dataArray(31) As Long
    Dim SimStr(7) As String
    Dim SimArray() As Long
        
    Dim InDSPWave As New DSPWave
    Dim PVTP_Wave As New DSPWave
    Dim PVTP_PLUS_1_Wave As New DSPWave
    Dim PVTP_PLUS_2_Wave As New DSPWave
    Dim PVTP_PLUS_3_Wave As New DSPWave
    Dim PVTP_PLUS_4_Wave As New DSPWave
    Dim PVTP_MINUS_1_Wave As New DSPWave
    
    Dim ParallelStream As New DSPWave ''''20190604Error
    Dim PVT_1to0 As Integer ''''20190604Error
    
    Dim PVTx As New SiteDouble
    Dim PVTx_Stored As New SiteDouble
    
    
    Dim AnySite_1to0_Check As New SiteBoolean
    'Dim All_1_to_0 As new SiteBoolean
    
    
    
    InDSPWave = GetStoredCaptureData(argv(0))
    
    Dim keyname As String
    
    keyname = argv(1)
    
    k = TheExec.Flow.var("SrcCodeIndx").value
    
    PVTPNR_int = InDSPWave.Element(0)
        
    If k = 0 Then
        ParallelStream.CreateConstant 0, 1, DspLong
        
        For Each site In TheExec.sites
            DigCapStrs = str(InDSPWave.Element(0))
        Next site
        PVT_1to0 = 999
        
    ElseIf k <= 32 Then
    
   
        For Each site In TheExec.sites
            If right(DigCapStrs, 1) = "1" And InDSPWave.Element(0) = 0 Then
                PVT_1to0 = k
                ParallelStream.Element(0) = k
                
            End If
            DigCapStrs = DigCapStrs & Trim(str(InDSPWave.Element(0)))
            
        Next site
        
        
        If k = 32 Then
            TheExec.Flow.TestLimit PVT_1to0, ForceResults:=tlForceNone 'eng_forceflow_transfer
            
            TheExec.Datalog.WriteComment "         0         1         2         3"
            For Each site In TheExec.sites
                TheExec.Datalog.WriteComment "site(" & site & "):" & DigCapStrs(site)
                PVTP_Wave = ParallelStream.ConvertStreamTo(tldspSerial, 6, 0, Bit0IsMsb)
            Next site
            
            Stop
            
            AddStoredCaptureData keyname, PVTP_Wave
            
        End If
    Else
    
            
    End If
    Exit Function
    
    PVTx = PVTPNR_int.compare(EqualTo, 0)
    PVTx = PVTx.Abs.Multiply(k)
    
    
    If k = 32 Then
        
        For Each site In TheExec.sites
            TheExec.Datalog.WriteComment "site(" & site & "):" & DigCapStrs(site)
        Next site
        Stop

        
    Else
    
    
        'Stop
    End If
    
    
    
    
    

    
    
    

End Function
Public Function Calc_PVTN(argc As Long, argv() As String) As Long
    
    
    Dim site As Variant
    Dim PVTPNR_int As New SiteLong
    Dim i, j, k As Integer
    Dim WaveeStr As String
    Dim dataArray(31) As Long
    Dim SimStr(7) As String
    Dim SimArray() As Long
     
    Dim InDSPWave As New DSPWave
    Dim PVTN_Wave As New DSPWave
    Dim PVTN_PLUS_1_Wave As New DSPWave
    Dim PVTN_PLUS_2_Wave As New DSPWave
    Dim PVTN_PLUS_3_Wave As New DSPWave
    Dim PVTN_PLUS_4_Wave As New DSPWave
    Dim PVTN_MINUS_1_Wave As New DSPWave

    If TheExec.TesterMode = testModeOffline Then
        For Each site In TheExec.sites
            For i = 0 To UBound(dataArray)
                dataArray(i) = 0
                If i > 15 + site Then dataArray(i) = 1
            Next i
            InDSPWave.data = dataArray
        Next site
    Else
        InDSPWave = GetStoredCaptureData(argv(0))
    End If
    
    Call rundsp.NAND_PVTN(InDSPWave, PVTPNR_int, PVTN_Wave, PVTN_PLUS_1_Wave, PVTN_PLUS_2_Wave, PVTN_PLUS_3_Wave, PVTN_PLUS_4_Wave, PVTN_MINUS_1_Wave)
    
    AddStoredCaptureData argv(1), PVTN_Wave
    AddStoredCaptureData "PVTN_PLUS_1", PVTN_PLUS_1_Wave
    AddStoredCaptureData "PVTN_PLUS_2", PVTN_PLUS_2_Wave
    AddStoredCaptureData "PVTN_PLUS_3", PVTN_PLUS_3_Wave
    AddStoredCaptureData "PVTN_PLUS_4", PVTN_PLUS_4_Wave
    AddStoredCaptureData "PVTN_MINUS_1", PVTN_MINUS_1_Wave
    
    'If NDLog = False Then
        For Each site In TheExec.sites
            Call Wave2Str_Single(PVTN_Wave, WaveeStr)
            TheExec.Datalog.WriteComment "Site: " & site & " The transition is point " & PVTPNR_int
            TheExec.Datalog.WriteComment "Site: " & site & " The PVTN Binary Code =  " & WaveeStr
    
            Call Wave2Str_Single(InDSPWave, WaveeStr)
            TheExec.Datalog.WriteComment "Site: " & site & ", Capture bits " & InDSPWave.SampleSize & " = " & WaveeStr
        Next site
    'End If
    
    ''2017/08/11 , updated by Kaino for CZ2 naming
    'TheExec.Flow.TestLimit PVTPNR_int, , , , , , , , , , , , , , , ForceResults:=tlForceFlow
    If TheExec.Flow.enableWord("CZ2_PRINT_EN") = False Then
        TheExec.Flow.TestLimit PVTPNR_int, , , , , , , , , , , , , , , ForceResults:=tlForceFlow
    Else
        Report_TestLimit_by_CZ_Format resultVal:=PVTPNR_int, ForceResults:=tlForceFlow, MeasType:="C", PinName:="JTAG_TDO"
    End If
        
    'Exit Function

End Function


Public Function Calc_ONE_MAX_COUNT(argc As Integer, argv() As String) As Long
    'Calc_ONE_MAX_COUNT
    Dim InputKey() As String
  
    Dim site As Variant
    Dim arg As Long
    Dim i As Integer
    Dim WaveeStr As String
    
    Dim Input_Dspwave() As New DSPWave
    Dim SampleSize As Integer
    
    Dim Temp_ContinuousOne As New SiteLong
    Dim Max_ContinuousOne As New SiteLong
    Dim MDLL_Calc As New SiteDouble
    Dim EachRCapDspWave As New DSPWave
    
    '/* ------------------------------ */
    ReDim Input_Dspwave(argc)
    ReDim InputKey(argc)
    'NAND_T9BISTWRV1818_PP_TURA0_S_FULP_AN_AN01_PFF_JTG_CAL_V1818_SI_BISTWR_T9_HV
    '(0) leading
    '(1) trailing
    '(2) nis_bist_wr_bitmap
    'NAND_T10BISTRDV1818_PP_TURA0_S_FULP_AN_AN01_PFF_JTG_CAL_V1818_SI_BISTRD_T10_HV
    '(2) nis_bist_rd_pos_bitmap
    '(2) nis_bist_rd_neg_bitmap
    
    For arg = 0 To argc - 1
        InputKey(arg) = LCase(argv(arg))
        Input_Dspwave(arg) = GetStoredCaptureData(InputKey(arg))
        For Each site In TheExec.sites
            Call Wave2Str_Single(Input_Dspwave(arg), WaveeStr)
            TheExec.Datalog.WriteComment "Site(" & site & "):" & InputKey(arg) & " = " & WaveeStr
        Next site
    Next arg
    
    Dim leading As Integer
    Dim nis_bist_x_bitmap As Integer
    'Stop
'    If arg = 2 Then
'        nis_bist_x_bitmap = 0
'        leading = 1
'
'    ElseIf arg = 3 Then
'        nis_bist_x_bitmap = 0
'        leading = 2
'    Else
'        Stop
'    End If
    
    leading = argc - 1
    

    For nis_bist_x_bitmap = 0 To argc - 2
        'Stop
        Call rundsp.ConvertToLongAndSerialToParrel(Input_Dspwave(leading), 9, EachRCapDspWave)  'leading
            For Each site In TheExec.sites
                Temp_ContinuousOne = 0
                Max_ContinuousOne = 0
                For i = 0 To Input_Dspwave(nis_bist_x_bitmap).SampleSize - 1
                    If Input_Dspwave(nis_bist_x_bitmap).Element(i) = 1 Then
                        Temp_ContinuousOne = Temp_ContinuousOne + 1
                    Else
                        If Temp_ContinuousOne > Max_ContinuousOne Then
                            Max_ContinuousOne = Temp_ContinuousOne
                        End If
                        Temp_ContinuousOne = 0
                    End If
                Next i
                If Temp_ContinuousOne > Max_ContinuousOne Then
                        Max_ContinuousOne = Temp_ContinuousOne
                        Temp_ContinuousOne = 0
                End If
             TheExec.Datalog.WriteComment "Site(" & site & "):" & "Decimal of Leading" & " = " & EachRCapDspWave.Element(0)
             TheExec.Datalog.WriteComment "Site(" & site & "):" & "Max Continuous One of Bitmap (" & InputKey(nis_bist_x_bitmap) & ") = " & Max_ContinuousOne
             TheExec.Datalog.WriteComment "Site(" & site & "):" & "Formula" & " = " & Max_ContinuousOne & "/" & EachRCapDspWave.Element(0) & "*" & 360
             If EachRCapDspWave.Element(0) = 0 Then
                TheExec.Datalog.WriteComment "Site(" & site & "):" & "Decimal of Leading" & " = 0, Set Decimal of Leading=99999999"
                EachRCapDspWave.Element(0) = 99999999
             End If
             MDLL_Calc = Round((Max_ContinuousOne * 360) / EachRCapDspWave.Element(0), 5)
         Next site
        'Report_TestLimit_by_CZ_Format resultVal:=Max_ContinuousOne, MeasType:="C", UserVar5:="MaxOneCount", UserVar7:=InputKey(arg + 1), scaletype:=scaleNoScaling, ForceResults:=tlForceFlow
        'Report_TestLimit_by_CZ_Format resultVal:=MDLL_Calc, MeasType:="C", UserVar7:=InputKey(arg + 1), scaletype:=scaleNoScaling, ForceResults:=tlForceFlow
   
         
        Dim TestNameInput As String
        Dim gl_FlowForLoop_DigSrc_SweepCode_temp As String
        gl_FlowForLoop_DigSrc_SweepCode_temp = gl_FlowForLoop_DigSrc_SweepCode
        
        gl_FlowForLoop_DigSrc_SweepCode = Replace(InputKey(nis_bist_x_bitmap), "_", vbNullString)
        TestNameInput = Report_TName_From_Instance(CalcC, vbNullString, Tname:="MaxOneCount", ForceResult:=tlForceNone) 'eng_forceflow_transfer
        
            
        TheExec.Flow.TestLimit resultVal:=Max_ContinuousOne, Tname:=TestNameInput, ForceResults:=tlForceNone 'eng_forceflow_transfer
        TestNameInput = Report_TName_From_Instance(CalcC, vbNullString, ForceResult:=tlForceFlow)
        TheExec.Flow.TestLimit resultVal:=MDLL_Calc, Tname:=TestNameInput, ForceResults:=tlForceFlow
        
        gl_FlowForLoop_DigSrc_SweepCode = gl_FlowForLoop_DigSrc_SweepCode_temp
   
   Next nis_bist_x_bitmap
   
End Function



Public Function Calc_Abs_Res(argc As Long, argv() As String)

    Dim Calc_Operand1 As Double
    Dim Calc_Operand2 As New PinListData
    Dim Calc_Res As New PinListData
    Dim TestNameInput As String
    Dim p As Variant
    Dim Temp_index As Long
        
    If InStr(UCase(argv(0)), UCase("VDD")) <> 0 Then
        'Calc_Operand1 = TheExec.Specs.DC.Item(Mid(argv(0), 2)).ContextValue
        Call HIP_Evaluate_ForceVal_New(argv(0)) 'add for +-*/ calc. by CW 190816
        Calc_Operand1 = argv(0)
        'Calc_Operand1 = TheExec.specs.DC.Item(argv(0) & "_VAR_H").ContextValue
    Else
        Calc_Operand1 = CDbl(argv(0))
    End If
    Calc_Operand2 = GetStoredMeasurement(argv(1))
    
    Calc_Res = Calc_Operand2.Copy
    Calc_Res = Calc_Operand2.Math.Abs.Invert.Multiply(Calc_Operand1)
    
    

    Temp_index = TheExec.Flow.TestLimitIndex
    
    For p = 0 To Calc_Res.Pins.Count - 1
        
        
        TestNameInput = Report_TName_From_Instance(CalcC, Calc_Res.Pins(p), , 0)
        'TestNameInput = Report_TName_From_Instance("Calc", Calc_Res.Pins(p), , 0)
        TheExec.Flow.TestLimit Calc_Res.Pins(p), , , , , , unitCustom, customUnit:="ohm", Tname:=TestNameInput, ForceResults:=tlForceFlow
        
        
        If argv(1) <> "ip0" Then TheExec.Flow.TestLimitIndex = Temp_index 'modify for AMPH T58
        
    Next p
    
    TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1

End Function



Public Function Calc_ROSC_Freq(argc As Long, argv() As String)
    
'''''''''''''''''''''input Variable info'''''''''''''''''
'Calc:Calc_ROSC_Freq;
'CalcArg:24000000,ringclk_count_val,refclk_count_val;
'
'24MHz             value to calc    > argv(0)
'ringclk_count_val dictionary name  > argv(1)
'refclk_count_val  dictionary name  > argv(2)
'
'
''''''''''''''''''''' Algorithm'''''''''''''''''
'Calculate 24Mhz*(ringclk_count_val/refclk_count_val)
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Input_Freq As Long
    Dim Input_ringclk_count_val As New DSPWave
    Dim Input_refclk_count_val As New DSPWave
    
    Dim Input_ringclk_count_val_Dec As New DSPWave
    Dim Input_refclk_count_val_Dec As New DSPWave
    
    Dim Output_Calc_Freq As New DSPWave
    
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    Dim site As Variant

    Input_Freq = CLng(argv(0))
    Input_ringclk_count_val = GetStoredCaptureData(argv(1))
    Input_refclk_count_val = GetStoredCaptureData(argv(2))
    
    Input_ringclk_count_val_Dec.CreateConstant 0, 1, DspDouble
    Input_refclk_count_val_Dec.CreateConstant 0, 1, DspDouble
    Output_Calc_Freq.CreateConstant 0, 1, DspDouble
    
    'Convert DSPwave from Binary value to Decimal value
    Call HardIP_Bin2Dec(Input_ringclk_count_val_Dec, Input_ringclk_count_val)
    Call HardIP_Bin2Dec(Input_refclk_count_val_Dec, Input_refclk_count_val)
        
    For Each site In TheExec.sites
'        If TheExec.TesterMode = testModeOffline Then
'            If Site Mod 2 = 0 Then
'                Output_Calc_Freq.Element(0) = 1000000 * CLng(TheExec.Flow.var("SrcCodeIndx").Value)
'            Else
'                Output_Calc_Freq.Element(0) = 2000000 * CLng(TheExec.Flow.var("SrcCodeIndx").Value)
'            End If
'        Else
            If Input_refclk_count_val_Dec.Element(0) <= 0 Then
                Output_Calc_Freq.Element(0) = -1
            Else
                Output_Calc_Freq.Element(0) = Input_Freq * (Input_ringclk_count_val_Dec.Element(0) / Input_refclk_count_val_Dec.Element(0))
            End If
'        End If
    Next site


    TestNameInput = Report_TName_From_Instance(CalcC, "X", , 0)

     
    TheExec.Flow.TestLimit resultVal:=Output_Calc_Freq.Element(0), Tname:=TestNameInput, unit:=unitHz, ForceResults:=tlForceFlow
        

End Function


Public Function Calc_SweepV_COMPE(argc As Long, argv() As String) As Long
    Dim site As Variant
    Dim k As Integer
    
    Dim DigSrcWave As New DSPWave
    
    
    
    Dim FlowSweepVar As String              '0
    Dim FlowFrom As String                  '1
    Dim FlowTo  As String                   '2
    
    Dim PinName As String                   '3
    Dim VFrom As String                    '4
    Dim VTo    As String                   '5
    Dim Vstep   As String                   '6
    
    Dim DigCapKey1 As String                '7
    Dim DigCapKey2 As String                '8
    
    Dim DigCapWave1 As New DSPWave          '7
    Dim DigCapWave2 As New DSPWave          '8
    
    Dim DatalogArg As String                '9
    
    Dim DatalogOut As New SiteDouble
    
    
    Dim Savekeyname As String
    Dim PS As New DSPWave
    Dim SweepFrom As Integer
    Dim SweepTo As Integer
    
    
    
    FlowSweepVar = argv(0)
    FlowFrom = argv(1)
    FlowTo = argv(2)
    PinName = argv(3)
    VFrom = argv(4)
    VTo = argv(5)
    Vstep = argv(6)
    
    DigCapKey1 = argv(7)
    DigCapKey2 = argv(8)
    DatalogArg = argv(9)
    
'    DigCapKey1 = "comp1"
'    DigCapKey2 = "comp2"
    
    DigCapWave1 = GetStoredCaptureData(DigCapKey1)
    DigCapWave2 = GetStoredCaptureData(DigCapKey2)
    
    Call HIP_Evaluate_ForceVal_New(VFrom)
    Call HIP_Evaluate_ForceVal_New(VTo)
    Call HIP_Evaluate_ForceVal_New(Vstep)
    Call HIP_Evaluate_ForceVal_New(DatalogArg)
        

    k = TheExec.Flow.var(FlowSweepVar).value


    If k = FlowFrom Then

        For Each site In TheExec.sites
            DigCapStrsCompeA = str(DigCapWave1.Element(0))
            DigCapStrsCompeB = str(DigCapWave2.Element(0))
        Next site
        CompeVref = -99.9
        CompeDigcapSwap = -1

    ElseIf (FlowFrom < FlowTo And k <= FlowTo) Or (FlowFrom > FlowTo And k >= FlowTo) Then


        For Each site In TheExec.sites
            If right(DigCapStrsCompeA, 1) = "0" And DigCapWave1.Element(0) = 1 Then
               ' Stop
                
                CompeVref = CDbl(VFrom) + k * CDbl(Vstep)
                CompeDigcapSwap = k
                
            ElseIf right(DigCapStrsCompeB, 1) = "0" And DigCapWave2.Element(0) = 1 Then
                'Stop
                CompeVref = CDbl(VFrom) + k * CDbl(Vstep)
                CompeDigcapSwap = k

            End If
            
            
            
'            DigCapStrs = DigCapStrs & Trim(Str(DigCapWave.Element(0)))
            
            
            
            DigCapStrsCompeA = DigCapStrsCompeA & Trim(str(DigCapWave1.Element(0)))
            DigCapStrsCompeB = DigCapStrsCompeB & Trim(str(DigCapWave2.Element(0)))

        Next site


        If k = FlowTo Then
            
            

            
            Dim TestNameInput As String
            Dim gl_FlowForLoop_DigSrc_SweepCode_temp As String
            
            
            PS.CreateConstant 0, 1, DspLong

            If FlowTo > FlowFrom Then
                TheExec.Datalog.WriteComment "DigCap:  0         1         2         3         4         5         6         7         8"
            ElseIf SweepTo = 0 Then
                TheExec.Datalog.WriteComment "DigCap:  8         7         6         5         4         3         2         1         0"
            End If
            For Each site In TheExec.sites
                TheExec.Datalog.WriteComment "site(" & site & "):" & DigCapStrsCompeA(site)
                TheExec.Datalog.WriteComment "site(" & site & "):" & DigCapStrsCompeB(site)
                
                
   
            Next site

           'AddStoredCaptureData Savekeyname, DigSrcWave
            
            
            
            
            gl_Sweep_Name = vbNullString
            TestNameInput = Report_TName_From_Instance(CalcC, "X")
            
            TheExec.Flow.TestLimit CompeVref, Tname:=TestNameInput, unit:=unitVolt, ForceResults:=tlForceFlow
            
            TestNameInput = Report_TName_From_Instance(CalcC, "X")
            DatalogOut = CompeVref.Subtract(CDbl(DatalogArg))
            
            For Each site In TheExec.sites
                If CompeVref = -99.9 Then
                    DatalogOut = -99.9
                End If
            Next site
            TheExec.Flow.TestLimit DatalogOut, Tname:=TestNameInput, unit:=unitVolt, ForceResults:=tlForceFlow
           
        End If
    Else

    End If

    
    
End Function


Public Function Calc_ADDRIO_Find_Closest_Result_To_0_Alg1(argc As Integer, argv() As String) As Long
    
    On Error GoTo errHandler
    

'Alg::Calc_ADDRIO_Find_Closest_Result_To_0(sn1,ADDR_P2M_CK_P,ADDRIO_Norm_Y_T1,50,0,8,ADDR_RX_ZCPU,ADDR_TX_ZCPU)

''''''input argument'''''''''
'MeasR result: sn1
'PinName : ADDR_P2M_CK_P
'polynom_result(each code):ADDRIO_Norm_Y_T1()
'Offset:50
'Closest to:0
'StoreDictionary_Bits:8
'StoreDictionary_Name: ADDR_RX_ZCPU,ADDR_TX_ZCPU

    
'1.  Use T1P1 pattern, set code as 35 then measure R from ADDR_P2M_CK_P pins, Get R1 result
'2.  UseT1 coefficient and calculate R1 result * Norm Y (trim code ) -50 , Total 64 result.
'3.  Find the result closets to 0 then the index as the trim code. Need to fuse in FT2.
    

    Dim site As Variant
    Dim i As Long
    Dim MeasR_Results As New PinListData:: MeasR_Results = GetStoredMeasurement(argv(0))
    Dim PinName As String:: PinName = argv(1)
    Dim Offset_Value As Long:: Offset_Value = CDbl(argv(3))
    Dim Closest_to_What_Value As Long:: Closest_to_What_Value = CDbl(argv(4))
    
    Dim Temp_Value_Array() As New SiteVariant
    
    
    If UCase(argv(2)) Like "*T1*" Then
        ReDim Temp_Value_Array(UBound(ADDRIO_Norm_Y_T1))
        
        For i = 0 To UBound(ADDRIO_Norm_Y_T1)
            Temp_Value_Array(i) = MeasR_Results.Math.Multiply(ADDRIO_Norm_Y_T1(i)).Subtract(Offset_Value)
        Next i
        
        
    
    ElseIf UCase(argv(2)) Like "*T2*" Then
        ReDim Temp_Value_Array(UBound(ADDRIO_Norm_Y_T2))
        
        For i = 0 To UBound(ADDRIO_Norm_Y_T2)
            Temp_Value_Array(i) = MeasR_Results.Math.Multiply(ADDRIO_Norm_Y_T2(i)).Subtract(Offset_Value)
        Next i
    
    End If
    
    Dim Temp_Closest_Previous As Double
    Dim Temp_Closest As Double
    Dim Closest_Code As Double
    Dim Closest_Code_Final As New DSPWave
    Dim Closest_Code_Final_Bin As New DSPWave
    
    Closest_Code_Final.CreateConstant 0, 1, DspLong
    For Each site In TheExec.sites.Active
        Temp_Closest_Previous = 999999
        Closest_Code = 999999
        For i = 0 To UBound(Temp_Value_Array)
        
            Temp_Closest = Abs(Temp_Value_Array(i) - Closest_to_What_Value)

            If Temp_Closest <= Temp_Closest_Previous Then
                Temp_Closest_Previous = Temp_Closest
                Closest_Code = i
            End If

        Next i
        Closest_Code_Final.Element(0) = Closest_Code
        Closest_Code_Final_Bin = Closest_Code_Final.ConvertStreamTo(tldspSerial, argv(5), 0, Bit0IsMsb)
        
        If gl_Disable_HIP_debug_log = False Then Call TheExec.Datalog.WriteComment("Site(" & site & ")  Closest Code = " & Closest_Code)
        
    Next site
    
    Call AddStoredCaptureData(argv(6), Closest_Code_Final_Bin)
    Call AddStoredCaptureData(argv(7), Closest_Code_Final_Bin)
    
   
    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in Calc_ADDRIO_Find_Closest_Result_To_0"
    If AbortTest Then Exit Function Else Resume Next
    
End Function




Public Function Public_AddStoredCaptureData(keyname As String, ByRef obj As DSPWave) As Long
'**************************************************
'SeaHawk Edited by 20190606
'**************************************************
    keyname = LCase(keyname)
    If gl_DictDSPWave.Exists(keyname) Then
        gl_DictDSPWave.Remove (keyname)
    End If
    gl_DictDSPWave.Add keyname, obj
    
    '20220106, Add for Real value validation( reverse bit)
    Dim vsite As Variant
    Dim objvalue As New DSPWave
    If gB_efuse_DicValue_Chk_Flag = True Then ''Efuse_DicValue_Chk
        For Each vsite In TheExec.sites
            If obj.SampleSize = 1 Then
                objvalue = obj.ConvertDataTypeTo(DspLong)
                TheExec.Datalog.WriteComment "Site(" + CStr(vsite) + ")" + "@@@ Key_Name:= " + keyname + " Value:= " + CStr(objvalue.Element(0))
            Else
                objvalue = obj.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspParallel, obj.SampleSize, 0)
                TheExec.Datalog.WriteComment "Site(" + CStr(vsite) + ")" + "@@@ Key_Name:= " + keyname + " Value:= " + CStr(objvalue.Element(0))
            End If
        Next vsite
    End If
    ''Efuse_DicValue_Chk <==
End Function

Public Function Public_GetStoredCaptureData(keyname As String) As Variant
'**************************************************
'SeaHawk Edited by 20190606
'**************************************************
    keyname = LCase(keyname)
    If Not gl_DictDSPWave.Exists(keyname) Then
        TheExec.ErrorLogMessage "Stored capture data " & keyname & " not found."
    Else
        Set Public_GetStoredCaptureData = gl_DictDSPWave(keyname)
    End If
End Function

Public Function MTRTMPS_Gain_AVG(InWf As DSPWave, StoreName_GainMean As String, Integer_Bit As Long) As Long
Dim DSP_Gain_Mean As New DSPWave
Dim DSP_Gain_Mean_Array(0) As Double
Dim DSP_Gain_Mean_Fuse As New DSPWave
Dim DSP_Gain_Mean_Fuse_Array(0) As Double
Dim TestNameInput As String
'Dim High_limit As Double: High_limit = Bin2Dec_rev(String(Integer_Bit - 1, "1"))
'Dim Low_limit As Double: Low_limit = -2 ^ (Integer_Bit - 1)

    For Each site In TheExec.sites.Active
        DSP_Gain_Mean_Array(0) = InWf.CalcMean
        DSP_Gain_Mean.data = DSP_Gain_Mean_Array
'        If DSP_OffSet_Mean_Array(0) < Low_limit Then
'            DSP_OffSet_Mean_Fuse_Array(0) = 2 ^ (Integer_Bit) + FormatNumber(Low_limit)
'        ElseIf DSP_OffSet_Mean_Array(0) >= Low_limit And DSP_OffSet_Mean_Array(0) < 0 Then
'            DSP_OffSet_Mean_Fuse_Array(0) = 2 ^ (Integer_Bit) + FormatNumber(DSP_OffSet_Mean_Array(0))
'        ElseIf DSP_OffSet_Mean_Array(0) < High_limit And DSP_OffSet_Mean_Array(0) >= 0 Then
'            DSP_OffSet_Mean_Fuse_Array(0) = FormatNumber(DSP_OffSet_Mean_Array(0))
'        Else
'            DSP_OffSet_Mean_Fuse_Array(0) = FormatNumber(High_limit)
'        End If
'        DSP_OffSet_Mean_Fuse.Data = DSP_OffSet_Mean_Fuse_Array
    Next site
    
    TestNameInput = Report_TName_From_Instance("C", "X", , 0, 0)
    TheExec.Flow.TestLimit resultVal:=DSP_Gain_Mean.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
    
'    Call AddStoredCaptureData(StoreName_OffSetMean, DSP_OffSet_Mean_Fuse)
End Function
Public Function Calc_MetrologyGR_RPSRSPARE0(argc As Integer, argv() As String) As Long
    Dim i As Long
    Dim MeasValue() As New PinListData: ReDim MeasValue(argc - 1)
    Dim RPSR_SPARE0 As New SiteDouble
    Dim site As Variant
    For i = 0 To argc - 1
        MeasValue(i) = GetStoredMeasurement(argv(i))
    Next i
    For Each site In TheExec.sites
        RPSR_SPARE0 = (MeasValue(0).Pins(0).value - MeasValue(1).Pins(0).value) / 0.000001
    Next site
    TestNameInput = Report_TName_From_Instance("CalcR", vbNullString)
    TheExec.Flow.TestLimit resultVal:=RPSR_SPARE0, Tname:=TestNameInput, ForceResults:=tlForceFlow
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_MetrologyGR_RPSRSPARE0"
    If AbortTest Then Exit Function Else Resume Next
End Function
'Added by Oscar for MTRBTS, From Canary/Swiftlet, 20201023
Public Function Calc_MetrologyBTS_Select(argc As Integer, argv() As String) As Long
    Dim DSP_MTRBTS_OUT As New DSPWave: DSP_MTRBTS_OUT = GetStoredCaptureData(argv(0))
    Dim Select_Bit As Long: Select_Bit = argv(1)
    Dim DSP_MTRBTS_Select_Binary As New DSPWave
    Dim DSP_MTRBTS_Select_Dec As New DSPWave
    Dim TestNameInput As String
    For Each site In TheExec.sites.Active
        DSP_MTRBTS_Select_Binary = DSP_MTRBTS_OUT.Select(DSP_MTRBTS_OUT.SampleSize - Select_Bit, 1, Select_Bit).Copy
        DSP_MTRBTS_Select_Dec = DSP_MTRBTS_Select_Binary.ConvertStreamTo(tldspParallel, DSP_MTRBTS_Select_Binary.SampleSize, 0, Bit0IsMsb)
    Next site
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.Flow.TestLimit resultVal:=DSP_MTRBTS_Select_Dec.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
    Call AddStoredCaptureData(argv(2), DSP_MTRBTS_Select_Dec)
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_MetrologyBTS_Select"
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Calc_MetrologyTMPS_Max_Min(argc As Integer, argv() As String) As Long
    
    Dim Temperature As New SiteDouble
    Dim Temperature_Sensor() As String: Temperature_Sensor = Split(argv(0), "+")
    Dim Temperature_Array(0) As Double
    Dim DSP_Temperature() As New DSPWave: ReDim DSP_Temperature(UBound(Temperature_Sensor))
    Dim DSP_MTRSNS_Temperature() As New DSPWave: ReDim DSP_MTRSNS_Temperature(UBound(Temperature_Sensor))
    
    Dim DSP_MTRSNS_Temperature_Maximum As New DSPWave
    Dim DSP_MTRSNS_Temperature_Minimum As New DSPWave
    
    Dim i As Long
    Dim site As Variant
    Dim TestNameInput As String
    Dim Temperature_Dictionary() As String
    
    
    DSP_MTRSNS_Temperature_Maximum.CreateConstant 0, UBound(Temperature_Sensor) + 1
    DSP_MTRSNS_Temperature_Minimum.CreateConstant 0, UBound(Temperature_Sensor) + 1
    
    For i = 0 To UBound(Temperature_Sensor)
        Temperature = GetStoredData(Temperature_Sensor(i) + "_para")
        For Each site In TheExec.sites
            Temperature_Array(0) = Temperature / 64
            DSP_Temperature(i).data = Temperature_Array
            DSP_MTRSNS_Temperature_Maximum.Element(i) = DSP_MTRSNS_Temperature_Maximum.Element(i) + DSP_Temperature(i).Element(0)
            DSP_MTRSNS_Temperature_Minimum.Element(i) = DSP_MTRSNS_Temperature_Minimum.Element(i) + DSP_Temperature(i).Element(0)
        Next site
    Next i
    
    For Each site In TheExec.sites
        DSP_MTRSNS_Temperature_Maximum(site).Element(0) = DSP_MTRSNS_Temperature_Maximum(site).CalcMaximumValue
        DSP_MTRSNS_Temperature_Minimum(site).Element(0) = DSP_MTRSNS_Temperature_Minimum(site).CalcMinimumValue
    Next site
    
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, , , , , , , tlForceNone) 'eng_forceflow_transfer
    TheExec.Flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Maximum.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature_Maximum_", , , vbTextCompare), ForceResults:=tlForceNone, customUnit:="C" 'eng_forceflow_transfer
    TheExec.Flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Minimum.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature_Minimum_", , , vbTextCompare), ForceResults:=tlForceNone, customUnit:="C" 'eng_forceflow_transfer

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_MetrologyTMPS_Max_Min"
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Calc_MetrologyBTS_Temperature_Max_Min_Avg(argc As Integer, argv() As String) As Long

    Dim Temperature As New SiteDouble
    Dim Temperature_Sensor() As String: Temperature_Sensor = Split(argv(0), "+")
    Dim Temperature_Array(0) As Double
    Dim DSP_Temperature() As New DSPWave: ReDim DSP_Temperature(UBound(Temperature_Sensor))
    Dim DSP_MTRSNS_Temperature() As New DSPWave: ReDim DSP_MTRSNS_Temperature(UBound(Temperature_Sensor))
    Dim DSP_MTRSNS_Temperature_eFuse() As New DSPWave: ReDim DSP_MTRSNS_Temperature_eFuse(UBound(Temperature_Sensor))
    
    Dim DSP_MTRSNS_Temperature_Average As New DSPWave
    Dim DSP_MTRSNS_Temperature_Maximum As New DSPWave
    Dim DSP_MTRSNS_Temperature_Minimum As New DSPWave
    
    
    Dim i As Long
    Dim site As Variant
    Dim TestNameInput As String
    Dim Temperature_Dictionary() As String
    Dim Sensor_Num() As String
    
    DSP_MTRSNS_Temperature_Average.CreateConstant 0, UBound(Temperature_Sensor) + 1
    DSP_MTRSNS_Temperature_Maximum.CreateConstant 0, UBound(Temperature_Sensor) + 1
    DSP_MTRSNS_Temperature_Minimum.CreateConstant 0, UBound(Temperature_Sensor) + 1
    
    
    For i = 0 To UBound(Temperature_Sensor)
        Temperature = GetStoredData(Temperature_Sensor(i) + "_para")
        For Each site In TheExec.sites
            Temperature_Array(0) = Temperature / 64
            DSP_Temperature(i).data = Temperature_Array
            
            DSP_MTRSNS_Temperature_Average.Element(i) = DSP_MTRSNS_Temperature_Average.Element(i) + DSP_Temperature(i).Element(0)
            DSP_MTRSNS_Temperature_Maximum.Element(i) = DSP_MTRSNS_Temperature_Maximum.Element(i) + DSP_Temperature(i).Element(0)
            DSP_MTRSNS_Temperature_Minimum.Element(i) = DSP_MTRSNS_Temperature_Minimum.Element(i) + DSP_Temperature(i).Element(0)
            
        Next site
    Next i
    
    For Each site In TheExec.sites
        DSP_MTRSNS_Temperature_Average(site).Element(0) = FormatNumber(DSP_MTRSNS_Temperature_Average(site).CalcMean, 0)
        DSP_MTRSNS_Temperature_Maximum(site).Element(0) = FormatNumber(DSP_MTRSNS_Temperature_Maximum(site).CalcMaximumValue, 4)
        DSP_MTRSNS_Temperature_Minimum(site).Element(0) = FormatNumber(DSP_MTRSNS_Temperature_Minimum(site).CalcMinimumValue, 4)
    Next site
    
    
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, , , , , , , tlForceNone) 'eng_forceflow_transfer
    
    If UCase(TheExec.Flow.CurrentFlowSheetName) = "FLOW_TMPS" Or UCase(TheExec.Flow.CurrentFlowSheetName) = "FLOW_TMPS_NO_RELAY" Or UCase(TheExec.Flow.CurrentFlowSheetName) = "FLOW_TMPS_TSNS" Then
        Dim TestNameInput_Ary() As String
        TestNameInput_Ary = Split(TestNameInput, "_")
        TestNameInput_Ary(8) = glb_MTRRecord & "-" & glb_MTRBTSCnt
        TestNameInput = Join(TestNameInput_Ary, "_")
    End If
    
    TheExec.Flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Average.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature_Average_", , , vbTextCompare), ForceResults:=tlForceNone, customUnit:="C" 'eng_forceflow_transfer
    TheExec.Flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Maximum.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature_Maximum_", , , vbTextCompare), ForceResults:=tlForceNone, customUnit:="C" 'eng_forceflow_transfer
    TheExec.Flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Minimum.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature_Minimum_", , , vbTextCompare), ForceResults:=tlForceNone, customUnit:="C" 'eng_forceflow_transfer

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_MetrologyBTS_Temperature_Max_Min_Avg"
    If AbortTest Then Exit Function Else Resume Next

End Function












