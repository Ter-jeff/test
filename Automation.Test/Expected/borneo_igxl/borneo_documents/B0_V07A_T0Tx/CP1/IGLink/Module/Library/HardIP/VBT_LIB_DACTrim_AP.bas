Attribute VB_Name = "VBT_LIB_DACTrim_AP"
#Const isUFP = True
Option Explicit
Public DACInitialFlag As Boolean
Public DACTrimValue As New Dictionary
Public DACTargetStr As New Dictionary
'Public StoreSeq As Integer

'Public Function DAC_Trim_VFI(Optional patset As Pattern, Optional TestSequence As String, Optional CPUA_Flag_In_Pat As Boolean, _
'Optional DisableComparePins As PinList, Optional DisableConnectPins As PinList, Optional DisableFRC As Boolean = False, Optional FRCPortName As String, _
'Optional MeasV_PinS As String, _
'Optional MeasF_PinS_SingleEnd As String, Optional MeasF_Interval As String, Optional MeasF_EventSourceWithTerminationMode As EventSourceWithTerminationMode, Optional MeasF_Flag_MeasureThreshold As Boolean = False, Optional MeasF_ThresholdPercentage As Double = 0.5, Optional MeasF_WaitTime As String, _
'Optional MeasI_pinS As String, Optional MeasI_Range As String, Optional MeasI_WaitTime As String, _
'Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, _
'Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As Long, Optional DigSrc_Sample_Size As Long, _
'Optional TrimTarget As Double, Optional TrimDataWidth As Long, Optional TrimStartInt As String = "0", Optional TrimStopInt As String = "0", _
'Optional TrimSearchMethod As SearchMethod = 0, Optional DigSrc_Equation As String, Optional TrimStoreName As String, Optional FuseName As String, Optional FuseType As String, Optional ReverS As Boolean = False, _
'Optional DigSrc_Assignment As String, Optional DigSrc_FlowForLoopIntegerName As String = "", _
'Optional SpecialCalcValSetting As CalculateMethodSetup = 0, _
'Optional InstSpecialSetting As InstrumentSpecialSetup = 0, _
'Optional CUS_Str_MainProgram As String = "", Optional CUS_Str_DigCapData As String = "", Optional CUS_Str_DigSrcData As String = "", _
'Optional Flag_SingleLimit As Boolean = False, Optional TestLimitPerPin_VFI As String = "FFF", _
'Optional MeasF_PinS_Differential As String, Optional ForceFunctional_Flag As Boolean = False, _
'Optional Meas_StoreName As String, Optional Calc_Eqn As String, _
'Optional Interpose_PrePat As String, Optional Interpose_PreMeas As String, Optional Interpose_PostTest As String, Optional InitStart As Long, Optional TrimCalcName As String = "", Optional TrimTargetCondition As TargetCondition = 0, Optional CharSetName As String, _
'Optional ForceV_Val As String, Optional ForceI_Val As String, _
'Optional Validating_ As Boolean) As Long
''Optional MeasF_WalkingStrobe_Flag As Boolean, Optional MeasF_WalkingStrobe_StartV As Double, Optional MeasF_WalkingStrobe_EndV As Double, Optional MeasF_WalkingStrobe_StepVoltage As Double, Optional MeasF_WalkingStrobe_BothVohVolDiffV As Double, Optional MeasF_WalkingStrobe_interval As Double, Optional MeasF_WalkingStrobe_miniFreq As Double, _
'
'    Dim PatCount As Long
'    Dim PattArray() As String
'
'    If Validating_ Then
'        Call PrLoadPattern(patset.Value)
'        Exit Function    ' Exit after validation
'    End If
'
'    Dim i As Long, j As Long, k As Long
'    Dim TestOptLen As Integer
'    Dim TestSequenceArray() As String, MeasPinAry_V() As String, MeasPinAry_F() As String, MeasPinAry_I() As String, MeasPinAry_IRange() As String
'    Dim MeasPinAry_F_Differential() As String
'    Dim MeasureF_Pin_Differential As New PinList
'    Dim Ts As Variant, TestOption As Variant, Site As Variant
'    Dim TestSeqNum As Integer
'    Dim MeasureV_pin As New PinList, MeasureF_Pin_SingleEnd As New PinList, MeasureI_pin As New PinList
'    Dim MeasureI_Pin_CurrentRange As String
'    Dim testnum As Long
'    Dim InDSPWave As New DSPWave, OutDspWave As New DSPWave
'    Dim ShowDec As String, ShowOut As String
'    Dim patt As Variant
'    Dim Pat As String
'    Dim HighLimitVal() As Double, LowLimitVal() As Double
'    Dim MeasureV_Pin_PPMU As String, MeasureV_Pin_UVI80 As String
'    Dim d_MeasF_Interval As Double
'    Dim FreqPinsCheckType() As String
'    Dim ThisPinType As String
'    Dim MeasF_EventSource As FreqCtrEventSrcSel
'    Dim MeasF_EnableVtMode As Boolean
'    Dim Split_F_Str() As String
'
'    ''20160906 - Return measurement to directionary if needed
'    Dim Rtn_MeasVolt As New PinListData, Rtn_MeasCurr As New PinListData, Rtn_MeasFreq As New PinListData
'    Dim MeasStoreName_Ary() As String
'    Dim Interpose_PreMeas_Ary() As String
'
'    ''    Dim RTN_InterposeString As String
'    Dim CheckDSPWave As New DSPWave
'
'    '' DACTrim Varable
'    Dim Meas_ToTestlimit() As New PinListData
'    Dim TrimNamePrint As String
'    Dim Trimname As String
'    Dim SiteSelect As New SiteBoolean
'    Dim TrimAssign As New SiteVariant
'    Dim LFlag As Boolean
'    Dim LastInterval As New SiteDouble
'    Dim DACTrimTestName As String
'    Dim TrimSeqNum As Integer
'    Dim outDec As New SiteLong
'    Dim outVal As New SiteDouble
'    Dim TrimStart_() As String
'    Dim TrimStop_() As String
'    Dim SeqCount As Integer
'    Dim TrimStart As Long
'    Dim TrimStop As Long
'
'    On Error GoTo errHandler
'
'    '''''Trim pre process init''''''
'        glb_Disable_CurrRangeSetting_Print = False
'    If Not DACInitialFlag Then Call DAC_Trim_Initial
'    DACTrimValue.RemoveAll
'    If TrimSearchMethod = Transitions Then TrimTargetCondition = Transition
'    If TrimStoreName = "" Then TrimStoreName = "VERIFICATION"
'    Call SearchPathSplit(TrimStartInt, TrimStopInt, TrimStart_(), TrimStop_(), SeqCount)
'    Call GetTrimCalcNameType(TrimCalcName, Meas_StoreName, Calc_Eqn)
'    ByPassTestLimit = True 'ByPassLimit
'    Trimname = UCase(TrimStoreName)
'    TrimNamePrint = Trimname
'    SiteSelect = TheExec.sites.Selected
'    RV = ReverS
'    SearchDone = False
'
'   ' If CompareMethod = "Transition" Then SearchDone = False
'
'    Call tl_PinListDataSort(True)
'
'    ''20170322-Store MeasF mid value for VT
'    Dim SplitFreqVtValue() As String
'    Dim DictKey_StoreVT As String
'    Dim Dict_VT_Value As New SiteDouble
'
'    '' 20160201 - Check input argumenets whether have "@" in the first character. Add it If no "@" in the beginning. Then remove it to process fomat.
'    Call VFI_AnalyzedInputStringByAt(MeasV_PinS, MeasF_PinS_SingleEnd, MeasI_pinS, MeasI_Range, MeasF_PinS_Differential, ForceV_Val, ForceI_Val)
'
'    Dim ForceV_Val_Ary() As String
'    Dim ForceI_Val_Ary() As String
'    Dim MeasurePin_ForceV_Val As String
'    Dim MeasurePin_ForceI_Val As String
'
'    Call VFI_ProcessInputString(TestSequence, MeasV_PinS, MeasI_pinS, MeasF_PinS_SingleEnd, MeasF_PinS_Differential, MeasI_Range, Meas_StoreName, Interpose_PreMeas, _
'                                            ForceV_Val, ForceI_Val, _
'                                            TestSequenceArray(), MeasPinAry_V(), MeasPinAry_I(), MeasPinAry_F(), _
'                                            MeasPinAry_F_Differential(), MeasPinAry_IRange(), MeasStoreName_Ary(), Interpose_PreMeas_Ary(), ForceV_Val_Ary(), ForceI_Val_Ary())
'
'''    '' 20150121 - Range Check
'''    If Range_Check_Enable_Word = True Then
'''        If TheExec.DataManager.MemberIndex = 0 Then
'''            'gl_UseLimitCheck_Counter = 0
'''        End If
'''    End If
'
'    Call Freq_ProcessEventSourceTerminationMode(MeasF_EventSourceWithTerminationMode, MeasF_EventSource, MeasF_EnableVtMode)
'
'    ''20141219 Get use-limit from flow table
'    Call GetFlowSingleUseLimit(HighLimitVal, LowLimitVal)
'
'    ''20161130-Get test name from flow table
'    Dim FlowTestNme() As String
'    If TPModeAsCharz_GLB Then
'        gl_CZ_FlowTestName_Counter = 0
'        Call GetFlowTestName(FlowTestNme)
'    End If
'
'    ''========================================================================================
'    Dim Store_Rtn_Meas() As New PinListData
'    Dim SoreMaxNum As Long
'    Dim StoreIndex As Long
'    ''20170123-Get how many store name in MeasStoreName_Ary
'    If Meas_StoreName <> "" Then
'        SoreMaxNum = 0
'        For i = 0 To UBound(MeasStoreName_Ary)
'            If MeasStoreName_Ary(i) <> "" Then
'                SoreMaxNum = SoreMaxNum + 1
'            End If
'        Next i
'         ReDim Store_Rtn_Meas(SoreMaxNum - 1) As New PinListData
'         StoreIndex = 0
'     End If
'    ''========================================================================================
'
'    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
'
'''    Dim Loop_count As Long
'''    Dim Loop_Init As Long
'''    Dim Loop_Max As Long
'''    Dim Loop_BitNum As Long
'''    Dim Loop_RegName As String
'''    Dim SplitLoop_RegName() As String
'''    Dim Split_Loop_DigSrc_Str() As String
'''    Dim BinStr As String
'''    Dim Loop_SplitByComma() As String
'''    Dim Loop_SplitByEqual() As String
'''
'''    Loop_Init = 0
'''    Loop_Max = 0
'''
'''    If CUS_Str_MainProgram <> "" And InStr(UCase(CUS_Str_MainProgram), UCase("Loop_DigSrc")) <> 0 Then
'''        Split_Loop_DigSrc_Str = Split(CUS_Str_MainProgram, ";")
'''        Loop_Init = Split_Loop_DigSrc_Str(1)
'''        Loop_Max = Split_Loop_DigSrc_Str(2)
'''        Loop_BitNum = Split_Loop_DigSrc_Str(3)
'''        Loop_RegName = Split_Loop_DigSrc_Str(4)
'''        SplitLoop_RegName = Split(Loop_RegName, ":")
'''    End If
'''
'''    Dim loop_i As Long, Loop_j As Long
'''    Dim Temp_Equal_Str As String
'''    Dim Final_Comma_Str As String
'''    Temp_Equal_Str = ""
'''    Final_Comma_Str = ""
'''
'''    For Loop_count = Loop_Init To Loop_Max
'''        If CUS_Str_MainProgram <> "" And InStr(UCase(CUS_Str_MainProgram), UCase("Loop_DigSrc")) <> 0 Then
'''            BinStr = Dec2BinStr32Bit_Rev(Loop_BitNum, Loop_count)
'''            Loop_SplitByComma = Split(DigSrc_Assignment, ";")
'''
'''            For loop_i = 0 To UBound(Loop_SplitByComma)
'''                Loop_SplitByEqual = Split(Loop_SplitByComma(loop_i), "=")
'''                For Loop_j = 0 To UBound(SplitLoop_RegName)
'''                    If UCase(Loop_SplitByEqual(0)) = UCase(SplitLoop_RegName(Loop_j)) Then
'''                        Loop_SplitByEqual(1) = BinStr
'''                        Temp_Equal_Str = Loop_SplitByEqual(0) & "=" & Loop_SplitByEqual(1)
'''                        Exit For
'''                    Else
'''                        Temp_Equal_Str = Loop_SplitByEqual(0) & "=" & Loop_SplitByEqual(1)
'''                    End If
'''                Next Loop_j
'''                If loop_i = 0 Then
'''                    Final_Comma_Str = Temp_Equal_Str
'''                Else
'''                    Final_Comma_Str = Final_Comma_Str & ";" & Temp_Equal_Str
'''                End If
'''            Next loop_i
'''            DigSrc_Assignment = Final_Comma_Str
'''        End If
'
'    '' 20160923 - Add Interpose_PrePat entry point
'    If Interpose_PrePat <> "" Then
'        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
'    End If
'
'    Call HardIP_InitialSetupForPatgen
'
'    ''20161205 - Force_Flow_Shmoo_Condition
'    If TheExec.sites.Item(Site).SiteVariableValue("Flow_Shmoo_DevCharSetup") <> "" Then Force_Flow_Shmoo_Condition 'Do Flow Shmoo
'
'    If patset.Value <> "" Then
'        TheHdw.Patterns(patset).Load
'        Call PATT_GetPatListFromPatternSet(patset.Value, PattArray, PatCount)
'    Else
'        ReDim PattArray(0)
'        PattArray(0) = ""
'    End If
'
'    If (DisableConnectPins <> "") Then TheHdw.Digital.Pins(DisableConnectPins).Disconnect
'    If (DisableComparePins <> "") Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = True
'
'    ''20161107-Return sweep test name
'    Dim Rtn_SweepTestName As String
'    Rtn_SweepTestName = ""
'
'Dim q As Integer
'For q = 0 To SeqCount
'    TrimStart = CLng(TrimStart_(q))
'    TrimStop = CLng(TrimStop_(q))
'    TrimStoreName = TrimNamePrint
'    Trimname = TrimNamePrint
'
'L: LFlag = DACTrim_DigSrc_Data(TrimTarget, TrimStart, TrimStop, TrimSearchMethod, DigSrc_Equation, TrimStoreName, TrimDataWidth, TrimAssign, InDSPWave, , TrimSeqNum, DigSrc_Assignment, InitStart, TrimTargetCondition): If TheExec.sites.Selected.Count = 0 Then GoTo L2
'    StoreIndex = 0
'
'    For Each patt In PattArray
'        If patt <> "" Then
'            Pat = CStr(patt)
'            TheHdw.Patterns(Pat).Load
'
'            Call DACDigSrcDspWave(Pat, DigSrc_pin, "Meas_src", InDSPWave)
'
'''        If TPModeAsCharz_GLB = True Then
'''            If Rtn_SweepTestName <> "" Then
'''''                Rtn_SweepTestName = "_" & Rtn_SweepTestName
'''                For i = 0 To UBound(FlowTestNme)
'''                    FlowTestNme(i) = Replace(LCase(FlowTestNme(i)), "sweepcode", Rtn_SweepTestName)
'''                Next i
'''            Else
'''                'Call SimulateFlowForSweep(FlowShmooString_GLB)
''''                If FlowShmooString_GLB <> "" Then
''''                    For i = 0 To UBound(FlowTestNme)
''''                        FlowTestNme(i) = Replace(LCase(FlowTestNme(i)), "sweepvoltage", FlowShmooString_GLB)
''''                    Next i
''''                End If
'''            End If
'''        End If
'            Set OutDspWave = Nothing
'            Call GeneralDigCapSetting(Pat, DigCap_Pin, DigCap_Sample_Size, OutDspWave)
'
'            If CUS_Str_MainProgram <> "" Then
'               If InStr(UCase(CUS_Str_MainProgram), "MTR_UVI80_SETUP") <> 0 Then
'                    Call MTR_UVI80_Setup
'               End If
'            End If
'
'        '' 20160713 - If no cpuflags in the test item modify the code to run pattern by using .test
'            If (CPUA_Flag_In_Pat) Then
'                Call TheHdw.Patterns(Pat).start
'            Else
'                Call TheHdw.Patterns(Pat).test(pfAlways, 0)
'            End If
'
'        End If
'        TestSeqNum = 0
'        TrimSeqNum = 0
'
'        For Each Ts In TestSequenceArray
'
'            ''20150907 - Only need CPUA_Flag_In_Pat to do control
'            If (CPUA_Flag_In_Pat) Then
'                Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0) 'Meas during CPUA loop
'            Else
'                Call TheHdw.Digital.Patgen.HaltWait 'Pattern without CPUA loop
'            End If
'
'            ''20160923 - Add Interpose_PreMeas entry point by each sequence
'            If Interpose_PreMeas <> "" Then
'                If UBound(Interpose_PreMeas_Ary) = 0 Then
'                    Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
'                ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
'                    Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
'                End If
'            End If
'
'            TestOptLen = Len(Ts)
''            ReDim Preserve Meas_ToTestlimit(TrimSeqNum)
'
'            For k = 1 To TestOptLen
'
'                TestOption = Mid(Ts, k, 1)
'
'                '' 20160106 - If "ForceFunctional_Flag" = True to let TestOption = "N" to make the test instance only run functional test
'                If ForceFunctional_Flag = True Then
'                    TestOption = "N"
'                End If
'
'                '' 20170523 - Add force I value for UVI80
'                If TestOption = "V" Then
'                    Call Decide_Measure_Pin(TestSeqNum, MeasPinAry_V, MeasureV_pin)
'                    Call Decide_ForceVal(TestSeqNum, ForceI_Val_Ary, MeasurePin_ForceI_Val)
'                End If
'                If TestOption = "F" And MeasF_PinS_SingleEnd <> "" Then Call Decide_Measure_Pin(TestSeqNum, MeasPinAry_F, MeasureF_Pin_SingleEnd)
'                If TestOption = "F" And MeasF_PinS_Differential <> "" Then Call Decide_Measure_Pin(TestSeqNum, MeasPinAry_F_Differential, MeasureF_Pin_Differential)
'                If TestOption = "I" Then
'                    Call Decide_Measure_Pin(TestSeqNum, MeasPinAry_I, MeasureI_pin)
'                    Call Decide_MeasureI_CurrentRange(TestSeqNum, MeasPinAry_IRange, MeasureI_Pin_CurrentRange)
'                    Call Decide_ForceVal(TestSeqNum, ForceV_Val_Ary, MeasurePin_ForceV_Val)
'                End If
'
'                For Each Site In TheExec.sites.Active
'                    testnum = TheExec.sites.Item(Site).TestNumber
'                Next Site
'
'                Select Case UCase(TestOption)
'
'                    Case "V"
'                        ReDim Preserve Meas_ToTestlimit(TrimSeqNum)
'                        Call Trim_DiscriminateMeasureV_TestCondition(MeasureV_pin)
'                        Call DiscriminateMeasureV_PinType(MeasureV_pin, MeasureV_Pin_PPMU, MeasureV_Pin_UVI80)
'''                        Call start_profile_DCVI(MeasureV_Pin, 0.01, 1000000, 1024, "capture_signal")
'
'                        Call HardIP_MeasureVolt(MeasureV_pin, TestLimitPerPin_VFI, TestSeqNum, k, Pat, Flag_SingleLimit, HighLimitVal(0), LowLimitVal(0), FlowTestNme, InstSpecialSetting, CUS_Str_MainProgram, SpecialCalcValSetting, Meas_ToTestlimit(TrimSeqNum), Rtn_SweepTestName, _
'                                                            MeasurePin_ForceI_Val)
'''                        Call Plot_profile_DCVI(MeasureV_Pin, "capture_signal")
'
'                        If TheExec.TesterMode = testModeOffline Then
'                            'If TrimStoreName Like "VERIFICATION*" Then Stop
'                            Dim Pin As Variant
'                                For Each Site In TheExec.sites
'                                    For Each Pin In Meas_ToTestlimit(TrimSeqNum).Pins
'                                        If Not ReverS Then
'                                            Meas_ToTestlimit(TrimSeqNum).Pins(Pin).Value = BinStr2Dec(GetTrimCode(CStr(TrimAssign))) * 0.1 + Site * 1.1 + TestSeqNum * 0.01
'                                        Else
'                                            Meas_ToTestlimit(TrimSeqNum).Pins(Pin).Value = InverseDec(BinStr2Dec(GetTrimCode(CStr(TrimAssign))), TrimDataWidth) * 0.1 + Site * 1.1 + TestSeqNum * 0.01
'                                        End If
'                                    Next Pin
'                                Next Site
'                        End If
'
'                        ''20160906 - Check store measurement or not
'                        If Meas_StoreName <> "" Then
'                            If MeasStoreName_Ary(TestSeqNum) <> "" Then
'                                Store_Rtn_Meas(StoreIndex) = Meas_ToTestlimit(TrimSeqNum)
'                                Call StoreDataAllType(MeasStoreName_Ary(TestSeqNum), Store_Rtn_Meas(StoreIndex))
'                                StoreIndex = StoreIndex + 1
'                            End If
'                        End If
'
'                    Case "F"
'                        ReDim Preserve Meas_ToTestlimit(TrimSeqNum)
'
'                        If MeasF_Interval <> "" Then
'                            d_MeasF_Interval = CDbl(MeasF_Interval)
'                        Else
'                            d_MeasF_Interval = pc_Def_VFI_FreqInterval
'                        End If
'
'                        If MeasureF_Pin_SingleEnd <> "" Then
'
'                            If MeasF_Flag_MeasureThreshold = True Then
'                                If MeasF_ThresholdPercentage = 0 Then MeasF_ThresholdPercentage = pc_Def_VFI_FreqThresholdPercentage
'                                Call Freq_PPMU_Meas_VOH(MeasureF_Pin_SingleEnd, MeasF_ThresholdPercentage, MeasF_EnableVtMode, MeasF_EventSource)
'                            End If
'
'                            If MeasF_EnableVtMode = True Then
'                                TheHdw.Digital.Pins(MeasureF_Pin_SingleEnd).Levels.DriverMode = tlDriverModeVt
'                            End If
'
'                            '' 20151113 - Modify get instrument type for actural waveform plot
'                            FreqPinsCheckType() = Split(MeasureF_Pin_SingleEnd, ",")
'                            ThisPinType = GetInstrument(CStr(FreqPinsCheckType(0)), 0)
'                            If ThisPinType = "HSD-U" Then
'                                Call HardIP_FrequencyMeasure(MeasureF_Pin_SingleEnd, False, TestLimitPerPin_VFI, LowLimitVal(0), HighLimitVal(0), TestSeqNum, Pat, Flag_SingleLimit, d_MeasF_Interval, FlowTestNme, MeasF_WaitTime, MeasF_EventSource, SpecialCalcValSetting, Meas_ToTestlimit(TrimSeqNum), Rtn_SweepTestName)
'                            Else
'                                Call HardIP_FrequencyMeasure_Dctime(MeasureF_Pin_SingleEnd, TestLimitPerPin_VFI, LowLimitVal(0), HighLimitVal(0), TestSeqNum, Pat, Flag_SingleLimit, d_MeasF_Interval, MeasF_WaitTime)
'                            End If
'
'                         '' 20151019 - Enable differential frequency counter
'                        ElseIf MeasureF_Pin_Differential <> "" Then
'
'                            If CUS_Str_DigSrcData <> "" Then
'                                SplitFreqVtValue = Split(CUS_Str_DigSrcData, ":")
'                                If UCase(SplitFreqVtValue(0)) = "SETUP_STORE_VT" Then
'                                    DictKey_StoreVT = SplitFreqVtValue(1)
'                                    Dict_VT_Value = GetStoreDataAllType(DictKey_StoreVT)
'
'                                    For Each Site In TheExec.sites
'                                        TheHdw.Digital.Pins(MeasureF_Pin_Differential).DifferentialLevels.Value(chDiff_Vt) = Dict_VT_Value(Site)
'                                        'TheExec.Datalog.WriteComment ("Site= " & Site & " Set " & MeasureF_Pin_Differential & " Diff_Vt = " & Dict_VT_Value(Site))
'                                    Next Site
'                                End If
'                            End If
'
'                            Call HardIP_FrequencyMeasure(MeasureF_Pin_Differential, True, TestLimitPerPin_VFI, LowLimitVal(0), HighLimitVal(0), TestSeqNum, Pat, Flag_SingleLimit, d_MeasF_Interval, FlowTestNme, MeasF_WaitTime, MeasF_EventSource, , Meas_ToTestlimit(TrimSeqNum), Rtn_SweepTestName, CUS_Str_MainProgram)
'                        End If
'
'                        ''20160906 - Check store measurement or not
'                        If Meas_StoreName <> "" Then
'                            If MeasStoreName_Ary(TestSeqNum) <> "" Then
'                                Store_Rtn_Meas(StoreIndex) = Meas_ToTestlimit(TrimSeqNum)
'                                Call StoreDataAllType(MeasStoreName_Ary(TestSeqNum), Store_Rtn_Meas(StoreIndex))
'                                StoreIndex = StoreIndex + 1
'                            End If
'                        End If
'
'                    Case "I"
'                        ReDim Preserve Meas_ToTestlimit(TrimSeqNum)
'                        If DisableFRC = True Then FreeRunClk_Disable (FRCPortName)
'
'                        Call HardIP_MeasureCurrent(MeasureI_pin, LowLimitVal(0), HighLimitVal(0), MeasureI_Pin_CurrentRange, Flag_SingleLimit, TestSeqNum, Pat, TestLimitPerPin_VFI, FlowTestNme, MeasI_WaitTime, SpecialCalcValSetting, Meas_ToTestlimit(TrimSeqNum), Rtn_SweepTestName, MeasurePin_ForceV_Val)
'
'                        ''20160906 - Check store measurement or not
'                        If Meas_StoreName <> "" Then
'                            If MeasStoreName_Ary(TestSeqNum) <> "" Then
'                                Store_Rtn_Meas(StoreIndex) = Meas_ToTestlimit(TrimSeqNum)
'                                Call StoreDataAllType(MeasStoreName_Ary(TestSeqNum), Store_Rtn_Meas(StoreIndex))
'                                StoreIndex = StoreIndex + 1
'                            End If
'                        End If
'
'                    Case "N"
'''                        TheExec.Datalog.WriteComment ("test sequence is N")
'                          TrimSeqNum = TrimSeqNum - 1
'
'                    Case Else
'
'                        TheExec.Datalog.WriteComment "Error Test Option, please select V,I or F"
'
'                End Select
'
'                If TheExec.sites.Active.Count = 0 Then Exit Function
'            Next k
'
'            ''20161206-Restore force condiction after measurement
'''            Call SetForceCondition("RESTORE")
'
'            If Interpose_PreMeas <> "" Then
'                If UBound(Interpose_PreMeas_Ary) = 0 Then
'                    Call SetForceCondition("RESTOREPREMEAS")
'                ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
'                    Call SetForceCondition("RESTOREPREMEAS")
'                End If
'            End If
'
'            TestSeqNum = TestSeqNum + 1
'            TrimSeqNum = TrimSeqNum + 1
'
'            If (CPUA_Flag_In_Pat) Then Call TheHdw.Digital.Patgen.Continue(0, cpuA)
'
'        Next Ts
'
'        TestSeqNum = TestSeqNum - 1
'        TrimSeqNum = TrimSeqNum - 1
'
'        If DebugPrintEnable = True Then TheExec.Datalog.WriteComment "  Pattern(" & PatCount & "): " & Pat & ""
'
'        TheHdw.Digital.Patgen.HaltWait ' Haltwait at patten end
'
'        PatCount = PatCount + 1
'
'        '' 20160923 - Add Interpose_PostTest entry point
'        Call SetForceCondition(Interpose_PostTest)
'
'        If DigCap_Sample_Size <> 0 Then
'
'            Dim DigCapPinAry() As String, NumberPins As Long
'            Call TheExec.DataManager.DecomposePinList(DigCap_Pin, DigCapPinAry(), NumberPins)
'
'            If NumberPins > 1 Then
'''                Call CreateSimulateDataDSPWave_Parallel(OutDspWave, DigCap_Sample_Size)
'''                Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDspWave, NumberPins)
'''                Call DigCapDataProcessByDSP_Parallel(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, NumberPins)
'
'            ElseIf NumberPins = 1 Then
'                Call CreateSimulateDataDSPWave(OutDspWave, DigCap_Sample_Size, DigCap_DataWidth)
'                Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDspWave, NumberPins)
'                Call DigCapDataProcessByDSP(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, DigCap_DataWidth, , True)
'            End If
'        End If
'
'    Next patt
'
'    If MeasureV_pin <> "" Then
'        Call EndSetupForMeasureVoltPins(MeasureV_Pin_PPMU, MeasureV_Pin_UVI80)
'    End If
'
'    If DisableConnectPins <> "" Then TheHdw.Digital.Pins(DisableConnectPins).Connect
'    If DisableComparePins <> "" Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = False
'
'    If DisableFRC = True Then
'        Call ReStart_FRC(FRCPortName)
'    End If
'
'    '' 20160907 - Process calculate equation by dictionary.
'    If Calc_Eqn <> "" Then
'        Call ProcessCalcEquation(Calc_Eqn)
'    End If
'
'''    If Calc_Eqn <> "" Then
'''            Dim SplitBySemi() As String
'''            Dim CalcIndex As Variant
'''            SplitBySemi() = Split(Calc_Eqn, ";")
'''            For i = 0 To UBound(SplitBySemi)
'''                ReDim Preserve Meas_ToTestlimit(TrimSeqNum + 1)
'''
'''                Call DACTrim_ProcessCalcEquation(Calc_Eqn, SplitBySemi(i), Meas_ToTestlimit(TrimSeqNum + 1), TrimSeqNum)
'''            Next i
'''    End If
'    glb_Disable_CurrRangeSetting_Print = True
'    Call CalcTarget_V2(Meas_ToTestlimit, TrimSeqNum, Trimname, TrimAssign, LastInterval, TrimTarget, TestSequenceArray, TrimCalcName, TrimTargetCondition): If LFlag Then GoTo L
'L3: Next q
'L2: If TrimReturn(q, SeqCount, TrimSearchMethod) Then GoTo L3
'    If Not TrimStoreName Like "VERIFICATION*" Or LFlag Then Trimname = "VERIFICATION": TrimStoreName = "VERIFICATION": TheExec.sites.Selected = SiteSelect: GoTo L
'
'    Call TrimLimit_V2(TrimNamePrint, TrimDataWidth, outDec, outVal, TestSequenceArray, Meas_ToTestlimit, TrimSeqNum, , DigSrc_pin)
'
'    If DigCap_Sample_Size <> 0 Then
'        Call CreateSimulateDataDSPWave(OutDspWave, DigCap_Sample_Size, DigCap_DataWidth)
'        Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDspWave, NumberPins)
'        Call DigCapDataProcessByDSP(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, DigCap_DataWidth)
'    End If
'
'    Call HardIP_WriteFuncResult
'
'    Call DACTrimPrint(TrimSeqNum, TrimNamePrint, TrimTarget)
'
'    If FuseName <> "" And FuseType <> "" Then Call FuseUpdate2(FuseName, FuseType, outDec)
'
'    ByPassTestLimit = False
'
'    DebugPrintFunc patset.Value  ' print all debug information
'
'    If TheExec.sites.Item(Site).SiteVariableValue("Flow_Shmoo_DevCharSetup") <> "" Then  'Do Flow Shmoo
'        If Flow_Shmoo_Port_Name <> "" Then Restart_All_Freerun_Clk
'    End If
'
'    If Interpose_PrePat <> "" Then
'        Call SetForceCondition("RESTOREPREPAT")
'    End If
'
'''    ''=============================== CharSetName ====================================
'''
'''     If Theexec.DevChar.Setups.IsRunning = False And CharSetName <> "" Then
'''         Dim p As Variant, p_ary() As String, p_cnt As Long, ApplyPins As String, Setup_mode As String
'''         'If TheExec.DevChar.Setups(CharSetName).TestMethod.Value = tlDevCharTestMethod_Reburst Then TheExec.Datalog.WriteComment "[PrintCharCondition:" & PrintCharSetup(Interpose_PrePat_GLB) & ",Test]"
'''         Setup_mode = Theexec.DevChar.Setups(CharSetName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.Name
'''         If (LCase(Setup_mode) <> "vid" And LCase(Setup_mode) <> "vicm") Then
'''             ApplyPins = Theexec.DevChar.Setups(CharSetName).Shmoo.Axes(tlDevCharShmooAxis_X).ApplyTo.Pins
'''             Theexec.DataManager.DecomposePinList ApplyPins, p_ary, p_cnt
'''             For Each p In p_ary
'''                 Theexec.DevChar.Setups(CharSetName).Shmoo.Axes(tlDevCharShmooAxis_X).ApplyTo.Pins = p
'''                 run_shmoo CharSetName
'''             Next p
'''             Theexec.DevChar.Setups(CharSetName).Shmoo.Axes(tlDevCharShmooAxis_X).ApplyTo.Pins = ApplyPins
'''         Else
'''             run_shmoo CharSetName
'''         End If
'''     End If
'''
'''    Next Loop_count
'''    ''================================================================================
'    glb_Disable_CurrRangeSetting_Print = False
'    Exit Function
'
'errHandler:
'    TheExec.Datalog.WriteComment "error in Meas_FreqVoltCurr_Universal_func"
'        glb_Disable_CurrRangeSetting_Print = False
'    If AbortTest Then Exit Function Else Resume Next
'
'End Function
Public Sub DAC_Trim_Initial()
    DACTrimValue.RemoveAll
    DACTargetStr.RemoveAll
    DACInitialFlag = True
End Sub

Public Sub FuseUpdate(FuseName As String, Trimname As String)
Dim TargetIdx As New SiteVariant
    If DACTargetStr.Exists(Trimname) Then TargetIdx = DACTargetStr.item(Trimname): Call DACTargetStr.Add(FuseName, TargetIdx)
    
End Sub
Public Sub FuseUpdate2(FuseName As String, FuseType As String, CodeDec As SiteLong)
    Dim patPassed As New SiteBoolean
    Dim site As Variant
    patPassed = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
    
    '20210406 Add for new Efuse
    Dim opbank As eFuseBdfBank
    Dim field As eFuseBdfField
    Set opbank = GetBdfBank(FuseType)
    Set field = opbank.Fields(FuseName)
    opbank.SetEfuse field.name, CodeDec, patPassed, , , , True
    
'    For Each site In TheExec.sites
'        Call auto_eFuse_SetPatTestPass_Flag(FuseType, FuseName, CBool(patPassed))
'        Call auto_eFuse_SetWriteDecimal(FuseType, FuseName, CLng(CodeDec))
'    Next site
End Sub

Public Sub SearchPathSplit(ST As String, SP As String, StartInt() As String, StopInt() As String, SeqCount As Integer)
Dim ST_ As String
Dim SP_ As String
    If ST = "" Then ST = "0"
    If SP = "" Then SP = "0"
    StartInt = Split(ST, ",")
    StopInt = Split(SP, ",")
    ST_ = UBound(StartInt)
    SP_ = UBound(StopInt)
    
    If UBound(StartInt) > UBound(StopInt) Then
        TheExec.ErrorLogMessage ("Stop Sequence Path loss")
        SeqCount = SP_
    ElseIf UBound(StartInt) < UBound(StopInt) Then
        TheExec.ErrorLogMessage ("Start Sequence Path loss")
        SeqCount = ST_
    Else
        SeqCount = UBound(StartInt)
    End If

End Sub
